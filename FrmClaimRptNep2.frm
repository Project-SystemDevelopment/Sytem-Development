VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmClaimRptNep2 
   Caption         =   "Claim Report 2 (Jupiter)"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame22 
      Caption         =   "Selection Criteria"
      Height          =   5535
      Left            =   0
      TabIndex        =   86
      Top             =   840
      Width           =   11865
      Begin VB.CheckBox chkAgeband 
         Caption         =   "Check1"
         Height          =   255
         Left            =   11445
         TabIndex        =   205
         Top             =   2700
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom2 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7350
         TabIndex        =   199
         Top             =   1600
         Width           =   255
      End
      Begin MSMask.MaskEdBox TxtPayEffDate2 
         Height          =   315
         Left            =   10320
         TabIndex        =   51
         Top             =   330
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
         Left            =   10320
         TabIndex        =   53
         Top             =   600
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
         Left            =   10320
         TabIndex        =   55
         Top             =   840
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
         Left            =   10320
         TabIndex        =   57
         Top             =   1125
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
         Left            =   8520
         TabIndex        =   50
         Top             =   330
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
         Left            =   8520
         TabIndex        =   52
         Top             =   600
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
         Left            =   8520
         TabIndex        =   54
         Top             =   840
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
         Left            =   8520
         TabIndex        =   56
         Top             =   1125
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboRace 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1725
         Width           =   1320
      End
      Begin VB.TextBox txtAgentCode 
         Height          =   285
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   20
         Top             =   330
         Width           =   2800
      End
      Begin VB.TextBox txtPolNo 
         Height          =   285
         Left            =   4200
         MaxLength       =   25
         TabIndex        =   21
         Top             =   585
         Width           =   2800
      End
      Begin VB.TextBox txtDOCName 
         Height          =   285
         Left            =   4200
         MaxLength       =   100
         TabIndex        =   22
         Top             =   810
         Width           =   2800
      End
      Begin VB.TextBox TxtPatientName 
         Height          =   285
         Left            =   4200
         MaxLength       =   60
         TabIndex        =   23
         Top             =   1065
         Width           =   2800
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
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   1320
         Width           =   2800
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   1605
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
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   1845
         Width           =   2800
      End
      Begin VB.ComboBox cboHOSGrp 
         Height          =   315
         Left            =   4200
         TabIndex        =   27
         Top             =   2115
         Width           =   2800
      End
      Begin MSMask.MaskEdBox TxtDOA2 
         Height          =   315
         Left            =   6000
         TabIndex        =   29
         Top             =   2415
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
         Left            =   6000
         TabIndex        =   31
         Top             =   2700
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
         Left            =   6000
         TabIndex        =   33
         Top             =   2985
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOW2 
         Height          =   315
         Left            =   6000
         TabIndex        =   35
         Top             =   3240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOA1 
         Height          =   315
         Left            =   4200
         TabIndex        =   28
         Top             =   2415
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
         Left            =   4200
         TabIndex        =   30
         Top             =   2700
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
         Left            =   4200
         TabIndex        =   32
         Top             =   2985
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOW1 
         Height          =   315
         Left            =   4200
         TabIndex        =   34
         Top             =   3270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox ChkPolNo 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   129
         Top             =   630
         Width           =   255
      End
      Begin VB.CheckBox ChkDoctor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   128
         Top             =   840
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
         Left            =   2640
         TabIndex        =   127
         Top             =   2865
         Width           =   375
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
         TabIndex        =   126
         Top             =   4575
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
         Left            =   2640
         TabIndex        =   125
         Top             =   2610
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
         TabIndex        =   124
         Top             =   2310
         Width           =   375
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
         TabIndex        =   123
         Top             =   4860
         Width           =   375
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
         Left            =   7125
         TabIndex        =   122
         Top             =   345
         Width           =   255
      End
      Begin VB.CheckBox ChkHLT 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   121
         Top             =   345
         Width           =   255
      End
      Begin VB.CheckBox ChkINS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   120
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox ChkMemType 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   119
         Top             =   855
         Width           =   255
      End
      Begin VB.CheckBox ChkAdultChild 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   118
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox ChkRace 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   117
         Top             =   1770
         Width           =   255
      End
      Begin VB.CheckBox ChkSex 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   116
         Top             =   2025
         Width           =   255
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
         TabIndex        =   115
         Top             =   4305
         Width           =   255
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
         TabIndex        =   114
         Top             =   4020
         Width           =   375
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
         TabIndex        =   113
         Top             =   3750
         Width           =   375
      End
      Begin VB.CheckBox ChkPolStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   112
         Top             =   3150
         Width           =   255
      End
      Begin VB.CheckBox ChkCaseStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   111
         Top             =   3435
         Width           =   255
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
         Left            =   11445
         TabIndex        =   110
         Top             =   1440
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
         Left            =   11445
         TabIndex        =   109
         Top             =   2160
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
         Left            =   11445
         TabIndex        =   108
         Top             =   345
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
         Left            =   11445
         TabIndex        =   107
         Top             =   1080
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
         Left            =   11445
         TabIndex        =   106
         Top             =   840
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
         Left            =   11445
         TabIndex        =   105
         Top             =   600
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
         Left            =   11445
         TabIndex        =   104
         Top             =   1680
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
         Left            =   11445
         TabIndex        =   103
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox ChkBordxDate 
         Caption         =   "Check1"
         Height          =   255
         Left            =   11445
         TabIndex        =   102
         Top             =   2415
         Width           =   255
      End
      Begin VB.TextBox txtPOLYear1 
         Height          =   285
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   58
         Top             =   1410
         Width           =   1000
      End
      Begin VB.TextBox txtPOLYear2 
         Height          =   285
         Left            =   10320
         MaxLength       =   2
         TabIndex        =   59
         Top             =   1410
         Width           =   1000
      End
      Begin VB.TextBox TxtMBMNo1 
         Height          =   285
         Left            =   8520
         MaxLength       =   15
         TabIndex        =   60
         Top             =   1665
         Width           =   1000
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   61
         Top             =   1665
         Width           =   1000
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   285
         Left            =   8520
         MaxLength       =   15
         TabIndex        =   62
         Top             =   1890
         Width           =   1000
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   285
         Left            =   10320
         MaxLength       =   15
         TabIndex        =   63
         Top             =   1920
         Width           =   1000
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   285
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   64
         Top             =   2145
         Width           =   1000
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   285
         Left            =   10320
         MaxLength       =   10
         TabIndex        =   65
         Top             =   2145
         Width           =   1000
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
         Left            =   7125
         TabIndex        =   101
         Top             =   3300
         Width           =   255
      End
      Begin VB.CheckBox ChkRepCat 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   100
         Top             =   4245
         Width           =   255
      End
      Begin VB.CheckBox ChkRep 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   99
         Top             =   4005
         Width           =   255
      End
      Begin VB.CheckBox ChkFinal 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   98
         Top             =   3795
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
         Left            =   7125
         TabIndex        =   97
         Top             =   5085
         Width           =   375
      End
      Begin VB.CheckBox ChkFirst 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   96
         Top             =   3525
         Width           =   255
      End
      Begin VB.CheckBox ChkHOSGrp 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   95
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox ChkHOS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   94
         Top             =   1875
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   93
         Top             =   1600
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   92
         Top             =   1350
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
         Left            =   7125
         TabIndex        =   91
         Top             =   2985
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
         Left            =   7125
         TabIndex        =   90
         Top             =   2700
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
         Left            =   7125
         TabIndex        =   89
         Top             =   2460
         Width           =   255
      End
      Begin VB.CheckBox ChkClaimType 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   88
         Top             =   4530
         Width           =   255
      End
      Begin VB.CheckBox ChkClaimNature 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7125
         TabIndex        =   87
         Top             =   4830
         Width           =   255
      End
      Begin VB.TextBox txtPreDiag1 
         Height          =   285
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   36
         Top             =   3525
         Width           =   1000
      End
      Begin VB.TextBox txtPreDiag2 
         Height          =   285
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   37
         Top             =   3525
         Width           =   1000
      End
      Begin VB.TextBox txtFinalDiag1 
         Height          =   285
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   38
         Top             =   3780
         Width           =   1000
      End
      Begin VB.TextBox txtFinalDiag2 
         Height          =   285
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   39
         Top             =   3780
         Width           =   1000
      End
      Begin VB.TextBox txtRepudCat1 
         Height          =   285
         Left            =   4200
         MaxLength       =   4
         TabIndex        =   40
         Top             =   4035
         Width           =   1000
      End
      Begin VB.TextBox txtRepudCat2 
         Height          =   285
         Left            =   6000
         MaxLength       =   4
         TabIndex        =   41
         Top             =   4035
         Width           =   1000
      End
      Begin VB.TextBox txtRepudCode1 
         Height          =   285
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   42
         Top             =   4290
         Width           =   1000
      End
      Begin VB.TextBox TxtClaimType1 
         Height          =   285
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   44
         Top             =   4545
         Width           =   1000
      End
      Begin VB.TextBox txtRepudCode2 
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   43
         Top             =   4290
         Width           =   1000
      End
      Begin VB.TextBox TxtClaimType2 
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   45
         Top             =   4545
         Width           =   1000
      End
      Begin VB.TextBox TxtClaimNature1 
         Height          =   285
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   46
         Top             =   4800
         Width           =   1000
      End
      Begin VB.TextBox TxtClaimNature2 
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   47
         Top             =   4800
         Width           =   1000
      End
      Begin VB.TextBox txtPlan1 
         Height          =   315
         Left            =   4200
         MaxLength       =   4
         TabIndex        =   48
         Top             =   5055
         Width           =   1000
      End
      Begin VB.TextBox txtPlan2 
         Height          =   315
         Left            =   6000
         MaxLength       =   4
         TabIndex        =   49
         Top             =   5055
         Width           =   1000
      End
      Begin MSMask.MaskEdBox TxtBordxDate1 
         Height          =   315
         Left            =   8520
         TabIndex        =   66
         Top             =   2400
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordxDate2 
         Height          =   315
         Left            =   10320
         TabIndex        =   67
         Top             =   2400
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "S - Sihat"
         Top             =   330
         Width           =   1320
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
         TabIndex        =   4
         Top             =   585
         Width           =   1320
      End
      Begin VB.ComboBox CboMemType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   825
         Width           =   1320
      End
      Begin VB.ComboBox cboSuppType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1110
         Width           =   1320
      End
      Begin VB.ComboBox CboSex 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   2010
         Width           =   1320
      End
      Begin VB.ComboBox CboRelation 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   2295
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
         TabIndex        =   11
         Top             =   2580
         Width           =   1320
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   2865
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
         TabIndex        =   13
         Top             =   3120
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
         TabIndex        =   14
         Top             =   3405
         Width           =   1320
      End
      Begin VB.ComboBox CboExpired 
         Height          =   315
         Left            =   1200
         TabIndex        =   15
         Top             =   3705
         Width           =   1320
      End
      Begin VB.ComboBox CboPreMode 
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   3990
         Width           =   1320
      End
      Begin VB.ComboBox CboState 
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Top             =   4275
         Width           =   1320
      End
      Begin VB.ComboBox cboRenewalInd 
         Height          =   315
         Left            =   1200
         TabIndex        =   18
         Top             =   4545
         Width           =   1320
      End
      Begin VB.ComboBox cboTakeOver 
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Top             =   4830
         Width           =   1320
      End
      Begin VB.ComboBox CboAdultChild 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   1395
         Width           =   1320
      End
      Begin VB.TextBox txtAgeband 
         Height          =   285
         Left            =   8520
         TabIndex        =   201
         Top             =   2680
         Width           =   1005
      End
      Begin VB.TextBox txtAgeband2 
         Height          =   285
         Left            =   10320
         TabIndex        =   202
         Top             =   2680
         Width           =   1005
      End
      Begin VB.Label Label74 
         Caption         =   "To"
         Height          =   255
         Left            =   9830
         TabIndex        =   204
         Top             =   2700
         Width           =   255
      End
      Begin VB.Label Label68 
         Caption         =   "Ageband"
         Height          =   255
         Left            =   7650
         TabIndex        =   203
         Top             =   2700
         Width           =   855
      End
      Begin VB.Label Label69 
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
         Left            =   7400
         TabIndex        =   200
         Top             =   225
         Width           =   135
      End
      Begin VB.Label Label31 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   120
         TabIndex        =   198
         Top             =   1185
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   3120
         TabIndex        =   197
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "State"
         Height          =   255
         Left            =   120
         TabIndex        =   196
         Top             =   4305
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Prem. Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   195
         Top             =   4065
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Expired Status"
         Height          =   255
         Left            =   120
         TabIndex        =   194
         Top             =   3825
         Width           =   1095
      End
      Begin VB.Label Label19 
         Height          =   255
         Left            =   2520
         TabIndex        =   193
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Relationship"
         Height          =   255
         Left            =   120
         TabIndex        =   192
         Top             =   2385
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Race"
         Height          =   255
         Left            =   120
         TabIndex        =   191
         Top             =   1785
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Adult / Child"
         Height          =   255
         Left            =   120
         TabIndex        =   190
         Top             =   1545
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   120
         TabIndex        =   189
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Sex"
         Height          =   255
         Left            =   120
         TabIndex        =   188
         Top             =   2145
         Width           =   855
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
         TabIndex        =   187
         Top             =   360
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
         TabIndex        =   186
         Top             =   585
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
         TabIndex        =   185
         Top             =   2625
         Width           =   1020
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
         TabIndex        =   184
         Top             =   3225
         Width           =   1020
      End
      Begin VB.Label Label53 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   120
         TabIndex        =   183
         Top             =   2865
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
         TabIndex        =   182
         Top             =   3465
         Width           =   945
      End
      Begin VB.Label Label39 
         Caption         =   "Doctor"
         Height          =   255
         Left            =   3120
         TabIndex        =   181
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label44 
         Caption         =   "Renewal "
         Height          =   255
         Left            =   120
         TabIndex        =   180
         Top             =   4665
         Width           =   1095
      End
      Begin VB.Label Label45 
         Caption         =   "Take Over"
         Height          =   255
         Left            =   120
         TabIndex        =   179
         Top             =   4905
         Width           =   1095
      End
      Begin VB.Label Label55 
         Caption         =   "Agent Code"
         Height          =   255
         Left            =   3120
         TabIndex        =   178
         Top             =   360
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
         Left            =   7125
         TabIndex        =   177
         Top             =   225
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
         Left            =   2640
         TabIndex        =   176
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Label62 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   3120
         TabIndex        =   175
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label56 
         Caption         =   "Pol. Year"
         Height          =   255
         Left            =   7650
         TabIndex        =   174
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9735
         TabIndex        =   173
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label60 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   7650
         TabIndex        =   172
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9735
         TabIndex        =   171
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9735
         TabIndex        =   170
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "DOB"
         Height          =   255
         Left            =   7650
         TabIndex        =   169
         ToolTipText     =   "Payor Expiry Date"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9735
         TabIndex        =   168
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label40 
         Caption         =   "Date Join"
         Height          =   255
         Left            =   7650
         TabIndex        =   167
         ToolTipText     =   "Payor Expiry Date"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   7650
         TabIndex        =   166
         ToolTipText     =   "Payor Effective Date"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   7650
         TabIndex        =   165
         ToolTipText     =   "Payor Expiry Date"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9735
         TabIndex        =   164
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9735
         TabIndex        =   163
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "Member No"
         Height          =   255
         Left            =   7650
         TabIndex        =   162
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9735
         TabIndex        =   161
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label66 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   7650
         TabIndex        =   160
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9735
         TabIndex        =   159
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label70 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   7650
         TabIndex        =   158
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9735
         TabIndex        =   157
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label75 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5415
         TabIndex        =   156
         Top             =   3315
         Width           =   375
      End
      Begin VB.Label Label76 
         Caption         =   "DOW"
         Height          =   255
         Left            =   3120
         TabIndex        =   155
         ToolTipText     =   "Worksheet Date"
         Top             =   3315
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Repud Cat"
         Height          =   255
         Left            =   3120
         TabIndex        =   154
         Top             =   4125
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Final Diag"
         Height          =   255
         Left            =   3120
         TabIndex        =   153
         Top             =   3855
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "Hosp. Grp"
         Height          =   255
         Left            =   3120
         TabIndex        =   152
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   3120
         TabIndex        =   151
         Top             =   1965
         Width           =   975
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5415
         TabIndex        =   150
         Top             =   4380
         Width           =   375
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5415
         TabIndex        =   149
         Top             =   4125
         Width           =   375
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5415
         TabIndex        =   148
         Top             =   3855
         Width           =   375
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5415
         TabIndex        =   147
         Top             =   3540
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Repud Code"
         Height          =   255
         Left            =   3120
         TabIndex        =   146
         Top             =   4380
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Pre. Diag"
         Height          =   255
         Left            =   3120
         TabIndex        =   145
         Top             =   3540
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   3120
         TabIndex        =   144
         Top             =   1665
         Width           =   1095
      End
      Begin VB.Label Label51 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   3120
         TabIndex        =   143
         Top             =   5085
         Width           =   855
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5415
         TabIndex        =   142
         Top             =   5085
         Width           =   375
      End
      Begin VB.Label Label38 
         Caption         =   "Payor"
         Height          =   225
         Left            =   3120
         TabIndex        =   141
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "DORC"
         Height          =   255
         Left            =   3120
         TabIndex        =   140
         ToolTipText     =   "Member Value Date"
         Top             =   3045
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5415
         TabIndex        =   139
         Top             =   3045
         Width           =   375
      End
      Begin VB.Label Label47 
         Caption         =   "DOA"
         Height          =   255
         Left            =   3120
         TabIndex        =   138
         ToolTipText     =   "Date of Admission"
         Top             =   2535
         Width           =   975
      End
      Begin VB.Label Label48 
         Caption         =   "DOD"
         Height          =   255
         Left            =   3120
         TabIndex        =   137
         ToolTipText     =   "Date of Discharged"
         Top             =   2805
         Width           =   975
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5415
         TabIndex        =   136
         Top             =   2565
         Width           =   375
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5415
         TabIndex        =   135
         Top             =   2805
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Claim  Type"
         Height          =   255
         Left            =   3120
         TabIndex        =   134
         Top             =   4620
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5415
         TabIndex        =   133
         Top             =   4620
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Claim Nature"
         Height          =   255
         Left            =   3120
         TabIndex        =   132
         Top             =   4845
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   195
         Left            =   5415
         TabIndex        =   131
         Top             =   4845
         Width           =   375
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
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   11445
         TabIndex        =   130
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Type"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.ComboBox cboReportType 
         Height          =   315
         Left            =   600
         TabIndex        =   1
         Top             =   300
         Width           =   4000
      End
      Begin VB.ComboBox cboReportSelection 
         Height          =   315
         Left            =   7080
         TabIndex        =   2
         Top             =   300
         Width           =   4000
      End
      Begin VB.Label Label22 
         Caption         =   "  Report Selection"
         Height          =   255
         Left            =   6600
         TabIndex        =   85
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   75
      Top             =   7320
      Width           =   11775
      Begin VB.TextBox lblTotalRecord 
         Height          =   285
         Left            =   1680
         TabIndex        =   207
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   206
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtDay 
         Height          =   285
         Left            =   1440
         TabIndex        =   78
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   68
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   73
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   74
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   72
         Top             =   160
         Width           =   975
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   77
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   76
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7080
         TabIndex        =   70
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   7920
         TabIndex        =   71
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6240
         TabIndex        =   69
         Top             =   160
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   255
         Left            =   3600
         TabIndex        =   79
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
      Begin VB.Label Label63 
         Caption         =   "of"
         Height          =   255
         Left            =   1320
         TabIndex        =   81
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label64 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   2640
         TabIndex        =   80
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView ListView2 
      CausesValidation=   0   'False
      Height          =   855
      Left            =   0
      TabIndex        =   82
      Top             =   6480
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1508
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
      NumItems        =   33
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
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
         Alignment       =   1
         SubItemIndex    =   8
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   28
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   29
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   30
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Object.Width           =   1764
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Fgrid 
      Height          =   375
      Left            =   0
      TabIndex        =   83
      Top             =   6960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   36
      Cols            =   5
      BackColorFixed  =   -2147483638
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claims Reports 2 (Jupiter)"
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
      TabIndex        =   84
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmClaimRptNep2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AA As Integer
Private FinalDiag, Description As String
Private Rnb, ICU, HSS, OT, MRI, PREXRAY, PRESpecFee, Surg, Anaest, NSPreSpec, total, SubTotal As String
Private Phys, Post, Op, Ambul, Physio, Kid, Organ, Tax  As String
Private Guardian, Medical, Tel, Admin, Others1, Others2 As String
Private strAppend As String
Private intExit As Integer
Private Current As String
Private GrpCompany As String
Private HOspitals As String
Private m_blnDirection(1 To 33) As Boolean

Private Sub ListViewHeader1()

    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOA"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "DOD"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "105(MRI)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub cboPayorCode_Click()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset
Dim con As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String

If Len(Mid(cboPayorCode, 1, 2)) Then
    If ChkGrpCom2.Value = 0 Then
        CboGrpCom.Clear
    Else
        CboGrpCom.Clear
     '   medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"
        
        'strAppendCase = "3"
        'StrAppend1 = " AND PAYCode = '" & Mid(cboPayorCode, 1, 2) & "'"
        'Set con.ActiveConnection = medixneptuneconn
        '    con.CommandText = "ClaimReport"
        '    con.CommandType = adCmdStoredProc
        '    con.CommandTimeout = 300
        '    Set prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 255, StrAppend1)
        '    con.Parameters.Append prm1
        '    Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
        '    con.Parameters.Append prm2
        '    PAYGRP.CursorLocation = adUseClient
        '    PAYGRP.CursorType = adOpenForwardOnly
        '    PAYGRP.CacheSize = 100
        '    Set PAYGRP = con.Execute
    strsql = "select CoGRPName from CompanyGrp where CoGRPPaycode = '" & Mid(cboPayorCode, 1, 2) & "'"
    PAYGRP.Open strsql, medixneptuneconn, adOpenForwardOnly, adLockReadOnly

        If PAYGRP.recordCount Then
            Do Until PAYGRP.EOF
               CboGrpCom.AddItem Trim(PAYGRP!CoGRPName)
               PAYGRP.MoveNext
            Loop
        End If
        PAYGRP.Close
        Set PAYGRP = Nothing
    '    medixneptuneconn.Close
   '     Set medixneptuneconn = Nothing
    End If
End If
End Sub

Private Sub ChkGrpCom2_Click()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset
Dim con As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String

If Len(Mid(cboPayorCode, 1, 2)) Then
    If ChkGrpCom2.Value = 0 Then
        CboGrpCom.Clear
    Else
        CboGrpCom.Clear
        'medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"
        
        'strAppendCase = "3"
        'StrAppend1 = " AND PAYCode = '" & Mid(cboPayorCode, 1, 2) & "'"
        'Set con.ActiveConnection = medixneptuneconn
        '    con.CommandText = "ClaimReport"
        '    con.CommandType = adCmdStoredProc
        '    con.CommandTimeout = 300
        '    Set prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 255, StrAppend1)
        '    con.Parameters.Append prm1
        '    Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
        '    con.Parameters.Append prm2
        '    PAYGRP.CursorLocation = adUseClient
        '    PAYGRP.CursorType = adOpenForwardOnly
        '    PAYGRP.CacheSize = 100
        '    Set PAYGRP = con.Execute
        strsql = "select CoGRPName from CompanyGrp where CoGRPPaycode = '" & Mid(cboPayorCode, 1, 2) & "'"
        PAYGRP.Open strsql, medixneptuneconn, adOpenForwardOnly, adLockReadOnly
        
        If PAYGRP.recordCount Then
            Do Until PAYGRP.EOF
               CboGrpCom.AddItem Trim(PAYGRP!CoGRPName)
               PAYGRP.MoveNext
            Loop
        End If
        PAYGRP.Close
        Set PAYGRP = Nothing
    '    medixneptuneconn.Close
    '    Set medixneptuneconn = Nothing
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

Private Sub cboHLTCode_Change()
    If InStr(cboHLTCode.Text, "null") Then
       cboHLTCode.Enabled = False
    End If
    
    If Len(Trim(cboHLTCode)) Then
       ' medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"
        LoadPayor (cboHLTCode)
       ' medixneptuneconn.Close
      '  Set medixneptuneconn = Nothing
    End If
End Sub

Private Sub cboHLTCode_Click()
    If Len(Trim(cboHLTCode)) Then
       ' medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"
        LoadPayor (cboHLTCode)
       ' medixneptuneconn.Close
       ' Set medixneptuneconn = Nothing
    End If
End Sub

Private Sub LoadPayor(Optional HLTCode As String)
    Dim Pay As New ADODB.Recordset
    Dim strAppend As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    
    strAppend = strAppend & "And HLTCOde = '" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    
    Set cmd.ActiveConnection = medixneptuneconn
    cmd.CommandText = "SearchPayorHLTType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Pay.CursorLocation = adUseClient
    Pay.CursorType = adOpenForwardOnly
    Pay.CacheSize = 100
    Set Pay = cmd.Execute
    
    cboPayorCode.Clear
    
    Do Until Pay.EOF
        With Pay
            cboPayorCode.AddItem !PAYCode & " - " & !PAYCompanyName
            Pay.MoveNext
        End With
    Loop
    
    Pay.Close
    Set Pay = Nothing
End Sub


Private Sub CmdBack_Click()
    
    If cboReportType = cboReportType.List(3) Then
        Fgrid.Height = 780
        Fgrid.Left = 0
        Fgrid.Top = 6480
        Fgrid.Width = 11775
        Frame22.Visible = True
        Frame1.Visible = True
    Else
        ListView2.Height = 780
        ListView2.Left = 0
        ListView2.Top = 6480
        ListView2.Width = 11775
        Frame22.Visible = True
        Frame1.Visible = True
    End If
    
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub CmdView_Click()
    
    If cboReportType = cboReportType.List(3) Then
        Fgrid.Height = 6200
        Fgrid.Left = 0
        Fgrid.Top = 300
        Fgrid.Width = 11775
        ListView2.Visible = False
        Frame22.Visible = False
        Frame1.Visible = False
    Else
        ListView2.Height = 6200
        ListView2.Left = 0
        ListView2.Top = 300
        ListView2.Width = 11775
        Frame22.Visible = False
        Frame1.Visible = False
        Fgrid.Visible = False
    End If
End Sub

Private Sub Form_Load()
    intExit = 0
      '  medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"
        Call ComboListing
   ' medixneptuneconn.Close
   ' Set medixneptuneconn = Nothing
    Call DoInitialSettings
    ListView2.Visible = True
    cboReportType.ListIndex = 0
End Sub

' =================== Command Button Click =============================

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    
    If cboReportType = cboReportType.List(3) Then
        Call ExportToExcelWorksheet
    Else
        If ListView2.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(ListView2)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
On Error GoTo errRun
Dim strAppend As String
    strAppend = ""
    CmdPrint.Enabled = False
    CmdClear.Enabled = False
    cmdExit.Enabled = False
    lblCurrent = ""
    lblTotalRecord = ""
    cmdSearch.Enabled = False
    
    If cboReportType = cboReportType.List(3) Then
        Fgrid.Height = 6200
        Fgrid.Left = 0
        Fgrid.Top = 1000
        Fgrid.Width = 11775
        ListView2.Visible = False
        Frame22.Visible = False
    Else
        ListView2.Height = 6200
        ListView2.Left = 0
        ListView2.Top = 1000
        ListView2.Width = 11775
        Frame22.Visible = False
        Fgrid.Visible = False
    End If
    
    Screen.MousePointer = vbHourglass
   ' medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"
   ' medixneptuneconnWS.Open "File Name=" & BOSPath & "Conn\medineptuneconnectWS.UDL"
    medixneptuneconnA.Open "File Name=" & BOSPath & "Conn\medineptuneconnectA.UDL"
    medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"
    medixneptuneconnWS.Open "File Name=" & BOSPath & "Conn\medineptuneconnectWS.UDL"
    medixneptuneconnPayment.Open "File Name=" & BOSPath & "Conn\medineptuneconnectPayment.UDL"
        Call SearchRecords
    medixneptuneconnA.Close
    Set medixneptuneconnA = Nothing
    medixneptuneconn.Close
    Set medixneptuneconn = Nothing
    medixneptuneconnWS.Close
    Set medixneptuneconnWS = Nothing
    medixneptuneconnPayment.Close
    Set medixneptuneconnPayment = Nothing
    '  medixneptuneconn.Close
  ''  Set medixneptuneconn = Nothing
   ' medixneptuneconnWS.Close
   ' Set medixneptuneconnWS = Nothing
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
    medixneptuneconnA.Close
    Set medixneptuneconnA = Nothing
    medixneptuneconn.Close
    Set medixneptuneconn = Nothing
    medixneptuneconnWS.Close
    Set medixneptuneconnWS = Nothing
    medixneptuneconnPayment.Close
    Set medixneptuneconnPayment = Nothing
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
            ctl.Value = False
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    
    End If
Next ctl
End Function

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.ListItems.Clear
    
    For FCol = 1 To 4
        Fgrid.col = FCol
        For Frow = 1 To 29
            Fgrid.row = Frow
            Fgrid = "-"
        Next
    Next
End Sub

Private Sub CboReportType_Click()

    ListView2.ListItems.Clear
    If cboReportType = cboReportType.List(0) Then
        cboReportSelection = cboReportSelection.List(0)
        Call cdActual
    ElseIf cboReportType = cboReportType.List(1) Then
        cboReportSelection = cboReportSelection.List(0)
        Call cdApproved
    ElseIf cboReportType = cboReportType.List(2) Then
        cboReportSelection = cboReportSelection.List(0)
        Call cdApprovedBTBG
    ElseIf cboReportType = cboReportType.List(3) Then
        Call cdBoth
    End If
End Sub

Private Sub cboReportSelection_Click()

    ListView2.ListItems.Clear
    If cboReportSelection = cboReportSelection.List(0) Then
        Call rptCaseNo
    ElseIf cboReportSelection = cboReportSelection.List(1) Then
        Call rptClaimNo
    ElseIf cboReportSelection = cboReportSelection.List(2) Then
        Call rptInsTypePlan
    ElseIf cboReportSelection = cboReportSelection.List(3) Then
        Call rptInsTypeAgeband
    ElseIf cboReportSelection = cboReportSelection.List(4) Then
        Call rptInsTypeSex
    ElseIf cboReportSelection = cboReportSelection.List(5) Then
        Call rptAgeband
    ElseIf cboReportSelection = cboReportSelection.List(6) Then
        Call rptSex
    ElseIf cboReportSelection = cboReportSelection.List(7) Then
        Call rptAgebandSex
    ElseIf cboReportSelection = cboReportSelection.List(8) Then
        Call rptDiagCode
    ElseIf cboReportSelection = cboReportSelection.List(9) Then
        Call rptHospital
    ElseIf cboReportSelection = cboReportSelection.List(10) Then
        Call rptClaimType
    ElseIf cboReportSelection = cboReportSelection.List(11) Then
        Call rptClaimNature
    ElseIf cboReportSelection = cboReportSelection.List(12) Then
        Call rptClaimNatureSex
    ElseIf cboReportSelection = cboReportSelection.List(13) Then
        Call rptPayor
    ElseIf cboReportSelection = cboReportSelection.List(14) Then
        Call rptPayorInsTypePlan
    ElseIf cboReportSelection = cboReportSelection.List(15) Then
        Call rptInsTypeRel
    ElseIf cboReportSelection = cboReportSelection.List(16) Then
        Call rptInsType
    ElseIf cboReportSelection = cboReportSelection.List(17) Then
        Call rptBordxNo
    ElseIf cboReportSelection = cboReportSelection.List(18) Then
        Call rptPolicyYear
    ElseIf cboReportSelection = cboReportSelection.List(19) Then
        Call rptTakeover
    End If

End Sub
' =================== Command Button Click / Change ==========================


'============== ComboBox Listing =================================
Private Sub ComboListing()
Dim Pay As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset
Dim state As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim HosGrp As New ADODB.Recordset
Dim cmd As New ADODB.Command
    

    
    Set cmd.ActiveConnection = medixneptuneconn
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
    
    Set cmd.ActiveConnection = medixneptuneconn
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
    
    Set cmd.ActiveConnection = medixneptuneconn
    cmd.CommandText = "SearchState"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    state.CursorLocation = adUseClient
    state.CursorType = adOpenForwardOnly
    state.CacheSize = 100
    Set state = cmd.Execute
        
    Do Until state.EOF
        With state
            CboState.AddItem !STADescription
            state.MoveNext
        End With
    Loop
    state.Close
    Set state = Nothing
        
    Set cmd.ActiveConnection = medixneptuneconn
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
    
    Set cmd.ActiveConnection = medixneptuneconn
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
        
    Set cmd.ActiveConnection = medixneptuneconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Pay.CursorLocation = adUseClient
    Pay.CursorType = adOpenForwardOnly
    Pay.CacheSize = 100
    Set Pay = cmd.Execute
        
    Do Until Pay.EOF
        With Pay
            cboPayorCode.AddItem !PAYCode & " - " & !PAYCompanyName
            Pay.MoveNext
        End With
    Loop
    Pay.Close
    Set Pay = Nothing
    
        
    
    cboHLTCode.AddItem "S" & " - " & "Sihat"
    cboHLTCode.AddItem "N" & " - " & "Non-Sihat"
    
    CboMemType.AddItem "Both"
    CboMemType.AddItem "P" & " - " & "Principal"
    CboMemType.AddItem "S" & " - " & "Supplementary"
    CboMemType.Text = "Both"
    
    CboAdultChild.AddItem "A - Adult"
    CboAdultChild.AddItem "C - Child"
    
        
    CboRace.AddItem "C - Chinese"
    CboRace.AddItem "I - Indian"
    CboRace.AddItem "M - Malay"
    CboRace.AddItem "O - Other"
            
    CboSex.AddItem "M" & " - " & "Male"
    CboSex.AddItem "F" & " - " & "Female"
    
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
        
    cboRenewalInd.AddItem "N - New"
    cboRenewalInd.AddItem "P - Partial"
    cboRenewalInd.AddItem "R - ReNew"
    cboRenewalInd.AddItem "T - Timely"
    
    cboTakeOver.AddItem "N - None"
    cboTakeOver.AddItem "Y - Take Over"
    cboTakeOver.AddItem "C - Conversion"
           
    cboSuppType.AddItem "Both"
    cboSuppType.AddItem "C" & " - " & "Combined Limit"
    cboSuppType.AddItem "S" & " - " & "Separated Limit"
    cboSuppType.Text = "Both"
    
    cboReportType.AddItem "Cost Details - Actual"
    cboReportType.AddItem "Cost Details - Approved"
    cboReportType.AddItem "Cost Details - Approved (BTBG)"
    cboReportType.AddItem "Cost Details - Both"
    
    cboReportSelection.AddItem "By Case No"
    cboReportSelection.AddItem "By Claim No"
    cboReportSelection.AddItem "By Ins. Type & Plan"
    cboReportSelection.AddItem "By Ins. Type & Ageband"
    cboReportSelection.AddItem "By Ins. Type & Sex"
    cboReportSelection.AddItem "By Ageband"
    cboReportSelection.AddItem "By Sex"
    cboReportSelection.AddItem "By Ageband & Sex"
    cboReportSelection.AddItem "By Diag. Code"
    cboReportSelection.AddItem "By Hospital"
    cboReportSelection.AddItem "By Claim Type"
    cboReportSelection.AddItem "By Claim Nature"
    cboReportSelection.AddItem "By Claim Nature & Sex"
    cboReportSelection.AddItem "By Payor"
    cboReportSelection.AddItem "By Payor, Ins. Type & Plan"
    cboReportSelection.AddItem "By Ins. Type & Rel."
    cboReportSelection.AddItem "By Ins. Type"
    cboReportSelection.AddItem "By Bordx No"
    cboReportSelection.AddItem "By Policy Year"
    cboReportSelection.AddItem "By Takeover"
End Sub
'============== ComboBox Listing =================================

'=========== Option Button, Combo Box Change/KeyPress ===================
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
 
    For i = 0 To cboReportType.ListCount - 1
 
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

End Sub


Private Sub cdActual()

    Fgrid.Visible = False
    ListView2.Visible = True
    ListView2.Height = 780
    ListView2.Left = 0
    ListView2.Top = 6600
    ListView2.Width = 11775
    Frame22.Visible = True
   
    cboReportSelection.Visible = True
    
    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOA"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "DOD"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "105(MRI)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptAgebandSex()

    ListView2.ColumnHeaders(1).Text = "Ageband"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(MRI)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    
    For AA = 28 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptDiagCode()

    ListView2.ColumnHeaders(1).Text = "Final Diag"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(MRI)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    
    For AA = 28 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptHospital()

    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(MRI)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Total"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End Sub

Private Sub rptClaimType()

    ListView2.ColumnHeaders(1).Text = "Claim Type"
    ListView2.ColumnHeaders(2).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(MRI)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Total"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptClaimNature()

    ListView2.ColumnHeaders(1).Text = "Claim Nature"
    ListView2.ColumnHeaders(2).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(MRI)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Total"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptClaimNatureSex()

    ListView2.ColumnHeaders(1).Text = "Claim Nature"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(MRI)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    
    For AA = 28 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptPayor()

    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(MRI)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Total"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptPayorInsTypePlan()

    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "105(MRI)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Total"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    
    For AA = 29 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptInsTypeRel()

    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Relation"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(MRI)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    
    For AA = 28 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptInsType()

    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(MRI)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Total"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub cdApproved()
    
    Fgrid.Visible = False
    ListView2.Visible = True
    ListView2.Height = 780
    ListView2.Left = 0
    ListView2.Top = 6600
    ListView2.Width = 11775
    Frame22.Visible = True

    cboReportSelection.Visible = True
    
    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOA"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "DOD"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "105(MRI)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptBordxNo()

    ListView2.ColumnHeaders(1).Text = "Bordx No"
    ListView2.ColumnHeaders(2).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(MRI)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Total"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptPolicyYear()

    ListView2.ColumnHeaders(1).Text = "Policy Year"
    ListView2.ColumnHeaders(2).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(MRI)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Total"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptTakeover()

    ListView2.ColumnHeaders(1).Text = "Takeover"
    ListView2.ColumnHeaders(2).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(MRI)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Total"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub cdBoth()
    
    ListView2.Visible = False
    Fgrid.Visible = True
    Fgrid.Height = 780
    Fgrid.Left = 0
    Fgrid.Top = 6600
    Fgrid.Width = 11775
    Frame22.Visible = True

    cboReportSelection.Visible = False
    
End Sub

Private Sub cdApprovedBTBG()
    Fgrid.Visible = False
    ListView2.Visible = True
    ListView2.Height = 780
    ListView2.Left = 0
    ListView2.Top = 6600
    ListView2.Width = 11775
    Frame22.Visible = True
    
    cboReportSelection.Visible = True
    
    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOA"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "DOD"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "105(MRI)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptCaseNo()

    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOA"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "DOD"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "105(MRI)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End Sub

Private Sub rptClaimNo()

    ListView2.ColumnHeaders(1).Text = "Claim No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "DOA"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "DOD"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "105(MRI)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Total"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnRight
    
    For AA = 31 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

End Sub

Private Sub rptInsTypePlan()

    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(MRI)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    
    For AA = 28 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptInsTypeAgeband()

    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Ageband"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(MRI)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    
    For AA = 28 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptInsTypeSex()

    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "105(MRI)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    
    For AA = 28 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptAgeband()

    ListView2.ColumnHeaders(1).Text = "Ageband"
    ListView2.ColumnHeaders(2).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(MRI)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Total"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End Sub

Private Sub rptSex()

    ListView2.ColumnHeaders(1).Text = "Sex"
    ListView2.ColumnHeaders(2).Text = "101(Room & Board)"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "102(Intensive Care Unit)"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "103(Hosp. Supplies & Services)"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "104(Operating Theatre)"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "105(MRI)"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "201(Pre-Hosp. X-Ray and Lab Charge)"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "202(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "203(Surgical Fees)"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "204(Anaesthetist's Fee)"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "301(Pre-Hosp. Specialist Fee)"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "302(Daily In-Hosp. Physician's Visit)"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "303(Post-Hospitalisation Treatment)"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "401(Out-Patient Accident)"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "402(Ambulance Fees)"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "403(Physiotherapy Treatment)"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(17).Text = "404(Kidney Dialysis & Cancer Treatment)"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "501(Organ Transplantation)"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(19).Text = "601(Govt. Hosp. Income Benefit)"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "701(Child's Guardian Allowance)"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(21).Text = "801(Medical Report)"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "802(Telephone Charges)"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(23).Text = "803(Admin/Registration)"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "901(OTHERS 1)"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "901(OTHERS 2)"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Total"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    
    For AA = 27 To 33
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
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
Dim sql As String
Dim strAppend As String

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
    
    If Len(Mid(CboAdultChild, 1, 1)) Then
        strAppend = strAppend & " AND AdultChild = '" & Trim(Mid(CboAdultChild, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboRace, 1, 1)) Then
        strAppend = strAppend & " AND RACCode = '" & Trim(Mid(CboRace, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboSex, 1, 1)) Then
        strAppend = strAppend & " AND Sex = '" & Trim(Mid(CboSex, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboRelation, 1, 1)) Then
        strAppend = strAppend & " AND Relation = '" & Trim(Mid(CboRelation, 1, 1)) & "'"
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

    If Len(Mid(CboCaseStatus, 1, 1)) Then
        strAppend = strAppend & " AND CASStatus = '" & Trim(Mid(CboCaseStatus, 1, 1)) & "'"
    End If
    
    If Len(Trim(Mid(CboExpired, 1, 1))) Then
          strAppend = strAppend & "AND MBMExpiredStatus = '" & Trim(Mid(CboExpired, 1, 1)) & "' "
    End If
    
    If Len(Mid(CboPreMode, 1, 1)) Then
        strAppend = strAppend & " AND MBMInstallment = '" & Trim(Mid(CboPreMode, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboState)) Then
          strAppend = strAppend & "AND STACode= '" & Trim(CboState) & "' "
    End If
    
    If Mid(cboRenewalInd, 1, 1) = "P" Then
        strAppend = strAppend & " AND (MBMRenewFlag = 'P')"
    ElseIf Mid(cboRenewalInd, 1, 1) = "N" Then
        strAppend = strAppend & " AND (MBMRenewFlag Is null OR MBMRenewFlag = '' OR MBMRenewFlag = 'N') "
    ElseIf Mid(cboRenewalInd, 1, 1) = "R" Then
        strAppend = strAppend & " AND (MBMRenewFlag = 'R')"
    ElseIf Mid(cboRenewalInd, 1, 1) = "T" Then
        strAppend = strAppend & " AND (MBMRenewFlag = 'T')"
    End If
    
    If Len(Mid(cboTakeOver, 1, 1)) Then
        strAppend = strAppend & " And MBMTakeOver = '" & Trim(Mid(cboTakeOver, 1, 1)) & "'"
    End If
    
    If Len(Trim(txtAgentCode)) Then
       strAppend = strAppend + " And  MBMAgentCode like '%" & ChkStr(txtAgentCode) & "%' "
    End If
    
    If Len(Trim(txtAgeband)) Then
       strAppend = strAppend + " And  Ageband >= '" & Trim(txtAgeband) & "'"
    End If
    
    If Len(Trim(txtAgeband2)) Then
       strAppend = strAppend + " And  Ageband <= '" & Trim(txtAgeband2) & "'"
    End If
    
    If Len(Trim(txtPolNo)) Then
       strAppend = strAppend + " And  MBMPolicyNo like '%" & ChkStr(txtPolNo) & "%' "
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Trim(Mid(cboPayorCode, 1, 2)) & "'"
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
        
    If (TxtDOA1) <> "__/__/____" And IsDate(TxtDOA1) = True _
        And (TxtDOA2) <> "__/__/____" And IsDate(TxtDOA2) = True Then
        sql = ""
        sql = " AND CasDateAdmitted >= '" & Format(Trim(TxtDOA1), "mm/dd/yyyy") & _
              "' And CasDateAdmitted <= '" & Format(Trim(TxtDOA2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtDOA1)) > 0 And IsDate(TxtDOA1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(TxtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOA2)) > 0 And IsDate(TxtDOA2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
       
    If (txtDOD1) <> "__/__/____" And IsDate(txtDOD1) = True _
        And (txtDOD2) <> "__/__/____" And IsDate(txtDOD2) = True Then
        sql = ""
        sql = " AND CasDischargeDate >= '" & Format(Trim(txtDOD1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDOD2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDOD1)) > 0 And IsDate(txtDOD1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD2)) > 0 And IsDate(txtDOD2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    'If (TxtDOR1) <> "__/__/____" And IsDate(TxtDOR1) = True _
    '    And (TxtDOR2) <> "__/__/____" And IsDate(TxtDOR2) = True Then
    '    Sql = ""
    '    Sql = " AND CaseRegDate >= '" & Format(Trim(TxtDOR1), "mm/dd/yyyy") & _
    '          "' And CaseRegDate <= '" & Format(Trim(TxtDOR2), "mm/dd/yyyy") & "'"
    '    strAppend = strAppend + Sql
    'End If
    
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
    
    If Len(Trim(TxtClaimType1)) Then
        strAppend = strAppend & " AND PayTypeCode >= '" & Trim(TxtClaimType1) & "'"
    End If
    
    If Len(Trim(TxtClaimType2)) Then
        strAppend = strAppend & " AND PayTypeCode <= '" & Trim(TxtClaimType2) & "'"
    End If
    
    If Len(Trim(TxtClaimNature1)) Then
        strAppend = strAppend & " AND ClaimNatureCode >= '" & Trim(TxtClaimNature1) & "'"
    End If
    
    If Len(Trim(TxtClaimNature2)) Then
        strAppend = strAppend & " AND ClaimNatureCode <= '" & Trim(TxtClaimNature2) & "'"
    End If
    
    If Len(Trim(txtPlan1)) Then
        strAppend = strAppend & " And PLNCode >= '" & Trim(txtPlan1) & "'"
    End If
    
    If Len(Trim(txtPlan2)) Then
        strAppend = strAppend & " And PLNCode <= '" & Trim(txtPlan2) & "'"
    End If
    
    If Len(Trim(txtPolicyNo1)) Then
        strAppend = strAppend & " AND MBMPolicyNo >= '" & Trim(txtPolicyNo1) & "'"
    End If
    
    If Len(Trim(txtPolicyNo2)) Then
        strAppend = strAppend & " AND MBMPolicyNo <= '" & Trim(txtPolicyNo2) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
        strAppend = strAppend & " AND BordxRefNo >= '" & Trim(TxtBordxNo1) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
        strAppend = strAppend & " AND BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
    End If
    
    
    If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend + " And MBMPayorEffDate >= '" & Format(TxtPayEffDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
    End If
       
    If (TxtPayEffDate1) <> "__/__/____" And IsDate(TxtPayEffDate1) = True _
        And (TxtPayEffDate2) <> "__/__/____" And IsDate(TxtPayEffDate2) = True Then
        sql = ""
        sql = " AND MBMPayorEffDate >= '" & Format(Trim(TxtPayEffDate1), "mm/dd/yyyy") & _
              "' And MBMPayorEffDate <= '" & Format(Trim(TxtPayEffDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If (TxtPayExpDate1) <> "__/__/____" And IsDate(TxtPayExpDate1) = True _
        And (TxtPayExpDate2) <> "__/__/____" And IsDate(TxtPayExpDate2) = True Then
        sql = ""
        sql = " AND MBMPayorExpDate >= '" & Format(Trim(TxtPayExpDate1), "mm/dd/yyyy") & _
              "' And MBMPayorExpDate <= '" & Format(Trim(TxtPayExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtPayExpDate1)) And IsDate(TxtPayExpDate1) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate >= '" & Format(TxtPayExpDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayExpDate2)) And IsDate(TxtPayExpDate2) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate <= '" & Format(TxtPayExpDate2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDateJoin1) <> "__/__/____" And IsDate(txtDateJoin1) = True _
        And (txtDateJoin2) <> "__/__/____" And IsDate(txtDateJoin2) = True Then
        sql = ""
        sql = " AND MBMFirstJoinedDate >= '" & Format(Trim(TxtExpDate1), "mm/dd/yyyy") & _
              "' And MBMFirstJoinedDate <= '" & Format(Trim(TxtExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDateJoin1)) And IsDate(txtDateJoin1) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate >= '" & Format(txtDateJoin1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDateJoin2)) And IsDate(txtDateJoin2) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate <= '" & Format(txtDateJoin2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDOB1) <> "__/__/____" And IsDate(txtDOB1) = True _
        And (txtDOB2) <> "__/__/____" And IsDate(txtDOB2) = True Then
        sql = ""
        sql = " AND DOB >= '" & Format(Trim(txtDOB1), "mm/dd/yyyy") & _
              "' And DOB <= '" & Format(Trim(txtDOB2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDOB1)) And IsDate(txtDOB1) = True Then
        strAppend = strAppend & " AND DOB >= '" & Format(txtDOB1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOB2)) And IsDate(txtDOB2) = True Then
        strAppend = strAppend & " AND DOB <= '" & Format(txtDOB2, "mm/dd/yyyy") & "'"
    End If
    
    If (TxtBordxDate1) <> "__/__/____" And IsDate(TxtBordxDate1) = True _
        And (TxtBordxDate2) <> "__/__/____" And IsDate(TxtBordxDate2) = True Then
        sql = ""
        sql = " AND BordxDate >= '" & Format(Trim(TxtBordxDate1), "mm/dd/yyyy") & _
              "' And BordxDate <= '" & Format(Trim(TxtBordxDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtBordxDate1)) And IsDate(TxtBordxDate1) = True Then
        strAppend = strAppend & " AND BordxDate >= '" & Format(TxtBordxDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxDate2)) And IsDate(TxtBordxDate2) = True Then
        strAppend = strAppend & " AND BordxDate <= '" & Format(TxtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOW1)) > 0 And IsDate(txtDOW1) = True Then
        strAppend = strAppend + " And ClaimOpenDate >= '" & Format(txtDOW1, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(txtDOW2)) > 0 And IsDate(txtDOW2) = True Then
        strAppend = strAppend + " And ClaimOpenDate <= '" & Format(txtDOW2, "mm/dd/yyyy") & "'"
    End If

    If ChkDOW.Value = 1 Then
        strAppend = strAppend & " AND (ClaimOpenDate Is null OR ClaimOpenDate = '')"
    End If
    
    If Len(Trim(txtPOLYear1)) Then
        strAppend = strAppend & " AND MBMPolicyYear >= '" & Trim(txtPOLYear1) & "'"
    End If
    
    If Len(Trim(txtPOLYear2)) Then
        strAppend = strAppend & " AND MBMPolicyYear <= '" & Trim(txtPOLYear2) & "'"
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
    
    If chkAgeband.Value = 1 Then
        strAppend = strAppend & " AND (Ageband Is null OR Ageband = '')"
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
    
    If ChkAdultChild.Value = 1 Then
        strAppend = strAppend & " AND (AdultChild Is null OR AdultChild = '')"
    End If
    
    If ChkRace.Value = 1 Then
        strAppend = strAppend & " AND (RACCode Is null OR RACCode = '')"
    End If
    
    If ChkSex.Value = 1 Then
        strAppend = strAppend & " AND (Sex Is null OR Sex = '')"
    End If
    
    If ChkRelation.Value = 1 Then
        strAppend = strAppend & " AND (Relation Is null OR Relation = '')"
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
    
    If ChkPreMode.Value = 1 Then
        strAppend = strAppend & " AND (MBMInstallment Is null OR MBMInstallment = '')"
    End If
    
    If ChkState.Value = 1 Then
        strAppend = strAppend & " AND (STACode Is null OR STACode = '')"
    End If
    
    If ChkRenewal.Value = 1 Then
        strAppend = strAppend & " AND (MBMRenewFlag Is null OR MBMRenewFlag = '')"
    End If
    
    If ChkTakeOver.Value = 1 Then
        strAppend = strAppend & " AND (MBMTakeOver Is null OR MBMTakeOver = '')"
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
    
    If ChkGrpCom2.Value = 1 Then
        strAppend = strAppend
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
        strAppend = strAppend & " AND (CASRepCode1 Is null OR CASRepCode1 = '')"
    End If
    
    If ChkRepCat.Value = 1 Then
        strAppend = strAppend & " AND (RepCatCode Is null OR RepCatCode = '')"
    End If
    
    If ChkPlan.Value = 1 Then
        strAppend = strAppend & " AND (PLNCode Is null OR PLNCode = '')"
    End If
    
    If ChkBordx.Value = 1 Then
        strAppend = strAppend & " AND (BordxRefNo Is null OR BordxRefNo = '')"
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
    
    If ChkPolYear.Value = 1 Then
        strAppend = strAppend & " AND (MBMPolicyYear Is null OR MBMPolicyYear = '')"
    End If
    
    If ChkClaimType.Value = 1 Then
        strAppend = strAppend & " AND (PayTypeCode Is null OR PayTypeCode = '')"
    End If
    
    If ChkClaimNature.Value = 1 Then
        strAppend = strAppend & " AND (ClaimNatureCode Is null OR ClaimNatureCode = '')"
    End If
    
    If ChkBordxDate.Value = 1 Then
        strAppend = strAppend & " AND (BordxDate Is null OR BordxDate = '')"
    End If
    
    If cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(0) Then
        Call ActualCase(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(1) Then
        Call ActualClaim(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(2) Then
        Call ActualInsPlan(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(3) Then
        Call ActualInsAge(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(4) Then
        Call ActualInsSex(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(5) Then
        Call ActualAgeband(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(6) Then
        Call ActualSex(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(7) Then
        Call ActualAgebandSex(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(8) Then
        Call ActualFinalDiag(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(9) Then
        Call ActualHos(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(10) Then
        'Call PlanAgebandListing(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(11) Then
        Call ActualClaimNature(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(12) Then
        Call ActualClaimNatureSex(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(13) Then
        Call ActualPayor(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(14) Then
        Call ActualPayorInsPlan(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(15) Then
        Call ActualInsRelation(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(16) Then
        Call ActualIns(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(17) Then
        Call ActualBordxNo(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(18) Then
        Call ActualPolYear(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(19) Then
        Call ActualTakeover(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(0) Then
        Call ApprovedCase(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(1) Then
        Call ApprovedClaim(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(2) Then
        Call ApprovedInsPlan(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(3) Then
        Call ApprovedInsAge(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(4) Then
        Call ApprovedInsSex(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(5) Then
        Call ApprovedAgeband(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(6) Then
        Call ApprovedSex(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(7) Then
        Call ApprovedAgebandSex(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(8) Then
        Call ApprovedFinalDiag(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(9) Then
        Call ApprovedHos(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(10) Then
        'Call PlanAgebandListing(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(11) Then
        Call ApprovedClaimNature(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(12) Then
        Call ApprovedClaimNatureSex(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(13) Then
        Call ApprovedPayor(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(14) Then
        Call ApprovedPayorInsPlan(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(15) Then
        Call ApprovedInsRelation(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(16) Then
        Call ApprovedIns(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(17) Then
        Call ApprovedBordxNo(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(18) Then
        Call ApprovedPolYear(strAppend)
    ElseIf cboReportType = cboReportType.List(1) And cboReportSelection = cboReportSelection.List(19) Then
        Call ApprovedTakeover(strAppend)
    ElseIf cboReportType = cboReportType.List(3) Then
        Call CostDetailActual(strAppend)
        Call CostDetailApproved(strAppend)
        Call CostDetailHOS(strAppend)
        Call CostDetailMBM(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(0) Then
        Call ApprovedCaseBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(1) Then
        Call ApprovedClaimBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(2) Then
        Call ApprovedInsPlanBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(3) Then
        Call ApprovedInsAgeBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(4) Then
        Call ApprovedInsSexBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(5) Then
        Call ApprovedAgebandBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(6) Then
        Call ApprovedSexBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(7) Then
        Call ApprovedAgebandSexBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(8) Then
        Call ApprovedFinalDiagBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(9) Then
        Call ApprovedHosBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(10) Then
        'Call PlanAgebandListingBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(11) Then
        Call ApprovedClaimNatureBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(12) Then
        Call ApprovedClaimNatureSexBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(13) Then
        Call ApprovedPayorBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(14) Then
        Call ApprovedPayorInsPlanBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(15) Then
        Call ApprovedInsRelationBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(16) Then
        Call ApprovedInsBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(17) Then
        Call ApprovedBordxNoBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(18) Then
        Call ApprovedPolYearBTBG(strAppend)
    ElseIf cboReportType = cboReportType.List(2) And cboReportSelection = cboReportSelection.List(19) Then
        Call ApprovedTakeoverBTBG(strAppend)
    End If
End Function

Private Sub FinalICD(StrAppend1, strAppendCase, Record)
Dim rst3 As New ADODB.Recordset
Dim con As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter

Set con.ActiveConnection = medixneptuneconn
    con.CommandText = "ClaimReport"
    con.CommandType = adCmdStoredProc
    con.CommandTimeout = 300
    Set Prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    con.Parameters.Append Prm1
    Set Prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    con.Parameters.Append Prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = con.Execute
                        
    If cboReportType.ListIndex = 0 Then
        If rst3.EOF = True Then
            rst3.Close
            Set rst3 = Nothing
        Else
            Description = rst3!ICDDescription
            rst3.Close
            Set rst3 = Nothing
        End If
        
    ElseIf cboReportType.ListIndex = 1 Then
        If rst3.EOF = True Then
            rst3.Close
            Set rst3 = Nothing
        Else
            Description = rst3!ICDDescription
            rst3.Close
            Set rst3 = Nothing
        End If
    
    ElseIf cboReportType.ListIndex = 2 Then
        If rst3.EOF = True Then
            rst3.Close
            Set rst3 = Nothing
        Else
            Description = rst3!ICDDescription
            rst3.Close
            Set rst3 = Nothing
        End If
    End If
    
    If Len(Description) Then
        Record.SubItems(1) = Trim(Description)
    End If

End Sub

Private Sub ActualCase(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
                  
    strAppendCase = "1"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!CasNumber) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber)
                End If
                                              
                If Len(rst2!MBMNumber) Then
                   Record.SubItems(1) = rst2!MBMNumber
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(2) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(3) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(4) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(5) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(7) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(8) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(9) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(10) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(11) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(12) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(13) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(14) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(15) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(16) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(17) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(18) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(19) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(20) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(21) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(22) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(23) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(24) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(25) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(26) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(27) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(28) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(29) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualClaim(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "2"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!CasNumber & "-" & rst2!claimversion) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber & "-" & rst2!claimversion)
                End If
                                              
                If Len(rst2!MBMNumber) Then
                   Record.SubItems(1) = rst2!MBMNumber
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(2) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(3) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(4) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(5) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(7) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(8) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(9) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(10) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(11) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(12) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(13) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(14) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(15) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(16) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(17) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(18) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(19) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(20) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(21) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(22) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(23) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(24) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(25) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(26) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(27) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(28) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(29) = Format(total, "#,##0.00")
                
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


Private Sub ActualInsPlan(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "3"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                End If
                                              
                If Len(rst2!PLNCode) Then
                   Record.SubItems(1) = Trim(rst2!PLNCode)
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualInsAge(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "4"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = Trim(rst2!AgeBand)
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualAgeband(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    
    ListView2.ListItems.Clear
        
    strAppendCase = "6"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!AgeBand)
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualInsSex(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
                   
    strAppendCase = "5"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                End If
                                              
                If Len(rst2!Sex) Then
                   Record.SubItems(1) = Trim(rst2!Sex)
                Else
                   Record.SubItems(1) = "Empty"
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualSex(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "7"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!Sex) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!Sex)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualAgebandSex(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
                    
    strAppendCase = "8"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!AgeBand)
                End If
                                                
                If Len(rst2!Sex) Then
                   Record.SubItems(1) = rst2!Sex
                Else
                   Record.SubItems(1) = "Empty"
                End If
                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualFinalDiag(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1
Dim Record As ListItem
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
    
    strAppendCase = "9"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            FinalDiag = ""
            Description = ""
            With rst2
                If Len(rst2!FinalDiag) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!FinalDiag)
                    FinalDiag = rst2!FinalDiag
                    
                    strAppendCase = "2"
                    StrAppend1 = " AND ICDCode = '" & Trim(FinalDiag) & "'"
                    Call FinalICD(StrAppend1, strAppendCase, Record)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualHos(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "10"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!HosName) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!HosName)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualClaimNature(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "11"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualClaimNatureSex(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
               
    strAppendCase = "12"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                              
                If Len(rst2!Sex) Then
                   Record.SubItems(1) = rst2!Sex
                Else
                   Record.SubItems(1) = "Empty"
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualPayor(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "13"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualPayorInsPlan(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "14"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    'rst2.Open sql, medixneptuneconnWS, adOpenForwardOnly, adLockBatchOptimistic
        
            
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(1) = rst2!INSCode
                End If
                
                If Len(rst2!PLNCode) Then
                   Record.SubItems(2) = rst2!PLNCode
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(4) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(5) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(6) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(7) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(9) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(10) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(11) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(12) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(13) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(14) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(15) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(16) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(17) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(18) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(19) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(20) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(21) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(22) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(23) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(24) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(26) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(27) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualInsRelation(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "15"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!Relation) Then
                   Record.SubItems(1) = rst2!Relation
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualIns(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "16"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualBordxNo(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "17"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!BordxRefNo) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!BordxRefNo)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualPolYear(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "18"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!MBMPolicyYear) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!MBMPolicyYear)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ActualTakeover(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "19"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!MBMTakeover) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!MBMTakeover)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

'====================== Approved ===========================
Private Sub ApprovedCase(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "20"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!CasNumber) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber)
                End If
                                              
                If Len(rst2!MBMNumber) Then
                   Record.SubItems(1) = rst2!MBMNumber
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(2) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(3) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(4) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(5) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(7) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(8) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(9) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(10) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(11) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(12) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(13) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(14) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(15) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(16) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(17) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(18) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(19) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(20) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(21) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(22) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(23) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(24) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(25) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(26) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(27) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(28) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(29) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedClaim(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "21"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!CasNumber & "-" & rst2!claimversion) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber & "-" & rst2!claimversion)
                End If
                                              
                If Len(rst2!MBMNumber) Then
                   Record.SubItems(1) = rst2!MBMNumber
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(2) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(3) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(4) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(5) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(7) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(8) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(9) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(10) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(11) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(12) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(13) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(14) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(15) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(16) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(17) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(18) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(19) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(20) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(21) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(22) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(23) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(24) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(25) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(26) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(27) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(28) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(29) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedInsPlan(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "22"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                End If
                                              
                If Len(rst2!PLNCode) Then
                   Record.SubItems(1) = Trim(rst2!PLNCode)
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedInsAge(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
               
    strAppendCase = "23"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = Trim(rst2!AgeBand)
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedAgeband(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "25"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!AgeBand)
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedInsSex(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "24"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                End If
                                              
                If Len(rst2!Sex) Then
                   Record.SubItems(1) = Trim(rst2!Sex)
                Else
                   Record.SubItems(1) = "Empty"
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedSex(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "26"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!Sex) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!Sex)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedAgebandSex(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String

Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "27"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
        
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!AgeBand)
                End If
                                                
                If Len(rst2!Sex) Then
                   Record.SubItems(1) = rst2!Sex
                Else
                   Record.SubItems(1) = "Empty"
                End If
                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedFinalDiag(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
    
    strAppendCase = "28"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
           
     Do
    
        Do Until rst2.EOF
            FinalDiag = ""
            Description = ""
            With rst2
                If Len(rst2!FinalDiag) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!FinalDiag)
                    FinalDiag = rst2!FinalDiag
                    
                    strAppendCase = "2"
                    StrAppend1 = " AND ICDCode = '" & Trim(FinalDiag) & "'"
                    Call FinalICD(StrAppend1, strAppendCase, Record)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedHos(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
                   
    strAppendCase = "29"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!HosName) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!HosName)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedClaimNature(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "30"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedClaimNatureSex(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "31"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                              
                If Len(rst2!Sex) Then
                   Record.SubItems(1) = rst2!Sex
                Else
                   Record.SubItems(1) = "Empty"
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedPayor(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "32"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedPayorInsPlan(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "33"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(1) = rst2!INSCode
                End If
                
                If Len(rst2!PLNCode) Then
                   Record.SubItems(2) = rst2!PLNCode
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(4) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(5) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(6) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(7) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(9) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(10) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(11) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(12) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(13) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(14) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(15) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(16) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(17) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(18) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(19) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(20) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(21) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(22) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(23) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(24) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(26) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(27) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedInsRelation(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "34"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!Relation) Then
                   Record.SubItems(1) = rst2!Relation
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedIns(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "35"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedBordxNo(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "36"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!BordxRefNo) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!BordxRefNo)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedPolYear(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "37"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
           
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!MBMPolicyYear) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!MBMPolicyYear)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedTakeover(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "38"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!MBMTakeover) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!MBMTakeover)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Sub DoInitialSettings()
    Fgrid.row = 0
    Fgrid.col = 0
    Fgrid = "Benefits"
    Fgrid.col = 1
    Fgrid = "Actual"
    Fgrid.col = 2
    Fgrid = "Approved"
    Fgrid.col = 3
    Fgrid = "Pay2Hosp"
    Fgrid.col = 4
    Fgrid = "Pay2/fr Mem"

    Fgrid.col = 0
    
    Fgrid.ColAlignment(0) = 1
    Fgrid.ColAlignment(1) = 6
    Fgrid.ColAlignment(2) = 6
    Fgrid.ColAlignment(3) = 6
    Fgrid.ColAlignment(4) = 6
    Fgrid.ColWidth(0) = 3500
    Fgrid.ColWidth(1) = 1500
    Fgrid.ColWidth(2) = 1500
    Fgrid.ColWidth(3) = 1500
    Fgrid.ColWidth(4) = 1500
    Fgrid.row = 1
    Fgrid = "101 - Room & Board"
    
    Fgrid.row = 2
    Fgrid = "102 - Intensive Care Unit"
    
    Fgrid.row = 3
    Fgrid = "103 - Hospital Supplies & Services"
    
    Fgrid.row = 4
    Fgrid = "104 - Operating Theatre"
    
    Fgrid.row = 5
    Fgrid = "105 - MRI"
    
    Fgrid.row = 6
    Fgrid = "201 - Pre-Hospital X-ray and lab. charges"
    
    Fgrid.row = 7
    Fgrid = "202 - Pre-Hospital Specialist Fee"
    
    Fgrid.row = 8
    Fgrid = "203 - Surgical Fees"
    
    Fgrid.row = 9
    Fgrid = "204 - Anaesthetist's Fee"
    
    Fgrid.row = 10
    Fgrid = "301 - Pre-Hospital Specialist Fee"
    
    Fgrid.row = 11
    Fgrid = "302 - Daily In-Hospital Physician's Visit"
    
    Fgrid.row = 12
    Fgrid = "303 - Post-Hospitalisation Treatment"
        
    Fgrid.row = 13
    Fgrid = "304 - Medical Benefit"
    
    Fgrid.row = 14
    Fgrid = "401 - Out-Patient Accident"
    
    Fgrid.row = 15
    Fgrid = "402 - Ambulance Fees"
    
    Fgrid.row = 16
    Fgrid = "403 - Physiotherapy Treatment"
    
    Fgrid.row = 17
    Fgrid = "404 - Kidney Dialysis & Cancer Treatment"
    
    Fgrid.row = 18
    Fgrid = "501 - Organ Transplantation"
    
    Fgrid.row = 19
    Fgrid = "601 - Govt. Hospitals Income Benefit"
    
    Fgrid.row = 20
    Fgrid = "602 - Govt. Service Tax"
    
    Fgrid.row = 21
    Fgrid = "701 - Child's Guardian Allowance"
    
    Fgrid.row = 22
    Fgrid = "801 - Medical Report"
    
    Fgrid.row = 23
    Fgrid = "802 - Telephone Charges"
    
    Fgrid.row = 24
    Fgrid = "803 - Admin/Registration"
    
    Fgrid.row = 25
    Fgrid = "804 - Nursing Care"
    
    Fgrid.row = 26
    Fgrid = "901 - Others"
    
    Fgrid.row = 27
    Fgrid = "902 - Others Not Co-pay"
    
    Fgrid.row = 28
    Fgrid = "903 - Service Tax R&B"
    
    Fgrid.row = 29
    Fgrid = "904 - Service Tax ICU"
    
    Fgrid.row = 30
    Fgrid = "905 - Service Tax Lodger"
            
    Fgrid.row = 31
    Fgrid.CellFontBold = True
    Fgrid.CellFontSize = 8
    Fgrid = "Sub-Total"
    
    Fgrid.row = 32
    Fgrid.CellFontBold = True
    Fgrid.CellFontSize = 8
    Fgrid = "Excess Collected"
    
    Fgrid.row = 33
    Fgrid.CellFontBold = True
    Fgrid.CellFontSize = 8
    Fgrid = "Limit Exceeded"
    
    Fgrid.row = 34
    Fgrid.CellFontBold = True
    Fgrid.CellFontSize = 8
    Fgrid = "Final Total"
    
    Fgrid.row = 35
    Fgrid.CellFontBold = True
    Fgrid.CellFontSize = 8
    Fgrid = "Bordereaux Amount"
 
End Sub

Private Sub CostDetailActual(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim strCount
    
    total = 0
        
    strCount = " select count(*) as totalRecord From VIEW_CLMReport_Act where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        strCount = strCount + strAppend
    Else
        strCount = strCount + strAppend
    End If
    
    Set rstCount = medixneptuneconnWS.Execute(strCount)
    strAppendCase = "39"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                Fgrid.col = 1
                Fgrid.row = 1
                Fgrid.ColAlignment(1) = 6
                If Len(!CLMBoardAmt) Then
                    Fgrid = !CLMBoardAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 2
                If Len(!CLMICUAmt) Then
                    Fgrid = !CLMICUAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 3
                If Len(!CLMServices) Then
                    Fgrid = !CLMServices
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 4
                If Len(!CLMTheatre) Then
                    Fgrid = !CLMTheatre
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 5
                If Len(!CLMMRI) Then
                    Fgrid = !CLMMRI
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 6
                If Len(!CLMMPreHSAmt) Then
                    Fgrid = !CLMMPreHSAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 7
                If Len(!CLMPreHSAmt) Then
                    Fgrid = !CLMPreHSAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 8
                If Len(!CLMSurgicalAmt) Then
                    Fgrid = !CLMSurgicalAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 9
                If Len(!CLMAna) Then
                    Fgrid = !CLMAna
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 10
                If Len(!CLMNSPreHos) Then
                    If Len(!CLMPreHosp) Then
                        Fgrid = !CLMNSPreHos + !CLMPreHosp
                    Else
                        Fgrid = !CLMNSPreHos
                    End If
                ElseIf Len(!CLMPreHosp) Then
                    Fgrid = !CLMPreHosp
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 11
                If Len(!CLMPhysician) Then
                    Fgrid = !CLMPhysician
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 12
                If Len(!CLMPost) Then
                    If Len(!CLMAccPostAmt) Then
                        Fgrid = !CLMPost + !CLMAccPostAmt
                    Else
                        Fgrid = !CLMPost
                    End If
                ElseIf Len(!CLMAccPostAmt) Then
                    Fgrid = !CLMAccPostAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 13
                'If Len(!CLMCOMPreHSSum) Then
                '    Fgrid = !CLMCOMPreHSSum
                'Else
                    Fgrid = 0
                'End If
                Fgrid.row = 14
                If Len(!CLMEmergency) Then
                    Fgrid = !CLMEmergency
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 15
                If Len(!CLMAmbulance) Then
                    Fgrid = !CLMAmbulance
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 16
                If Len(!CLMOutpatient) Then
                    Fgrid = !CLMOutpatient
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 17
                If Len(!CLMMonthly) Then
                    Fgrid = !CLMMonthly
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 18
                If Len(!CLMOrgan) Then
                    Fgrid = !CLMOrgan
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 19
                If Len(!CLMCashAmt) Then
                    Fgrid = !CLMCashAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 20
                If Len(!CLMTax) Then
                    Fgrid = !CLMTax
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 21
                If Len(!CLMGuardianAmt) Then
                    Fgrid = !CLMGuardianAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 22
                If Len(!CLMMedical) Then
                    Fgrid = !CLMMedical
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 23
                If Len(!CLMPhone) Then
                    Fgrid = !CLMPhone
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 24
                If Len(!CLMAdmin) Then
                    Fgrid = !CLMAdmin
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 25
                If Len(!CLMHomeNCAmt) Then
                    Fgrid = !CLMHomeNCAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 26
                If Len(!CLMOthers1) Then
                    Fgrid = !CLMOthers1
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 27
                If Len(!CLMOthersNotCoPay) Then
                    Fgrid = !CLMOthersNotCoPay
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 28
                If Len(!CLMBoardTaxAmt) Then
                    Fgrid = !CLMBoardTaxAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 29
                If Len(!CLMICUTaxAmt) Then
                    Fgrid = !CLMICUTaxAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 30
                If Len(!CLMLodgerTaxAmt) Then
                    Fgrid = !CLMLodgerTaxAmt
                Else
                    Fgrid = 0
                End If
                                                
                total = CDbl(IIf(IsNumeric(!CLMBoardAmt) = True, !CLMBoardAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMICUAmt) = True, !CLMICUAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMServices) = True, !CLMServices, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTheatre) = True, !CLMTheatre, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMRI) = True, !CLMMRI, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMPreHSAmt) = True, !CLMMPreHSAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPreHSAmt) = True, !CLMPreHSAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMSurgicalAmt) = True, !CLMSurgicalAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAna) = True, !CLMAna, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMNSPreHos) = True, !CLMNSPreHos, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPreHosp) = True, !CLMPreHosp, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPhysician) = True, !CLMPhysician, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPost) = True, !CLMPost, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAccPostAmt) = True, !CLMAccPostAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMEmergency) = True, !CLMEmergency, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMOutpatient) = True, !CLMOutpatient, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAmbulance) = True, !CLMAmbulance, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMonthly) = True, !CLMMonthly, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMOrgan) = True, !CLMOrgan, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMCashAmt) = True, !CLMCashAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTax) = True, !CLMTax, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMGuardianAmt) = True, !CLMGuardianAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMedical) = True, !CLMMedical, 0))
                        
                total = total + CDbl(IIf(IsNumeric(!CLMPhone) = True, !CLMPhone, 0)) _
                              + CDbl(IIf(IsNumeric(!CLMAdmin) = True, !CLMAdmin, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMHomeNCAmt) = True, !CLMHomeNCAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMOthers1) = True, !CLMOthers1, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMOthersNotCoPay) = True, !CLMOthersNotCoPay, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMBoardTaxAmt) = True, !CLMBoardTaxAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMICUTaxAmt) = True, !CLMICUTaxAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMLodgerTaxAmt) = True, !CLMLodgerTaxAmt, 0))
                
                Fgrid.row = 31
                Fgrid = Format(total, "#,##0.00")
                
                'Current = Current + 1
                'lblTotalRecord = Current
                                
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
    rst2.Close
    Set rst2 = Nothing
    'CmdPrint.Enabled = True
    'CmdClear.Enabled = True
    'cmdExit.Enabled = True
End Sub

Private Sub CostDetailApproved(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim strCount
    
    total = 0
    
    strCount = " select count(*) as totalRecord From VIEW_CLMReport_App where 0 = 0 "
    If Len(Trim(strAppend)) Then
        strCount = strCount + strAppend
    Else
        strCount = strCount + strAppend
    End If
    
    Set rstCount = medixneptuneconnWS.Execute(strCount)
    strAppendCase = "40"
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        Do Until rst2.EOF
            
            With rst2
                Fgrid.col = 2
                Fgrid.row = 1
                Fgrid.ColAlignment(1) = 6
                If Len(!CLMBoardAmt) Then
                    Fgrid = !CLMBoardAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 2
                If Len(!CLMICUAmt) Then
                    Fgrid = !CLMICUAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 3
                If Len(!CLMServices) Then
                    Fgrid = !CLMServices
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 4
                If Len(!CLMTheatre) Then
                    Fgrid = !CLMTheatre
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 5
                If Len(!CLMMRI) Then
                    Fgrid = !CLMMRI
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 6
                If Len(!CLMMPreHSAmt) Then
                    Fgrid = !CLMMPreHSAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 7
                If Len(!CLMPreHSAmt) Then
                    Fgrid = !CLMPreHSAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 8
                If Len(!CLMSurgicalAmt) Then
                    Fgrid = !CLMSurgicalAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 9
                If Len(!CLMAna) Then
                    Fgrid = !CLMAna
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 10
                If Len(!CLMNSPreHos) Then
                    If Len(!CLMPreHosp) Then
                        Fgrid = !CLMNSPreHos + !CLMPreHosp
                    Else
                        Fgrid = !CLMNSPreHos
                    End If
                ElseIf Len(!CLMPreHosp) Then
                    Fgrid = !CLMPreHosp
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 11
                If Len(!CLMPhysician) Then
                    Fgrid = !CLMPhysician
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 12
                If Len(!CLMPost) Then
                    If Len(!CLMAccPostAmt) Then
                        Fgrid = !CLMPost + !CLMAccPostAmt
                    Else
                        Fgrid = !CLMPost
                    End If
                ElseIf Len(!CLMAccPostAmt) Then
                    Fgrid = !CLMAccPostAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 13
                'If Len(!CLMCOMPreHSSum) Then
                '    Fgrid = !CLMCOMPreHSSum
                'Else
                    Fgrid = 0
                'End If
                Fgrid.row = 14
                If Len(!CLMEmergency) Then
                    Fgrid = !CLMEmergency
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 15
                If Len(!CLMAmbulance) Then
                    Fgrid = !CLMAmbulance
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 16
                If Len(!CLMOutpatient) Then
                    Fgrid = !CLMOutpatient
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 17
                If Len(!CLMMonthly) Then
                    Fgrid = !CLMMonthly
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 18
                If Len(!CLMOrgan) Then
                    Fgrid = !CLMOrgan
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 19
                If Len(!CLMCashAmt) Then
                    Fgrid = !CLMCashAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 20
                If Len(!CLMTax) Then
                    Fgrid = !CLMTax
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 21
                If Len(!CLMGuardianAmt) Then
                    Fgrid = !CLMGuardianAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 22
                If Len(!CLMMedical) Then
                    Fgrid = !CLMMedical
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 23
                If Len(!CLMPhone) Then
                    Fgrid = !CLMPhone
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 24
                If Len(!CLMAdmin) Then
                    Fgrid = !CLMAdmin
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 25
                If Len(!CLMHomeNCAmt) Then
                    Fgrid = !CLMHomeNCAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 26
                If Len(!CLMOthers1) Then
                    Fgrid = !CLMOthers1
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 27
                If Len(!CLMOthersNotCoPay) Then
                    Fgrid = !CLMOthersNotCoPay
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 28
                If Len(!CLMBoardTaxAmt) Then
                    Fgrid = !CLMBoardTaxAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 29
                If Len(!CLMICUTaxAmt) Then
                    Fgrid = !CLMICUTaxAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 30
                If Len(!CLMLodgerTaxAmt) Then
                    Fgrid = !CLMLodgerTaxAmt
                Else
                    Fgrid = 0
                End If
                                                
                total = CDbl(IIf(IsNumeric(!CLMBoardAmt) = True, !CLMBoardAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMICUAmt) = True, !CLMICUAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMServices) = True, !CLMServices, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTheatre) = True, !CLMTheatre, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMRI) = True, !CLMMRI, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMPreHSAmt) = True, !CLMMPreHSAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPreHSAmt) = True, !CLMPreHSAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMSurgicalAmt) = True, !CLMSurgicalAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAna) = True, !CLMAna, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMNSPreHos) = True, !CLMNSPreHos, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPreHosp) = True, !CLMPreHosp, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPhysician) = True, !CLMPhysician, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPost) = True, !CLMPost, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAccPostAmt) = True, !CLMAccPostAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMEmergency) = True, !CLMEmergency, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMOutpatient) = True, !CLMOutpatient, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAmbulance) = True, !CLMAmbulance, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMonthly) = True, !CLMMonthly, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMOrgan) = True, !CLMOrgan, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMCashAmt) = True, !CLMCashAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTax) = True, !CLMTax, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMGuardianAmt) = True, !CLMGuardianAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMedical) = True, !CLMMedical, 0))
                        
                total = total + CDbl(IIf(IsNumeric(!CLMPhone) = True, !CLMPhone, 0)) _
                              + CDbl(IIf(IsNumeric(!CLMAdmin) = True, !CLMAdmin, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMHomeNCAmt) = True, !CLMHomeNCAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMOthers1) = True, !CLMOthers1, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMOthersNotCoPay) = True, !CLMOthersNotCoPay, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMBoardTaxAmt) = True, !CLMBoardTaxAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMICUTaxAmt) = True, !CLMICUTaxAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMLodgerTaxAmt) = True, !CLMLodgerTaxAmt, 0))
                
                Fgrid.row = 31
                Fgrid = Format(total, "#,##0.00")
                
                'Current = Current + 1
                'lblTotalRecord = Current
                                
            End With
            rst2.MoveNext
            Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CostDetailHOS(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim strCount
    
    total = 0
            
    strCount = " select count(*) as totalRecord From VIEW_CLMReport_Hos where 0 = 0 "
    If Len(Trim(strAppend)) Then
        strCount = strCount + strAppend
    Else
        strCount = strCount + strAppend
    End If
    
    Set rstCount = medixneptuneconnWS.Execute(strCount)
    strAppendCase = "41"
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        Do Until rst2.EOF
            
            With rst2
                Fgrid.col = 3
                Fgrid.row = 1
                Fgrid.ColAlignment(1) = 6
                If Len(!CLMBoardAmt) Then
                    Fgrid = !CLMBoardAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 2
                If Len(!CLMICUAmt) Then
                    Fgrid = !CLMICUAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 3
                If Len(!CLMServices) Then
                    Fgrid = !CLMServices
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 4
                If Len(!CLMTheatre) Then
                    Fgrid = !CLMTheatre
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 5
                If Len(!CLMMRI) Then
                    Fgrid = !CLMMRI
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 6
                If Len(!CLMMPreHSAmt) Then
                    Fgrid = !CLMMPreHSAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 7
                If Len(!CLMPreHSAmt) Then
                    Fgrid = !CLMPreHSAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 8
                If Len(!CLMSurgicalAmt) Then
                    Fgrid = !CLMSurgicalAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 9
                If Len(!CLMAna) Then
                    Fgrid = !CLMAna
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 10
                If Len(!CLMNSPreHos) Then
                    If Len(!CLMPreHosp) Then
                        Fgrid = !CLMNSPreHos + !CLMPreHosp
                    Else
                        Fgrid = !CLMNSPreHos
                    End If
                ElseIf Len(!CLMPreHosp) Then
                    Fgrid = !CLMPreHosp
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 11
                If Len(!CLMPhysician) Then
                    Fgrid = !CLMPhysician
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 12
                If Len(!CLMPost) Then
                    If Len(!CLMAccPostAmt) Then
                        Fgrid = !CLMPost + !CLMAccPostAmt
                    Else
                        Fgrid = !CLMPost
                    End If
                ElseIf Len(!CLMAccPostAmt) Then
                    Fgrid = !CLMAccPostAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 13
                'If Len(!CLMCOMPreHSSum) Then
                '    Fgrid = !CLMCOMPreHSSum
                'Else
                    Fgrid = 0
                'End If
                Fgrid.row = 14
                If Len(!CLMEmergency) Then
                    Fgrid = !CLMEmergency
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 15
                If Len(!CLMAmbulance) Then
                    Fgrid = !CLMAmbulance
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 16
                If Len(!CLMOutpatient) Then
                    Fgrid = !CLMOutpatient
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 17
                If Len(!CLMMonthly) Then
                    Fgrid = !CLMMonthly
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 18
                If Len(!CLMOrgan) Then
                    Fgrid = !CLMOrgan
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 19
                If Len(!CLMCashAmt) Then
                    Fgrid = !CLMCashAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 20
                If Len(!CLMTax) Then
                    Fgrid = !CLMTax
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 21
                If Len(!CLMGuardianAmt) Then
                    Fgrid = !CLMGuardianAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 22
                If Len(!CLMMedical) Then
                    Fgrid = !CLMMedical
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 23
                If Len(!CLMPhone) Then
                    Fgrid = !CLMPhone
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 24
                If Len(!CLMAdmin) Then
                    Fgrid = !CLMAdmin
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 25
                If Len(!CLMHomeNCAmt) Then
                    Fgrid = !CLMHomeNCAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 26
                If Len(!CLMOthers1) Then
                    Fgrid = !CLMOthers1
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 27
                If Len(!CLMOthersNotCoPay) Then
                    Fgrid = !CLMOthersNotCoPay
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 28
                If Len(!CLMBoardTaxAmt) Then
                    Fgrid = !CLMBoardTaxAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 29
                If Len(!CLMICUTaxAmt) Then
                    Fgrid = !CLMICUTaxAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 30
                If Len(!CLMLodgerTaxAmt) Then
                    Fgrid = !CLMLodgerTaxAmt
                Else
                    Fgrid = 0
                End If
                                                
                total = CDbl(IIf(IsNumeric(!CLMBoardAmt) = True, !CLMBoardAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMICUAmt) = True, !CLMICUAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMServices) = True, !CLMServices, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTheatre) = True, !CLMTheatre, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMRI) = True, !CLMMRI, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMPreHSAmt) = True, !CLMMPreHSAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPreHSAmt) = True, !CLMPreHSAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMSurgicalAmt) = True, !CLMSurgicalAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAna) = True, !CLMAna, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMNSPreHos) = True, !CLMNSPreHos, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPreHosp) = True, !CLMPreHosp, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPhysician) = True, !CLMPhysician, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPost) = True, !CLMPost, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAccPostAmt) = True, !CLMAccPostAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMEmergency) = True, !CLMEmergency, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMOutpatient) = True, !CLMOutpatient, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAmbulance) = True, !CLMAmbulance, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMonthly) = True, !CLMMonthly, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMOrgan) = True, !CLMOrgan, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMCashAmt) = True, !CLMCashAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTax) = True, !CLMTax, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMGuardianAmt) = True, !CLMGuardianAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMedical) = True, !CLMMedical, 0))
                        
                total = total + CDbl(IIf(IsNumeric(!CLMPhone) = True, !CLMPhone, 0)) _
                              + CDbl(IIf(IsNumeric(!CLMAdmin) = True, !CLMAdmin, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMHomeNCAmt) = True, !CLMHomeNCAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMOthers1) = True, !CLMOthers1, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMOthersNotCoPay) = True, !CLMOthersNotCoPay, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMBoardTaxAmt) = True, !CLMBoardTaxAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMICUTaxAmt) = True, !CLMICUTaxAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMLodgerTaxAmt) = True, !CLMLodgerTaxAmt, 0))
                
                Fgrid.row = 31
                Fgrid = Format(total, "#,##0.00")
                
                'Current = Current + 1
                'lblTotalRecord = Current
                                
            End With
            rst2.MoveNext
            Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CostDetailMBM(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim strCount
    
    total = 0
            
    strCount = " select count(*) as totalRecord From VIEW_CLMReport_MBM where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        strCount = strCount + strAppend
    Else
        strCount = strCount + strAppend
    End If
    
    Set rstCount = medixneptuneconnWS.Execute(strCount)
    strAppendCase = "42"
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        Do Until rst2.EOF
            
            With rst2
                Fgrid.col = 4
                Fgrid.row = 1
                Fgrid.ColAlignment(1) = 6
                If Len(!CLMBoardAmt) Then
                    Fgrid = !CLMBoardAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 2
                If Len(!CLMICUAmt) Then
                    Fgrid = !CLMICUAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 3
                If Len(!CLMServices) Then
                    Fgrid = !CLMServices
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 4
                If Len(!CLMTheatre) Then
                    Fgrid = !CLMTheatre
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 5
                If Len(!CLMMRI) Then
                    Fgrid = !CLMMRI
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 6
                If Len(!CLMMPreHSAmt) Then
                    Fgrid = !CLMMPreHSAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 7
                If Len(!CLMPreHSAmt) Then
                    Fgrid = !CLMPreHSAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 8
                If Len(!CLMSurgicalAmt) Then
                    Fgrid = !CLMSurgicalAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 9
                If Len(!CLMAna) Then
                    Fgrid = !CLMAna
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 10
                If Len(!CLMNSPreHos) Then
                    If Len(!CLMPreHosp) Then
                        Fgrid = !CLMNSPreHos + !CLMPreHosp
                    Else
                        Fgrid = !CLMNSPreHos
                    End If
                ElseIf Len(!CLMPreHosp) Then
                    Fgrid = !CLMPreHosp
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 11
                If Len(!CLMPhysician) Then
                    Fgrid = !CLMPhysician
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 12
                If Len(!CLMPost) Then
                    If Len(!CLMAccPostAmt) Then
                        Fgrid = !CLMPost + !CLMAccPostAmt
                    Else
                        Fgrid = !CLMPost
                    End If
                ElseIf Len(!CLMAccPostAmt) Then
                    Fgrid = !CLMAccPostAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 13
                'If Len(!CLMCOMPreHSSum) Then
                '    Fgrid = !CLMCOMPreHSSum
                'Else
                    Fgrid = 0
                'End If
                Fgrid.row = 14
                If Len(!CLMEmergency) Then
                    Fgrid = !CLMEmergency
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 15
                If Len(!CLMAmbulance) Then
                    Fgrid = !CLMAmbulance
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 16
                If Len(!CLMOutpatient) Then
                    Fgrid = !CLMOutpatient
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 17
                If Len(!CLMMonthly) Then
                    Fgrid = !CLMMonthly
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 18
                If Len(!CLMOrgan) Then
                    Fgrid = !CLMOrgan
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 19
                If Len(!CLMCashAmt) Then
                    Fgrid = !CLMCashAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 20
                If Len(!CLMTax) Then
                    Fgrid = !CLMTax
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 21
                If Len(!CLMGuardianAmt) Then
                    Fgrid = !CLMGuardianAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 22
                If Len(!CLMMedical) Then
                    Fgrid = !CLMMedical
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 23
                If Len(!CLMPhone) Then
                    Fgrid = !CLMPhone
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 24
                If Len(!CLMAdmin) Then
                    Fgrid = !CLMAdmin
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 25
                If Len(!CLMHomeNCAmt) Then
                    Fgrid = !CLMHomeNCAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 26
                If Len(!CLMOthers1) Then
                    Fgrid = !CLMOthers1
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 27
                If Len(!CLMOthersNotCoPay) Then
                    Fgrid = !CLMOthersNotCoPay
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 28
                If Len(!CLMBoardTaxAmt) Then
                    Fgrid = !CLMBoardTaxAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 29
                If Len(!CLMICUTaxAmt) Then
                    Fgrid = !CLMICUTaxAmt
                Else
                    Fgrid = 0
                End If
                Fgrid.row = 30
                If Len(!CLMLodgerTaxAmt) Then
                    Fgrid = !CLMLodgerTaxAmt
                Else
                    Fgrid = 0
                End If
                                                
                total = CDbl(IIf(IsNumeric(!CLMBoardAmt) = True, !CLMBoardAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMICUAmt) = True, !CLMICUAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMServices) = True, !CLMServices, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTheatre) = True, !CLMTheatre, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMRI) = True, !CLMMRI, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMPreHSAmt) = True, !CLMMPreHSAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPreHSAmt) = True, !CLMPreHSAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMSurgicalAmt) = True, !CLMSurgicalAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAna) = True, !CLMAna, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMNSPreHos) = True, !CLMNSPreHos, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPreHosp) = True, !CLMPreHosp, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPhysician) = True, !CLMPhysician, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMPost) = True, !CLMPost, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAccPostAmt) = True, !CLMAccPostAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMEmergency) = True, !CLMEmergency, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMOutpatient) = True, !CLMOutpatient, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMAmbulance) = True, !CLMAmbulance, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMonthly) = True, !CLMMonthly, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMOrgan) = True, !CLMOrgan, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMCashAmt) = True, !CLMCashAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMTax) = True, !CLMTax, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMGuardianAmt) = True, !CLMGuardianAmt, 0)) _
                        + CDbl(IIf(IsNumeric(!CLMMedical) = True, !CLMMedical, 0))
                        
                total = total + CDbl(IIf(IsNumeric(!CLMPhone) = True, !CLMPhone, 0)) _
                              + CDbl(IIf(IsNumeric(!CLMAdmin) = True, !CLMAdmin, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMHomeNCAmt) = True, !CLMHomeNCAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMOthers1) = True, !CLMOthers1, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMOthersNotCoPay) = True, !CLMOthersNotCoPay, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMBoardTaxAmt) = True, !CLMBoardTaxAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMICUTaxAmt) = True, !CLMICUTaxAmt, 0)) _
                                + CDbl(IIf(IsNumeric(!CLMLodgerTaxAmt) = True, !CLMLodgerTaxAmt, 0))
                
                Fgrid.row = 31
                Fgrid = Format(total, "#,##0.00")
                
                'Current = Current + 1
                'lblTotalRecord = Current
                                
            End With
            rst2.MoveNext
            Loop
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Public Function ExportToExcelWorksheet()
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET

    Screen.MousePointer = vbHourglass
    
    'Set objExcel = New excel.Application
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized

    Set objWorkbook = objExcel.Workbooks.Add '(LocCardPath & "Tracking\SihatMalaysia Stats-NIAM REPORT.xlt")
    Set objWorksheet = objWorkbook.Sheets(1)
    'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
        
    '============== Benefit =======================
    Fgrid.col = 0
    Fgrid.row = 0
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(1, "A") = Fgrid
    End If
    Fgrid.row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(2, "A") = Fgrid
    End If
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(3, "A") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "A") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "A") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "A") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "A") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "A") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "A") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "A") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "A") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "A") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "A") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "A") = Fgrid
    End If
    Fgrid.row = 14
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "A") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "A") = Fgrid
    End If
    Fgrid.row = 16
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "A") = Fgrid
    End If
    Fgrid.row = 17
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "A") = Fgrid
    End If
    Fgrid.row = 18
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(19, "A") = Fgrid
    End If
    Fgrid.row = 19
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "A") = Fgrid
    End If
    Fgrid.row = 20
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(21, "A") = Fgrid
    End If
    Fgrid.row = 21
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(22, "A") = Fgrid
    End If
    Fgrid.row = 22
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(23, "A") = Fgrid
    End If
    Fgrid.row = 23
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(24, "A") = Fgrid
    End If
    Fgrid.row = 24
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(25, "A") = Fgrid
    End If
    Fgrid.row = 25
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(26, "A") = Fgrid
    End If
    Fgrid.row = 26
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(27, "A") = Fgrid
    End If
    Fgrid.row = 27
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(28, "A") = Fgrid
    End If
    Fgrid.row = 28
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(29, "A") = Fgrid
    End If
    Fgrid.row = 29
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(30, "A") = Fgrid
    End If
    
    '================= ACTUAL AMT ====================
    Fgrid.col = 1
    Fgrid.row = 0
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(1, "B") = Fgrid
    End If
    Fgrid.row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(2, "B") = Fgrid
    End If
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(3, "B") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "B") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "B") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "B") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "B") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "B") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "B") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "B") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "B") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "B") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "B") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "B") = Fgrid
    End If
    Fgrid.row = 14
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "B") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "B") = Fgrid
    End If
    Fgrid.row = 16
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "B") = Fgrid
    End If
    Fgrid.row = 17
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "B") = Fgrid
    End If
    Fgrid.row = 18
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(19, "B") = Fgrid
    End If
    Fgrid.row = 19
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "B") = Fgrid
    End If
    Fgrid.row = 20
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(21, "B") = Fgrid
    End If
    Fgrid.row = 21
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(22, "B") = Fgrid
    End If
    Fgrid.row = 22
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(23, "B") = Fgrid
    End If
    Fgrid.row = 23
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(24, "B") = Fgrid
    End If
    Fgrid.row = 24
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(25, "B") = Fgrid
    End If
    Fgrid.row = 25
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(26, "B") = Fgrid
    End If
    Fgrid.row = 26
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(27, "B") = Fgrid
    End If
    Fgrid.row = 27
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(28, "B") = Fgrid
    End If
    Fgrid.row = 28
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(29, "B") = Fgrid
    End If
    Fgrid.row = 29
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(30, "B") = Fgrid
    End If
    
    '================= APPROVED AMT ====================
    Fgrid.col = 2
    Fgrid.row = 0
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(1, "C") = Fgrid
    End If
    Fgrid.row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(2, "C") = Fgrid
    End If
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(3, "C") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "C") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "C") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "C") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "C") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "C") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "C") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "C") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "C") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "C") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "C") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "C") = Fgrid
    End If
    Fgrid.row = 14
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "C") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "C") = Fgrid
    End If
    Fgrid.row = 16
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "C") = Fgrid
    End If
    Fgrid.row = 17
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "C") = Fgrid
    End If
    Fgrid.row = 18
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(19, "C") = Fgrid
    End If
    Fgrid.row = 19
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "C") = Fgrid
    End If
    Fgrid.row = 20
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(21, "C") = Fgrid
    End If
    Fgrid.row = 21
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(22, "C") = Fgrid
    End If
    Fgrid.row = 22
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(23, "C") = Fgrid
    End If
    Fgrid.row = 23
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(24, "C") = Fgrid
    End If
    Fgrid.row = 24
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(25, "C") = Fgrid
    End If
    Fgrid.row = 25
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(26, "C") = Fgrid
    End If
    Fgrid.row = 26
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(27, "C") = Fgrid
    End If
    Fgrid.row = 27
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(28, "C") = Fgrid
    End If
    Fgrid.row = 28
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(29, "C") = Fgrid
    End If
    Fgrid.row = 29
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(30, "C") = Fgrid
    End If
    
    '================= PAY2HOSP AMT ====================
    Fgrid.col = 3
    Fgrid.row = 0
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(1, "D") = Fgrid
    End If
    Fgrid.row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(2, "D") = Fgrid
    End If
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(3, "D") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "D") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "D") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "D") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "D") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "D") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "D") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "D") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "D") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "D") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "D") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "D") = Fgrid
    End If
    Fgrid.row = 14
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "D") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "D") = Fgrid
    End If
    Fgrid.row = 16
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "D") = Fgrid
    End If
    Fgrid.row = 17
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "D") = Fgrid
    End If
    Fgrid.row = 18
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(19, "D") = Fgrid
    End If
    Fgrid.row = 19
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "D") = Fgrid
    End If
    Fgrid.row = 20
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(21, "D") = Fgrid
    End If
    Fgrid.row = 21
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(22, "D") = Fgrid
    End If
    Fgrid.row = 22
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(23, "D") = Fgrid
    End If
    Fgrid.row = 23
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(24, "D") = Fgrid
    End If
    Fgrid.row = 24
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(25, "D") = Fgrid
    End If
    Fgrid.row = 25
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(26, "D") = Fgrid
    End If
    Fgrid.row = 26
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(27, "D") = Fgrid
    End If
    Fgrid.row = 27
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(28, "D") = Fgrid
    End If
    Fgrid.row = 28
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(29, "D") = Fgrid
    End If
    Fgrid.row = 29
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(30, "D") = Fgrid
    End If
    
    '================= PAY2/FR MBM AMT ====================
    Fgrid.col = 4
    Fgrid.row = 0
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(1, "E") = Fgrid
    End If
    Fgrid.row = 1
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(2, "E") = Fgrid
    End If
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(3, "E") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(4, "E") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(5, "E") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(6, "E") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "E") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "E") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "E") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "E") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "E") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "E") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "E") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "E") = Fgrid
    End If
    Fgrid.row = 14
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "E") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "E") = Fgrid
    End If
    Fgrid.row = 16
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "E") = Fgrid
    End If
    Fgrid.row = 17
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "E") = Fgrid
    End If
    Fgrid.row = 18
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(19, "E") = Fgrid
    End If
    Fgrid.row = 19
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "E") = Fgrid
    End If
    Fgrid.row = 20
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(21, "E") = Fgrid
    End If
    Fgrid.row = 21
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(22, "E") = Fgrid
    End If
    Fgrid.row = 22
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(23, "E") = Fgrid
    End If
    Fgrid.row = 23
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(24, "E") = Fgrid
    End If
    Fgrid.row = 24
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(25, "E") = Fgrid
    End If
    Fgrid.row = 25
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(26, "E") = Fgrid
    End If
    Fgrid.row = 26
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(27, "E") = Fgrid
    End If
    Fgrid.row = 27
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(28, "E") = Fgrid
    End If
    Fgrid.row = 28
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(29, "E") = Fgrid
    End If
    Fgrid.row = 29
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(30, "E") = Fgrid
    End If
        
    Screen.MousePointer = vbDefault
    objWorksheet.Columns.AutoFit
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    
    Exit Function
ExitFunction:
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Screen.MousePointer = vbDefault
    objExcel.Quit
    Exit Function
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    GoTo ExitFunction
    
End Function

Private Sub ApprovedAgebandBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "48"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!AgeBand)
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedAgebandSexBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "50"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!AgeBand) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!AgeBand)
                End If
                                                
                If Len(rst2!Sex) Then
                   Record.SubItems(1) = rst2!Sex
                Else
                   Record.SubItems(1) = "Empty"
                End If
                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedBordxNoBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "59"
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
     Do
    
        Do Until rst2.EOF

            With rst2
                If Len(rst2!BordxRefNo) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!BordxRefNo)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedCaseBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "43"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!CasNumber) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber)
                End If
                                              
                If Len(rst2!MBMNumber) Then
                   Record.SubItems(1) = rst2!MBMNumber
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(2) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(3) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(4) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(5) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(7) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(8) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(9) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(10) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(11) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(12) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(13) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(14) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(15) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(16) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(17) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(18) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(19) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(20) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(21) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(22) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(23) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(24) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(25) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(26) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(27) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(28) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(29) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedClaimBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "44"
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!CasNumber & "-" & rst2!claimversion) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber & "-" & rst2!claimversion)
                End If
                                              
                If Len(rst2!MBMNumber) Then
                   Record.SubItems(1) = rst2!MBMNumber
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(2) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(3) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(4) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(5) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(7) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(8) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(9) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(10) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(11) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(12) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(13) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(14) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(15) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(16) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(17) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(18) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(19) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(20) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(21) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(22) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(23) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(24) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(25) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(26) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(27) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(28) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(29) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedClaimNatureBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "53"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedClaimNatureSexBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "54"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                              
                If Len(rst2!Sex) Then
                   Record.SubItems(1) = rst2!Sex
                Else
                   Record.SubItems(1) = "Empty"
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedFinalDiagBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "51"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
                   
     Do
    
        Do Until rst2.EOF
            FinalDiag = ""
            Description = ""
            With rst2
                If Len(rst2!FinalDiag) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!FinalDiag)
                    FinalDiag = rst2!FinalDiag
                    
                    strAppendCase = "2"
                    StrAppend1 = " AND INDCode = '" & Trim(FinalDiag) & "'"
                    Call FinalICD(StrAppend1, strAppendCase, Record)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedHosBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "52"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!HosName) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!HosName)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedInsBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "58"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedInsPlanBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "45"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                End If
                                              
                If Len(rst2!PLNCode) Then
                   Record.SubItems(1) = Trim(rst2!PLNCode)
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedInsRelationBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "57"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!Relation) Then
                   Record.SubItems(1) = rst2!Relation
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedInsSexBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "47"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                End If
                                              
                If Len(rst2!Sex) Then
                   Record.SubItems(1) = Trim(rst2!Sex)
                Else
                   Record.SubItems(1) = "Empty"
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedPayorBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "55"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedPayorInsPlanBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "56"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(1) = rst2!INSCode
                End If
                
                If Len(rst2!PLNCode) Then
                   Record.SubItems(2) = rst2!PLNCode
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(4) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(5) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(6) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(7) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(9) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(10) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(11) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(12) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(13) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(14) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(15) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(16) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(17) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(18) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(19) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(20) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(21) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(22) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(23) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(24) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(26) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(27) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedPolYearBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "60"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            'lblTotalRecord = rstCount!totalrecord
            With rst2
                If Len(rst2!MBMPolicyYear) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!MBMPolicyYear)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedSexBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "49"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
           
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!Sex) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!Sex)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedTakeoverBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "61"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!MBMTakeover) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!MBMTakeover)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(1) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(3) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(4) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(5) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(6) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(8) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(9) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(10) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(11) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(12) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(13) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(14) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(15) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(16) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(17) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(18) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(19) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(20) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(21) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(22) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(23) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(25) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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

Private Sub ApprovedInsAgeBTBG(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim strCount

    Rnb = 0
    ICU = 0
    HSS = 0
    OT = 0
    MRI = 0
    PREXRAY = 0
    PRESpecFee = 0
    Surg = 0
    Anaest = 0
    NSPreSpec = 0
    Phys = 0
    Post = 0
    Op = 0
    Ambul = 0
    Physio = 0
    Kid = 0
    Organ = 0
    Tax = 0
    Guardian = 0
    Medical = 0
    Tel = 0
    Admin = 0
    Others1 = 0
    Others2 = 0
    total = 0
    SubTotal = 0
    Current = 0
    
    ListView2.ListItems.Clear
            
    strAppendCase = "46"
    
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "ClaimReport2"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
            
     Do
    
        Do Until rst2.EOF
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!INSCode)
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = Trim(rst2!AgeBand)
                End If
                                                
                If Len(rst2!CLMBoardAmt) Then
                   Record.SubItems(2) = Format(rst2!CLMBoardAmt, "#,##0.00")
                   If rst2!CLMBoardAmt = "-" Then
                        Rnb = 0
                   Else
                        Rnb = rst2!CLMBoardAmt
                   End If
                End If
                
                If Len(rst2!CLMICUAmt) Then
                   Record.SubItems(3) = Format(rst2!CLMICUAmt, "#,##0.00")
                    If rst2!CLMICUAmt = "-" Then
                        ICU = 0
                    Else
                        ICU = rst2!CLMICUAmt
                    End If
                End If
                
                If Len(rst2!CLMServices) Then
                   Record.SubItems(4) = Format(rst2!CLMServices, "#,##0.00")
                    If rst2!CLMServices = "-" Then
                        HSS = 0
                    Else
                        HSS = rst2!CLMServices
                    End If
                End If
                
                If Len(rst2!CLMTheatre) Then
                   Record.SubItems(5) = Format(rst2!CLMTheatre, "#,##0.00")
                    If rst2!CLMTheatre = "-" Then
                        OT = 0
                    Else
                        OT = rst2!CLMTheatre
                    End If
                End If
                
                If Len(rst2!CLMMRI) Then
                   Record.SubItems(6) = Format(rst2!CLMMRI, "#,##0.00")
                    If rst2!CLMMRI = "-" Then
                        MRI = 0
                    Else
                        MRI = rst2!CLMMRI
                    End If
                End If
                
                If Len(rst2!CLMPreHSAmt) Then
                   Record.SubItems(7) = Format(rst2!CLMPreHSAmt, "#,##0.00")
                    If rst2!CLMPreHSAmt = "-" Then
                        PREXRAY = 0
                    Else
                        PREXRAY = rst2!CLMPreHSAmt
                    End If
                End If
                
                If Len(rst2!CLMPreHosp) Then
                   Record.SubItems(8) = Format(rst2!CLMPreHosp, "#,##0.00")
                    If rst2!CLMPreHosp = "-" Then
                        PRESpecFee = 0
                    Else
                        PRESpecFee = rst2!CLMPreHosp
                    End If
                End If
                
                If Len(rst2!CLMSurgicalAmt) Then
                   Record.SubItems(9) = Format(rst2!CLMSurgicalAmt, "#,##0.00")
                    If rst2!CLMSurgicalAmt = "-" Then
                        Surg = 0
                    Else
                        Surg = rst2!CLMSurgicalAmt
                    End If
                End If
                
                If Len(rst2!CLMAna) Then
                   Record.SubItems(10) = Format(rst2!CLMAna, "#,##0.00")
                    If rst2!CLMAna = "-" Then
                        Anaest = 0
                    Else
                        Anaest = rst2!CLMAna
                    End If
                End If
                
                If Len(rst2!CLMNSPreHos) Then
                   Record.SubItems(11) = Format(rst2!CLMNSPreHos, "#,##0.00")
                    If rst2!CLMNSPreHos = "-" Then
                        NSPreSpec = 0
                    Else
                        NSPreSpec = rst2!CLMNSPreHos
                    End If
                End If
                
                If Len(rst2!CLMPhysician) Then
                   Record.SubItems(12) = Format(rst2!CLMPhysician, "#,##0.00")
                    If rst2!CLMPhysician = "-" Then
                        Phys = 0
                    Else
                        Phys = rst2!CLMPhysician
                    End If
                End If
                
                If Len(rst2!CLMPost) Then
                   Record.SubItems(13) = Format(rst2!CLMPost, "#,##0.00")
                    If rst2!CLMPost = "-" Then
                        Post = 0
                    Else
                        Post = rst2!CLMPost
                    End If
                End If
                
                If Len(rst2!CLMEmergency) Then
                   Record.SubItems(14) = Format(rst2!CLMEmergency, "#,##0.00")
                    If rst2!CLMEmergency = "-" Then
                        Op = 0
                    Else
                        Op = rst2!CLMEmergency
                    End If
                End If
                
                If Len(rst2!CLMAmbulance) Then
                   Record.SubItems(15) = Format(rst2!CLMAmbulance, "#,##0.00")
                    If rst2!CLMAmbulance = "-" Then
                        Ambul = 0
                    Else
                        Ambul = rst2!CLMAmbulance
                    End If
                End If
                
                If Len(rst2!CLMOutpatient) Then
                   Record.SubItems(16) = Format(rst2!CLMOutpatient, "#,##0.00")
                    If rst2!CLMOutpatient = "-" Then
                        Physio = 0
                    Else
                        Physio = rst2!CLMOutpatient
                    End If
                End If
                
                If Len(rst2!CLMMonthly) Then
                   Record.SubItems(17) = Format(rst2!CLMMonthly, "#,##0.00")
                    If rst2!CLMMonthly = "-" Then
                        Kid = 0
                    Else
                        Kid = rst2!CLMMonthly
                    End If
                End If
                                
                If Len(rst2!CLMOrgan) Then
                   Record.SubItems(18) = Format(rst2!CLMOrgan, "#,##0.00")
                    If rst2!CLMOrgan = "-" Then
                        Organ = 0
                    Else
                        Organ = rst2!CLMOrgan
                    End If
                End If
                                
                If Len(rst2!CLMTax) Then
                   Record.SubItems(19) = Format(rst2!CLMTax, "#,##0.00")
                    If rst2!CLMTax = "-" Then
                        Tax = 0
                    Else
                        Tax = rst2!CLMTax
                    End If
                End If
                
                If Len(rst2!CLMGuardianAmt) Then
                   Record.SubItems(20) = Format(rst2!CLMGuardianAmt, "#,##0.00")
                    If rst2!CLMGuardianAmt = "-" Then
                        Guardian = 0
                    Else
                        Guardian = rst2!CLMGuardianAmt
                    End If
                End If
                
                If Len(rst2!CLMMedical) Then
                   Record.SubItems(21) = Format(rst2!CLMMedical, "#,##0.00")
                    If rst2!CLMMedical = "-" Then
                        Medical = 0
                    Else
                        Medical = rst2!CLMMedical
                    End If
                End If
                
                If Len(rst2!CLMPhone) Then
                   Record.SubItems(22) = Format(rst2!CLMPhone, "#,##0.00")
                    If rst2!CLMPhone = "-" Then
                        Tel = 0
                    Else
                        Tel = rst2!CLMPhone
                    End If
                End If
                
                If Len(rst2!CLMAdmin) Then
                   Record.SubItems(23) = Format(rst2!CLMAdmin, "#,##0.00")
                    If rst2!CLMAdmin = "-" Then
                        Admin = 0
                    Else
                        Admin = rst2!CLMAdmin
                    End If
                End If
                
                If Len(rst2!CLMOthers1) Then
                   Record.SubItems(24) = Format(rst2!CLMOthers1, "#,##0.00")
                    If rst2!CLMOthers1 = "-" Then
                        Others1 = 0
                    Else
                        Others1 = rst2!CLMOthers1
                    End If
                End If
                
                If Len(rst2!CLMOthers2) Then
                   Record.SubItems(25) = Format(rst2!CLMOthers2, "#,##0.00")
                    If rst2!CLMOthers2 = "-" Then
                        Others2 = 0
                    Else
                        Others2 = rst2!CLMOthers2
                    End If
                End If
                
                total = CDbl(Rnb) + CDbl(ICU) + CDbl(HSS) + CDbl(OT) + CDbl(MRI) + CDbl(PREXRAY) + CDbl(PRESpecFee) _
                        + CDbl(Surg) + CDbl(Anaest) + CDbl(NSPreSpec) + CDbl(Phys) + CDbl(Post) + CDbl(Op) _
                        + CDbl(Ambul) + CDbl(Physio) + CDbl(Kid) + CDbl(Organ) + CDbl(Tax) + CDbl(Guardian) _
                        + CDbl(Medical) + CDbl(Tel) + CDbl(Admin) + CDbl(Others1) + CDbl(Others2)
                        
                Record.SubItems(26) = Format(total, "#,##0.00")
                
                Current = Current + 1
                lblTotalRecord = Current
                                
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
  
