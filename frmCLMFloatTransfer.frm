VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCLMFloatTransfer 
   Caption         =   "Claim Float Transfer"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   11730
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CheckAll 
      Caption         =   "Check All"
      Height          =   255
      Left            =   120
      TabIndex        =   96
      Top             =   3120
      Width           =   975
   End
   Begin VB.Frame FrameReceiptPV 
      Height          =   1095
      Left            =   120
      TabIndex        =   50
      Top             =   5880
      Width           =   11505
      Begin VB.TextBox txtFTLNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1125
         MaxLength       =   15
         TabIndex        =   27
         Top             =   200
         Width           =   1215
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   7680
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   240
         Width           =   3735
      End
      Begin MSMask.MaskEdBox txtFTLDate 
         Height          =   285
         Left            =   1125
         TabIndex        =   28
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtFTRecAmt 
         Height          =   285
         Left            =   1125
         MaxLength       =   10
         TabIndex        =   29
         Top             =   710
         Width           =   1215
      End
      Begin VB.ComboBox cboAccPeriodYR 
         Height          =   315
         ItemData        =   "frmCLMFloatTransfer.frx":0000
         Left            =   9720
         List            =   "frmCLMFloatTransfer.frx":0088
         TabIndex        =   39
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtPVNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   285
         Left            =   5685
         TabIndex        =   31
         Top             =   255
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboAccPeriodMth 
         Height          =   315
         ItemData        =   "frmCLMFloatTransfer.frx":0194
         Left            =   9960
         List            =   "frmCLMFloatTransfer.frx":01BC
         TabIndex        =   38
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtchqNo 
         Height          =   285
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   32
         Top             =   480
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtchqDate 
         Height          =   285
         Left            =   5685
         TabIndex        =   33
         Top             =   495
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtReceiptNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         MaxLength       =   18
         TabIndex        =   34
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtRECDate 
         Height          =   285
         Left            =   5685
         TabIndex        =   35
         Top             =   750
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Chq No"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   94
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Chq Date"
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   93
         Top             =   510
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "A/C Year"
         Height          =   255
         Left            =   10080
         TabIndex        =   88
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "A/C Mth"
         Height          =   255
         Left            =   10200
         TabIndex        =   87
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "PV Date"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   84
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "PV No"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   83
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Rec Date"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   82
         Top             =   750
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "FTL No"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   81
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "FTL Amt"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
         Height          =   255
         Left            =   7080
         TabIndex        =   53
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Rec No"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   52
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "FTL Date"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   120
      TabIndex        =   41
      Top             =   240
      Width           =   11535
      Begin VB.CheckBox chkAppCode 
         Caption         =   "Check1"
         Height          =   195
         Left            =   5040
         TabIndex        =   103
         Top             =   1560
         Width           =   195
      End
      Begin VB.CheckBox ChkApprovalDate 
         Caption         =   "Check1"
         Height          =   195
         Left            =   11160
         TabIndex        =   102
         Top             =   1680
         Width           =   195
      End
      Begin VB.TextBox txtBordxFr 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   0
         Top             =   255
         Width           =   1335
      End
      Begin VB.TextBox txtBordxTo 
         Height          =   285
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   1
         Top             =   255
         Width           =   1335
      End
      Begin VB.TextBox txtFTRECNoTo 
         Height          =   285
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   3
         Top             =   510
         Width           =   1335
      End
      Begin VB.TextBox txtFTRECNoFr 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   2
         Top             =   510
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   7800
         TabIndex        =   15
         Top             =   225
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   315
         Left            =   9720
         TabIndex        =   16
         Top             =   225
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPVNoFr 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   4
         Top             =   765
         Width           =   1335
      End
      Begin VB.TextBox txtPVNoTo 
         Height          =   285
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   5
         Top             =   765
         Width           =   1335
      End
      Begin VB.TextBox txtFTNoFr 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox txtFTNoTo 
         Height          =   285
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox txtTUNoFr 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1275
         Width           =   1335
      End
      Begin VB.TextBox txtTUNoTo 
         Height          =   285
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1275
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtRecDateFr 
         Height          =   315
         Left            =   7800
         TabIndex        =   17
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRecDateTo 
         Height          =   315
         Left            =   9720
         TabIndex        =   18
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPVDateFr 
         Height          =   315
         Left            =   7800
         TabIndex        =   19
         Top             =   750
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPVDateTo 
         Height          =   315
         Left            =   9720
         TabIndex        =   20
         Top             =   750
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFTDateFR 
         Height          =   315
         Left            =   7800
         TabIndex        =   21
         Top             =   1035
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFTDateTo 
         Height          =   315
         Left            =   9720
         TabIndex        =   22
         Top             =   1035
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTUDateFR 
         Height          =   315
         Left            =   7800
         TabIndex        =   23
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTUDateTo 
         Height          =   315
         Left            =   9720
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtAppCode1 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1530
         Width           =   1335
      End
      Begin VB.TextBox txtAppCode2 
         Height          =   285
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1530
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtAppDate2 
         Height          =   315
         Left            =   9720
         TabIndex        =   26
         Top             =   1610
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtAppDate1 
         Height          =   315
         Left            =   7800
         TabIndex        =   25
         Top             =   1610
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   1790
         Width           =   3255
      End
      Begin VB.ComboBox CboFTStatus 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   2070
         Width           =   3255
      End
      Begin VB.ComboBox cboCashType 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   2360
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Approval Code"
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
         Index           =   5
         Left            =   120
         TabIndex        =   101
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   9240
         TabIndex        =   100
         Top             =   1650
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3165
         TabIndex        =   99
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Approval Date"
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
         Index           =   2
         Left            =   6240
         TabIndex        =   98
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Cash Type"
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
         Index           =   1
         Left            =   120
         TabIndex        =   97
         Top             =   2420
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Rec Date from"
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
         Index           =   3
         Left            =   6240
         TabIndex        =   86
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Index           =   3
         Left            =   9240
         TabIndex        =   85
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Req Date from"
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
         Index           =   2
         Left            =   6240
         TabIndex        =   80
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Index           =   2
         Left            =   9240
         TabIndex        =   79
         Top             =   1365
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "PV Date from"
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
         Index           =   1
         Left            =   6240
         TabIndex        =   78
         Top             =   825
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Index           =   1
         Left            =   9240
         TabIndex        =   77
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Req No from"
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
         Index           =   3
         Left            =   120
         TabIndex        =   76
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "To"
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
         Index           =   3
         Left            =   3165
         TabIndex        =   75
         Top             =   1275
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "FT No from"
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
         Index           =   2
         Left            =   120
         TabIndex        =   74
         Top             =   1035
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "To"
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
         Index           =   2
         Left            =   3165
         TabIndex        =   73
         Top             =   1035
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "PV No from"
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
         Index           =   1
         Left            =   120
         TabIndex        =   72
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "To"
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
         Index           =   1
         Left            =   3165
         TabIndex        =   71
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "To"
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
         Index           =   1
         Left            =   3165
         TabIndex        =   70
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "Bordx No From"
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
         Index           =   1
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label46 
         Caption         =   "To"
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
         Left            =   9240
         TabIndex        =   49
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label45 
         Caption         =   "Bordx Date from"
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
         Left            =   6240
         TabIndex        =   48
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
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
         TabIndex        =   47
         Top             =   1830
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "FT Status"
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
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   2100
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "To"
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
         Index           =   0
         Left            =   3165
         TabIndex        =   45
         Top             =   510
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Recp No from"
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
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Index           =   0
         Left            =   9240
         TabIndex        =   43
         Top             =   825
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "FT Date from"
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
         Index           =   0
         Left            =   6240
         TabIndex        =   42
         Top             =   1080
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstCLMFTListing 
      Height          =   2580
      Left            =   120
      TabIndex        =   40
      Top             =   3360
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   4551
      SortKey         =   26
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   27
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Bordx Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "App Status"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Float Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "FT No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "FT Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "FT Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "PV No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "PV Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Chq No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Chq Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Rec No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Rec Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "TU No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "TU Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Acc Mth"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Acc Yr"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Remarks"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "PAYCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "NoBordx"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "ClaimNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "ClaimVersion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Payor App Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Payor App Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "CashType"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   55
      Top             =   6960
      Width           =   11535
      Begin VB.TextBox lbltotalAmt 
         Height          =   285
         Left            =   3720
         TabIndex        =   106
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox lblSelectedAmt 
         Height          =   225
         Left            =   3720
         TabIndex        =   105
         Top             =   210
         Width           =   1455
      End
      Begin VB.TextBox lblcurrent 
         Height          =   285
         Left            =   840
         TabIndex        =   104
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrintPV 
         Caption         =   "Print P&V"
         Height          =   330
         Left            =   8520
         TabIndex        =   95
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   5880
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S&earch"
         Height          =   330
         Left            =   7200
         TabIndex        =   64
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   7920
         TabIndex        =   63
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10800
         TabIndex        =   62
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10200
         TabIndex        =   61
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   60
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   59
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Height          =   330
         Left            =   6480
         TabIndex        =   58
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "&Print to File"
         Height          =   330
         Left            =   9240
         TabIndex        =   57
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   6480
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Bordx Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   67
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   65
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label lblMLAmt 
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
      Left            =   2520
      TabIndex        =   92
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "MedicaLife Float Balance :"
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
      TabIndex        =   91
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label lblMGAmt 
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
      TabIndex        =   90
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label lblMG 
      Caption         =   "MedicaGen Float Balance :"
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
      Left            =   4560
      TabIndex        =   89
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claim Float Transfer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   11685
   End
End
Attribute VB_Name = "frmCLMFloatTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim tmpPayCode As String
Dim tmpchkPAYCode As String
Dim tmpFTRecNo As String
Dim tmpFTPVNo As String
Dim intExit As Integer
Dim tmpPrevFloatAmt As Double
Dim tmpChkFloatAmt As Double
Dim tmpFloatBal As Double
Dim tmpFTFloatBLAmt As Double
Dim tmpchkItems As Integer

Dim tmpCASNumber As String
Dim tmpVersion As String
Dim tmpRecNo As String
Dim tmpRecDate As String
Dim tmpRecAmt As Double
Dim tmpChequeNo As String
Dim tmpBordxNo As String
Dim tmpcheckAmt As Double
Dim m_blnDirection(1 To 27) As Boolean



Private Sub CboFTStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboFTStatus.Text, CboFTStatus.SelStart) + Chr(KeyAscii) + Mid(CboFTStatus.Text, CboFTStatus.SelStart + CboFTStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboFTStatus.ListCount - 1
 
        If NewText = LCase(Left(CboFTStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboFTStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
                CboFTStatus.Text = ValidValue
                CboFTStatus.SelStart = 0
                CboFTStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
                CboPayor.Text = ValidValue
                CboPayor.SelStart = 0
                CboPayor.SelLength = Len(ValidValue)
             End If
        End If

End Sub


Private Sub CheckAll_Click()
Dim LstCnt As Integer

Dim counttransfered As Integer
counttransfered = 0

    LstCnt = lstCLMFTListing.ListItems.Count
    
        tmpchkItems = 0
        txtFTRecAmt = CDbl(0)
        lblSelectedAmt = CDbl(0)
                
        If LstCnt = 0 Then
            CheckAll.Value = 0
        Else
            If CheckAll.Value = 1 Then
                For LstCnt = 1 To LstCnt
                    If lstCLMFTListing.ListItems(LstCnt).Checked = False And lstCLMFTListing.ListItems(LstCnt).Text = "Transfered" Then
                        'MsgBox "This item has been transfered! ", vbCritical
                        'tmpchkItems = tmpchkItems - 1
                         counttransfered = counttransfered + 1
                    ElseIf lstCLMFTListing.ListItems(LstCnt).Checked = False Then
                        lstCLMFTListing.ListItems(LstCnt).Checked = True
                        lstCLMFTListing.ListItems(LstCnt).Text = "Transfer"
                        
                        txtFTRecAmt = CDbl(txtFTRecAmt) + CDbl(lstCLMFTListing.ListItems(LstCnt).SubItems(3))
                         lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(lstCLMFTListing.ListItems(LstCnt).SubItems(3))
                         tmpchkItems = tmpchkItems + 1
                    ElseIf lstCLMFTListing.ListItems(LstCnt).Checked = True Then
                        lstCLMFTListing.ListItems(LstCnt).Checked = False
                        lstCLMFTListing.ListItems(LstCnt).Text = ""
                        'tmpchkItems = tmpchkItems - 1
                    End If
                Next LstCnt
            Else
                For LstCnt = 1 To LstCnt
                    If lstCLMFTListing.ListItems(LstCnt).Checked = True Then
                        lstCLMFTListing.ListItems(LstCnt).Checked = False
                        lstCLMFTListing.ListItems(LstCnt).Text = ""
                    End If
                Next LstCnt
                tmpchkItems = 0
                txtFTRecAmt = CDbl(0)
                lblSelectedAmt = CDbl(0)
            End If
            
            If lstCLMFTListing.ListItems.Count = counttransfered And counttransfered <> 0 Then
                MsgBox "All items have been transferred! ", vbCritical
            ElseIf lstCLMFTListing.ListItems.Count <> counttransfered And counttransfered <> 0 Then
                MsgBox "Some ( " + counttransfered + " )item have been transferred! ", vbCritical
            End If
        End If

End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    CboFTStatus.ListIndex = 0
    lstCLMFTListing.ListItems.Clear
    lblSelectedAmt = 0
    txtFTRecAmt = 0
    EditFlag = False
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    CmdEdit.Visible = False
    CmdUpdate.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdEXPtoExcel_Click()
    Screen.MousePointer = vbHourglass
    
    If lstCLMFTListing.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstCLMFTListing)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPrintPV_Click()
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim strsql As String
Dim rstTU As New ADODB.Recordset
Dim TUCnt As Integer
Dim PAY As New ADODB.Recordset
Dim strsqlPay As String
Dim tmpAdd1 As String
Dim tmpAdd2 As String
Dim tmpAdd3 As String
Dim tmpPayName As String
Dim strsqlFLT As String
Dim rstFLT As New ADODB.Recordset

    Screen.MousePointer = vbHourglass

    If txtFTLNo = "" Then
    
        MsgBox "Please Enter or Generate Rec No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\FloatTrans.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        If Err.Number > 0 Then
            MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
            GoTo ExitFunction
        End If
        
        On Error GoTo HANDLE_ERROR
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        tmpPayCode = IIf(Mid(txtFTLNo, 1, 3) = "FTL", "XL", "XG")

        'strsqlPay = "Select PAYCode, PAYCompanyName,PAYAddress1,PAYAddress2,PAYAddress3 from PAY where PAYCode = '" & tmpPAYCode & "'"
        'PAY.Open strsqlPay, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
    
        'Do Until PAY.EOF
        '    tmpPayName = PAY!PAYCompanyName
        '    tmpAdd1 = PAY!PAYAddress1
        '    tmpAdd2 = PAY!PAYAddress2
        '    tmpAdd3 = PAY!PAYAddress3
        'PAY.MoveNext
        'Loop
        'PAY.Close
        'Set PAY = Nothing
        strsqlFLT = ""
        'strsqlFLT = "Select CLMFTNo, CLMFTDate,CLMFTAmt,CLMNoodBordx,CLMPVNo,CLMPVDate,CLMRecNo from CLMFloatTransfer where CLMFTNo = '" & Trim(txtFTLNo) & "'"
        strsqlFLT = "Select CLMFTNo, CONVERT(varchar(10), CLMFTDate, 103) as CLMFTDate ,CLMFTAmt,CLMNoodBordx,CLMPVNo,CONVERT(varchar(10), CLMPVDate, 103) as CLMPVDate ,CLMRecNo from CLMFloatTransfer where CLMFTNo = '" & Trim(txtFTLNo) & "'"
        rstTU.Open strsqlFLT, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
        
        TUCnt = 0
        Do While Not rstTU.EOF
            TUCnt = TUCnt + 1
            If TUCnt = 1 Then
            '    objWorksheet.Cells(9, "C") = tmpPayName
            '    objWorksheet.Cells(10, "C") = IIf(Trim(tmpAdd1) <> "", Trim(tmpAdd1), "")
            '    objWorksheet.Cells(11, "C") = IIf(Trim(tmpAdd2) <> "", Trim(tmpAdd2), "")
            '    objWorksheet.Cells(12, "C") = IIf(Trim(tmpAdd3) <> "", Trim(tmpAdd3), "")
                objWorksheet.Cells(9, "J") = rstTU!CLMPVNo
                'objWorksheet.Cells(10, "J") = rstTU!CLMPVDate
                objWorksheet.Cells(10, "J") = Format(rstTU!CLMPVDate, "dd/MM/yyyy")
                objWorksheet.Cells(11, "J") = txtchqNo
                
                'If tmpPAYCode = "XL" Then
                '    objWorksheet.Cells(35, "C") = Format(lblMLAmt, "#,##0.00")
                'Else
                '    objWorksheet.Cells(35, "C") = Format(lblMLAmt, "#,##0.00")
                'End If
                objWorksheet.Cells(36, "C") = Logon
            End If
            
            objWorksheet.Cells(13 + CDbl(TUCnt), "A") = TUCnt
            objWorksheet.Cells(13 + CDbl(TUCnt), "B") = rstTU!CLMFTNo
            objWorksheet.Cells(13 + CDbl(TUCnt), "C") = Format(rstTU!CLMFTDate, "dd/MM/yyyy")
            objWorksheet.Cells(13 + CDbl(TUCnt), "D") = rstTU!CLMNoodBordx
            objWorksheet.Cells(13 + CDbl(TUCnt), "E") = rstTU!CLMRecNo
            objWorksheet.Cells(13 + CDbl(TUCnt), "J") = Format(rstTU!CLMFTAmt, "#,##0.00")

            rstTU.MoveNext
        Loop
        rstTU.Close
        Set rstTU = Nothing
        
        Set objWorksheet = objWorkbook.Sheets(2)
        Set objWorksheet = objWorkbook.Worksheets.Item(2)
        
        strsqlFLT = ""
        'strsqlFLT = "Select CLMFTNo, CLMFTDate,CLMFTAmt,CLMNoodBordx,CLMTUNo,CLMTUDate,BordxRefNo,BordxDate,BordxClaimAmt,CLMPayorAppDate,CLMPayorAppCode from View_CLMFloatTransfer where CLMFTNo = '" & Trim(txtFTLNo) & "' order by BordxRefNo asc"
        strsqlFLT = "Select CLMFTNo, CONVERT(varchar(10), CLMFTDate, 103) as CLMFTDate,CLMFTAmt,CLMNoodBordx,CLMTUNo,CONVERT(varchar(10), CLMTUDate, 103) as CLMTUDate ,BordxRefNo,CONVERT(varchar(10), BordxDate, 103) as BordxDate ,BordxClaimAmt,CONVERT(varchar(10), CLMPayorAppDate, 103) as CLMPayorAppDate ,CLMPayorAppCode from View_CLMFloatTransfer where CLMFTNo = '" & Trim(txtFTLNo) & "' order by BordxRefNo asc"
        rstTU.Open strsqlFLT, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
        
        Dim sumBordxClaimAmt As Double
        sumBordxClaimAmt = 0
        TUCnt = 0
     
        Do While Not rstTU.EOF
            TUCnt = TUCnt + 1
            If TUCnt = 1 Then
                If Len(Trim(cboCashType)) Then
                    If Mid(cboCashType, 1, 1) = "M" Then
                        objWorksheet.Cells(1, "B") = " - Member "
                    ElseIf Mid(cboCashType, 1, 1) = "H" Then
                        objWorksheet.Cells(1, "B") = " - Hospital "
                    End If
                End If
             End If
            objWorksheet.Cells(3 + CDbl(TUCnt), "A") = TUCnt
            objWorksheet.Cells(3 + CDbl(TUCnt), "B") = rstTU!CLMFTNo
            objWorksheet.Cells(3 + CDbl(TUCnt), "C") = Format(rstTU!CLMFTDate, "dd/MM/yyyy")
            objWorksheet.Cells(3 + CDbl(TUCnt), "D") = rstTU!BordxRefNo
            objWorksheet.Cells(3 + CDbl(TUCnt), "E") = Format(rstTU!BordxDate, "dd/MM/yyyy")
            objWorksheet.Cells(3 + CDbl(TUCnt), "F") = rstTU!CLMPayorAppCode
            objWorksheet.Cells(3 + CDbl(TUCnt), "G") = Format(rstTU!CLMPayorAppDate, "mm/dd/yyyy")
            objWorksheet.Cells(3 + CDbl(TUCnt), "H") = rstTU!BordxClaimAmt
            sumBordxClaimAmt = CDbl(sumBordxClaimAmt) + CDbl(rstTU!BordxClaimAmt)
            rstTU.MoveNext
        Loop
        objWorksheet.Cells(3 + CDbl(TUCnt + 2), "H") = sumBordxClaimAmt
        rstTU.Close
        Set rstTU = Nothing
        
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing
        'cmdCredit.Enabled = False
        objExcel.DisplayAlerts = True
        'objWorksheet.Columns.AutoFit
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        Exit Sub
HANDLE_ERROR:
        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        medixmasterconnWS.Close
        Set medixmasterconnWS = Nothing
        GoTo ExitFunction
End Sub

Private Sub CmdSave_Click()
Dim AA As Integer
Dim ItmChk As Boolean
Dim BB As Integer
Dim tmpSql As String
Dim startTrans As Boolean
Dim bSucess  As Boolean
Dim DiffPay As Boolean
Dim strsqlRem1 As String
Dim strsqlRem2 As String
Dim strsqlRem3 As String
Dim REM1 As New ADODB.Recordset
Dim REM2 As New ADODB.Recordset
Dim REM3 As New ADODB.Recordset
Dim RemLen1 As Integer
Dim RemLen2 As Integer
Dim RemLen3 As Integer
Dim RemLen4 As Integer
Dim RemLen5 As Integer
Dim RemLen6 As Integer
Dim RemLen7 As Integer
Dim Remarks1 As String
Dim Remarks2 As String
Dim Remarks3 As String
Dim Remarks4 As String
Dim Remarks5 As String
Dim Remarks6 As String
Dim Remarks7 As String
Dim AddRemark As String
Dim strsqlREM As String
    ' Update Claim Table
    RemLen1 = 0
    RemLen2 = 0
    RemLen3 = 0
    RemLen4 = 0
    Remarks1 = ""
    Remarks2 = ""
    Remarks3 = ""
    Remarks4 = ""
    ItmChk = False
    DiffPay = False
    
        AA = lstCLMFTListing.ListItems.Count
        'For BB = 1 To lstCLMFTListing.ListItems.Count
        '    If lstCLMFTListing.ListItems(BB).Checked = True Then
        '        If ItmChk = False Then
        '            ItmChk = True
        '        End If
                    'If tmpchkPAYCode <> "" Then
                    '    If tmpchkPAYCode <> lstCLMFTListing.ListItems(BB).SubItems(20) Then
                    '        DiffPay = True
                    '        Exit For
                    '    End If
                    'End If
                    'tmpchkPAYCode = lstCLMFTListing.ListItems(BB).SubItems(20)
                    'tmpPayCode = Trim(IIf(lstCLMFTListing.ListItems(BB).SubItems(20) = "XG", "XG", "XL"))
                    
                    
                'Exit For
         '   End If
        'Next BB
    'Else
        AA = lstCLMFTListing.ListItems.Count
        For BB = 1 To lstCLMFTListing.ListItems.Count
            If lstCLMFTListing.ListItems(BB).Checked = True Then
    
                ItmChk = True
                Exit For
            End If
        Next BB
    'End If
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        'Call SearchFloatAmt("And FTPayCode = '" & Trim(tmpPayCode) & "'")
        'tmpFloatBal = CDbl(tmpChkFloatAmt) - CDbl(txtFTRecAmt)
    'medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
    
    If AA = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf ItmChk = False Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf txtFTLDate = "__/__/____" Then
        MsgBox "Please Enter FT Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtFTLDate.SetFocus
    ElseIf Trim(txtchqNo) = "" Then
       MsgBox "Please Enter Cheque No!", vbCritical
       txtchqNo.SetFocus
    ElseIf txtchqDate = "__/__/____" Then
        MsgBox "Please Enter Cheque Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtchqDate.SetFocus
    'ElseIf CDbl(txtFTRecAmt) = "" Then
    '    MsgBox "Please Transfer Amount!", vbCritical
    '    txtReceiptRemark.SetFocus
    'ElseIf Trim(txtRECDate) = "__/__/____" Then
    '    MsgBox "Please Enter Receipt Date ! ", vbOKOnly + vbCritical, DialogBoxTitle
    '    TxtReceiptNo.SetFocus
    ElseIf Trim(txtPVDate) = "__/__/____" Then
        MsgBox "Please Enter PV Date ! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtPVDate.SetFocus
    'ElseIf DiffPay = True Then
    '    MsgBox "Different Payors for this transaction is not allowed! ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf CDbl(lblSelectedAmt) <> CDbl(txtFTRecAmt) Then
        MsgBox "Please check Float Transfer Amount! ", vbOKOnly + vbCritical, DialogBoxTitle
    'ElseIf CDbl(tmpFloatBal) <= 0 Then
    '    MsgBox "Insufficient Float! ", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Screen.MousePointer = vbHourglass
        'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        'medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

        startTrans = True
        medixmasterconnWS.beginTrans
        medixmasterconnPayment.beginTrans
        medixmasterconn.beginTrans
        
            'tmpRecNo = GenerateSRecNo("RC")
            tmpFTRecNo = GenerateFTRecNo("XL")
            tmpFTPVNo = GenerateFTPVNo("XL")
            
            Call SaveFTFloat
            For BB = 1 To lstCLMFTListing.ListItems.Count
                If lstCLMFTListing.ListItems(BB).Checked = True Then
                    
                    tmpBordxNo = lstCLMFTListing.ListItems(BB).SubItems(1)
                    tmpcheckAmt = lstCLMFTListing.ListItems(BB).SubItems(3)
                    tmpCASNumber = lstCLMFTListing.ListItems(BB).SubItems(22)
                    tmpVersion = lstCLMFTListing.ListItems(BB).SubItems(23)
                    tmpRecAmt = CDbl(txtFTRecAmt)
                    
                    Call SaveFTBordx(tmpBordxNo)
                    'Call SaveReceiptPayor
                    If Len(tmpCASNumber) Then
                        AddRemark = ""
                        AddRemark = vbNewLine & Date & "   Claim Float Transfer  -  " & tmpFTRecNo & " on " & txtFTLDate & "                                         " & Logon & " "
                        'AddRemark = AddRemark & vbNewLine & Date & "   Claim Float Receipt  -  " & tmpRecNo & " on " & txtRECDate & "                                " & Logon & " "
                        AddRemark = ChkStr(Replace(AddRemark, Chr(13), "~"))
                        strsqlRem1 = "Select * from CASRemarks where CASNumber = '" & Trim(tmpCASNumber) & "'"
                        REM1.Open strsqlRem1, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If REM1.recordCount > 0 Then
                                RemLen1 = Len(REM1!CASORemarks)
                                Remarks1 = REM1!CASORemarks & AddRemark
                                Remarks1 = ChkStr(Replace(Remarks1, "'", "''"))
                                RemLen2 = Len(REM1!CASORemarksTwo)
                                Remarks2 = REM1!CASORemarksTwo & AddRemark
                                Remarks2 = ChkStr(Replace(Remarks2, "'", "''"))
                            Else
                                RemLen1 = 0
                                RemLen2 = 0
                                Remarks1 = ""
                                Remarks2 = ""
                            End If
                        REM1.Close
                        Set REM1 = Nothing
                        
                        strsqlRem2 = "Select * from CASRemarksTwo where CASNumber= '" & Trim(tmpCASNumber) & "'"
                        REM2.Open strsqlRem2, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If REM2.recordCount > 0 Then
                                RemLen3 = Len(REM2!CASORemarks)
                                Remarks3 = REM2!CASORemarks & AddRemark
                                Remarks3 = ChkStr(Replace(Remarks3, "'", "''"))
                                RemLen4 = Len(REM2!CASORemarksTwo)
                                Remarks4 = REM2!CASORemarksTwo & AddRemark
                                Remarks4 = ChkStr(Replace(Remarks4, "'", "''"))
                            Else
                                RemLen3 = 0
                                Remarks4 = 0
                                Remarks3 = ""
                                Remarks4 = ""
                            End If
                        REM2.Close
                        Set REM2 = Nothing
                        'REM3.Close
                        'Set REM3 = Nothing
                        strsqlRem3 = ""
                        strsqlRem3 = "Select * from CASRemarkThree where CASNumber= '" & Trim(tmpCASNumber) & "'"
                        REM3.Open strsqlRem3, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If REM3.recordCount Then
                            RemLen5 = Len(REM3!CasRemarks)
                            Remarks5 = REM3!CasRemarks & AddRemark
                            Remarks5 = ChkStr(Replace(Remarks5, "'", "''"))
                            RemLen6 = Len(REM3!CASRemarksFour)
                            Remarks6 = REM3!CASRemarksFour & AddRemark
                            Remarks6 = ChkStr(Replace(Remarks6, "'", "''"))
                            RemLen7 = Len(REM3!CASRemarksFive)
                            Remarks7 = REM3!CASRemarksFive & AddRemark
                            Remarks7 = ChkStr(Replace(Remarks7, "'", "''"))
                        Else
                            RemLen5 = 0
                            Remarks5 = ""
                            RemLen6 = 0
                            Remarks6 = ""
                            RemLen7 = 0
                            Remarks7 = ""
                        End If
                        REM3.Close
                        Set REM3 = Nothing
                        
                        strsqlREM = ""
                        If (RemLen7 <> 0 And RemLen7 < 3700) Or RemLen6 > 3700 Then
                        'Remarks7 = "Test Msg"
                            strsqlREM = " update CASRemarkThree set " & _
                                     " CASRemarksFive = '" & Trim(CStr(Remarks7)) & "'" & _
                                     " where CASNumber = '" & Trim(tmpCASNumber) & "'"
                        ElseIf (RemLen6 <> 0 And RemLen6 < 3700) Or RemLen5 > 3700 Then
                            strsqlREM = " update CASRemarkThree set " & _
                                     " CASRemarksFour = '" & Trim(Remarks6) & "'" & _
                                     " where CASNumber = '" & Trim(tmpCASNumber) & "'"
                        ElseIf (RemLen5 <> 0 And RemLen5 < 3700) Or RemLen4 > 1800 Then
                            strsqlREM = " update CASRemarkThree set " & _
                                     " CasRemarks = '" & Remarks5 & "'" & _
                                     " where CASNumber = '" & Trim(tmpCASNumber) & "'"
                        ElseIf (RemLen4 <> 0 And RemLen4 < 1800) Or RemLen3 > 1800 Then
                            strsqlREM = " update CASRemarksTwo set " & _
                                     " CASORemarksTwo = '" & Remarks4 & "'" & _
                                     " where CASNumber = '" & Trim(tmpCASNumber) & "'"
                        ElseIf (RemLen3 <> 0 And RemLen3 < 1800) Or RemLen2 > 1800 Then
                            strsqlREM = " update CASRemarksTwo set " & _
                                     " CASORemarks = '" & Remarks3 & "' " & _
                                     " where CASNumber = '" & Trim(tmpCASNumber) & "'"
                        ElseIf (RemLen2 <> 0 And RemLen2 < 1800) Or RemLen1 > 1800 Then
                            strsqlREM = " update CasRemarks set " & _
                                     " CASORemarksTwo = '" & Mid(Remarks2, 1, 1799) & "'" & _
                                     " where CASNumber = '" & Trim(tmpCASNumber) & "'"
                        ElseIf RemLen1 <> 0 And RemLen1 < 1800 Then
                            strsqlREM = " update CasRemarks set " & _
                                     " CASORemarks = '" & Mid(Remarks1, 1, 1799) & "'" & _
                                     " where CASNumber = '" & Trim(tmpCASNumber) & "'"
                        Else
                            strsqlREM = " update CasRemarks set " & _
                                     " CASORemarks = '" & Remarks1 & "'" & _
                                     " where CASNumber = '" & Trim(tmpCASNumber) & "'"

                        End If
                        medixmasterconn.Execute strsqlREM

                    End If
                    
                End If
            Next BB
            Call UpdateFloatAmt
            medixmasterconnWS.commitTrans
            medixmasterconnPayment.commitTrans
            medixmasterconn.commitTrans
            startTrans = False
            txtFTLNo = tmpFTRecNo
            'TxtReceiptNo = tmpRecNo
            txtPVNo = tmpFTPVNo
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            If tmpFTFloatBLAmt < 300000 Then
                MsgBox "Float Balance is " & tmpFTFloatBLAmt & " for " & tmpPayCode, vbOKOnly + vbInformation, DialogBoxTitle
            End If
            
            tmpSql = ""
            tmpSql = " and  CLMFTNo = '" & Trim(txtFTLNo) & "'"
            lstCLMFTListing.ListItems.Clear
            Call CLMFloatListing(tmpSql)
        EditFlag = False
        
        
        
       ' medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
       ' medixmasterconnPayment.Close
       ' Set medixmasterconnPayment = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        Screen.MousePointer = vbDefault
        'Call ClearContraDetail
        'Call ClearRecDetail
        lblSelectedAmt = 0
        lbltotalAmt = 0
    End If
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If bSucess = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnPayment.RollbackTrans
    End If
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
    MsgBox Err.Description
    
End Sub

Private Sub CmdSearch_Click()
On Error GoTo errRun
Dim strAppend As String
    strAppend = ""
    CmdSearch.Enabled = False
    'cmdprint.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    Call ClearFrameCase
    Screen.MousePointer = vbHourglass
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Call SearchRecords
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
    Screen.MousePointer = vbDefault
    CmdSearch.Enabled = True
    CmdEdit.Visible = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    Exit Sub

errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstCLMFTListing.ListItems.Clear
    lblcurrent = ""
    lbltotalAmt = ""
    lblSelectedAmt = ""
    'cmdprint.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    CmdEdit.Visible = True
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
End Sub

Private Function SearchRecords()
Dim sql As String
Dim strAppend As String

    strAppend = ""
    'GrpCompany = ""
    'HOspitals = ""
    
    If Len(Trim(txtBordxFr)) Then
        strAppend = strAppend + " AND  BordxRefNo >= '" & Trim(txtBordxFr) & "'"
    End If
    If Len(Trim(txtBordxTo)) Then
        strAppend = strAppend + " AND  BordxRefNo <= '" & Trim(txtBordxTo) & "'"
    End If
    
    If Len(Trim(txtFTRECNoFr)) Then
        strAppend = strAppend + " AND CLMRecNo >= '" & Trim(txtFTRECNoFr) & "'"
    End If
    If Len(Trim(txtFTRECNoTo)) Then
        strAppend = strAppend + " AND CLMRecNo <= '" & Trim(txtFTRECNoTo) & "'"
    End If
    
    If Len(Trim(txtPVNoFr)) Then
        strAppend = strAppend + " AND   CLMPVNo >= '" & Trim(txtPVNoFr) & "'"
    End If
    If Len(Trim(txtPVNoTo)) Then
        strAppend = strAppend + " AND   CLMPVNo <= '" & Trim(txtPVNoTo) & "'"
    End If
    
    If Len(Trim(txtFTNoFr)) Then
        strAppend = strAppend + " AND CLMFTNo >= '" & Trim(txtFTNoFr) & "'"
    End If
    If Len(Trim(txtFTNoTo)) Then
        strAppend = strAppend + " AND CLMFTNo <= '" & Trim(txtFTNoTo) & "'"
    End If
    
    
    If Len(Trim(txtTUNoFr)) Then
        strAppend = strAppend + " AND  CLMTUNo >= '" & Trim(txtTUNoFr) & "'"
    End If
    
    If Len(Trim(txtTUNoTo)) Then
        strAppend = strAppend + " AND  CLMTUNo <= '" & Trim(txtTUNoTo) & "'"
    End If
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " AND  PAYCode = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
    End If
    
    If Len(Trim(CboFTStatus)) Then
        If Mid(CboFTStatus, 1, 1) = "0" Then
            strAppend = strAppend + " AND  (FLTFTNo Is null OR FLTFTNo = '') "
        ElseIf Mid(CboFTStatus, 1, 1) = "1" Then
            strAppend = strAppend + " AND  (FLTFTNo Is NOT null) "
        End If
    End If
    
    If Len(Trim(txtBordxDateFr)) > 0 And IsDate(txtBordxDateFr) = True Then
        strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtBordxDateTo)) > 0 And IsDate(txtBordxDateTo) = True Then
        strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtFTDateFR)) > 0 And IsDate(txtFTDateFR) = True Then
        strAppend = strAppend + " And CLMFTDate >= '" & Format(txtFTDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtFTDateTo)) > 0 And IsDate(txtFTDateTo) = True Then
        strAppend = strAppend + " And CLMFTDate <= '" & Format(txtFTDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDateFr)) > 0 And IsDate(txtPVDateFr) = True Then
        strAppend = strAppend + " And CLMPVDate >= '" & Format(txtPVDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDateTo)) > 0 And IsDate(txtPVDateTo) = True Then
        strAppend = strAppend + " And CLMPVDate <= '" & Format(txtPVDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtTUDateFR)) > 0 And IsDate(txtTUDateFR) = True Then
        strAppend = strAppend + " And CLMTUDate >= '" & Format(txtTUDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtTUDateTo)) > 0 And IsDate(txtTUDateTo) = True Then
        strAppend = strAppend + " And CLMTUDate <= '" & Format(txtTUDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRECDateFr)) > 0 And IsDate(txtRECDateFr) = True Then
        strAppend = strAppend + " And CLMRecDate >= '" & Format(txtRECDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRECDateTo)) > 0 And IsDate(txtRECDateTo) = True Then
        strAppend = strAppend + " And CLMRecDate <= '" & Format(txtRECDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(cboCashType)) Then
        strAppend = strAppend + " And INVPayTo = '" & Mid(cboCashType, 1, 1) & "'"
    End If
    
    If ChkApprovalDate.Value = 1 Then
        strAppend = strAppend + " AND (CLMPayorAppDate Is null OR CLMPayorAppDate = '')"
    End If
    
    If chkAppCode.Value = 1 Then
        strAppend = strAppend + " AND (CLMPayorAppCode Is null OR CLMPayorAppCode = '')"
    End If
    
    If Len(Trim(txtAppCode1)) > 0 And Trim(txtAppCode1) = True Then
        strAppend = strAppend + " And CLMPayorAppCode >= '" & Trim(txtAppCode1) & "'"
    End If
    
    If Len(Trim(txtAppCode2)) > 0 And Trim(txtAppCode2) = True Then
        strAppend = strAppend + " And CLMPayorAppCode <= '" & Trim(txtAppCode2) & "'"
    End If
    
    If Len(Trim(txtAppDate1)) > 0 And IsDate(txtAppDate1) = True Then
        strAppend = strAppend + " And CLMPayorAppDate >= '" & Format(txtAppDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtAppDate2)) > 0 And IsDate(txtAppDate2) = True Then
        strAppend = strAppend + " And CLMPayorAppDate <= '" & Format(txtAppDate2, "mm/dd/yyyy") & "'"
    End If
    
        Call CLMFloatListing(strAppend)
        Call SearchFloatAmt("")
End Function

Private Sub CLMFloatListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim Current As String
Dim Check As Boolean
Dim strsql As String
Dim RecFound As Boolean
Dim cmd As New ADODB.Command
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Record As ListItem
Dim BordxAPPStatus As String
Dim strAppendCase As Integer
    Current = 0
    'If ChkScan.value = 0 Then
        lstCLMFTListing.ListItems.Clear
        lblcurrent = 0
        lblSelectedAmt = 0
        lbltotalAmt = 0
    'End If
        strAppendCase = 1
    Set cmd1.ActiveConnection = medixmasterconnWS
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_ClaimFloatTransfer"
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("@strAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd1.Parameters.Append Prm2
    
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    rst2.CursorLocation = adUseClient
    Set rst2 = cmd1.Execute
    RecFound = False
    Do
        Do Until rst2.EOF

        RecFound = True
        
        With rst2
            If Len(Trim(!CLMFTNo)) Then
                Set Record = lstCLMFTListing.ListItems.Add(, , "Transfered")
            Else
                Set Record = lstCLMFTListing.ListItems.Add(, , "")
            End If
            Record.SubItems(1) = IIf(Len(!BordxRefNo), !BordxRefNo, "")
            Record.SubItems(2) = IIf(Len(!BordxDate), Format(!BordxDate, "dd/mm/yyyy"), "")
            Record.SubItems(3) = Format(IIf(IsNumeric(!BordxClaimAmt), !BordxClaimAmt, 0), "#,##0.00")
            Record.SubItems(4) = IIf(Len(!CLMPayorAppStatus), !CLMPayorAppStatus, "")
            Record.SubItems(5) = IIf(Trim(!PLNFloatStatus) <> "", Trim(!PLNFloatStatus), "")
            Record.SubItems(6) = IIf(Len(!CLMFTNo), Format(!CLMFTNo, "dd/mm/yyyy"), "")
            Record.SubItems(7) = IIf(Trim(!CLMFTDate) <> "", Format(!CLMFTDate, "dd/mm/yyyy"), "")
            Record.SubItems(8) = Format(IIf(Len(!CLMFTAmt), !CLMFTAmt, 0), "#,##0.00")
            Record.SubItems(9) = IIf(Trim(Len(!CLMPVNo)), !CLMPVNo, "")
            Record.SubItems(10) = IIf(Trim(!CLMPVDate) <> "", Format(!CLMPVDate, "dd/mm/yyyy"), "")
            Record.SubItems(11) = IIf(Trim(!CLMCHQNo) <> "", Trim(!CLMCHQNo), "")
            Record.SubItems(12) = IIf(Len(!CLMCHQDate), Format(!CLMCHQDate, "dd/mm/yyyy"), "")
            Record.SubItems(13) = IIf(Len(!CLMRecNo), !CLMRecNo, "")
            Record.SubItems(14) = IIf(Trim(!CLMRecDate) <> "", Format(!CLMRecDate, "dd/mm/yyyy"), "")
            Record.SubItems(15) = IIf(Len(!CLMTUNo), !CLMTUNo, "")
            Record.SubItems(16) = IIf(Len(!CLMTUDate), Format(!CLMTUDate, "dd/mm/yyyy"), "")
            Record.SubItems(17) = IIf(Trim(!CLMAccMonth) <> "", !CLMAccMonth, "")
            Record.SubItems(18) = IIf(Len(!CLMAccYear), !CLMAccYear, "")
            Record.SubItems(19) = IIf(Len(!CLMRemarks), Format(!CLMRemarks, "dd/mm/yyyy"), "")
            Record.SubItems(20) = IIf(Len(!PAYCode), !PAYCode, "")
            Record.SubItems(21) = IIf(IsNumeric(!CLMNoodBordx), !CLMNoodBordx, 0)
            Record.SubItems(22) = IIf(Len(!CasNumber), !CasNumber, "")
            Record.SubItems(23) = IIf(IsNumeric(!claimversion), !claimversion, 0)
            Record.SubItems(24) = IIf(Len(!CLMPayorAppCode), !CLMPayorAppCode, "")
            Record.SubItems(25) = IIf(Len(!CLMPayorAppDate), !CLMPayorAppDate, "")
            Record.SubItems(26) = IIf(Len(!InvPayTo), !InvPayTo, "")
            
            
            
            lbltotalAmt = CDbl(lbltotalAmt) + Format(CDbl(!BordxClaimAmt), "#,##0.00")
            Current = Current + 1
            lblcurrent = Current
            'CurrBordxNo = !BordxRefNo
            'If PrevBordxNo <> CurrBordxNo Then
            '    PrevBordxNo = !BordxRefNo
            '    lbltotalAmt = CDbl(lbltotalAmt) + CDbl(Format(!BordxClaimAmt, "#,##0.00"))
            'End If
            
            'lblTotal = CDbl(lblTotal) + CDbl(Format(IIf(IsNumeric(!CheckAmt) = True, !CheckAmt, 0), "#,##0.00"))
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
    rst2.Close
    Set rst2 = Nothing
    
    If RecFound = False Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        lblcurrent = 0
    End If
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub ProcessCheckItems()
Dim LstCnt As Integer
    LstCnt = lstCLMFTListing.ListItems.Count

        'If lstCnt = 0 Then
        '    chkCheckItems.value = 0
        'Else

            For LstCnt = 1 To LstCnt
                If lstCLMFTListing.ListItems(LstCnt).Text <> "Transfered" Then
                    If lstCLMFTListing.ListItems(LstCnt).Checked = False Then
                        lstCLMFTListing.ListItems(LstCnt).Checked = True
                        lstCLMFTListing.ListItems(LstCnt).Text = "Transfer"
                    ElseIf lstCLMFTListing.ListItems(LstCnt).Checked = True Then
                        lstCLMFTListing.ListItems(LstCnt).Checked = False
                        lstCLMFTListing.ListItems(LstCnt).Text = ""
                    End If
                End If
            Next LstCnt
        'End If

End Sub

Private Sub CmdStop_Click()
    intExit = 1
End Sub

Private Sub cmdUpdate_Click()
Dim AA As Integer
Dim ItmChk As Boolean
Dim BB As Integer
Dim tmpSql As String
Dim tmpBordxNo As String
Dim startTrans As Boolean
Dim bSucess  As Boolean
    
        AA = lstCLMFTListing.ListItems.Count
        tmpPayCode = lstCLMFTListing.SelectedItem.SubItems(20)
        
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            Call SearchFloatAmt("And FTPayCode = '" & Trim(tmpPayCode) & "'")
            'Call SearchFloatAmt("FTPayCode = '" & Trim(tmpPAYCode) & "'")
            tmpFloatBal = CDbl(tmpChkFloatAmt) + CDbl(tmpPrevFloatAmt) - CDbl(txtFTRecAmt)
      '  medixmasterconnWS.Close
      '  Set medixmasterconnWS = Nothing
        
        
    If AA = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    'ElseIf ItmChk = False Then
    '    MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf txtFTLDate = "__/__/____" Then
        MsgBox "Please Enter FT Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtFTLDate.SetFocus
    'ElseIf Trim(txtChequeNo) = "" Then
    '   MsgBox "Please Enter Cheque No!", vbCritical
    '   txtChequeNo.SetFocus
    ElseIf txtFTRecAmt = "" Then
        MsgBox "Please Transfer Amount!", vbCritical
        txtFTRecAmt.SetFocus
    'ElseIf Trim(txtRECDate) = "__/__/____" Then
    '    MsgBox "Please Enter Receipt Date ! ", vbOKOnly + vbCritical, DialogBoxTitle
    '    TxtReceiptNo.SetFocus
    ElseIf Trim(txtPVDate) = "__/__/____" Then
        MsgBox "Please Enter PV Date ! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtPVDate.SetFocus
    Else
        Screen.MousePointer = vbHourglass
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        startTrans = True
        medixmasterconnWS.beginTrans

            Call UpdateFTFloat
            'For BB = 1 To lstCLMFTListing.listitems.Count
                'If lstCLMFTListing.listitems(BB).Checked = True Then
                    tmpBordxNo = lstCLMFTListing.SelectedItem.SubItems(1)
                    tmpPayCode = lstCLMFTListing.SelectedItem.SubItems(20)
                    Call UpdateFTBordx(tmpBordxNo)
                'End If
            'Next BB
            Call UpdateFloatAmt
            medixmasterconnWS.commitTrans
            startTrans = False

            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            If tmpFloatBal < 300000 Then
                MsgBox "Float Balance is " & tmpFloatBal & " for " & tmpPayCode, vbOKOnly + vbInformation, DialogBoxTitle
            End If
            tmpSql = " and  CLMFTNo = '" & Trim(txtFTLNo) & "'"
            lstCLMFTListing.ListItems.Clear
            Call CLMFloatListing(tmpSql)
        EditFlag = False
        CmdEdit.Visible = True
       ' medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
        Screen.MousePointer = vbDefault
        CmdEdit.Visible = True
        lblSelectedAmt = 0
        lbltotalAmt = 0
    End If
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If bSucess = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnPayment.RollbackTrans
    End If
    
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    MediConnCISDB.Close
    Set MediConnCISDB = Nothing
    MsgBox Err.Description
    
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

Private Sub Form_Load()
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        EditFlag = False
        Call ComboListing
        txtFTRecAmt = 0
        lblSelectedAmt = 0
        lbltotalAmt = 0
 '   medixmasterconn.Close
 '   Set medixmasterconn = Nothing
End Sub

Private Sub lstCLMFTListing_Click()
    If lstCLMFTListing.SelectedItem.Text = "Transfered" Then
        tmpPrevFloatAmt = lstCLMFTListing.SelectedItem.SubItems(8)
        Call LvDisplayData
    End If
End Sub

Private Sub lstCLMFTListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
    Case 14
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(15)
    Case 16
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
    Case 18
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
    'End Select

    Case 22
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(22)
    Case 23
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(23)
    Case 24
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(24)
    Case 25
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(25)
    Case 26
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(26)
    Case 27
        SortListView lstCLMFTListing, ColumnHeader.Index, ldtString, m_blnDirection(21)
    End Select
End Sub

Private Sub lstCLMFTListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Text <> "Transfered" Then
        If Item.Checked = True Then
            Item.Text = "Transfer"
            txtFTRecAmt = CDbl(txtFTRecAmt) + CDbl(Item.SubItems(3))
            lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(Item.SubItems(3))
            tmpchkItems = tmpchkItems + 1
        ElseIf Item.Checked = False Then
            Item.Text = ""
            If txtFTRecAmt > 0 Then
                txtFTRecAmt = CDbl(txtFTRecAmt) - CDbl(Item.SubItems(3))
                lblSelectedAmt = CDbl(lblSelectedAmt) - CDbl(Item.SubItems(3))
                tmpchkItems = tmpchkItems - 1
            End If
        End If
    Else
        Item.Checked = False
        MsgBox "This Item has been transfered", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset

    PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

    With CboFTStatus
        .AddItem "0 - Not Transfered"
        .AddItem "1 - Transfered"
        .Text = "0 - Not Transfered"
    End With
    
    With cboCashType
        .AddItem "H - Hospital"
        .AddItem "M - Member"
    End With
    
End Sub

Private Sub SaveFTFloat()
Dim strsqlFT As String
Dim tmpFTLDate As String
Dim tmpPVDate As String
Dim tmpRecDate As String
Dim tmpDate As String
Dim tmpChqDate As String

    tmpFTLDate = Format(txtFTLDate, "yyyy/mm/dd")
    tmpPVDate = Format(txtPVDate, "yyyy/mm/dd")
    tmpRecDate = Format(txtRECDate, "yyyy/mm/dd")
    tmpDate = Format(Date, "yyyy/mm/dd")
    tmpChqDate = Format(txtchqDate, "yyyy/mm/dd")
    
    strsqlFT = ""
    'strsqlFT = "Insert Into CLMFloatTransfer(CLMFTNo, CLMFTDate," & _
    '            "CLMFTAmt, CLMPVNo, CLMPVDate, " & _
    '            "CLMRecNo, CLMRecDate," & _
    '            "CLMAccMonth, CLMAccYear, CLMRemarks,CLMFTUser,CLMFTRegDate,CLMNoodBordx, CLMCHQNo, CLMCHQDate)" & _
    '            "Values('" & tmpFTRecNo & "','" & tmpFTLDate & "', " & _
    '            "" & txtFTRecAmt & ",'" & tmpFTPVNo & "','" & tmpPVDate & "', " & _
    '            "'" & tmpRecNo & "','" & tmpRecDate & "', " & _
    '            "'" & cboAccPeriodMth & "','" & cboAccPeriodYR & "' , '" & txtRemark & "', " & _
    '            "'" & Logon & "','" & tmpDate & "', " & tmpchkItems & ",'" & txtchqNo & "', '" & tmpChqDate & "')"
    'medixmasterconnWS.Execute strsqlFT

    strsqlFT = "Insert Into CLMFloatTransfer(CLMFTNo, CLMFTDate," & _
                "CLMFTAmt, CLMPVNo, CLMPVDate, " & _
                "CLMAccMonth, CLMAccYear, CLMRemarks,CLMFTUser,CLMFTRegDate,CLMNoodBordx, CLMCHQNo, CLMCHQDate)" & _
                "Values('" & tmpFTRecNo & "','" & tmpFTLDate & "', " & _
                "" & txtFTRecAmt & ",'" & tmpFTPVNo & "','" & tmpPVDate & "', " & _
                "'" & cboAccPeriodMth & "','" & cboAccPeriodYR & "' , '" & TxtRemark & "', " & _
                "'" & Logon & "','" & tmpDate & "', " & tmpchkItems & ",'" & txtchqNo & "', '" & tmpChqDate & "')"
    medixmasterconnWS.Execute strsqlFT

End Sub

Private Sub SaveFTBordx(tmpBordxNo As String)
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
Dim strsqlCLM As String
Dim savedstatus As Boolean
    strsqlCLM = ""
    strsqlCLM = " Select  BordxRefNo,BordxDate, FLTFTNo,FLTFTDate, " & _
                "FLTPVNo, FLTPVDate,  FLTRecNo, FLTRecDate from Bordx " & _
                " Where BordxRefNo = '" & Trim(tmpBordxNo) & "'"
    rst.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                    
        With rst
            !FLTFTNo = Trim(tmpFTRecNo)
            !FLTFTDate = Format(txtFTLDate, "dd/mm/yyyy")
            !FLTPVNo = tmpFTPVNo
            !FLTPVDate = Format(txtPVDate, "dd/mm/yyyy")
            !FLTRecNo = tmpFTRecNo
            !FLTRecDate = Format(txtPVDate, "dd/mm/yyyy")
        End With
        rst.Update
        
        rst.Close
        Set rst = Nothing
End Sub

Private Sub UpdateFTFloat()
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
Dim savedstatus As Boolean
    strsqlFT = ""
    'rst.Close
    'Set rst = Nothing
    strsqlFT = " Select  CLMFTNo,CLMFTDate, CLMFTAmt, " & _
                "CLMPVNo, CLMPVDate, CLMCHQNo," & _
                "CLMCHQDate, CLMRecNo, CLMRecDate, " & _
                "CLMTUNo, CLMTUDate,CLMAccMonth, CLMAccYear, " & _
                "CLMRemarks,CLMFTUser,CLMFTRegDate from CLMFloatTransfer " & _
                " Where CLMFTNo = '" & Trim(txtFTLNo) & "'"
    
    rst.Open strsqlFT, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        If rst.recordCount = 0 Then
            rst.AddNew
            savedstatus = False
        ElseIf rst.recordCount > 0 Then
            savedstatus = True
        End If
            With rst
                !CLMFTNo = Trim(txtFTLNo)
                !CLMFTDate = Format(txtFTLDate, "dd/mm/yyyy")
                !CLMFTAmt = txtFTRecAmt
                !CLMPVNo = txtPVNo
                !CLMPVDate = Format(txtPVDate, "dd/mm/yyyy")
                '!CLMRecNo = TxtReceiptNo
                '!CLMRecDate = Format(txtRECDate, "dd/mm/yyyy")
                !CLMAccMonth = IIf(Len(Trim(cboAccPeriodMth)), Trim(cboAccPeriodMth), "")
                !CLMAccYear = IIf(Len(Trim(cboAccPeriodYR)), Trim(cboAccPeriodYR), "")
                !CLMRemarks = IIf(Len(TxtRemark), TxtRemark, "")
                !CLMFTUser = Logon
                !CLMFTRegDate = Date
                !CLMCHQNo = txtchqNo
                !CLMCHQDate = Format(txtchqDate, "dd/mm/yyyy")
            End With
            rst.Update
        
        rst.Close
        Set rst = Nothing
End Sub

Private Sub UpdateFTBordx(tmpBordxNo As String)
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
Dim strsqlCLM As String
Dim savedstatus As Boolean
    strsqlCLM = ""
    strsqlCLM = " Select  BordxRefNo,BordxDate, FLTFTNo,FLTFTDate, " & _
                "FLTPVNo, FLTPVDate,  FLTRecNo, FLTRecDate from Bordx " & _
                " Where BordxRefNo = '" & Trim(tmpBordxNo) & "'"
    rst.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                    
        With rst
            !FLTFTNo = Trim(txtFTLNo)
            !FLTFTDate = Format(txtFTLDate, "dd/mm/yyyy")
            !FLTPVNo = txtPVNo
            !FLTPVDate = Format(txtPVDate, "dd/mm/yyyy")
            '!FLTRecNo = TxtReceiptNo
            '!FLTRecDate = Format(txtRECDate, "dd/mm/yyyy")
        End With
        rst.Update
        
        rst.Close
        Set rst = Nothing
End Sub

Private Sub UpdateFloatAmt()
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
Dim tmpFTFloatAmt As Double
    'strsqlFT = " Select FTPayCode, FTFloatAmt from CLMFloatAmt Where FTPayCode = '" & Trim(tmpPayCode) & "'"
    'rst.Open strsqlFT, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
     '   tmpFTFloatAmt = rst!FTFloatAmt
     '   If EditFlag = False Then
     '       tmpFTFloatBLAmt = CDbl(tmpFTFloatAmt) - CDbl(txtFTRecAmt)
     '   Else
     '       tmpFTFloatBLAmt = (CDbl(tmpFTFloatAmt) + CDbl(tmpPrevFloatAmt)) - CDbl(txtFTRecAmt)
     '   End If
        
      '  rst!FTFloatAmt = tmpFTFloatBLAmt
      '  'tmpFTFloatBLAmt = tmpFTFloatBLAmt
      '  rst.Update
        
       ' If tmpPayCode = "XG" Then
       '     lblMGAmt = tmpFTFloatBLAmt
       ' Else
       '     lblMLAmt = tmpFTFloatBLAmt
       ' End If
    'rst.Close
    'Set rst = Nothing

End Sub

Private Sub LvDisplayData(Optional Item As ListItem, Optional ItemCheck As Boolean)
Dim lstRec As ListItem

    Call ClearFrameCase
    If lstCLMFTListing.ListItems.Count Then
        If ItemCheck = True Then
            Set lstRec = Item
            lstCLMFTListing.SelectedItem.Selected = False
            lstRec.Selected = True
        Else
            Set lstRec = lstCLMFTListing.SelectedItem
        End If
            txtFTLNo = IIf(Len(Trim(lstRec.SubItems(6))), lstRec.SubItems(6), "")
            If Len(lstRec.SubItems(7)) Then
                txtFTLDate = Format(lstRec.SubItems(7), "DD/MM/YYYY")
            End If
            If Len(lstRec.SubItems(8)) Then
                txtFTRecAmt = lstRec.SubItems(8)
            End If
            If Len(lstRec.SubItems(9)) Then
                txtPVNo = lstRec.SubItems(9)
            End If
            If IsDate(lstRec.SubItems(10)) Then
                txtPVDate = Format(lstRec.SubItems(10), "DD/MM/YYYY")
            End If
            
            txtchqNo = IIf(Len(lstRec.SubItems(11)), lstRec.SubItems(11), "")
            
            txtchqDate = IIf(Len(lstRec.SubItems(12)), Format(lstRec.SubItems(12), "DD/MM/YYYY"), "__/__/____")
            
            If Len(lstRec.SubItems(13)) Then
                TxtReceiptNo = lstRec.SubItems(13)
            End If
            If IsDate(lstRec.SubItems(14)) Then
                txtRECDate = Format(lstRec.SubItems(14), "DD/MM/YYYY")
            End If
            If Len(lstRec.SubItems(17)) Then
                cboAccPeriodMth = lstRec.SubItems(17)
                cboAccPeriodYR = lstRec.SubItems(18)
            End If
            
            TxtRemark = IIf(Len(lstRec.SubItems(19)), lstRec.SubItems(19), "")
            CmdSave.Enabled = False
    End If
End Sub

Private Sub ClearFrameCase()
    txtFTLNo = ""
    txtFTLDate = "__/__/____"
    txtFTRecAmt = 0
    TxtReceiptNo = ""
    txtRECDate.Text = "__/__/____"
    txtPVNo = ""
    txtPVDate.Text = "__/__/____"
    cboAccPeriodMth.Text = ""
    cboAccPeriodYR.Text = ""
    TxtRemark.Text = ""
    tmpchkItems = 0
End Sub

Private Sub lstCLMFTListing_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstCLMFTListing.SelectedItem.Text = "Transfered" Then
        tmpPrevFloatAmt = lstCLMFTListing.SelectedItem.SubItems(8)
        Call LvDisplayData
    End If
End Sub

Private Sub lstCLMFTListing_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstCLMFTListing.SelectedItem.Text = "Transfered" Then
        tmpPrevFloatAmt = lstCLMFTListing.SelectedItem.SubItems(8)
        Call LvDisplayData
    End If

End Sub

Private Sub SearchFloatAmt(strsqlSearch As String)
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
    
    strsqlFT = " Select FTPayCode, FTFloatAmt from CLMFloatAmt where 0=0"
    strsqlFT = strsqlFT + strsqlSearch
    rst.Open strsqlFT, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        
        If rst.recordCount = 1 Then
            tmpPayCode = rst!FTPayCode
            tmpChkFloatAmt = rst!FTFloatAmt
        Else
            Do While Not rst.EOF
                If rst!FTPayCode = "XG" Then
                    lblMGAmt = Format(rst!FTFloatAmt, "#,##0.00")
                Else
                    lblMLAmt = Format(rst!FTFloatAmt, "#,##0.00")
                End If
            rst.MoveNext
            Loop
        End If
    rst.Close
    Set rst = Nothing
End Sub


Private Sub txtBordxDateFr_LostFocus()
    If IsDate(txtBordxDateFr) Then
    Else
        If txtBordxDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDateFr.SetFocus
        End If
    End If

End Sub

Private Sub txtBordxDateTo_LostFocus()
    If IsDate(txtBordxDateTo) Then
    Else
        If txtBordxDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDateTo.SetFocus
        End If
    End If

End Sub


Private Sub txtChqDate_LostFocus()
    If IsDate(txtchqDate) Then
    Else
        If txtchqDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
             txtchqDate.SetFocus
        End If
    End If
End Sub

Private Sub txtFTDateFR_LostFocus()
    If IsDate(txtFTDateFR) Then
    Else
        If txtFTDateFR <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtFTDateFR.SetFocus
        End If
    End If

End Sub



Private Sub txtFTDateTo_LostFocus()
    If IsDate(txtFTDateTo) Then
    Else
        If txtFTDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtFTDateTo.SetFocus
        End If
    End If
End Sub

Private Sub txtFTLDate_LostFocus()
    If IsDate(txtFTLDate) Then
    Else
        If txtFTLDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtFTLDate.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDate_LostFocus()
    If IsDate(txtPVDate) Then
    Else
        If txtPVDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
             txtPVDate.SetFocus
        End If
    End If

End Sub

Private Sub txtPVDateFr_LostFocus()
    If IsDate(txtPVDateFr) Then
    Else
        If txtPVDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDateFr.SetFocus
        End If
    End If

End Sub

Private Sub txtPVDateTo_LostFocus()
    If IsDate(txtPVDateTo) Then
    Else
        If txtPVDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDateTo.SetFocus
        End If
    End If
End Sub

Private Sub txtRECDate_LostFocus()
    If IsDate(txtRECDate) Then
    Else
        If txtRECDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
             txtRECDate.SetFocus
        End If
    End If
End Sub

Private Sub txtRecDateFr_LostFocus()
    If IsDate(txtRECDateFr) Then
    Else
        If txtRECDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRECDateFr.SetFocus
        End If
    End If

End Sub

Private Sub txtRecDateTo_LostFocus()
    If IsDate(txtRECDateTo) Then
    Else
        If txtRECDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRECDateTo.SetFocus
        End If
    End If
End Sub


Private Sub txtTUDateFR_LostFocus()
    If IsDate(txtTUDateFR) Then
    Else
        If txtTUDateFR <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtTUDateFR.SetFocus
        End If
    End If

End Sub

Private Sub txtTUDateTo_LostFocus()
    If IsDate(txtTUDateTo) Then
    Else
        If txtTUDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtTUDateTo.SetFocus
        End If
    End If
End Sub

Private Sub SaveReceiptPayor()
Dim strsqlREC As String
    'strsqlREC = "Insert Into ReceiptfrPay(CasNumber, ClaimVersion, ReceiptNo," & _
    '     "ReceiptDate, ReceiptAmt, ChequeNo, BordxNo, " & _
    '     "PayLastUpdateDate, PayLastUpdateUser," & _
    '     "RECPayAmt, CheckAmt) " & _
    '     "Values('" & Trim(tmpCASNumber) & "','" & Trim(tmpVersion) & "','" & Trim(tmpRecNo) & _
    '     "', '" & Format(txtRECDate, "yyyy/mm/dd") & "'," & tmpRecAmt & ",'" & Trim(txtchqNo) & "','" & tmpBordxNo & _
    '     "', '" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'," & tmpRecAmt & "," & tmpcheckAmt & ")"
    'medixmasterconnPayment.Execute strsqlREC
    
    strsqlREC = "Insert Into ReceiptfrPay(CasNumber, ClaimVersion," & _
         " ReceiptAmt, ChequeNo, BordxNo, " & _
         "PayLastUpdateDate, PayLastUpdateUser," & _
         "RECPayAmt, CheckAmt) " & _
         "Values('" & Trim(tmpCASNumber) & "','" & Trim(tmpVersion) & _
         "', " & tmpRecAmt & ",'" & Trim(txtchqNo) & "','" & tmpBordxNo & _
         "', '" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'," & tmpRecAmt & "," & tmpcheckAmt & ")"
    medixmasterconnPayment.Execute strsqlREC

End Sub
