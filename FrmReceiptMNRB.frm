VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReceiptMNRB 
   Caption         =   "Receipt From MNRB"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameCas 
      Enabled         =   0   'False
      Height          =   1620
      Left            =   5880
      TabIndex        =   85
      Top             =   5280
      Width           =   5895
      Begin VB.TextBox TxtAmtPaid 
         Height          =   285
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   31
         Top             =   160
         Width           =   900
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Update &Case"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4680
         TabIndex        =   42
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TxtClaimNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   30
         Top             =   160
         Width           =   1095
      End
      Begin VB.TextBox TxtContraVer 
         Height          =   285
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   39
         Top             =   915
         Width           =   255
      End
      Begin VB.TextBox TxtBankCharge 
         Height          =   285
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   32
         Top             =   160
         Width           =   735
      End
      Begin VB.TextBox TxtMisc 
         Height          =   285
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   35
         Top             =   400
         Width           =   735
      End
      Begin VB.TextBox TxtCasRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2640
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   915
         Width           =   1920
      End
      Begin VB.TextBox TxtMedixRec 
         Height          =   285
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   37
         Top             =   640
         Width           =   735
      End
      Begin VB.TextBox TxtGainLoss 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   33
         Top             =   400
         Width           =   1095
      End
      Begin VB.TextBox TxtDiscount 
         Height          =   285
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   34
         Top             =   420
         Width           =   900
      End
      Begin VB.TextBox TxtWriteOff 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   36
         Top             =   640
         Width           =   1095
      End
      Begin VB.TextBox TxtContraCase 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   38
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox TxtContraAmt 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   40
         Top             =   1155
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Amt Paid"
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
         Left            =   2400
         TabIndex        =   101
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Medix Rec."
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
         Left            =   4200
         TabIndex        =   97
         Top             =   680
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "Remark"
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
         Left            =   2640
         TabIndex        =   95
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label37 
         Caption         =   "Write Off"
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
         TabIndex        =   94
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label36 
         Caption         =   "Discount"
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
         Left            =   2400
         TabIndex        =   93
         Top             =   440
         Width           =   735
      End
      Begin VB.Label Label35 
         Caption         =   "Misc. Diff."
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
         Left            =   4200
         TabIndex        =   92
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label34 
         Caption         =   "Gain/Loss"
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
         TabIndex        =   91
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "Contra Amt"
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
         TabIndex        =   90
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "Bank Chrg"
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
         Left            =   4200
         TabIndex        =   89
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label41 
         Caption         =   "Claim No"
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
         TabIndex        =   88
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label42 
         Caption         =   "Contra Case"
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
         TabIndex        =   87
         Top             =   915
         Width           =   975
      End
      Begin VB.Label Label43 
         Caption         =   "--"
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
         Left            =   2160
         TabIndex        =   86
         Top             =   960
         Width           =   135
      End
   End
   Begin VB.Frame FrameReceiptPV 
      Enabled         =   0   'False
      Height          =   1605
      Left            =   120
      TabIndex        =   74
      Top             =   5280
      Width           =   5750
      Begin VB.TextBox txtStatus 
         Height          =   285
         Left            =   2400
         TabIndex        =   105
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtIndex 
         Height          =   285
         Left            =   5160
         TabIndex        =   104
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBordxNO 
         Height          =   285
         Left            =   4680
         TabIndex        =   103
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtReceiptNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1120
         MaxLength       =   15
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   285
         Left            =   3120
         TabIndex        =   26
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtReceiptAmt 
         Height          =   285
         Left            =   1120
         MaxLength       =   10
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtChequeNo 
         Height          =   285
         Left            =   1120
         MaxLength       =   20
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cboAccPeriodMth 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":0000
         Left            =   3120
         List            =   "FrmReceiptMNRB.frx":0028
         TabIndex        =   27
         Top             =   500
         Width           =   1095
      End
      Begin VB.ComboBox cboAccPeriodYR 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":0053
         Left            =   4680
         List            =   "FrmReceiptMNRB.frx":00DB
         TabIndex        =   28
         Top             =   500
         Width           =   975
      End
      Begin VB.TextBox txtReceiptRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3120
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   780
         Width           =   2535
      End
      Begin VB.TextBox txtPVNo 
         Height          =   285
         Left            =   1120
         MaxLength       =   20
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   285
         Left            =   1120
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label38 
         Caption         =   "Pay Chq Date"
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
         TabIndex        =   102
         Top             =   1245
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   " Year"
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
         Left            =   4240
         TabIndex        =   81
         Top             =   560
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "A/C Mth"
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
         Left            =   2400
         TabIndex        =   80
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "RecDate"
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
         Left            =   2400
         TabIndex        =   79
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Rec No"
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
         TabIndex        =   78
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
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
         Left            =   2400
         TabIndex        =   77
         Top             =   810
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Rec Amt"
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
         TabIndex        =   76
         Top             =   520
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Pay Chq No"
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
         TabIndex        =   75
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Pay PV No."
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
         TabIndex        =   100
         Top             =   1020
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   120
      TabIndex        =   58
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
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
         Left            =   5400
         TabIndex        =   6
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtWSNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtWSNoFr 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   2
         Top             =   580
         Width           =   1575
      End
      Begin VB.TextBox txtRECNoTo 
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   10
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtRECNoFr 
         Height          =   285
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   9
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   4
         Top             =   820
         Width           =   1575
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   5
         Top             =   820
         Width           =   1575
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   3615
      End
      Begin VB.OptionButton OptWks 
         Caption         =   "By Worksheet"
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
         Left            =   3000
         TabIndex        =   1
         Top             =   200
         Width           =   1815
      End
      Begin VB.OptionButton OptBordx 
         Caption         =   "By Bordx"
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
         Left            =   1800
         TabIndex        =   0
         Top             =   200
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtRECDateFr 
         Height          =   315
         Left            =   8160
         TabIndex        =   11
         Top             =   390
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRECDateTo 
         Height          =   315
         Left            =   9840
         TabIndex        =   12
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboBordxStatus 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   3615
      End
      Begin VB.ComboBox CboWSStatus 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   1320
         Width           =   3615
      End
      Begin MSMask.MaskEdBox txtWSDateFr 
         Height          =   315
         Left            =   8160
         TabIndex        =   13
         Top             =   630
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtWSDateTo 
         Height          =   315
         Left            =   9840
         TabIndex        =   14
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   8160
         TabIndex        =   15
         Top             =   870
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   315
         Left            =   9840
         TabIndex        =   16
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":01E7
         Left            =   8160
         List            =   "FrmReceiptMNRB.frx":020F
         TabIndex        =   17
         Top             =   1155
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":023A
         Left            =   9840
         List            =   "FrmReceiptMNRB.frx":0262
         TabIndex        =   18
         Top             =   1125
         Width           =   1575
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":028D
         Left            =   9840
         List            =   "FrmReceiptMNRB.frx":0315
         TabIndex        =   20
         Top             =   1395
         Width           =   1575
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmReceiptMNRB.frx":0421
         Left            =   8160
         List            =   "FrmReceiptMNRB.frx":04A9
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label44 
         Caption         =   "Bordx Date from"
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
         Left            =   6600
         TabIndex        =   109
         Top             =   920
         Width           =   1455
      End
      Begin VB.Label Label39 
         Caption         =   "To"
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
         Left            =   9480
         TabIndex        =   108
         Top             =   920
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "WS Date from"
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
         Left            =   6600
         TabIndex        =   107
         Top             =   680
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "To"
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
         Left            =   9480
         TabIndex        =   106
         Top             =   670
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "A/C Period Yr Fr"
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
         Left            =   6600
         TabIndex        =   73
         Top             =   1485
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "To"
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
         Left            =   9480
         TabIndex        =   72
         Top             =   1485
         Width           =   375
      End
      Begin VB.Label lblWSTo 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   71
         Top             =   585
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "WS No From"
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
         Left            =   240
         TabIndex        =   70
         Top             =   580
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Left            =   9480
         TabIndex        =   69
         Top             =   435
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Recp Date from"
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
         Left            =   6600
         TabIndex        =   68
         Top             =   435
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Recp No from"
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
         Left            =   6600
         TabIndex        =   67
         Top             =   160
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "To"
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
         Left            =   9480
         TabIndex        =   66
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Payment Status"
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
         Left            =   240
         TabIndex        =   65
         Top             =   1400
         Width           =   1455
      End
      Begin VB.Label lblBordxTo 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   64
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No from"
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
         Left            =   240
         TabIndex        =   63
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
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
         Left            =   240
         TabIndex        =   62
         Top             =   1120
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Receipt Type"
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
         Left            =   240
         TabIndex        =   61
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "To"
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
         Left            =   9480
         TabIndex        =   60
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period Mth Fr"
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
         Left            =   6600
         TabIndex        =   59
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   52
      Top             =   6840
      Width           =   11655
      Begin VB.TextBox lbltotalAmt 
         Height          =   195
         Left            =   2640
         TabIndex        =   113
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox lbltotal 
         Height          =   225
         Left            =   2640
         TabIndex        =   112
         Top             =   450
         Width           =   855
      End
      Begin VB.TextBox lblSelectedAmt 
         Height          =   285
         Left            =   4800
         TabIndex        =   111
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblcurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   110
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPrintPV 
         Caption         =   "Print &OR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         TabIndex        =   49
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "&Print to File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8640
         TabIndex        =   48
         Top             =   220
         Width           =   975
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6600
         TabIndex        =   44
         Top             =   220
         Width           =   720
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   6600
         TabIndex        =   45
         Top             =   220
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox TxtPStatus2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   55
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   54
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
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
         Height          =   330
         Left            =   10320
         TabIndex        =   50
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10920
         TabIndex        =   51
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         TabIndex        =   47
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S&earch"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7320
         TabIndex        =   46
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         TabIndex        =   43
         Top             =   220
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Total Rec. Amt"
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
         Left            =   1440
         TabIndex        =   99
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "Total Bordx Amt"
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
         Left            =   1440
         TabIndex        =   98
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Checked Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   57
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
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
         TabIndex        =   56
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lstBordx 
      Height          =   3120
      Left            =   120
      TabIndex        =   82
      Top             =   2160
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5503
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
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Bordx Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Receipt Amt"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "OS Amt"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Paying"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "OrgStatus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "PVNo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "PVDate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ReceiptNo"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ReceiptDate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "AC Period"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "ReceiptChq"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "PVRemark"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "ChqAmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "BordxDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "RCInd"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstWorksheet 
      Height          =   3120
      Left            =   120
      TabIndex        =   83
      Top             =   2160
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5503
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
      NumItems        =   32
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Worksheet No"
         Object.Width           =   2291
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amt Receipt"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "OS Amt"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Paying"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Receipt No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Receipt Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "OrgStatus "
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Bank Charge"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Contra Case No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Contra Case Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Gain Loss"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Misc Diff"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Discount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "Write Off Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "Medix Receipt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "ReceiptAmt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "ReceiptMth"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "ReceiptYr"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "CasRemark"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "PVNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "PVDate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "PVChqNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "RecIndex"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "TotRec'dAmt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "WS Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "BordxDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "RCInd"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Receipt from MNRB"
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
      TabIndex        =   96
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   84
      Top             =   7770
      Width           =   1695
   End
End
Attribute VB_Name = "FrmReceiptMNRB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intExit As Integer
Dim ScanCode As String
Dim tmpAuthStatus As String
Dim SaveFlag As Boolean
Dim tmpRecNo As String
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

Private m_blnDirection(1 To 31) As Boolean

Private Sub ChkScan_Click()
    If OptBordx.Value = True And ChkScan.Value = 0 Then
        lblBordxTo.Visible = True
        TxtBordx2.Visible = True
        ChkScan.Left = 5400
        ChkScan.Top = 840
        TxtBordx1.SetFocus
    ElseIf OptBordx.Value = True And ChkScan.Value = 1 Then
        lblBordxTo.Visible = False
        TxtBordx2.Visible = False
        ChkScan.Left = 3360
        ChkScan.Top = 840
        TxtBordx1.SetFocus
    ElseIf OptWks.Value = True And ChkScan.Value = 0 Then
        lblWSTo.Visible = True
        txtWSNoTo.Visible = True
        txtWSNoFr.SetFocus
        ChkScan.Left = 5400
        ChkScan.Top = 600
    ElseIf OptWks.Value = True And ChkScan.Value = 1 Then
        lblWSTo.Visible = False
        txtWSNoTo.Visible = False
        ChkScan.Left = 3360
        ChkScan.Top = 600
    End If
End Sub

Private Sub CmdAdd_Click()
Dim AskRec
    
If lstWorksheet.ListItems.Count <> 0 Then
    If lstWorksheet.SelectedItem.Checked = True Then
        If (Trim(TxtContraCase) <> "" And Trim(TxtContraVer) = "") Or (Trim(TxtContraCase) = "" And Trim(TxtContraVer) <> "") Then
            MsgBox "Invalid Contra Case..., Please Keyin Contra Case No. or Contra Version ", vbCritical + vbOKOnly, DialogBoxTitle
            TxtContraCase.SetFocus
        ElseIf (Trim(TxtContraCase) <> "" And Trim(TxtContraVer) <> "" And Trim(TxtContraAmt) = "") Then
            MsgBox "Please Keyin Contra Case Amount ", vbCritical + vbOKOnly, DialogBoxTitle
            TxtContraAmt.SetFocus
        Else
            AskRec = MsgBox("Update " & TxtClaimNo & " ?", vbYesNo + vbExclamation)
            If AskRec = vbYes Then
                If IsNumeric(TxtAmtPaid) Then
                    lblSelectedAmt = lblSelectedAmt - CDbl(Format(lstWorksheet.SelectedItem.SubItems(7), "#,##0.00"))
                    If (CDbl(TxtAmtPaid) + CDbl(lstWorksheet.SelectedItem.SubItems(28))) >= CDbl(lstWorksheet.SelectedItem.SubItems(4)) Then
                        lstWorksheet.SelectedItem.SubItems(1) = "Full"
                    Else
                        lstWorksheet.SelectedItem.SubItems(1) = "Partial"
                    End If
                    lstWorksheet.SelectedItem.SubItems(6) = CDbl(lstWorksheet.SelectedItem.SubItems(4)) - CDbl(lstWorksheet.SelectedItem.SubItems(5)) - CDbl(TxtAmtPaid)
                    lstWorksheet.SelectedItem.SubItems(7) = TxtAmtPaid
                    lblSelectedAmt = lblSelectedAmt + CDbl(Format(lstWorksheet.SelectedItem.SubItems(7), "#,##0.00"))
                End If
                lstWorksheet.SelectedItem.SubItems(11) = TxtBankCharge
                lstWorksheet.SelectedItem.SubItems(12) = Trim(TxtContraCase) & "-" & Trim(TxtContraVer)
                lstWorksheet.SelectedItem.SubItems(13) = TxtContraAmt
                lstWorksheet.SelectedItem.SubItems(14) = TxtGainLoss
                lstWorksheet.SelectedItem.SubItems(15) = TxtMisc
                lstWorksheet.SelectedItem.SubItems(16) = TxtDiscount
                lstWorksheet.SelectedItem.SubItems(17) = TxtWriteOff
                lstWorksheet.SelectedItem.SubItems(18) = TxtMedixRec
                lstWorksheet.SelectedItem.SubItems(19) = TxtCasRemark
                MsgBox "Updated .... ", vbInformation + vbOKOnly, DialogBoxTitle
            End If
            Call ClearRecDetail
            Call ClearContraDetail
        End If
    Else
        MsgBox "Please select a record ! ", vbOKOnly, DialogBoxTitle
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

Private Sub CmdEdit_Click()
    Screen.MousePointer = vbHourglass
    FrameReceiptPV.Enabled = False
    'TxtReceiptNo.Enabled = False
    FrameCas.Enabled = False
    If Trim(TxtReceiptNo) <> "" Then
        CmdUpdate.Visible = True
        CmdUpdate.Enabled = True
        CmdEdit.Visible = False
        FrameReceiptPV.Enabled = True
        If OptWks.Value = True Then
            FrameCas.Enabled = True
        End If
        'txtReceiptDate.SetFocus
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdUpdate_Click()
Dim AA As Integer
Dim tmpCASNo As String
Dim tmpVerNo As String
Dim ContraAmt As Double
Dim GainLoss As Double
Dim MiscDiff As Double
Dim Discount As Double
Dim WriteOff As Double
Dim MedixRec As Double
Dim BankCharge As Double
Dim tmpSql As String
    If OptBordx.Value = True Then
        AA = lstBordx.ListItems.Count
    Else
        AA = lstWorksheet.ListItems.Count
    End If
    If AA = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    ElseIf Trim(txtChequeNo) = "" Then
       MsgBox "Please Enter Cheque No!", vbCritical
       txtChequeNo.SetFocus
    ElseIf Trim(TxtReceiptNo) = "" Then
        MsgBox "Please Enter Receipt No ! ", vbOKOnly + vbCritical, DialogBoxTitle
        'TxtReceiptNo.SetFocus
    Else
        If Trim(txtPVNo) = "" Then
            MsgBox "Please Enter PV No ! ", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf txtPVDate = "__/__/____" Then
            MsgBox "Please Enter PV Date", vbOKOnly + vbCritical, DialogBoxTitle
        End If
        Screen.MousePointer = vbHourglass
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
      '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        SaveFlag = False
        If OptBordx.Value = True Then
            Call SaveRecByBordx(txtBordxNO)
            MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            tmpSql = " and ReceiptNo = '" & Trim(TxtReceiptNo) & "'"
            Call ClearRecDetail
            Call ClearContraDetail
            lstBordx.ListItems.Clear
            Call ReceiptBordxListing(tmpSql)
        Else
            AA = InStr(1, TxtClaimNo, "-")
            tmpCASNo = Mid(TxtClaimNo, 1, AA - 1)
            tmpVerNo = Mid(TxtClaimNo, AA + 1, Len(TxtClaimNo) - AA)
            ContraAmt = IIf(IsNumeric(TxtContraAmt), TxtContraAmt, 0)
            GainLoss = IIf(IsNumeric(TxtGainLoss), TxtGainLoss, 0)
            MiscDiff = IIf(IsNumeric(TxtMisc), TxtMisc, 0)
            Discount = IIf(IsNumeric(TxtDiscount), TxtDiscount, 0)
            WriteOff = IIf(IsNumeric(TxtWriteOff), TxtWriteOff, 0)
            MedixRec = IIf(IsNumeric(TxtMedixRec), TxtMedixRec, 0)
            BankCharge = IIf(IsNumeric(TxtBankCharge), TxtBankCharge, 0)
            If (CDbl(TxtAmtPaid) + CDbl(lstWorksheet.SelectedItem.SubItems(28)) - CDbl(lstWorksheet.SelectedItem.SubItems(5))) >= CDbl(lstWorksheet.SelectedItem.SubItems(4)) Then
                lstWorksheet.SelectedItem.SubItems(1) = "Full"
                txtStatus = "F"
            Else
                lstWorksheet.SelectedItem.SubItems(1) = "Partial"
                txtStatus = "P"
            End If

            Call SaveRecByWS(tmpCASNo, CInt(tmpVerNo), txtBordxNO, Trim(txtStatus), CDbl(TxtAmtPaid), BankCharge, GainLoss, MiscDiff, Discount, WriteOff, MedixRec, TxtCasRemark, txtIndex)
            MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            tmpSql = " and ReceiptNo = '" & Trim(TxtReceiptNo) & "'"
            Call ClearRecDetail
            Call ClearContraDetail
            lstWorksheet.ListItems.Clear
            Call ReceiptWorksheetListing(tmpSql)
        End If
    '    medixmasterconnWS.Close
    '    Set medixmasterconnWS = Nothing
    '    medixmasterconnPayment.Close
    '    Set medixmasterconnPayment = Nothing
        Screen.MousePointer = vbDefault
        CmdUpdate.Visible = False
        CmdEdit.Visible = True
    End If
End Sub

Private Sub Form_Load()
Dim USR As New ADODB.Recordset
Dim strsqlUSR As String
    SaveFlag = False
    tmpAuthStatus = ""
    strsqlUSR = "Select USRCode, AuthorisedStatus From USR where USRCode = '" & Trim(Logon) & "'"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    USR.Open strsqlUSR, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    Do While Not USR.EOF
        tmpAuthStatus = IIf(Len(USR!AuthorisedStatus), USR!AuthorisedStatus, "")
        USR.MoveNext
    Loop
    USR.Close
    Set USR = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
        Call PayorComboListing
   ' medixmasterconn.Close
 '   Set medixmasterconn = Nothing
'
    lstWorksheet.Visible = False
    txtWSNoFr.Enabled = False
    txtWSNoTo.Enabled = False
    FrameCas.Enabled = False
    '==============================
    lblBordxTo.Visible = False
    TxtBordx2.Visible = False
    ChkScan.Left = 3360
    ChkScan.Top = 840
    '==============================
    CboBordxStatus.ListIndex = 0
    CboWSStatus.Visible = False
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    lblSelectedAmt = 0
    LblTotalAmt = 0
End Sub

'************** Start Command Button ****************************

Private Sub CmdSearch_Click()
Dim AA As Integer
Dim tmpFound As Boolean
Dim tmpTxtBox As TextBox

    If Mid(TxtBordx1, 1, 2) = "LO" Or Mid(txtWSNoFr, 1, 2) = "LO" Then
        If Mid(TxtBordx1, 1, 2) = "LO" Then
            ScanCode = TxtBordx1
            TxtBordx1 = ""
        ElseIf Mid(txtWSNoFr, 1, 2) = "LO" Then
            ScanCode = txtWSNoFr
            txtWSNoFr = ""
        End If
        Call SearchForm(ScanCode)
    Else
        Call EnabledCommand(False)
        Screen.MousePointer = vbHourglass
        If ChkScan.Value = 1 Then
            AA = lstBordx.ListItems.Count
            If OptWks.Value = True Then
                AA = lstWorksheet.ListItems.Count
            End If
            If OptWks.Value = True And Trim(txtWSNoFr) = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                txtWSNoFr.SetFocus
                Call EnabledCommand(True)
            ElseIf OptBordx.Value = True And Trim(TxtBordx1) = "" Then
                MsgBox "Please Enter Bordx No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtBordx1.SetFocus
                Call EnabledCommand(True)
            Else
                Screen.MousePointer = vbHourglass
                If AA <> 0 Then
                    For AA = 1 To AA
                        If OptWks.Value = True Then
                            tmpFound = CaseNoCheck(AA)
                        Else
                            tmpFound = BordxNoCheck(AA)
                        End If
                        If tmpFound = True Then Exit For
                    Next AA
                End If
                If tmpFound = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    Call EnabledCommand(True)
                Else
                    Call SearchRecord
                End If
                Screen.MousePointer = Default
            End If
            If OptWks.Value = True Then
                txtWSNoFr = ""
                txtWSNoFr.SetFocus
            Else
                TxtBordx1 = ""
                TxtBordx1.SetFocus
            End If
        Else
            Screen.MousePointer = vbHourglass
                Call SearchRecord
            Screen.MousePointer = Default
        End If
        Screen.MousePointer = Default
    End If
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
    cmdPrintPV.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub CmdClear_Click()
    Call ClearForm
    lstBordx.ListItems.Clear
    lstWorksheet.ListItems.Clear
    CboPayor.Clear
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call PayorComboListing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    Call ClearRecDetail
    '================================
    Call ClearContraDetail
    '================================
    CboBordxStatus.ListIndex = 0
    CboWSStatus.ListIndex = 0
    lblCurrent = 0
    lblTotal = 0
    txtReceiptAmt = ""
    lblSelectedAmt = 0
    LblTotalAmt = 0
    CmdSave.Enabled = False
    CmdUpdate.Visible = False
    CmdEdit.Visible = True
    CmdEdit.Enabled = False
    End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim AA As Integer
Dim ItmChk As Boolean
Dim BB As Integer
Dim tmpSql As String
Dim tmpBordxNo As String
    ItmChk = False
    ' Update Claim Table
    RemLen1 = 0
    RemLen2 = 0
    RemLen3 = 0
    RemLen4 = 0
    Remarks1 = ""
    Remarks2 = ""
    Remarks3 = ""
    Remarks4 = ""
    
    If OptBordx.Value = True Then
        AA = lstBordx.ListItems.Count
        For BB = 1 To lstBordx.ListItems.Count
            If lstBordx.ListItems(BB).Checked = True Then
                ItmChk = True
                Exit For
            End If
        Next BB
    Else
        AA = lstWorksheet.ListItems.Count
        For BB = 1 To lstWorksheet.ListItems.Count
            If lstWorksheet.ListItems(BB).Checked = True Then
                ItmChk = True
                Exit For
            End If
        Next BB
    End If
    
    If AA = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf ItmChk = False Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    ElseIf Trim(txtChequeNo) = "" Then
       MsgBox "Please Enter Cheque No!", vbCritical
       txtChequeNo.SetFocus
    ElseIf CDbl(txtReceiptAmt) <> CDbl(lblSelectedAmt) And Trim(txtReceiptRemark) = "" Then
        MsgBox "Please Enter Remark!", vbCritical
        txtReceiptRemark.SetFocus
    'ElseIf Trim(TxtReceiptNo) = "" Then
    '    MsgBox "Please Enter Receipt No ! ", vbOKOnly + vbCritical, DialogBoxTitle
    '    TxtReceiptNo.SetFocus
'    ElseIf txtPVDate = "__/__/____" Then
'        MsgBox "Please Enter PV Date", vbOKOnly + vbCritical, DialogBoxTitle
'        txtPVDate.SetFocus
'    ElseIf Trim(txtPVNo) = "" Then
'        MsgBox "Please Enter PV No ! ", vbOKOnly + vbCritical, DialogBoxTitle
'        txtPVNo.SetFocus
    Else
        Screen.MousePointer = vbHourglass
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
      '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

        SaveFlag = True
        tmpRecNo = GenerateSRecNo("RA")
        If OptBordx.Value = True Then
            For BB = 1 To lstBordx.ListItems.Count
                If lstBordx.ListItems(BB).Checked = True Then
                    tmpBordxNo = Trim(lstBordx.ListItems(BB).ListSubItems(2).Text)
                    SaveFlag = True
                    Call SaveRecByBordx(tmpBordxNo)
                End If
            Next BB
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            tmpSql = " and ReceiptNo = '" & Trim(tmpRecNo) & "'"
            lstBordx.ListItems.Clear
            TxtReceiptNo = tmpRecNo
            Call ReceiptBordxListing(tmpSql)
        Else
            Call ProcessWS
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            tmpSql = " and ReceiptNo = '" & Trim(tmpRecNo) & "'"
            lstBordx.ListItems.Clear
            TxtReceiptNo = tmpRecNo
            Call ReceiptWorksheetListing(tmpSql)
        End If
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
        LblTotalAmt = 0
    End If
End Sub

Private Sub cmdEXPtoExcel_Click()

    Screen.MousePointer = vbDefault
    
    If OptBordx.Value = True Then
        If lstBordx.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(lstBordx)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
    
    Else
    
        If lstWorksheet.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(lstWorksheet)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
    
    End If
        
    Screen.MousePointer = vbDefault
End Sub
'************** End Command Button ****************************

'************** Start SearchRecord ****************************
Private Function SearchRecord()
Dim strsql As String
Dim StrAppend As String
Dim sql As String
 
    StrAppend = ""
    SaveFlag = False
    If OptBordx.Value = True Then
       If Len(Trim(CboPayor)) Then
          StrAppend = StrAppend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
       End If
    Else
       If Len(Trim(CboPayor)) Then
          StrAppend = StrAppend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
       End If
    End If
    
    If Len(Trim(TxtBordx1)) Then
      StrAppend = StrAppend + " and BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
      StrAppend = StrAppend + " and BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
    End If
    
    If Mid(CboWSStatus, 1, 1) = 2 Then
      StrAppend = StrAppend + " and (PVRecMNRBStatus = '' or PVRecMNRBStatus is null) "
    End If
    
    If Mid(CboWSStatus, 1, 1) = 1 Then
     StrAppend = StrAppend + " and PVRecMNRBStatus = 'F'"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 1 Then
     StrAppend = StrAppend + " and BordxMNRBFrStatus = 'F'"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 2 Then
     StrAppend = StrAppend + " and (BordxMNRBFrStatus = '' or BordxMNRBFrStatus is null) "
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 3 Then
     StrAppend = StrAppend + " and BordxMNRBFrStatus = 'P'"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 4 Then
     StrAppend = StrAppend + " and (BordxMNRBFrStatus = 'P' or BordxMNRBFrStatus = 'F') "
    End If
    
    If Mid(CboBordxStatus, 1, 1) = 5 Then
     StrAppend = StrAppend + " and (BordxMNRBFrStatus = '' or BordxMNRBFrStatus is null or BordxMNRBFrStatus = 'P') "
    End If
    
    If Len(Trim(txtRecNoFr)) Then
       StrAppend = StrAppend + " AND ReceiptNo >= '" & Trim(txtRecNoFr) & "'"
    End If
       
    If Len(Trim(txtRecNoTo)) Then
       StrAppend = StrAppend + " AND ReceiptNo <= '" & Trim(txtRecNoTo) & "'"
    End If
        
    If Len(Trim(txtRecDateFr)) > 0 And IsDate(txtRecDateFr) = True Then
        StrAppend = StrAppend + " And ReceiptDate >= '" & Format(txtRecDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRecDateTo)) > 0 And IsDate(txtRecDateTo) = True Then
        StrAppend = StrAppend + " And ReceiptDate <= '" & Format(txtRecDateTo, "mm/dd/yyyy") & "'"
    End If
       
    If Len(cboACPeriodMthFr) Then
       StrAppend = StrAppend + " AND RECAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
       
    If Len(cboACPeriodMthTo) Then
       StrAppend = StrAppend + " AND RECAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If Len(cboACPeriodYRFr) Then
       StrAppend = StrAppend + " AND RECAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
       
    If Len(cboACPeriodYRTo) Then
       StrAppend = StrAppend + " AND RECAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkScan.Value = 1 And OptWks.Value = True Then
        If Len(txtWSNoFr) Then
           StrAppend = StrAppend + " AND CASNumber = '" & Trim(txtWSNoFr) & "'"
        End If
    ElseIf ChkScan.Value = 0 And OptWks.Value = True Then
        If Len(txtWSNoFr) Then
           StrAppend = StrAppend + " AND CASNumber >= '" & Trim(txtWSNoFr) & "'"
        End If
        If Len(txtWSNoTo) Then
           StrAppend = StrAppend + " AND CASNumber <= '" & Trim(txtWSNoTo) & "'"
        End If
        
        If Len(Trim(TxtBordx1)) Then
          StrAppend = StrAppend + " and BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
          StrAppend = StrAppend + " and BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
        
    End If
    
    If ChkScan.Value = 1 And OptBordx.Value = True Then
        If Len(Trim(TxtBordx1)) Then
          StrAppend = StrAppend + " and BordxRefNo = '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
    ElseIf ChkScan.Value = 0 And OptBordx.Value = True Then
        If Len(Trim(TxtBordx1)) Then
          StrAppend = StrAppend + " and BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
          StrAppend = StrAppend + " and BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
        
        If Len(txtWSNoFr) Then
           StrAppend = StrAppend + " AND CASNumber >= '" & Trim(txtWSNoFr) & "'"
        End If
        If Len(txtWSNoTo) Then
           StrAppend = StrAppend + " AND CASNumber <= '" & Trim(txtWSNoTo) & "'"
        End If
        
    End If
    
    If OptWks.Value = True Then
        If Len(Trim(txtWSDatefr)) > 0 And IsDate(txtWSDatefr) = True Then
            StrAppend = StrAppend + " And ClaimCloseDate >= '" & Format(txtWSDatefr, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtWSDateTo)) > 0 And IsDate(txtWSDateTo) = True Then
            StrAppend = StrAppend + " And ClaimCloseDate <= '" & Format(txtWSDateTo, "mm/dd/yyyy") & "'"
        End If
    End If
    
    If Len(Trim(txtBordxDateFr)) > 0 And IsDate(txtBordxDateFr) = True Then
        StrAppend = StrAppend + " And BordxDate >= '" & Format(txtBordxDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtBordxDateTo)) > 0 And IsDate(txtBordxDateTo) = True Then
        StrAppend = StrAppend + " And BordxDate <= '" & Format(txtBordxDateTo, "mm/dd/yyyy") & "'"
    End If
    
    
   ' medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

    If OptBordx.Value = True Then
       Call ReceiptBordxListing(StrAppend)
    ElseIf OptWks.Value = True Then
       Call ReceiptWorksheetListing(StrAppend)
    End If
   ' medixmasterconnPayment.Close
   ' Set medixmasterconnPayment = Nothing
  '  medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
 
End Function
'************** End SearchRecord ****************************

Private Function ProcessWS()
Dim CASNo As String
Dim CASVer As String
Dim status As String
Dim BordxNo As String
Dim BankCharge As Double
Dim ContraCase As String
Dim ContraVer As String
Dim ContraAmt As Double
Dim GainLoss As Double
Dim MiscDiff As Double
Dim Discount As Double
Dim WriteOff As Double
Dim MedixRec As Double
Dim CasRemarks As String
Dim ii As Integer
Dim PaymentAmt As Double
Dim ClmCheck1 As New ADODB.Recordset
Dim ClmCheck2 As New ADODB.Recordset
Dim strCheck1 As String
Dim strCheck2 As String
Dim ClmCnt As Integer
Dim FullCnt As Integer
Dim AA As Integer
Dim tmpIndex As String
           
    For ii = 1 To lstWorksheet.ListItems.Count
        If lstWorksheet.ListItems(ii).Checked = True Then
            ContraCase = ""
            ContraVer = ""
            CASNo = Trim(Mid(lstWorksheet.ListItems(ii).ListSubItems(3).Text, 1, IIf(InStr(1, lstWorksheet.ListItems(ii).ListSubItems(3).Text, "-") = 0, 10, InStr(1, lstWorksheet.ListItems(ii).ListSubItems(3).Text, "-") - 1)))
            CASVer = Trim(Mid(lstWorksheet.ListItems(ii).ListSubItems(3).Text, InStr(1, lstWorksheet.ListItems(ii).ListSubItems(3).Text, "-") + 1, Len(lstWorksheet.ListItems(ii).ListSubItems(3).Text) - InStr(1, lstWorksheet.ListItems(ii).ListSubItems(3).Text, "-")))
            BordxNo = Trim(lstWorksheet.ListItems(ii).ListSubItems(2).Text)
            status = Trim(Mid(lstWorksheet.ListItems(ii).ListSubItems(1).Text, 1, 1))
            PaymentAmt = IIf(IsNumeric(lstWorksheet.ListItems(ii).ListSubItems(7)), lstWorksheet.ListItems(ii).ListSubItems(7), 0)
            BankCharge = IIf(IsNumeric(lstWorksheet.ListItems(ii).SubItems(11)), lstWorksheet.ListItems(ii).SubItems(11), 0)
            tmpIndex = IIf(Len(lstWorksheet.ListItems(ii).ListSubItems(27)), lstWorksheet.ListItems(ii).ListSubItems(27), "")
            If Trim(lstWorksheet.ListItems(ii).SubItems(12)) <> "" Then
                AA = InStr(lstWorksheet.ListItems(ii).SubItems(12), "-")
                ContraCase = Mid(lstWorksheet.ListItems(ii).SubItems(12), 1, AA - 1)
                ContraVer = Mid(lstWorksheet.ListItems(ii).SubItems(12), AA + 1, Len(lstWorksheet.ListItems(ii).SubItems(12)) - AA)
            End If
            ContraAmt = IIf(IsNumeric(lstWorksheet.ListItems(ii).SubItems(13)), lstWorksheet.ListItems(ii).SubItems(13), 0)
            GainLoss = IIf(IsNumeric(lstWorksheet.ListItems(ii).SubItems(14)), lstWorksheet.ListItems(ii).SubItems(14), 0)
            MiscDiff = IIf(IsNumeric(lstWorksheet.ListItems(ii).SubItems(15)), lstWorksheet.ListItems(ii).SubItems(15), 0)
            Discount = IIf(IsNumeric(lstWorksheet.ListItems(ii).SubItems(16)), lstWorksheet.ListItems(ii).SubItems(16), 0)
            WriteOff = IIf(IsNumeric(lstWorksheet.ListItems(ii).SubItems(17)), lstWorksheet.ListItems(ii).SubItems(17), 0)
            MedixRec = IIf(IsNumeric(lstWorksheet.ListItems(ii).SubItems(18)), lstWorksheet.ListItems(ii).SubItems(18), 0)
            CasRemarks = IIf(Trim(lstWorksheet.ListItems(ii).SubItems(19)) <> "", lstWorksheet.ListItems(ii).SubItems(19), "")
                    
            Call SaveRecByWS(CASNo, CInt(CASVer), BordxNo, status, PaymentAmt, BankCharge, GainLoss, MiscDiff, Discount, WriteOff, MedixRec, CasRemarks, tmpIndex)
            Call UpdateContra(CASNo, CInt(CASVer), PaymentAmt, ContraCase, ContraVer, ContraAmt)
                        
            strCheck1 = " Select COUNT(*) AS ClmCnt " & _
                        " From Claim Where BordxRefno = '" & Trim(BordxNo) & "'"
        
            strCheck2 = " Select COUNT(*) as FullCnt From Claim " & _
                        " Where PVRecMNRBStatus = 'F' and BordxRefno = '" & Trim(BordxNo) & "'"
            
            Set ClmCheck1 = medixmasterconnWS.Execute(strCheck1)
            Set ClmCheck2 = medixmasterconnWS.Execute(strCheck2)
                                
            ClmCnt = ClmCheck1!ClmCnt
            FullCnt = ClmCheck2!FullCnt
                
            If ClmCnt = FullCnt Then
                Call UpdateBordxStatus(BordxNo, "F")
            Else
                Call UpdateBordxStatus(BordxNo, "P")
            End If
            
        If Len(CASNo) Then
            AddRemark = ""
            AddRemark = vbNewLine & Date & "   Claim Float Transfer  -  " & Trim(tmpRecNo) & " on " & txtReceiptDate & "                                    " & Logon & " "
            AddRemark = ChkStr(Replace(AddRemark, Chr(13), "~"))
            strsqlRem1 = "Select * from CASRemarks where CASNumber = '" & Trim(CASNo) & "'"
            REM1.Open strsqlRem1, medixmasterconn, adOpenKeyset, adLockOptimistic
                If REM1.recordCount > 0 Then
                    RemLen1 = Len(REM1!CASORemarks)
                    Remarks1 = REM1!CASORemarks & AddRemark
                    Remarks1 = ChkStr(Replace(Remarks1, "'", "''"))
                    RemLen2 = Len(REM1!CASORemarksTwo)
                    Remarks2 = REM1!CASORemarksTwo & AddRemark
                    Remarks2 = ChkStr(Replace(Remarks2, "'", "''"))
                End If
            REM1.Close
            Set REM1 = Nothing
            
            strsqlRem2 = "Select * from CASRemarksTwo where CASNumber= '" & Trim(CASNo) & "'"
            REM2.Open strsqlRem2, medixmasterconn, adOpenKeyset, adLockOptimistic
                RemLen3 = Len(REM2!CASORemarks)
                Remarks3 = REM2!CASORemarks & AddRemark
                Remarks3 = ChkStr(Replace(Remarks3, "'", "''"))
                RemLen4 = Len(REM2!CASORemarksTwo)
                Remarks4 = REM2!CASORemarksTwo & AddRemark
                Remarks4 = ChkStr(Replace(Remarks4, "'", "''"))
            REM2.Close
            Set REM2 = Nothing
            'REM3.Close
            'Set REM3 = Nothing
            strsqlRem3 = ""
            strsqlRem3 = "Select * from CASRemarkThree where CASNumber= '" & Trim(CASNo) & "'"
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
                strsqlREM = " update CASRemarkThree set " & _
                         " CASRemarksFive = '" & Remarks7 & "'" & _
                         " where CASNumber = '" & Trim(CASNo) & "'"
            ElseIf (RemLen6 <> 0 And RemLen6 < 3700) Or RemLen5 > 3700 Then
                strsqlREM = " update CASRemarkThree set " & _
                         " CASRemarksFour = '" & Remarks6 & "'" & _
                         " where CASNumber = '" & Trim(CASNo) & "'"
            ElseIf (RemLen5 <> 0 And RemLen5 < 3700) Or RemLen4 > 1800 Then
                strsqlREM = " update CASRemarkThree set " & _
                         " CasRemarks = '" & Remarks5 & "'" & _
                         " where CASNumber = '" & Trim(CASNo) & "'"
            ElseIf (RemLen4 <> 0 And RemLen4 < 1800) Or RemLen3 > 1800 Then
                strsqlREM = " update CASRemarksTwo set " & _
                         " CASORemarksTwo = '" & Remarks4 & "'" & _
                         " where CASNumber = '" & Trim(CASNo) & "'"
            ElseIf (RemLen3 <> 0 And RemLen3 < 1800) Or RemLen2 > 1800 Then
                strsqlREM = " update CASRemarksTwo set " & _
                         " CASORemarks = '" & Remarks3 & "' " & _
                         " where CASNumber = '" & Trim(CASNo) & "'"
            ElseIf (RemLen2 <> 0 And RemLen2 < 1800) Or RemLen1 > 1800 Then
                strsqlREM = " update CasRemarks set " & _
                         " CASORemarksTwo = '" & Remarks2 & "'" & _
                         " where CASNumber = '" & Trim(CASNo) & "'"
            ElseIf RemLen1 <> 0 And RemLen1 < 1800 Then
                strsqlREM = " update CasRemarks set " & _
                         " CASORemarks = '" & Remarks1 & "'" & _
                         " where CASNumber = '" & Trim(CASNo) & "'"
            Else
                strsqlREM = " update CasRemarks set " & _
                         " CASORemarks = '" & Remarks1 & "'" & _
                         " where CASNumber = '" & Trim(CASNo) & "'"

            End If
            medixmasterconn.Execute strsqlREM

        End If
            
        End If
    Next ii
End Function

Private Sub lstBordx_Click()
    Call LvDisplayBordx(, False)
End Sub

Private Sub lstBordx_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim RecAmt As Integer
Dim strsqlBOR As String
Dim BOR As New ADODB.Recordset
Dim BordxNo As String
Dim tmpBordxREcStatus As String

    If Trim(tmpAuthStatus) = "N" And Item.SubItems(6) = "F" Then
        Item.Selected = False
        MsgBox "Record has been paid !", vbCritical
        Item.Checked = False
        Exit Sub
    End If
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    tmpBordxREcStatus = ""
    BordxNo = Item.SubItems(2)
    
    strsqlBOR = "select BordxRefNo, BordxMNRBFrStatus from Bordx where BordxRefNo = '" & Trim(BordxNo) & "'"
    BOR.Open strsqlBOR, medixmasterconnWS, adOpenForwardOnly, adLockReadOnly
    
    Do While Not BOR.EOF
        tmpBordxREcStatus = IIf(Len(BOR!BordxMNRBFRStatus), Trim(BOR!BordxMNRBFRStatus), "")
        BOR.MoveNext
    Loop
    BOR.Close
    Set BOR = Nothing
  '  medixmasterconnWS.Close
 '   Set medixmasterconnWS = Nothing
    If tmpBordxREcStatus = "F" And Trim(tmpAuthStatus) <> "Y" Then
        MsgBox "Record has been paid !", vbCritical
        Item.Checked = False
        Exit Sub
    ElseIf tmpBordxREcStatus = "P" Then
        MsgBox "Record has been Partial Paid ! " & vbNewLine & "Please Pay by Receipt Type - Worksheet !", vbCritical
        Item.Checked = False
        Exit Sub
    End If
    If Item.Checked = True Then
        Call ClearRecDetail
        FrameReceiptPV.Enabled = True
        FrameCas.Enabled = True
        Item.SubItems(1) = "Full"
        If Item.SubItems(7) = "F" Then
            Item.SubItems(6) = Item.SubItems(4)
        Else
            Item.SubItems(6) = CDbl(Item.SubItems(5))
        End If
        
        Item.SubItems(5) = 0
        lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(Item.SubItems(6))
    Else
        lblSelectedAmt = lblSelectedAmt - CDbl(Item.SubItems(6)) 'CDbl(Format(Item.SubItems(3), "#,##0.00"))
        If Item.SubItems(7) = "F" Then
            Item.SubItems(1) = "Full"
        ElseIf Item.SubItems(7) = "P" Then
            Item.SubItems(1) = "Partial"
            Item.SubItems(5) = Item.SubItems(6)
        Else
            Item.SubItems(1) = ""
            Item.SubItems(5) = Item.SubItems(6)
        End If
        Item.SubItems(6) = 0
    End If
    
End Sub

Private Sub lstWorksheet_Click()
    Call LvDisplayWS(, False)
End Sub

Private Sub lstWorksheet_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CmdSave.Enabled = False
    If Trim(tmpAuthStatus) = "N" And Item.SubItems(10) = "F" Then
        MsgBox "Record has been paid !", vbCritical
        Item.Checked = False
        Exit Sub
    End If
    If Item.Checked = True Then
        Call ClearRecDetail
        FrameReceiptPV.Enabled = True
        FrameCas.Enabled = True
        CmdSave.Enabled = True
        If Item.SubItems(10) = "F" Then 'double received
            Item.SubItems(7) = Item.SubItems(4)
            Item.SubItems(27) = ""
        Else
            Item.SubItems(7) = Item.SubItems(6)
        End If
        Item.SubItems(6) = 0
        If CDbl(Item.SubItems(7)) + CDbl(Item.SubItems(5)) >= CDbl(Item.SubItems(4)) Then
            Item.SubItems(1) = "Full"
        Else
            Item.SubItems(1) = "Partial"
        End If
        lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(Item.SubItems(7))
    Else
        lstWorksheet.SelectedItem.Selected = False
        lblSelectedAmt = CDbl(lblSelectedAmt) - CDbl(Item.SubItems(7))
        Item.Selected = True
        Item.SubItems(7) = 0
        Item.SubItems(6) = CDbl(Item.SubItems(4)) - CDbl(Item.SubItems(5))
        If CDbl(Item.SubItems(4)) - CDbl(Item.SubItems(5)) <= 0 Then
            Item.SubItems(1) = "Full"
        ElseIf CDbl(Item.SubItems(4)) = CDbl(Item.SubItems(6)) Then
            Item.SubItems(1) = ""
        Else
            Item.SubItems(1) = "Partial"
        End If
        
    End If
    Call LvDisplayWS(Item, True)
    Call ClearRecDetail
End Sub

Private Sub ReceiptBordxListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim sql2 As String
Dim Current As String
Dim check As Boolean
Dim BordxAmt As Double
Dim RcvdAmt As Double
Dim RcvdCurrRec As New ADODB.Recordset
Dim strsql As String
Dim RecFound As Boolean
Dim Record As ListItem
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim PrevBordxNo As String
Dim CurrBordxNo As String
    
    Current = 0
    If ChkScan.Value = 0 Then
        lstBordx.ListItems.Clear
        lblTotal = 0
        lblSelectedAmt = 0
        LblTotalAmt = 0
    End If
    
    Set cmd1.ActiveConnection = medixmasterconnWS
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_Receipt"
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("ReceiptType", adVarChar, adParamInput, 50, "ReceiptMNRBByBordx")
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, StrAppend)
    cmd1.Parameters.Append Prm2
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    rst2.CursorLocation = adUseClient
    Set rst2 = cmd1.Execute
    RecFound = False
    Do
        Do While Not rst2.EOF
        RecFound = True
        TxtPStatus = ""
            
        If Len(rst2!BordxMNRBFRStatus) Then
            TxtPStatus = rst2!BordxMNRBFRStatus
        End If
        
        With rst2
            Set Record = lstBordx.ListItems.Add(, , "")
            'Record.SubItems(1) = IIf(Len(!BordxPAYCode), !BordxPAYCode, "")
            Record.SubItems(2) = IIf(Len(!BordxRefNo), !BordxRefNo, "")
            ' ----- Group BordxClaimAmt, CheckAmt and OUTSTANDING AMOUNT
            Record.SubItems(4) = Format(IIf(IsNumeric(!CheckAmt), !CheckAmt, 0), "#,##0.00")
            BordxAmt = Format(IIf(IsNumeric(!BordxClaimAmt), !BordxClaimAmt, 0), "#,##0.00")
            RcvdAmt = 0
            strsql = "SELECT BordxNo, SUM(CheckAmt) AS CheckAmt From ReceiptfrMNRB GROUP BY BordxNo HAVING BordxNo = '" & Trim(!BordxRefNo) & "'"
            RcvdCurrRec.Open strsql, medixmasterconnPayment, adOpenForwardOnly, adLockReadOnly
            Do While Not RcvdCurrRec.EOF
                RcvdAmt = IIf(IsNumeric(RcvdCurrRec!CheckAmt), RcvdCurrRec!CheckAmt, 0)
                RcvdCurrRec.MoveNext
            Loop
            RcvdCurrRec.Close
            Set RcvdCurrRec = Nothing
            Record.SubItems(3) = Format(BordxAmt, "#,##0.00")
            Record.SubItems(5) = Format(BordxAmt - RcvdAmt, "#,##0.00")
            Record.SubItems(7) = IIf(Len(!BordxMNRBFRStatus), !BordxMNRBFRStatus, "")
            Record.SubItems(1) = ""
            If TxtPStatus = "F" Then Record.SubItems(1) = "Full"
            If TxtPStatus = "P" Then Record.SubItems(1) = "Partial"
            Record.SubItems(8) = IIf(Trim(!PayPVNo) <> "", Trim(!PayPVNo), "")
            Record.SubItems(9) = IIf(Len(!PayPVDate), Format(!PayPVDate, "dd/mm/yyyy"), "")
            Record.SubItems(10) = IIf(Trim(!ReceiptNo) <> "", Trim(!ReceiptNo), "")
            Record.SubItems(11) = IIf(Len(!ReceiptDate), Format(!ReceiptDate, "dd/mm/yyyy"), "")
            Record.SubItems(12) = IIf(Trim(!RECAccPeriodMth & !RECAccPeriodYr) <> "", !RECAccPeriodMth & "-" & !RECAccPeriodYr, "")
            Record.SubItems(13) = IIf(Trim(!ChequeNo) <> "", Trim(!ChequeNo), "")
            Record.SubItems(14) = IIf(Trim(!Remark) <> "", Trim(!Remark), "")
            Record.SubItems(15) = IIf(IsNumeric(!ReceiptAmt), !ReceiptAmt, "")
            Record.SubItems(16) = IIf(Len(!BordxDate), !BordxDate, "")
            
            Current = Current + 1
            lblCurrent = Current
            CurrBordxNo = !BordxRefNo
            If PrevBordxNo <> CurrBordxNo Then
                PrevBordxNo = !BordxRefNo
                LblTotalAmt = CDbl(LblTotalAmt) + CDbl(Format(!BordxClaimAmt, "#,##0.00"))
            End If
            lblTotal = CDbl(lblTotal) + CDbl(Format(IIf(IsNumeric(!CheckAmt) = True, !CheckAmt, 0), "#,##0.00"))
        End With
        rst2.MoveNext
            
        DoEvents
            If intExit = 1 Then
                check = False
                Exit Do
            End If
         
        Loop
       
    Loop Until check = False
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    
    If RecFound = False Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        lblCurrent = 0
    End If
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
    cmdPrintPV.Enabled = True
    CmdSearch.Enabled = True
End Sub

'************** Start ReceiptWorksheetListing ****************************

Private Sub ReceiptWorksheetListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim Current As Integer
Dim check As Boolean
Dim tmpMBMBTBG As String
Dim tmpAppAmt As Double
Dim tmpPay2H As Double
Dim RecFound As Boolean
Dim StrRec As String
Dim ReceiptRec As New ADODB.Recordset
Dim TotRecAmt As Double
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpMgmtFee As Double
Dim Record As ListItem

    If ChkScan.Value = 0 Then
        LblTotalAmt = 0
        lblTotal = 0
        lstWorksheet.ListItems.Clear
    End If
    'VIEW_WKSReceiptMNRB
    Set cmd1.ActiveConnection = medixmasterconnWS
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_Receipt"
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("ReceiptType", adVarChar, adParamInput, 50, "ReceiptMNRBByWS")
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, StrAppend)
    cmd1.Parameters.Append Prm2
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    rst2.CursorLocation = adUseClient
    Set rst2 = cmd1.Execute
    RecFound = False
    Do
        Do Until rst2.EOF
        RecFound = True
        TxtPStatus = ""
        If Len(rst2!PVRecMNRBStatus) Then
            TxtPStatus = rst2!PVRecMNRBStatus
        End If
        
        With rst2
            Set Record = lstWorksheet.ListItems.Add(, , "")
            'Record.SubItems(1) = IIf(Len(!BordxPAYCode), !BordxPAYCode, "")
            Record.SubItems(1) = ""
            If TxtPStatus = "F" Then Record.SubItems(1) = "Full"
            If TxtPStatus = "P" Then Record.SubItems(1) = "Partial"
            Record.SubItems(10) = IIf(Trim(TxtPStatus) <> "", TxtPStatus, "")
            Record.SubItems(2) = IIf(Len(!BordxRefNo), !BordxRefNo, "")
            If Len(!CASNumber & "-" & !claimversion) Then
               Record.SubItems(3) = (!CASNumber & "-" & !claimversion)
            End If
            tmpMBMBTBG = IIf(Len(!MBMBTBG), Trim(!MBMBTBG), "")
            tmpAppAmt = IIf(IsNumeric(!CLMTotalApproved), !CLMTotalApproved, 0)
            tmpPay2H = IIf(IsNumeric(!CLMPayableByMedix2H), !CLMPayableByMedix2H, 0)
            tmpMgmtFee = IIf(IsNumeric(!CLMCaseMgmtFees), !CLMCaseMgmtFees, 0)
            Record.SubItems(4) = Format(GetReceiptAmt(tmpMBMBTBG, tmpAppAmt, tmpPay2H, tmpMgmtFee), "#,##0.00")
            
            TotRecAmt = 0
            StrRec = "SELECT CASNumber, SUM(CheckAmt) AS totAmt From ReceiptfrMNRB GROUP BY CASNumber,  ClaimVersion HAVING (CASNumber ='" & Trim(!CASNumber) & "' AND (ClaimVersion ='" & Trim(!claimversion) & "'))"
            ReceiptRec.Open StrRec, medixmasterconnPayment, adOpenForwardOnly, adLockReadOnly
            Do While Not ReceiptRec.EOF
                TotRecAmt = IIf(IsNumeric(ReceiptRec!TotAmt), ReceiptRec!TotAmt, 0)
                ReceiptRec.MoveNext
            Loop
            ReceiptRec.Close
            Set ReceiptRec = Nothing
            Record.SubItems(28) = Format(TotRecAmt, "#,##0.00")
            Record.SubItems(5) = Format(IIf(IsNumeric(!CheckAmt), !CheckAmt, 0), "#,##0.00")
            Record.SubItems(6) = Format(Record.SubItems(4) - Record.SubItems(28), "#,##0.00")
            'Record.SubItems(7) Amount Paying
            Record.SubItems(8) = IIf(Len(!ReceiptNo), !ReceiptNo, "")
            If Len(!ReceiptDate) Then
               Record.SubItems(9) = !ReceiptDate
            End If
            
            Record.SubItems(11) = IIf(IsNumeric(!MNRBBankCharge), Format(!MNRBBankCharge, "#,##0.00"), "")
            If Len(!ContraCaseNo & !ContraVersionNo) <> "" Then
                Record.SubItems(12) = (Trim(!ContraCaseNo) & "-" & Trim(!ContraVersionNo))
            End If
            Record.SubItems(13) = IIf(IsNumeric(!ContraCaseAmt), Format(!ContraCaseAmt, "#,##0.00"), "")
            Record.SubItems(14) = IIf(IsNumeric(!MNRBGainLoss), Format(!MNRBGainLoss, "#,##0.00"), "")
            Record.SubItems(15) = IIf(IsNumeric(!MNRBMisc), Format(!MNRBMisc, "#,##0.00"), "")
            Record.SubItems(16) = IIf(IsNumeric(!MNRBDiscount), Format(!MNRBDiscount, "#,##0.00"), "")
            Record.SubItems(17) = IIf(IsNumeric(!MNRBWriteOff), Format(!MNRBWriteOff, "#,##0.00"), "")
            Record.SubItems(18) = IIf(IsNumeric(!MNRBMedixReceipt), Format(!MNRBMedixReceipt, "#,##0.00"), "")
            Record.SubItems(19) = IIf(Len(!MNRBRemarks), Trim(!MNRBRemarks), "")
            Record.SubItems(20) = IIf(IsNumeric(!ReceiptAmt), Format(!ReceiptAmt, "#,##0.00"), "")
            Record.SubItems(21) = IIf(Len(!RECAccPeriodMth), Trim(!RECAccPeriodMth), "")
            Record.SubItems(22) = IIf(Len(!RECAccPeriodYr), Trim(!RECAccPeriodYr), "")
            Record.SubItems(23) = IIf(Len(!Remark), Trim(!Remark), "")
            Record.SubItems(24) = IIf(Len(!PayPVNo), Trim(!PayPVNo), "")
            Record.SubItems(25) = IIf(Len(!PayPVDate), Trim(!PayPVDate), "")
            Record.SubItems(26) = IIf(Len(!ChequeNo), Trim(!ChequeNo), "")
            Record.SubItems(27) = IIf(Len(!RecMNRBIndex), !RecMNRBIndex, "")
            Record.SubItems(29) = IIf(Len(!ClaimCloseDate), Trim(!ClaimCloseDate), "")
            Record.SubItems(30) = IIf(Len(!BordxDate), !BordxDate, "")
            Current = Current + 1
            lblCurrent = Current
            LblTotalAmt = CDbl(LblTotalAmt) + Record.SubItems(4)
            lblTotal = CDbl(lblTotal) + CDbl(Format(IIf(IsNumeric(!CheckAmt) = True, !CheckAmt, 0), "#,##0.00"))
        End With
        rst2.MoveNext
        
        DoEvents
            If intExit = 1 Then
                check = False
                Exit Do
            End If
        Loop
    Loop Until check = False
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    If RecFound = False Then
        MsgBox "No record found !", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    cmdEXPtoExcel.Enabled = True
    CmdClear.Enabled = True
    cmdPrintPV.Enabled = True
    CmdSearch.Enabled = True
    CmdExit.Enabled = True
End Sub

'************** End ReceiptWorksheetListing ****************************
Private Sub OptBordx_GotFocus()
    txtWSNoFr.Enabled = False
    txtWSNoTo.Enabled = False
    txtWSNoFr = ""
    txtWSNoTo = ""
End Sub

Private Sub TxtAmtPaid_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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

Private Sub txtChequeNo_LostFocus()
    txtChequeNo = ChkStr(txtChequeNo)
End Sub

Private Sub txtRecDateFr_LostFocus()
    If IsDate(txtRecDateFr) Then
    Else
        If txtRecDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRecDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtRecDateTo_LostFocus()
    If IsDate(txtRecDateTo) Then
    Else
        If txtRecDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRecDateTo.SetFocus
        End If
    End If
End Sub

Private Sub TxtReceiptNo_LostFocus()
    TxtReceiptNo = ChkStr(TxtReceiptNo)
End Sub

Private Sub txtPVNo_LostFocus()
    txtPVNo = ChkStr(txtPVNo)
End Sub

Private Sub txtReceiptRemark_LostFocus()
    txtReceiptRemark = ChkStr(txtReceiptRemark)
End Sub

Private Sub TxtWriteOff_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtMedixRec_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtBankCharge_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtBordx1_LostFocus()
    TxtBordx1 = ChkStr(TxtBordx1)
    CmdSearch.Default = False
End Sub

Private Sub TxtBordx2_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtBordx2_LostFocus()
    TxtBordx2 = ChkStr(TxtBordx2)
End Sub

Private Sub txtContraAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtContraVer_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtDiscount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtGainLoss_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtMisc_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

'*************************** Start Diable Non Integer Value
Private Sub TxtReceiptAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

'*************************** Start Combobox Change/ KeyPress

Private Sub CboPayor_Change()
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Dim strsql As String


  If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
    Next i
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboPayor.Text = ValidValue
            CboPayor.SelStart = 0
            CboPayor.SelLength = Len(ValidValue)
        End If
  End If
End Sub
'*************************** End Combobox Change/ KeyPress

Private Sub SaveRecByBordx(BordxNo As String)
Dim strsqlREC As String
Dim SearchDB As New ADODB.Recordset
Dim sql As String
Dim tmpMBMBTBG As String
Dim tmpAppAmt As Double
Dim tmpPay2H As Double
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpMgmtFee As Double
Dim tmpChkAmt As Double
Dim tmpReceiptAmt As Double
Dim strsqlCLM As String

    sql = " AND BordxRefNo = '" & Trim(BordxNo) & "'"
    Set cmd1.ActiveConnection = medixmasterconnWS
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_Receipt"
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("ReceiptType", adVarChar, adParamInput, 50, "ReceiptMNRBByWS")
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, sql)
    cmd1.Parameters.Append Prm2
    SearchDB.CursorType = adOpenForwardOnly
    SearchDB.CacheSize = 100
    SearchDB.CursorLocation = adUseClient
    Set SearchDB = cmd1.Execute

    Do While Not SearchDB.EOF
        tmpMBMBTBG = IIf(Len(SearchDB!MBMBTBG), Trim(SearchDB!MBMBTBG), "")
        tmpAppAmt = IIf(IsNumeric(SearchDB!CLMTotalApproved), SearchDB!CLMTotalApproved, 0)
        tmpPay2H = IIf(IsNumeric(SearchDB!CLMPayableByMedix2H), SearchDB!CLMPayableByMedix2H, 0)
        tmpMgmtFee = IIf(IsNumeric(SearchDB!CLMCaseMgmtFees), SearchDB!CLMCaseMgmtFees, 0)
        tmpChkAmt = Format(GetReceiptAmt(tmpMBMBTBG, tmpAppAmt, tmpPay2H, tmpMgmtFee), "#,##0.00")
        tmpReceiptAmt = IIf(IsNumeric(txtReceiptAmt), txtReceiptAmt, 0)
        If SaveFlag = True Then
            strsqlREC = "Insert Into ReceiptfrMNRB(CasNumber, ClaimVersion, ReceiptNo," & _
                 "ReceiptDate, ReceiptAmt, ChequeNo, BordxNo, " & _
                 "MNRBLastUpdateDate, MNRBLastUpdateUser," & _
                 "RECMNRBAmt, CheckAmt, PAYPVNo, PAYPVDate," & _
                 "RECAccPeriodMth, RECAccPeriodYR, Remark) " & _
                 "Values('" & Trim(SearchDB!CASNumber) & "','" & Trim(SearchDB!claimversion) & "','" & Trim(tmpRecNo) & _
                 "','" & Format(txtReceiptDate, "yyyy/mm/dd") & "'," & tmpReceiptAmt & _
                 ",'" & Trim(txtChequeNo) & "','" & BordxNo & _
                 "','" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'," & tmpReceiptAmt & _
                 "," & tmpChkAmt & ",'" & IIf(Len(txtPVNo), Trim(txtPVNo), "") & "','" & IIf(IsDate(txtPVDate), Format(txtPVDate, "yyyy/mm/dd"), "") & _
                 "','" & Trim(cboAccPeriodMth) & "','" & Trim(cboAccPeriodYR) & _
                 "','" & Trim(txtReceiptRemark) & "')"
        Else
            strsqlREC = "UPdate ReceiptfrMNRB SET ReceiptDate='" & Format(txtReceiptDate, "yyyy/mm/dd") & _
                 "', ReceiptAmt=" & tmpReceiptAmt & ", ChequeNo='" & Trim(txtChequeNo) & _
                 "', BordxNo='" & BordxNo & "', MNRBLastUpdateDate='" & Format(Date, "YYYY/MM/DD") & _
                 "', MNRBLastUpdateUser='" & Logon & "', RECMNRBAmt=" & tmpReceiptAmt & _
                 ", PAYPVNo='" & IIf(Len(txtPVNo), Trim(txtPVNo), "") & "', PAYPVDate='" & IIf(IsDate(txtPVDate), Format(txtPVDate, "yyyy/mm/dd"), "") & _
                 "', RECAccPeriodMth='" & Trim(cboAccPeriodMth) & "', RECAccPeriodYR='" & Trim(cboAccPeriodYR) & _
                 "', Remark='" & Trim(txtReceiptRemark) & _
                 "' WHERE ReceiptNo= '" & Trim(TxtReceiptNo) & _
                 "' AND CasNumber='" & Trim(SearchDB!CASNumber) & "' AND claimversion='" & Trim(SearchDB!claimversion) & "'"
                Call UPdateWholeRecInfo
        End If
        
        medixmasterconnPayment.Execute strsqlREC
        
        strsqlCLM = "Update Claim set PVRecMNRBStatus ='F' Where BordxRefNo = '" & Trim(BordxNo) & _
                    "' And CASNumber = '" & Trim(SearchDB!CASNumber) & "' AND ClaimVersion = '" & Trim(SearchDB!claimversion) & "'"
        medixmasterconnWS.Execute strsqlCLM
        
        If Len(SearchDB!CASNumber) Then
            AddRemark = ""
            AddRemark = vbNewLine & Date & "   Claim Receipt  -  " & Trim(tmpRecNo) & " on " & txtReceiptDate & "                                         " & Logon & " "
            AddRemark = ChkStr(Replace(AddRemark, Chr(13), "~"))
            strsqlRem1 = "Select * from CASRemarks where CASNumber = '" & Trim(SearchDB!CASNumber) & "'"
            REM1.Open strsqlRem1, medixmasterconn, adOpenKeyset, adLockOptimistic
                If REM1.recordCount > 0 Then
                    RemLen1 = Len(REM1!CASORemarks)
                    Remarks1 = REM1!CASORemarks & AddRemark
                    Remarks1 = ChkStr(Replace(Remarks1, "'", "''"))
                    RemLen2 = Len(REM1!CASORemarksTwo)
                    Remarks2 = REM1!CASORemarksTwo & AddRemark
                    Remarks2 = ChkStr(Replace(Remarks2, "'", "''"))
                End If
            REM1.Close
            Set REM1 = Nothing
            
            strsqlRem2 = "Select * from CASRemarksTwo where CASNumber= '" & Trim(SearchDB!CASNumber) & "'"
            REM2.Open strsqlRem2, medixmasterconn, adOpenKeyset, adLockOptimistic
                RemLen3 = Len(REM2!CASORemarks)
                Remarks3 = REM2!CASORemarks & AddRemark
                Remarks3 = ChkStr(Replace(Remarks3, "'", "''"))
                RemLen4 = Len(REM2!CASORemarksTwo)
                Remarks4 = REM2!CASORemarksTwo & AddRemark
                Remarks4 = ChkStr(Replace(Remarks4, "'", "''"))
            REM2.Close
            Set REM2 = Nothing
            'REM3.Close
            'Set REM3 = Nothing
            strsqlRem3 = ""
            strsqlRem3 = "Select * from CASRemarkThree where CASNumber= '" & Trim(SearchDB!CASNumber) & "'"
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
                strsqlREM = " update CASRemarkThree set " & _
                         " CASRemarksFive = '" & Remarks7 & "'" & _
                         " where CASNumber = '" & Trim(SearchDB!CASNumber) & "'"
            ElseIf (RemLen6 <> 0 And RemLen6 < 3700) Or RemLen5 > 3700 Then
                strsqlREM = " update CASRemarkThree set " & _
                         " CASRemarksFour = '" & Remarks6 & "'" & _
                         " where CASNumber = '" & Trim(SearchDB!CASNumber) & "'"
            ElseIf (RemLen5 <> 0 And RemLen5 < 3700) Or RemLen4 > 1800 Then
                strsqlREM = " update CASRemarkThree set " & _
                         " CasRemarks = '" & Remarks5 & "'" & _
                         " where CASNumber = '" & Trim(SearchDB!CASNumber) & "'"
            ElseIf (RemLen4 <> 0 And RemLen4 < 1800) Or RemLen3 > 1800 Then
                strsqlREM = " update CASRemarksTwo set " & _
                         " CASORemarksTwo = '" & Remarks4 & "'" & _
                         " where CASNumber = '" & Trim(SearchDB!CASNumber) & "'"
            ElseIf (RemLen3 <> 0 And RemLen3 < 1800) Or RemLen2 > 1800 Then
                strsqlREM = " update CASRemarksTwo set " & _
                         " CASORemarks = '" & Remarks3 & "' " & _
                         " where CASNumber = '" & Trim(SearchDB!CASNumber) & "'"
            ElseIf (RemLen2 <> 0 And RemLen2 < 1800) Or RemLen1 > 1800 Then
                strsqlREM = " update CasRemarks set " & _
                         " CASORemarksTwo = '" & Remarks2 & "'" & _
                         " where CASNumber = '" & Trim(SearchDB!CASNumber) & "'"
            ElseIf RemLen1 <> 0 And RemLen1 < 1800 Then
                strsqlREM = " update CasRemarks set " & _
                         " CASORemarks = '" & Remarks1 & "'" & _
                         " where CASNumber = '" & Trim(SearchDB!CASNumber) & "'"
            Else
                strsqlREM = " update CasRemarks set " & _
                         " CASORemarks = '" & Remarks1 & "'" & _
                         " where CASNumber = '" & Trim(SearchDB!CASNumber) & "'"

            End If
            medixmasterconn.Execute strsqlREM

        End If
        
        SearchDB.MoveNext
    Loop
    SearchDB.Close
    Set SearchDB = Nothing
    Call UpdateBordxStatus(BordxNo, "F")

End Sub

Private Sub SaveRecByWS(CASNo As String, CASVer As Integer, BordxNo As String, status As String, PaymentAmt As Double, BankCharge As Double, GainLoss As Double, MiscDiff As Variant, Discount As Double, WriteOff As Double, MedixRec As Double, Remarks As String, tmpIndex As String)
Dim strsqlREC1 As String
Dim strsqlCLM As String
Dim tmpReceiptAmt As Double
Dim tmpMth As Integer
Dim tmpYear As Integer
    tmpMth = CDec(cboAccPeriodMth)
    tmpYear = CDec(cboAccPeriodYR)
    tmpReceiptAmt = IIf(IsNumeric(txtReceiptAmt), txtReceiptAmt, 0)
    
    If SaveFlag = True Then
        strsqlREC1 = "Insert Into ReceiptfrMNRB(CasNumber,claimversion,ReceiptNo," & _
                 "ReceiptDate, ReceiptAmt, ChequeNo, Remark, BordxNo, " & _
                 "MNRBLastUpdateDate,MNRBLastUpdateUser," & _
                 "RECMNRBAmt, CheckAmt, PAYPVNo, PAYPVDate," & _
                 "RECAccPeriodMth, RECAccPeriodYR, MNRBBankCharge, MNRBGainLoss, MNRBMisc," & _
                 "MNRBDiscount, MNRBWriteOff, MNRBMedixReceipt, MNRBRemarks) " & _
                 "Values('" & Trim(CASNo) & "','" & Trim(CASVer) & "','" & Trim(tmpRecNo) & _
                 "','" & Format(txtReceiptDate, "yyyy/mm/dd") & "'," & tmpReceiptAmt & _
                 ",'" & Trim(txtChequeNo) & "','" & Trim(txtReceiptRemark) & "','" & BordxNo & _
                 "','" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'," & tmpReceiptAmt & _
                 "," & PaymentAmt & ",'" & IIf(Len(txtPVNo), Trim(txtPVNo), "") & "','" & IIf(IsDate(txtPVDate), Format(txtPVDate, "yyyy/mm/dd"), "") & _
                 "','" & Trim(cboAccPeriodMth) & "','" & Trim(cboAccPeriodYR) & "'," & BankCharge & _
                 "," & GainLoss & "," & MiscDiff & "," & Discount & "," & WriteOff & "," & MedixRec & _
                 ",'" & Trim(Remarks) & "')"
    Else
        strsqlREC1 = "UPdate ReceiptfrMNRB SET ReceiptDate='" & Format(txtReceiptDate, "yyyy/mm/dd") & _
                 "', ReceiptAmt=" & tmpReceiptAmt & ", ChequeNo ='" & Trim(txtChequeNo) & _
                 "', BordxNo='" & BordxNo & "', MNRBLastUpdateDate ='" & Format(Date, "YYYY/MM/DD") & _
                 "', MNRBLastUpdateUser ='" & Logon & "', RECMNRBAmt =" & tmpReceiptAmt & _
                 ", CheckAmt=" & PaymentAmt & ", PAYPVNo='" & IIf(Len(txtPVNo), Trim(txtPVNo), "") & "', PAYPVDate='" & IIf(IsDate(txtPVDate), Format(txtPVDate, "yyyy/mm/dd"), "") & _
                 "', RECAccPeriodMth = '" & cboAccPeriodMth & "', RECAccPeriodYR = '" & cboAccPeriodYR & _
                 "', MNRBRemarks='" & Trim(TxtCasRemark) & "', Remark='" & Trim(txtReceiptRemark) & _
                 "', MNRBGainLoss=" & GainLoss & ", MNRBMisc=" & MiscDiff & ", MNRBDiscount=" & Discount & _
                 ", MNRBWriteOff=" & WriteOff & ", MNRBMedixReceipt=" & MedixRec & _
                 " WHERE RecMNRBIndex= " & Trim(tmpIndex) & ""
                 Call UPdateWholeRecInfo
    End If
    medixmasterconnPayment.Execute strsqlREC1
    
    strsqlCLM = "Update Claim set PVRecMNRBStatus ='" & status & "' Where CASNumber = '" & Trim(CASNo) & "' AND ClaimVersion = '" & Trim(CASVer) & "'"
    medixmasterconnWS.Execute strsqlCLM
    
End Sub

Private Sub UpdateBordxStatus(BordxNo As String, status As String)
Dim strsqlBordx  As String

    strsqlBordx = "Update Bordx set BordxMNRBRecDate='" & Format(txtReceiptDate, "yyyy/mm/dd") & _
                  "', BordxMNRBFRStatus= '" & Trim(status) & "' where BordxrefNo = '" & Trim(BordxNo) & "'"
        
    medixmasterconnWS.Execute strsqlBordx
End Sub

Private Sub OptBordx_Click()
     If OptBordx.Value = True Then
        ChkScan.Value = 1
        FrameCas.Enabled = False
        lstBordx.Visible = True
        CboBordxStatus.Visible = True
        CboWSStatus.Visible = False
        CboBordxStatus.ListIndex = 0
        CboWSStatus.ListIndex = 0
        Call ClearRecDetail
        lstWorksheet.Visible = False
        txtWSNoFr.Enabled = False
        txtWSNoTo.Enabled = False
        txtWSNoTo.Visible = True
        lblWSTo.Visible = True
        txtWSNoFr = ""
        txtWSNoTo = ""
        ChkScan.Enabled = True
        lblBordxTo.Visible = False
        TxtBordx2.Visible = False
        TxtBordx1.SetFocus
        ChkScan.Left = 3360
        ChkScan.Top = 840
        'txtWSDateFr.Text = ""
        txtWSDatefr.Text = "__/__/____"
        txtWSDatefr.Enabled = False
        txtWSDateTo.Enabled = False
        
     End If
End Sub

Private Sub OptWks_Click()
    If OptWks.Value = True Then
        ChkScan.Value = 1
        FrameCas.Enabled = True
        lstWorksheet.Visible = True
        Call ClearRecDetail
        CboBordxStatus.Visible = False
        CboWSStatus.Visible = True
        CboWSStatus.ListIndex = 0
        CboBordxStatus.ListIndex = 0
        lstBordx.Visible = False
        TxtBordx2.Visible = True
        txtWSNoFr.Enabled = True
        lblWSTo.Visible = False
        lblBordxTo.Visible = True
        txtWSNoTo.Visible = False
        txtWSNoFr.SetFocus
        ChkScan.Enabled = True
        ChkScan.Value = 1
        ChkScan.Left = 3360
        ChkScan.Top = 600
        txtWSDatefr.Enabled = True
        txtWSDateTo.Enabled = True
    End If
End Sub
'*************************** End Option Click

'************** Start ComboListing ****************************
Private Sub ComboListing()
    
    With CboBordxStatus
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
        .AddItem "3 - Partial Paid"
        .AddItem "4 - Paid & Partial Paid"
        .AddItem "5 - Unpaid & Partial Paid"
    End With
        
    With CboWSStatus
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
    End With
    
End Sub

Private Sub PayorComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
   
    Do Until PAY.EOF
        With PAY
            CboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
  
End Sub
'************** End ComboListing ****************************

Private Sub txtReceiptAmt_LostFocus()
Dim tmpReceiptAmt As Double
    tmpReceiptAmt = IIf(IsNumeric(txtReceiptAmt), txtReceiptAmt, 0)
    If CDbl(lblSelectedAmt) <> 0 Then
        If CDbl(tmpReceiptAmt) <> CDbl(lblSelectedAmt) Then
            MsgBox "Receipt Amt & Checked Amt Are Different! " & vbNewLine & _
                   "Please Fill in Remark.", vbOKOnly + vbCritical
        End If
    End If
End Sub

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
        cboAccPeriodMth.Text = Month(txtReceiptDate)
        cboAccPeriodYR.Text = Format(txtReceiptDate, "yyyy")
    Else
        If txtReceiptDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDate.SetFocus
        End If
    End If
End Sub


Private Sub lstBordx_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstBordx, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstBordx, ColumnHeader.Index, ldtDateTime, m_blnDirection(12)
    Case 13
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    Case 14
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstBordx, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
    Case 17
        SortListView lstBordx, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
        
    End Select
End Sub

Private Sub lstWorksheet_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstWorksheet, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    Case 14
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    Case 15
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
    Case 16
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
    Case 17
        SortListView lstWorksheet, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
    Case 18
        SortListView lstWorksheet, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 30
        SortListView lstWorksheet, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
    Case 31
        SortListView lstWorksheet, ColumnHeader.Index, ldtDateTime, m_blnDirection(18)
        
    End Select
End Sub

Private Function ClearForm()
On Error Resume Next
Dim ctl As Control

For Each ctl In Me.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    ElseIf TypeOf ctl Is ListBox Then
            ctl.Clear
    ElseIf TypeOf ctl Is ComboBox Then
            ctl.Text = ""
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    End If
Next ctl
End Function

Private Function CaseNoCheck(tmpCnt As Integer) As Boolean
    CaseNoCheck = False
    If Trim(ChkStr(txtWSNoFr)) = Trim(Mid(lstWorksheet.ListItems(tmpCnt).SubItems(3), 1, IIf(InStr(1, lstWorksheet.ListItems(tmpCnt).SubItems(3), "-") = 0, 3, InStr(1, lstWorksheet.ListItems(tmpCnt).SubItems(3), "-") - 1))) Then
        CaseNoCheck = True
    End If
End Function

Private Function BordxNoCheck(tmpCnt As Integer) As Boolean
    BordxNoCheck = False
    If Trim(ChkStr(TxtBordx1)) = Trim(lstBordx.ListItems(tmpCnt).SubItems(2)) Then
        BordxNoCheck = True
    End If
End Function

Private Sub txtWSDateFr_LostFocus()
    If IsDate(txtWSDatefr) Then
    Else
        If txtWSDatefr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtWSDatefr.SetFocus
        End If
    End If

End Sub

Private Sub txtWSDateTo_LostFocus()
    If IsDate(txtWSDateTo) Then
    Else
        If txtWSDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtWSDateTo.SetFocus
        End If
    End If
End Sub

Private Sub txtWSNoFr_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub txtWSNoFr_LostFocus()
    CmdSearch.Default = False
    txtWSNoFr = ChkStr(txtWSNoFr)
End Sub

Private Sub cmdPrintPV_Click()
On Error Resume Next
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim i As Integer
Dim strsql As String
Dim REC As New ADODB.Recordset
Dim strCount
    
    Screen.MousePointer = vbHourglass
  '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

    If TxtReceiptNo = "" Then
    
        MsgBox "Please Enter or Generate Receipt No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
    Else
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
    
        strsql = " Select * From View_ReceiptFrMNRBRec Where ReceiptNo = '" & Trim(TxtReceiptNo) & "'"
        
        strCount = " Select count(*) as totalRecord From View_ReceiptFrMNRBRec Where ReceiptNo = '" & Trim(TxtReceiptNo) & "'"
        
        Set strCount = medixmasterconnPayment.Execute(strCount)
        REC.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                
        'If REC.recordCount Then
        If strCount!TotalRecord Then
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
        
            objExcel.WindowState = xlMaximized
    
        Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\OR-MNRBS.xlt")
                
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
        
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
            
            objWorksheet.Cells(8, "I") = Trim(REC!ReceiptNo)
            objWorksheet.Cells(10, "D") = "MNRB"
            objWorksheet.Cells(9, "O") = Format(REC!ReceiptDate, "dd/mm/yyyy")
            objWorksheet.Cells(12, "K") = Format(REC!ReceiptAmt, "#,##0.00")
            objWorksheet.Cells(14, "O") = Trim(REC!ChequeNo)
            
            For i = 1 To strCount!TotalRecord
            
                objWorksheet.Cells(18 + CDbl(i), "B") = i
                objWorksheet.Cells(18 + CDbl(i), "C") = Trim(REC!BordxNo)
                        
            REC.MoveNext
            
            Next i
            
            End If
            'cmdCredit.Enabled = False
            objExcel.DisplayAlerts = True
            'objWorksheet.Columns.AutoFit
            objExcel.Interactive = True
        '    medixmasterconnPayment.Close
       '     Set medixmasterconnPayment = Nothing

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
            GoTo ExitFunction
    End If
End Sub

Private Sub UpdateContra(CASNo As String, CASVer As String, PaymentAmt As Variant, ContraCase As String, ContraVer As String, ContraAmt As Variant)
Dim strsqlContra As String
Dim Contra As New ADODB.Recordset
    
            
    If ContraCase <> "" Then
        
        strsqlContra = " Select * from ContraAllocation " & _
                       " Where CaseNo = '" & Trim(CASNo) & "' And VersionNo = '" & Trim(CASVer) & "'"
                
        Contra.Open strsqlContra, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
        
        If Contra.recordCount = 0 Then
            Contra.AddNew
        End If
            
        With Contra
        
            !refno = ChkStr(TxtReceiptNo)
            !CaseNo = Trim(CASNo)
            !VersionNo = Trim(CASVer)
            !CaseAmt = Format(PaymentAmt, "#,##0.00")
            !ContraCaseNo = Trim(ContraCase)
            !ContraVersionNo = Trim(ContraVer)
            !ContraCaseAmt = Format(ContraAmt, "#,##0.00")
        
        End With
        Contra.Update
        Contra.Close
        Set Contra = Nothing
            
    End If
    
    'txtCHKAmt = 0
End Sub

Private Sub ClearRecDetail()
    TxtReceiptNo.Text = ""
    txtReceiptDate.Text = "__/__/____"
    txtReceiptAmt.Text = ""
    txtChequeNo.Text = ""
    txtReceiptRemark = ""
    cboAccPeriodMth.Text = ""
    cboAccPeriodYR.Text = ""
    txtPVDate.Text = "__/__/____"
    txtPVNo.Text = ""
    txtStatus = ""
    txtBordxNO = ""
    txtIndex = ""
End Sub

Private Function GetReceiptAmt(tmpMBMBTBG As String, tmpAppAmt As Double, tmpPay2H As Double, tmpMgmtFee As Double)
    GetReceiptAmt = tmpAppAmt + tmpMgmtFee
    If Trim(tmpMBMBTBG) = "Y" And Trim(tmpAppAmt) = 0 Then
       GetReceiptAmt = tmpPay2H
    ElseIf Trim(tmpMBMBTBG) = "" And Trim(tmpAppAmt) = 0 Then
       GetReceiptAmt = tmpPay2H
    End If
End Function

Private Sub ClearContraDetail()
    TxtBankCharge.Text = ""
    TxtContraCase.Text = ""
    TxtContraVer.Text = ""
    TxtContraAmt.Text = ""
    TxtGainLoss.Text = ""
    TxtMisc.Text = ""
    TxtDiscount.Text = ""
    TxtWriteOff.Text = ""
    TxtMedixRec.Text = ""
    TxtAmtPaid = ""
    TxtCasRemark.Text = ""
    TxtClaimNo = ""
    TxtMedixRec = ""
End Sub

Private Sub EnabledCommand(tmpStatus As Boolean)
    CmdExit.Enabled = tmpStatus
    CmdEdit.Enabled = tmpStatus
    CmdUpdate.Enabled = tmpStatus
    cmdEXPtoExcel.Enabled = tmpStatus
    CmdSave.Enabled = tmpStatus
    CmdClear.Enabled = tmpStatus
    cmdPrintPV.Enabled = tmpStatus
    CmdSearch.Enabled = tmpStatus
End Sub

Private Sub LvDisplayWS(Optional Item As ListItem, Optional ItemChk As Boolean)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
Dim AA As Integer
Dim tmpAmt As Double
Dim lstRec As ListItem
    Call ClearRecDetail
    Call ClearContraDetail
    If lstWorksheet.ListItems.Count Then
        If ItemChk = True Then
            Set lstRec = Item
            lstWorksheet.SelectedItem.Selected = False
            lstRec.Selected = True
        Else
            Set lstRec = lstWorksheet.SelectedItem
        End If
        txtPVNo = IIf(Len(Trim(lstRec.SubItems(8))), lstRec.SubItems(8), "")
        txtBordxNO = IIf(Len(Trim(lstRec.SubItems(2))), lstRec.SubItems(2), "")
        TxtClaimNo = lstRec.SubItems(3)
        txtStatus = IIf(Len(Trim(lstRec.SubItems(10))), lstRec.SubItems(10), "")
        TxtReceiptNo = lstRec.SubItems(8)
        If lstRec.SubItems(9) <> "" Then
            txtReceiptDate = lstRec.SubItems(9)
        End If
        txtReceiptAmt = lstRec.SubItems(20)
        txtChequeNo = lstRec.SubItems(26)
        
        cboAccPeriodMth = lstRec.SubItems(21)
        cboAccPeriodYR = lstRec.SubItems(22)
        txtPVNo = IIf(Len(lstRec.SubItems(24)), lstRec.SubItems(24), "")
        If Len(lstRec.SubItems(25)) Then
            txtPVDate = lstRec.SubItems(25)
        End If
        tmpAmt = IIf(IsNumeric(lstRec.SubItems(7)), lstRec.SubItems(7), 0)
        If tmpAmt <> 0 Then
            TxtAmtPaid = tmpAmt
        ElseIf IsNumeric(lstRec.SubItems(5)) Then
            TxtAmtPaid = lstRec.SubItems(5)
        End If
        
        If Trim(lstRec.SubItems(12)) <> "" Then
            AA = InStr(1, lstRec.SubItems(12), "-")
            TxtContraCase = Mid(lstRec.SubItems(12), 1, AA - 1)
            TxtContraVer = Mid(lstRec.SubItems(12), AA + 1, Len(lstRec.SubItems(12)) - AA)
        End If
        TxtBankCharge = lstRec.SubItems(11)
        TxtContraAmt = lstRec.SubItems(13)
        TxtGainLoss = lstRec.SubItems(14)
        TxtMisc = lstRec.SubItems(15)
        TxtDiscount = lstRec.SubItems(16)
        TxtWriteOff = lstRec.SubItems(17)
        TxtMedixRec = lstRec.SubItems(18)
        TxtCasRemark = lstRec.SubItems(19)
        txtReceiptRemark = lstRec.SubItems(23)
        txtIndex = IIf(Len(lstRec.SubItems(27)), lstRec.SubItems(27), "")
    End If
    CmdSave.Enabled = False
    CmdEdit.Visible = True
    CmdEdit.Enabled = False
    CmdUpdate.Visible = False
    FrameReceiptPV.Enabled = False
    'TxtReceiptNo.Enabled = False
    FrameCas.Enabled = False
    If CDbl(lblSelectedAmt) <> 0 Then
        FrameReceiptPV.Enabled = True
        'TxtReceiptNo.Enabled = True
        FrameCas.Enabled = True
        CmdSave.Enabled = True
    Else
        If Trim(lstRec.SubItems(10)) = "" Then
            FrameReceiptPV.Enabled = True
            'TxtReceiptNo.Enabled = True
            FrameCas.Enabled = True
            CmdSave.Enabled = True
        Else
            If Trim(TxtReceiptNo) <> "" Then
                'TxtReceiptNo.Enabled = False
                CmdEdit.Visible = True
                CmdUpdate.Visible = False
                CmdEdit.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub LvDisplayBordx(Optional Item As ListItem, Optional ItemChk As Boolean)
Dim AA As Integer
Dim lstRec As ListItem
    Call ClearRecDetail
    Call ClearContraDetail
    CmdSave.Enabled = False
    
    If lstBordx.ListItems.Count Then
        If ItemChk = True Then
            Set lstRec = Item
            lstBordx.SelectedItem.Selected = False
            lstRec.Selected = True
        Else
            Set lstRec = lstBordx.SelectedItem
        End If
        txtBordxNO = IIf(Len(Trim(lstRec.SubItems(2))), lstRec.SubItems(2), "")
        txtStatus = IIf(Len(Trim(lstRec.SubItems(7))), lstRec.SubItems(7), "")
        txtReceiptAmt = lstRec.SubItems(15)
        txtPVNo = IIf(Len(lstRec.SubItems(8)), lstRec.SubItems(8), "")
        If lstRec.SubItems(9) <> "" Then
            txtPVDate = lstRec.SubItems(9)
        End If
        TxtReceiptNo = lstRec.SubItems(10)
        If lstRec.SubItems(11) <> "" Then
            txtReceiptDate = lstRec.SubItems(11)
        End If
        If Trim(lstRec.SubItems(12)) <> "" Then
            AA = InStr(1, lstRec.SubItems(12), "-")
            cboAccPeriodMth = Mid(lstRec.SubItems(12), 1, AA - 1)
            cboAccPeriodYR = Mid(lstRec.SubItems(12), AA + 1, Len(lstRec.SubItems(12)) - AA)
        End If
        txtChequeNo = lstRec.SubItems(13)
        txtReceiptRemark = lstRec.SubItems(14)
    End If
    CmdSave.Enabled = False
    CmdEdit.Visible = True
    CmdEdit.Enabled = False
    CmdUpdate.Visible = False
    FrameReceiptPV.Enabled = False
    'TxtReceiptNo.Enabled = False
    FrameCas.Enabled = False
    If CDbl(lblSelectedAmt) <> 0 Then
        FrameReceiptPV.Enabled = True
        'TxtReceiptNo.Enabled = True
        CmdSave.Enabled = True
    Else
        If Trim(TxtReceiptNo) <> "" And Trim(txtStatus) = "F" Then
            CmdEdit.Visible = True
            CmdUpdate.Visible = False
            CmdEdit.Enabled = True
        End If
    End If
End Sub

Private Sub UPdateWholeRecInfo()
Dim strRec2 As String
Dim tmpReceiptAmt As Double
    tmpReceiptAmt = IIf(IsNumeric(txtReceiptAmt), txtReceiptAmt, 0)
    strRec2 = "UPdate ReceiptfrMNRB SET ReceiptDate='" & Format(txtReceiptDate, "yyyy/mm/dd") & _
             "', ReceiptAmt=" & tmpReceiptAmt & ", ChequeNo='" & Trim(txtChequeNo) & _
             "', MNRBLastUpdateDate='" & Format(Date, "YYYY/MM/DD") & _
             "', MNRBLastUpdateUser='" & Logon & "', RECMNRBAmt=" & tmpReceiptAmt & _
             ", PAYPVNo='" & Trim(txtPVNo) & "', PAYPVDate='" & Format(txtPVDate, "yyyy/mm/dd") & _
             "', RECAccPeriodMth='" & Trim(cboAccPeriodMth) & "', RECAccPeriodYR='" & Trim(cboAccPeriodYR) & _
             "', MNRBRemarks='" & Trim(txtReceiptRemark) & _
             "' WHERE ReceiptNo= '" & Trim(TxtReceiptNo) & "'"
    medixmasterconnPayment.Execute strRec2
End Sub




