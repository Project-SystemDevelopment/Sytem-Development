VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCLMFloatReceipt 
   Caption         =   "Claim Float Receipt"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11715
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   80
      TabIndex        =   46
      Top             =   240
      Width           =   11600
      Begin VB.TextBox txtFTNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   54
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txtFTNofr 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   53
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txtTUNoFr 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   52
         Top             =   435
         Width           =   1335
      End
      Begin VB.TextBox txtTUNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   51
         Top             =   435
         Width           =   1335
      End
      Begin VB.TextBox txtPVNoFr 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   50
         Top             =   655
         Width           =   1335
      End
      Begin VB.TextBox txtPVNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   49
         Top             =   655
         Width           =   1335
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   915
         Width           =   3255
      End
      Begin VB.ComboBox CboTopupStatus 
         Height          =   315
         Left            =   1800
         TabIndex        =   47
         Top             =   1200
         Width           =   3255
      End
      Begin MSMask.MaskEdBox txtFTDateFr 
         Height          =   315
         Left            =   7920
         TabIndex        =   55
         Top             =   195
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
         Left            =   9840
         TabIndex        =   56
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTUDateFr 
         Height          =   315
         Left            =   7920
         TabIndex        =   57
         Top             =   480
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
         Left            =   9840
         TabIndex        =   58
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
         Left            =   7920
         TabIndex        =   59
         Top             =   720
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
         Left            =   9840
         TabIndex        =   60
         Top             =   740
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
         Left            =   9840
         TabIndex        =   76
         Top             =   1010
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRecDateFr 
         Height          =   315
         Left            =   7920
         TabIndex        =   75
         Top             =   1010
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
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
         Left            =   9360
         TabIndex        =   78
         Top             =   1080
         Width           =   375
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
         Left            =   6360
         TabIndex        =   77
         Top             =   1080
         Width           =   1455
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
         Left            =   6360
         TabIndex        =   74
         Top             =   500
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
         Left            =   9360
         TabIndex        =   73
         Top             =   500
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
         Left            =   6360
         TabIndex        =   72
         Top             =   780
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
         Left            =   9360
         TabIndex        =   71
         Top             =   780
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
         Left            =   240
         TabIndex        =   70
         Top             =   450
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
         Left            =   3285
         TabIndex        =   69
         Top             =   435
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
         Left            =   240
         TabIndex        =   68
         Top             =   195
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
         Left            =   3285
         TabIndex        =   67
         Top             =   195
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
         Left            =   240
         TabIndex        =   66
         Top             =   680
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
         Left            =   3285
         TabIndex        =   65
         Top             =   660
         Width           =   375
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
         Left            =   240
         TabIndex        =   64
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Req Status"
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
         Left            =   240
         TabIndex        =   63
         Top             =   1200
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
         Index           =   0
         Left            =   9360
         TabIndex        =   62
         Top             =   240
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
         Left            =   6360
         TabIndex        =   61
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame FrameReceiptPV 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   11505
      Begin VB.TextBox txtFTLNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1125
         MaxLength       =   15
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   240
         Width           =   3735
      End
      Begin VB.ComboBox cboAccPeriodYR 
         Height          =   315
         ItemData        =   "frmCLMFloatReceipt.frx":0000
         Left            =   9720
         List            =   "frmCLMFloatReceipt.frx":0088
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtPVNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboAccPeriodMth 
         Height          =   315
         ItemData        =   "frmCLMFloatReceipt.frx":0194
         Left            =   9960
         List            =   "frmCLMFloatReceipt.frx":01BC
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtchqNo 
         Height          =   285
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtReceiptNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         MaxLength       =   18
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtFTLDate 
         Height          =   285
         Left            =   1125
         TabIndex        =   10
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   285
         Left            =   5685
         TabIndex        =   11
         Top             =   255
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtchqDate 
         Height          =   285
         Left            =   5685
         TabIndex        =   12
         Top             =   495
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRECDate 
         Height          =   285
         Left            =   5685
         TabIndex        =   13
         Top             =   750
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtFTRecAmt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1125
         MaxLength       =   10
         TabIndex        =   7
         Top             =   710
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Chq No"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Chq Date"
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   24
         Top             =   510
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "A/C Year"
         Height          =   255
         Left            =   10080
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "A/C Mth"
         Height          =   255
         Left            =   10200
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "PV Date"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "PV No"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Rec Date"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   19
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "FTL No"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "FTL Amt"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
         Height          =   255
         Left            =   7080
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Rec No"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "FTL Date"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CheckBox CheckAll 
      Caption         =   "Check All"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   6960
      Width           =   11535
      Begin VB.TextBox lbltotalAmt 
         Height          =   285
         Left            =   3720
         TabIndex        =   81
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox lblSelectedAmt 
         Height          =   285
         Left            =   3720
         TabIndex        =   80
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox lblcurrent 
         Height          =   375
         Left            =   960
         TabIndex        =   79
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   7200
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S&earch"
         Height          =   330
         Left            =   7920
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   8640
         TabIndex        =   34
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10800
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10200
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Height          =   330
         Left            =   7200
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "&Print to File"
         Height          =   330
         Left            =   9240
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   7200
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Bordx Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   39
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstFTListing 
      Height          =   3600
      Left            =   120
      TabIndex        =   45
      Top             =   2280
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6350
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Receipt No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Receipt Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "TU No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "TU Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Req Amt"
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
         Text            =   "FT Amount"
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
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Payor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Chq No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Chq Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claim Float Receipt"
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
      TabIndex        =   44
      Top             =   0
      Width           =   11805
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
      TabIndex        =   43
      Top             =   7920
      Visible         =   0   'False
      Width           =   2415
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
      TabIndex        =   42
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
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
      TabIndex        =   41
      Top             =   7920
      Visible         =   0   'False
      Width           =   2295
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
      TabIndex        =   40
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmCLMFloatReceipt"
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
Dim tmpCLMFTNo As String
Dim m_blnDirection(1 To 22) As Boolean

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

    LstCnt = lstFTListing.ListItems.Count
        tmpchkItems = 0
        txtFTRecAmt = CDbl(0)
        lblSelectedAmt = CDbl(0)
                
        If LstCnt = 0 Then
            CheckAll.Value = 0
        Else
            If CheckAll.Value = 1 Then
                For LstCnt = 1 To LstCnt
                    If lstFTListing.ListItems(LstCnt).Checked = False And lstFTListing.ListItems(LstCnt).Text = "Received" Then
                        'MsgBox "This item has been updated! ", vbCritical
                        counttransfered = counttransfered + 1
                    ElseIf lstFTListing.ListItems(LstCnt).Checked = False Then
                        lstFTListing.ListItems(LstCnt).Checked = True
                        lstFTListing.ListItems(LstCnt).Text = "Transfer"
                        
                        txtFTRecAmt = CDbl(txtFTRecAmt) + CDbl(lstFTListing.ListItems(LstCnt).SubItems(8))
                        lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(lstFTListing.ListItems(LstCnt).SubItems(8))
                        tmpchkItems = tmpchkItems + 1
                        
                                             
                    ElseIf lstFTListing.ListItems(LstCnt).Checked = True Then
                        lstFTListing.ListItems(LstCnt).Checked = False
                        lstFTListing.ListItems(LstCnt).Text = ""
                        'tmpchkItems = tmpchkItems - 1
                    End If
                Next LstCnt
            Else
                For LstCnt = 1 To LstCnt
                    If lstFTListing.ListItems(LstCnt).Checked = True Then
                        lstFTListing.ListItems(LstCnt).Checked = False
                        lstFTListing.ListItems(LstCnt).Text = ""
                    End If
                Next LstCnt
                 tmpchkItems = 0
                txtFTRecAmt = CDbl(0)
                lblSelectedAmt = CDbl(0)
            End If
            
            If lstFTListing.ListItems.Count = counttransfered And counttransfered <> 0 Then
                MsgBox "All items have been transferred! ", vbCritical
            ElseIf lstFTListing.ListItems.Count <> counttransfered And counttransfered <> 0 Then
                MsgBox "Some ( " + counttransfered + " )item have been transferred! ", vbCritical
            End If
            
        End If

End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lstFTListing.ListItems.Clear
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
    
    If lstFTListing.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstFTListing)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault

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
Dim strsqlCLM As String
Dim rst As New ADODB.Recordset

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
    tmpchkPAYCode = ""
        AA = lstFTListing.ListItems.Count
        For BB = 1 To lstFTListing.ListItems.Count
            If lstFTListing.ListItems(BB).Checked = True Then
                If ItmChk = False Then
                    ItmChk = True
                End If
                    If tmpchkPAYCode <> "" Then
                        If tmpchkPAYCode <> Trim(IIf(Mid(lstFTListing.ListItems(BB).SubItems(12), 1, 3) = "FTG", "XG", "XL")) Then
                            DiffPay = True
                            Exit For
                        End If
                    End If
                    tmpchkPAYCode = lstFTListing.ListItems(BB).SubItems(7)
                    tmpPayCode = Trim(IIf(Mid(lstFTListing.ListItems(BB).SubItems(7), 1, 3) = "FTG", "XG", "XL"))
            End If
        Next BB

   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        Call SearchFloatAmt("And FTPayCode = '" & Trim(tmpPayCode) & "'")
        tmpFloatBal = CDbl(tmpChkFloatAmt) - CDbl(txtFTRecAmt)
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    
    If AA = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf ItmChk = False Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    'ElseIf txtFTLDate = "__/__/____" Then
    '    MsgBox "Please Enter FT Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtFTLDate.SetFocus
    ElseIf Trim(txtchqNo) = "" Then
       MsgBox "Please Enter Cheque No!", vbCritical
    '   txtchqNo.SetFocus
    ElseIf txtchqDate = "__/__/____" Then
        MsgBox "Please Enter Cheque Date", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtchqDate.SetFocus
    'ElseIf CDbl(txtFTRecAmt) = "" Then
    '    MsgBox "Please Transfer Amount!", vbCritical
    '    txtReceiptRemark.SetFocus
    ElseIf Trim(txtRECDate) = "__/__/____" Then
        MsgBox "Please Enter Receipt Date ! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtRECDate.SetFocus
    'ElseIf Trim(txtPVDate) = "__/__/____" Then
    '    MsgBox "Please Enter PV Date ! ", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtPVDate.SetFocus
    ElseIf DiffPay = True Then
        MsgBox "Different Payors for this transaction is not allowed! ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf CDbl(lblSelectedAmt) <> CDbl(txtFTRecAmt) Then
        MsgBox "Please check Float Transfer Amount! ", vbOKOnly + vbCritical, DialogBoxTitle
    'ElseIf CDbl(tmpFloatBal) <= 0 Then
    '    MsgBox "Insufficient Float! ", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Screen.MousePointer = vbHourglass
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
       ' medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

        startTrans = True
        medixmasterconnWS.beginTrans
        medixmasterconnPayment.beginTrans
        medixmasterconn.beginTrans
        
            tmpRecNo = GenerateSRecNo("RC")
            'tmpFTRecNo = GenerateFTRecNo(tmpPayCode)
            'tmpFTPVNo = GenerateFTPVNo(tmpPayCode)
            
            For BB = 1 To lstFTListing.ListItems.Count
                If lstFTListing.ListItems(BB).Checked = True Then
                    tmpCLMFTNo = lstFTListing.ListItems(BB).SubItems(6)
                    Call SaveFTFloat
                    
                    strsqlCLM = " Select  BordxRefNo, BordxDate, FLTRecNo, FLTRecDate," & _
                                " FLTFTNo, FLTFTDate, FLTTUNo, FLTTUDate, FLTPVNo, FLTPVDate," & _
                                " CLMCHQNo, CLMCHQDate, CASNumber,ClaimVersion, CLMBordxAmt from View_Float_Receipt_Update " & _
                                " Where FLTFTNo = '" & Trim(lstFTListing.ListItems(BB).SubItems(6)) & "'"
                    rst.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                    
                    Do While Not rst.EOF
                        tmpRecAmt = CDbl(lstFTListing.ListItems(BB).SubItems(5))
                        tmpCASNumber = rst!CasNumber
                        tmpVersion = rst!claimversion
                        tmpChequeNo = rst!CLMCHQNo
                        tmpBordxNo = rst!BordxRefNo
                        tmpcheckAmt = rst!CLMBordxAmt
                        Call SaveFTBordx(tmpBordxNo)
                        Call SaveReceiptPayor
                        If Len(tmpCASNumber) Then
                            AddRemark = ""
                            AddRemark = vbNewLine & Date & "   Claim Float Transfer  -  " & tmpRecNo & " on " & txtRECDate & "                                         " & Logon & " "
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
                                strsqlREM = " update CASRemarkThree set " & _
                                         " CASRemarksFive = '" & Remarks7 & "'" & _
                                         " where CASNumber = '" & Trim(tmpCASNumber) & "'"
                            ElseIf (RemLen6 <> 0 And RemLen6 < 3700) Or RemLen5 > 3700 Then
                                strsqlREM = " update CASRemarkThree set " & _
                                         " CASRemarksFour = '" & Remarks6 & "'" & _
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
                    rst.MoveNext
                    Loop
                End If
            Next BB
            'Call UpdateFloatAmt
            medixmasterconnWS.commitTrans
            medixmasterconnPayment.commitTrans
            medixmasterconn.commitTrans
            startTrans = False
            'txtFTLNo = tmpFTRecNo
            TxtReceiptNo = tmpRecNo
            'txtPVNo = tmpFTPVNo
                        

            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            'If tmpFTFloatBLAmt < 300000 Then
            '    MsgBox "Float Balance is " & tmpFTFloatBLAmt & " for " & tmpPayCode, vbOKOnly + vbInformation, DialogBoxTitle
            'End If
            tmpSql = ""
            tmpSql = " and  CLMRecNo = '" & Trim(TxtReceiptNo) & "'"
            lstFTListing.ListItems.Clear
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
    
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
  ''  Set medixmasterconnMaint = Nothing
  '  MediConnCISDB.Close
  '  Set MediConnCISDB = Nothing
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
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Call SearchRecords
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    Screen.MousePointer = vbDefault
    CmdSearch.Enabled = True
    CmdEdit.Visible = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    Exit Sub

errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstFTListing.ListItems.Clear
    lblcurrent = ""
    lbltotalAmt = ""
    lblSelectedAmt = ""
    'cmdprint.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    CmdEdit.Visible = True
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
End Sub

Private Function SearchRecords()
Dim sql As String
Dim strAppend As String
Dim tmpPayCode As String
    strAppend = ""
    'GrpCompany = ""
    'HOspitals = ""
    
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
            
    If Len(Trim(txtPVNoFr)) Then
        strAppend = strAppend + " AND   CLMPVNo >= '" & Trim(txtPVNoFr) & "'"
    End If
    If Len(Trim(txtPVNoTo)) Then
        strAppend = strAppend + " AND   CLMPVNo <= '" & Trim(txtPVNoTo) & "'"
    End If
    
    If Len(Trim(CboPayor)) Then
        If Mid(CboPayor, 1, 2) = "XL" Then
            tmpPayCode = "TUL"
        Else
            tmpPayCode = "TUG"
        End If
        strAppend = strAppend + " AND  Substring(CLMTUNo,1,3) = '" & Trim(tmpPayCode) & "'"
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
    
    If Len(Trim(txtRECDateFr)) > 0 And IsDate(txtRECDateFr) = True Then
        strAppend = strAppend + " And CLMRecDate >= '" & Format(txtRECDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRECDateTo)) > 0 And IsDate(txtRECDateTo) = True Then
        strAppend = strAppend + " And CLMRecDate <= '" & Format(txtRECDateTo, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(CboTopupStatus)) Then
        If Mid(CboTopupStatus, 1, 1) = "0" Then
            strAppend = strAppend + " AND  (CLMRecNo Is null OR CLMRecNo = '') "
        ElseIf Mid(CboTopupStatus, 1, 1) = "1" Then
            strAppend = strAppend + " AND  (CLMRecNo Is NOT null) "
        End If
    End If
    
    strAppend = strAppend + " AND  (CLMTUNo is not null or CLMTUNo <> '')"
    
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
        lstFTListing.ListItems.Clear
        lblcurrent = 0
        lblSelectedAmt = 0
        lbltotalAmt = 0
    'End If
        strAppendCase = 1
    strsql = " Select  CLMFTNo,CLMFTDate,CLMFTAmt, CLMPVNo, CLMPVDate,CLMTUAmt,CLMCHQNo,CLMCHQDate,  " & _
                "CLMRecNo, CLMRecDate,CLMTUNo,CLMTUDate,CLMTURemarks from CLMFloatTransfer Where 0=0 "
    strsql = strsql + strAppend
    rst2.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    Do
        Do Until rst2.EOF

        RecFound = True
        
        With rst2
            If Len(Trim(!CLMRecNo)) Then
                Set Record = lstFTListing.ListItems.Add(, , "Received")
            Else
                Set Record = lstFTListing.ListItems.Add(, , "")
            End If
            Record.SubItems(1) = IIf(Len(!CLMRecNo), !CLMRecNo, "")
            Record.SubItems(2) = IIf(Len(!CLMRecDate), Format(!CLMRecDate, "dd/mm/yyyy"), "")
            Record.SubItems(3) = IIf(Len(!CLMTUNo), !CLMTUNo, "")
            Record.SubItems(4) = IIf(Len(!CLMTUDate), Format(!CLMTUDate, "dd/mm/yyyy"), "")
            Record.SubItems(5) = Format(IIf(Len(!CLMTUAmt), !CLMTUAmt, 0), "#,##0.00")
            Record.SubItems(6) = IIf(Len(!CLMFTNo), !CLMFTNo, "")
            Record.SubItems(7) = IIf(Len(!CLMFTDate), Format(!CLMFTDate, "dd/mm/yyyy"), "")
            Record.SubItems(8) = Format(IIf(Len(!CLMFTAmt), Trim(!CLMFTAmt), 0), "#,##0.00")
            Record.SubItems(9) = IIf(Trim(Len(!CLMPVNo)), !CLMPVNo, "")
            Record.SubItems(10) = IIf(Trim(Len(!CLMPVDate)), Format(!CLMPVDate, "dd/mm/yyyy"), "")
            Record.SubItems(11) = IIf(Len(!CLMTURemarks), !CLMTURemarks, "")
            Record.SubItems(12) = IIf(Len(!CLMFTNo), Mid(!CLMFTNo, 1, 3), "")
            Record.SubItems(13) = IIf(Len(!CLMCHQNo), !CLMCHQNo, "")
            Record.SubItems(14) = IIf(Len(!CLMCHQDate), !CLMCHQDate, "")
            lbltotalAmt = CDbl(lbltotalAmt) + Format(CDbl(!CLMFTAmt), "#,##0.00")
            Current = Current + 1
            lblcurrent = Current
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
    LstCnt = lstFTListing.ListItems.Count

        'If lstCnt = 0 Then
        '    chkCheckItems.value = 0
        'Else

            For LstCnt = 1 To LstCnt
                If lstFTListing.ListItems(LstCnt).Text <> "Received" Then
                    If lstFTListing.ListItems(LstCnt).Checked = False Then
                        lstFTListing.ListItems(LstCnt).Checked = True
                        lstFTListing.ListItems(LstCnt).Text = "Receive"
                    ElseIf lstFTListing.ListItems(LstCnt).Checked = True Then
                        lstFTListing.ListItems(LstCnt).Checked = False
                        lstFTListing.ListItems(LstCnt).Text = ""
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
    
        AA = lstFTListing.ListItems.Count
        'tmpPayCode = lstFTListing.SelectedItem.SubItems(20)
        
     '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            'Call SearchFloatAmt("And FTPayCode = '" & Trim(tmpPayCode) & "'")
            'Call SearchFloatAmt("FTPayCode = '" & Trim(tmpPAYCode) & "'")
            'tmpFloatBal = CDbl(tmpChkFloatAmt) + CDbl(tmpPrevFloatAmt) - CDbl(txtFTRecAmt)
    '    medixmasterconnWS.Close
     '   Set medixmasterconnWS = Nothing
        
        
    If AA = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    'ElseIf ItmChk = False Then
    '    MsgBox "Please select at least one record to be updated! ", vbCritical
    'ElseIf txtFTLDate = "__/__/____" Then
    '    MsgBox "Please Enter FT Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtFTLDate.SetFocus
    ElseIf Trim(txtchqNo) = "" Then
       MsgBox "Please Enter Cheque No!", vbCritical
       
    ElseIf txtchqDate = "__/__/____" Then
        MsgBox "Please Enter Cheque Date", vbOKOnly + vbCritical, DialogBoxTitle
       
    '   txtChequeNo.SetFocus
    'ElseIf txtFTRecAmt = "" Then
    '    MsgBox "Please Transfer Amount!", vbCritical
    '    txtFTRecAmt.SetFocus
    ElseIf Trim(txtRECDate) = "__/__/____" Then
        MsgBox "Please Enter Receipt Date ! ", vbOKOnly + vbCritical, DialogBoxTitle
        TxtReceiptNo.SetFocus
    'ElseIf Trim(txtPVDate) = "__/__/____" Then
    '    MsgBox "Please Enter PV Date ! ", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtPVDate.SetFocus
    Else
        Screen.MousePointer = vbHourglass
    '    medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        startTrans = True
        medixmasterconnWS.beginTrans

            'Call UpdateFTFloat
            'For BB = 1 To lstFTListing.listitems.Count
                'If lstFTListing.listitems(BB).Checked = True Then
                    tmpBordxNo = lstFTListing.SelectedItem.SubItems(1)
                    'tmpPayCode = lstFTListing.SelectedItem.SubItems(20)
                    'Call UpdateFTBordx(tmpBordxNo)
                'End If
            'Next BB
            'Call UpdateFloatAmt
            medixmasterconnWS.commitTrans
            startTrans = False

            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            'If tmpFloatBal < 300000 Then
            '    MsgBox "Float Balance is " & tmpFloatBal & " for " & tmpPayCode, vbOKOnly + vbInformation, DialogBoxTitle
            'End If
            tmpSql = " and  CLMRecNo = '" & Trim(TxtReceiptNo) & "'"
            lstFTListing.ListItems.Clear
            Call CLMFloatListing(tmpSql)
        EditFlag = False
        CmdEdit.Visible = True
      '  medixmasterconnWS.Close
      '  Set medixmasterconnWS = Nothing
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
    
 '   medixmasterconn.Close
 '   Set medixmasterconn = Nothing
 '   medixmasterconnMaint.Close
 '   Set medixmasterconnMaint = Nothing
 '   MediConnCISDB.Close
'    Set MediConnCISDB = Nothing
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
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        EditFlag = False
        Call ComboListing
        txtFTRecAmt = 0
        lblSelectedAmt = 0
        lbltotalAmt = 0
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

Private Sub lstFTListing_Click()
    If lstFTListing.SelectedItem.Text = "Received" Then
        tmpPrevFloatAmt = lstFTListing.SelectedItem.SubItems(8)
        Call LvDisplayData
    End If
End Sub

Private Sub lstFTListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstFTListing, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstFTListing, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
    Case 14
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(15)
    Case 16
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
    Case 18
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstFTListing, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
    End Select

End Sub

Private Sub lstFTListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Text <> "Received" Then
        If Item.Checked = True And Trim(Item.SubItems(3)) = "" Then
            MsgBox "Please check Top-up Request Number!", vbOKOnly + vbCritical, DialogBoxTitle
            Item.Checked = False
        ElseIf Item.Checked = True Then
            Item.Text = "Receive"
            txtFTRecAmt = CDbl(txtFTRecAmt) + CDbl(Item.SubItems(8))
            lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(Item.SubItems(8))
            tmpchkItems = tmpchkItems + 1
        ElseIf Item.Checked = False Then
            Item.Text = ""
            If txtFTRecAmt > 0 Then
                txtFTRecAmt = CDbl(txtFTRecAmt) - CDbl(Item.SubItems(8))
                lblSelectedAmt = CDbl(lblSelectedAmt) - CDbl(Item.SubItems(8))
                tmpchkItems = tmpchkItems - 1
            End If
        End If
    Else
        Item.Checked = False
        MsgBox "This Item has been updated", vbOKOnly + vbCritical, DialogBoxTitle
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

    With CboTopupStatus
        .AddItem "0 - Not Received"
        .AddItem "1 - Received"
        .Text = "0 - Not Received"
    End With
    
    'With cboCashType
    '    .AddItem "H - Hospital"
    '    .AddItem "M - Member"
    'End With
    
End Sub

Private Sub SaveFTFloat()
Dim strsqlFT As String
Dim tmpFTLDate As String
Dim tmpPVDate As String
Dim tmpRecDate As String
Dim tmpDate As String
Dim tmpChqDate As String
Dim rst As New ADODB.Recordset

    tmpRecDate = Format(txtRECDate, "yyyy/mm/dd")
    
    strsqlFT = " Select  CLMRecNo,CLMRecDate from CLMFloatTransfer Where CLMFTNo = '" & Trim(tmpCLMFTNo) & "'"
    rst.Open strsqlFT, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                    
        With rst
            !CLMRecNo = tmpRecNo
            !CLMRecDate = Format(txtRECDate, "dd/mm/yyyy")
        End With
        rst.Update
        
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub SaveFTBordx(tmpBordxNo As String)
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
Dim strsqlCLM As String
Dim savedstatus As Boolean
    strsqlCLM = ""
    strsqlCLM = " Select  BordxRefNo,BordxDate, FLTFTNo,FLTFTDate, " & _
                "FLTPVNo, FLTPVDate,  FLTRecNo, FLTRecDate,BordxPayFrStatus from Bordx " & _
                " Where BordxRefNo = '" & Trim(tmpBordxNo) & "'"
    rst.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                    
        With rst
            !BordxPAYFRStatus = "F"
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
                !CLMRecNo = TxtReceiptNo
                !CLMRecDate = Format(txtRECDate, "dd/mm/yyyy")
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
                "FLTPVNo, FLTPVDate,  FLTRecNo, FLTRecDate, BordxPayFrStatus from Bordx " & _
                " Where BordxRefNo = '" & Trim(tmpBordxNo) & "'"
    rst.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                    
        With rst
            !BordxPAYFRStatus = "F"
        End With
        rst.Update
        
        rst.Close
        Set rst = Nothing
End Sub

Private Sub UpdateFloatAmt()
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
Dim tmpFTFloatAmt As Double
    strsqlFT = " Select FTPayCode, FTFloatAmt from CLMFloatAmt Where FTPayCode = '" & Trim(tmpPayCode) & "'"
    rst.Open strsqlFT, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        tmpFTFloatAmt = rst!FTFloatAmt
        If EditFlag = False Then
            tmpFTFloatBLAmt = CDbl(tmpFTFloatAmt) - CDbl(txtFTRecAmt)
        Else
            tmpFTFloatBLAmt = (CDbl(tmpFTFloatAmt) + CDbl(tmpPrevFloatAmt)) - CDbl(txtFTRecAmt)
        End If
        
        rst!FTFloatAmt = tmpFTFloatBLAmt
        'tmpFTFloatBLAmt = tmpFTFloatBLAmt
        rst.Update
        
        If tmpPayCode = "XG" Then
            lblMGAmt = tmpFTFloatBLAmt
        Else
            lblMLAmt = tmpFTFloatBLAmt
        End If
    rst.Close
    Set rst = Nothing

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

Private Sub lstFTListing_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstFTListing.SelectedItem.Text = "Received" Then
        tmpPrevFloatAmt = lstFTListing.SelectedItem.SubItems(8)
        Call LvDisplayData
    End If
End Sub

Private Sub lstFTListing_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstFTListing.SelectedItem.Text = "Received" Then
        tmpPrevFloatAmt = lstFTListing.SelectedItem.SubItems(8)
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

Private Sub LvDisplayData(Optional Item As ListItem, Optional ItemCheck As Boolean)
Dim lstRec As ListItem

    Call ClearFrameCase
    If lstFTListing.ListItems.Count Then
        If ItemCheck = True Then
            Set lstRec = Item
            lstFTListing.SelectedItem.Selected = False
            lstRec.Selected = True
        Else
            Set lstRec = lstFTListing.SelectedItem
        End If
        
        TxtReceiptNo = IIf(Len(Trim(lstRec.SubItems(1))), lstRec.SubItems(1), "")
        txtRECDate = IIf(Len(Trim(lstRec.SubItems(2))), lstRec.SubItems(2), "")
        
        'txtTUNo = IIf(Len(Trim(lstRec.SubItems(3))), lstRec.SubItems(3), "")
        'If Len(lstRec.SubItems(4)) Then
        '    txtTUDate = Format(lstRec.SubItems(4), "DD/MM/YYYY")
        'End If
        If Len(lstRec.SubItems(5)) Then
            txtFTRecAmt = Format(lstRec.SubItems(5), "#,##0.00")
        End If
        txtFTLNo = IIf(Len(Trim(lstRec.SubItems(6))), lstRec.SubItems(6), "")
        txtFTLDate = IIf(Len(Trim(lstRec.SubItems(7))), lstRec.SubItems(7), "")
        txtPVNo = IIf(Len(Trim(lstRec.SubItems(9))), lstRec.SubItems(9), "")
        txtPVDate = IIf(Len(Trim(lstRec.SubItems(10))), lstRec.SubItems(10), "")
        TxtRemark = IIf(Len(lstRec.SubItems(11)), lstRec.SubItems(11), "")
        txtchqNo = IIf(Len(lstRec.SubItems(13)), lstRec.SubItems(13), "")
        txtchqDate = IIf(Len(lstRec.SubItems(14)), lstRec.SubItems(14), "")
        
        CmdSave.Enabled = False
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



Private Sub SaveReceiptPayor()
Dim strsqlREC As String
    strsqlREC = "Insert Into ReceiptfrPay(CasNumber, ClaimVersion, ReceiptNo," & _
         "ReceiptDate, ReceiptAmt, ChequeNo, BordxNo, " & _
         "PayLastUpdateDate, PayLastUpdateUser," & _
         "RECPayAmt, CheckAmt) " & _
         "Values('" & Trim(tmpCASNumber) & "','" & Trim(tmpVersion) & "','" & Trim(tmpRecNo) & _
         "', '" & Format(txtRECDate, "yyyy/mm/dd") & "'," & tmpRecAmt & ",'" & Trim(txtchqNo) & "','" & tmpBordxNo & _
         "', '" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'," & tmpRecAmt & "," & tmpcheckAmt & ")"
    medixmasterconnPayment.Execute strsqlREC
    
    'strsqlREC = "Insert Into ReceiptfrPay(CasNumber, ClaimVersion," & _
    '     " ReceiptAmt, ChequeNo, BordxNo, " & _
    '     "PayLastUpdateDate, PayLastUpdateUser," & _
    '     "RECPayAmt, CheckAmt) " & _
    '     "Values('" & Trim(tmpCASNumber) & "','" & Trim(tmpVersion) & _
    '     "', " & tmpRecAmt & ",'" & Trim(tmpChequeNo) & "','" & tmpBordxNo & _
    '     "', '" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "'," & tmpRecAmt & "," & tmpcheckAmt & ")"
    'medixmasterconnPayment.Execute strsqlREC

End Sub


