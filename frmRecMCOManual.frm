VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecMCOManual 
   Caption         =   "Receipt MCO/Others Manual"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   0
      TabIndex        =   43
      Top             =   6120
      Width           =   7815
      Begin VB.ComboBox cboACCType 
         Height          =   315
         Left            =   6240
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtChequeNo 
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
         Left            =   3960
         MaxLength       =   20
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtReceiptNo 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtRecRemarks 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   6240
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   530
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   495
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtChequeDate 
         Height          =   285
         Left            =   3960
         TabIndex        =   18
         Top             =   495
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtAmt 
         Height          =   285
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   16
         Top             =   750
         Width           =   1455
      End
      Begin VB.TextBox txtAccCode 
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
         Left            =   3960
         MaxLength       =   20
         TabIndex        =   19
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Acc Type"
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Acc Code"
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
         Index           =   2
         Left            =   3120
         TabIndex        =   50
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Chq. No"
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
         Index           =   1
         Left            =   3120
         TabIndex        =   49
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Chq. Date"
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
         Left            =   3120
         TabIndex        =   48
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   5400
         TabIndex        =   47
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Receipt Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Receipt Date"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   520
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Receipt No"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   250
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   0
      TabIndex        =   28
      Top             =   240
      Width           =   12135
      Begin VB.CheckBox chkCRNote 
         Caption         =   "Credit Note Search"
         Height          =   255
         Left            =   10320
         TabIndex        =   13
         Top             =   670
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   195
         Left            =   3360
         TabIndex        =   30
         Top             =   780
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.ComboBox CboRecType 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   180
         Width           =   3255
      End
      Begin VB.CheckBox chkRecpDate 
         Height          =   255
         Left            =   10320
         TabIndex        =   10
         Top             =   400
         Width           =   255
      End
      Begin VB.CheckBox ChkRecpNo 
         Height          =   255
         Left            =   10320
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   470
         Width           =   3255
      End
      Begin VB.TextBox txtRECNoFr 
         Height          =   285
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   5
         Top             =   135
         Width           =   1095
      End
      Begin VB.TextBox txtRECNoTo 
         Height          =   285
         Left            =   9120
         MaxLength       =   20
         TabIndex        =   6
         Top             =   135
         Width           =   1095
      End
      Begin VB.TextBox TxtInvoiceNo1 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   2
         Top             =   745
         Width           =   1455
      End
      Begin VB.TextBox TxtInvoiceNo2 
         Height          =   285
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   29
         Top             =   745
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtRECDateFr 
         Height          =   315
         Left            =   7560
         TabIndex        =   8
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRECDateTo 
         Height          =   315
         Left            =   9120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   3720
         TabIndex        =   4
         Top             =   1000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCRNoFr 
         Height          =   285
         Left            =   7560
         MaxLength       =   20
         TabIndex        =   11
         Top             =   650
         Width           =   1095
      End
      Begin VB.TextBox txtCRNoTo 
         Height          =   285
         Left            =   9120
         MaxLength       =   20
         TabIndex        =   12
         Top             =   650
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "CR No fr"
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
         Index           =   1
         Left            =   6360
         TabIndex        =   42
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "To"
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
         Index           =   1
         Left            =   8760
         TabIndex        =   41
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Left            =   8760
         TabIndex        =   40
         Top             =   165
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   39
         Top             =   780
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Invoice No fr"
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
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Invoice Date fr"
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
         Left            =   240
         TabIndex        =   37
         Top             =   1035
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   36
         Top             =   1075
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Rec. Type"
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
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   200
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Recp No fr"
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
         Left            =   6360
         TabIndex        =   34
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Recp Date fr"
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
         Left            =   6360
         TabIndex        =   33
         Top             =   435
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "To"
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
         Left            =   8760
         TabIndex        =   32
         Top             =   435
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Payor"
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
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   7920
      TabIndex        =   25
      Top             =   6120
      Width           =   4215
      Begin VB.CommandButton cmdUpdateCLM 
         Caption         =   "Update Case"
         Height          =   330
         Left            =   2640
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtInvNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPaidAmt 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   23
         Top             =   500
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Invoice No"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   235
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Amt Receipt"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   495
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstReceiptMCO 
      Height          =   4455
      Left            =   0
      TabIndex        =   68
      Top             =   1680
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7858
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
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Invoice No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Invoice Date"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Invoice Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Payor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Receipt No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Receipt Date"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Total Receipt Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Receipt Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Receipt Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Remark"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ReceiptIndex"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Chq Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Chq Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstReceiptOther 
      Height          =   4335
      Left            =   0
      TabIndex        =   69
      Top             =   1680
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Receipt No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Receipt Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Receipt Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Receipt Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Chq No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Chq Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Acc Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Remark"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ReceiptIndex"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   0
      TabIndex        =   52
      Top             =   7200
      Width           =   12135
      Begin VB.TextBox lblchkAmt 
         Height          =   285
         Left            =   4800
         TabIndex        =   75
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtTotalAmt 
         Height          =   285
         Left            =   2640
         TabIndex        =   74
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox lblCurRecord 
         Height          =   285
         Left            =   120
         TabIndex        =   73
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   6000
         TabIndex        =   62
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9840
         TabIndex        =   61
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Height          =   330
         Left            =   7200
         TabIndex        =   60
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   9240
         TabIndex        =   59
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10800
         TabIndex        =   58
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   11400
         TabIndex        =   57
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   330
         Left            =   8520
         TabIndex        =   56
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   6600
         TabIndex        =   55
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   7800
         TabIndex        =   54
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Height          =   330
         Left            =   7320
         TabIndex        =   53
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "chkAmt"
         Height          =   255
         Left            =   4080
         TabIndex        =   67
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Total Amt"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   66
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   1200
         TabIndex        =   65
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   960
         TabIndex        =   64
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1200
         TabIndex        =   63
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Receipt MCO/Others"
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
      TabIndex        =   72
      Top             =   0
      Width           =   12165
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable"
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
      TabIndex        =   71
      Top             =   7770
      Width           =   1695
   End
   Begin VB.Label lblIndex 
      Caption         =   "Index"
      Height          =   255
      Left            =   1800
      TabIndex        =   70
      Top             =   7800
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmRecMCOManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim EditFlag As Boolean
Private m_blnDirection(1 To 10) As Boolean
Dim strcriteria As String
Dim tmpItemCheck As Boolean
Dim listloop As Integer
Dim check As Boolean

Private Sub CboRecType_Change()
    If CboRecType = "M - MCO FEES" Then
        TxtInvoiceNo1.Enabled = True
        TxtInvoiceNo2.Enabled = True
        txtInvDate1.Enabled = True
        txtINVDate2.Enabled = True
        lstReceiptMCO.Visible = True
        lstReceiptOther.Visible = False
        CboPayor.Enabled = True
        ChkScan.Enabled = True
    Else
        TxtInvoiceNo1 = ""
        TxtInvoiceNo1.Enabled = False
        TxtInvoiceNo2 = ""
        TxtInvoiceNo2.Enabled = False
        txtInvDate1 = "__/__/____"
        txtInvDate1.Enabled = False
        txtINVDate2 = "__/__/____"
        txtINVDate2.Enabled = False
        txtCRNoFr = ""
        txtCRNoTo = ""
        lstReceiptMCO.Visible = False
        lstReceiptOther.Visible = True
        CboPayor.Enabled = False
        ChkScan.Enabled = False
        ChkScan.Value = "0"
    End If
    If Mid(CboRecType, 1, 1) = "F" Then
        txtAccCode.Enabled = False
        txtAccCode = "6720"
    ElseIf Mid(CboRecType, 1, 1) = "C" Then
        txtAccCode.Enabled = False
        txtAccCode = "2040"
    ElseIf Mid(CboRecType, 1, 1) = "A" Or Mid(CboRecType, 1, 1) = "O" Or Mid(CboRecType, 1, 1) = "M" Then
        txtAccCode.Enabled = True
        txtAccCode = ""
    End If
End Sub

Private Sub CboRecType_Click()
    If CboRecType = "M - MCO FEES" Then
        TxtInvoiceNo1.Enabled = True
        TxtInvoiceNo2.Enabled = True
        txtInvDate1.Enabled = True
        txtINVDate2.Enabled = True
        txtCRNoFr.Enabled = True
        txtCRNoTo.Enabled = True
        lstReceiptMCO.Visible = True
        lstReceiptOther.Visible = False
        CboPayor.Enabled = True
        ChkScan.Enabled = True
    Else
        TxtInvoiceNo1 = ""
        TxtInvoiceNo1.Enabled = False
        TxtInvoiceNo2 = ""
        TxtInvoiceNo2.Enabled = False
        txtInvDate1 = "__/__/____"
        txtInvDate1.Enabled = False
        txtINVDate2 = "__/__/____"
        txtINVDate2.Enabled = False
        txtCRNoFr = ""
        txtCRNoTo = ""
        txtCRNoFr.Enabled = False
        txtCRNoTo.Enabled = False
        lstReceiptMCO.Visible = False
        lstReceiptOther.Visible = True
        CboPayor.Enabled = False
        ChkScan.Enabled = False
        ChkScan.Value = "0"
    End If
    
    If Mid(CboRecType, 1, 1) = "F" Then
        txtAccCode.Enabled = False
        txtAccCode = "6720"
    ElseIf Mid(CboRecType, 1, 1) = "C" Then
        txtAccCode.Enabled = False
        txtAccCode = "2040"
    ElseIf Mid(CboRecType, 1, 1) = "A" Or Mid(CboRecType, 1, 1) = "O" Or Mid(CboRecType, 1, 1) = "M" Then
        txtAccCode.Enabled = True
        txtAccCode = ""
    End If
End Sub

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        lstReceiptMCO.ListItems.Clear
        Label14.Visible = False
        TxtInvoiceNo2.Visible = False
        ChkScan.Left = 3360
        TxtInvoiceNo1.SetFocus
    Else
        lstReceiptMCO.ListItems.Clear
        Label14.Visible = True
        ChkScan.Left = 5160
        TxtInvoiceNo2.Visible = True
        TxtInvoiceNo1.SetFocus
    End If

End Sub

Private Sub cmdAdd_Click()
    Call ClearReceipt
    CmdSave.Enabled = True
End Sub

Private Sub cmdDelete_Click()
Dim rsDeleteOther As New ADODB.Recordset
Dim strsql As String
Dim DeleteRecord
    
    CmdDelete.Enabled = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    
    If CboRecType = "M - MCO FEES" Then
        If lstReceiptMCO.ListItems.Count = 0 Then
            MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
            If DeleteRecord = vbYes Then
                strsql = "Delete from RECMCO Where RecMCOIndex = '" & Trim(lstReceiptMCO.SelectedItem.SubItems(9)) & "'"
                rsDeleteOther.Open strsql, medixmasterconnPayment, adOpenDynamic, adLockOptimistic
                MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
                'Call ClearForm(Me)
                Call ReceiptMCOListing
            End If
        End If
    Else
        If lstReceiptOther.ListItems.Count = 0 Then
            MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
            If DeleteRecord = vbYes Then
                strsql = "Delete from RECOthers Where RecOthersIndex = '" & Trim(lstReceiptOther.SelectedItem.SubItems(5)) & "'"
                rsDeleteOther.Open strsql, medixmasterconnPayment, adOpenDynamic, adLockOptimistic
                MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
                'Call ClearForm(Me)
                Call ReceiptOthersListing
            End If
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    CmdDelete.Enabled = True
End Sub

Private Sub cmdUpdateCLM_Click()
Dim AskRec
Dim AmtPaid As Double

AmtPaid = IIf(IsNumeric(txtPaidAmt), txtPaidAmt, 0)

If lstReceiptMCO.ListItems.Count <> 0 Then
    If lstReceiptMCO.SelectedItem.Checked = True Then
        If (Trim(txtINVNo) <> "" And Trim(AmtPaid) = "") Then
            MsgBox "Please Enter Invoice Receipt Amount ", vbCritical + vbOKOnly, DialogBoxTitle
            TxtContraAmt.SetFocus
        'ElseIf CDbl(AmtPaid) > CDbl(lstReceiptMCO.SelectedItem.SubItems(8)) Then
        '    MsgBox "Please Check Amount Receipt! ", vbCritical + vbOKOnly, DialogBoxTitle
        Else
            AskRec = MsgBox("Update " & txtINVNo & " ?", vbYesNo + vbExclamation)
            If AskRec = vbYes Then
                If IsNumeric(AmtPaid) Then
                    lstReceiptMCO.SelectedItem.SubItems(8) = AmtPaid
                End If
                MsgBox "Updated .... ", vbInformation + vbOKOnly, DialogBoxTitle
            End If
            txtINVNo = ""
            txtPaidAmt = ""
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

Private Sub cmdClear_Click()
    CmdSave.Enabled = True
    CboPayor = ""
    txtRecNoFr = ""
    txtRecNoTo = ""
    txtRecDateFr = "__/__/____"
    txtRecDateTo = "__/__/____"
    ChkRecpNo.Value = 0
    chkRecpDate.Value = 0
    TxtInvoiceNo1 = ""
    TxtInvoiceNo2 = ""
    txtInvDate1 = "__/__/____"
    txtINVDate2 = "__/__/____"
    TxtReceiptNo = ""
    txtReceiptDate = "__/__/____"
    txtAmt = ""
    TxtRecRemarks = ""
    lstReceiptMCO.ListItems.Clear
    lstReceiptOther.ListItems.Clear
    txtTotalAmt = 0
    lblCurRecord = 0
    lblTotal = 0
End Sub

Private Sub CmdEdit_Click()
    EditFlag = True
    CmdUpdate.Enabled = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

'************************Start ComboListing**************************

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim RecType As New ADODB.Recordset
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
    
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchRecType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    RecType.CursorLocation = adUseClient
    RecType.CursorType = adOpenForwardOnly
    RecType.CacheSize = 100
    Set RecType = cmd.Execute
    
    Do Until RecType.EOF
        With RecType
            CboRecType.AddItem !RecTypeCode & " - " & !RecTypeDescription
            RecType.MoveNext
        End With
    Loop
        CboRecType.Text = "M - MCO FEES"
    
    RecType.Close
    Set RecType = Nothing
    
    
    
End Sub
'************************Start ComboListing**************************

Private Function SearchRecord()
Dim strsql As String
Dim StrAppend As String
Dim sql As String
Dim tmpInvNo As String

    StrAppend = ""
                
    If CboRecType = "M - MCO FEES" Then
        If Len(Trim(CboPayor)) Then
            StrAppend = StrAppend + " And PayCode = '" & Mid(CboPayor, 1, 2) & "'"
        End If
        
        If ChkScan.Value = "1" Then
            tmpInvNo = TxtInvoiceNo1
            TxtInvoiceNo1 = ""
            If Len(Trim(tmpInvNo)) Then
                StrAppend = StrAppend + " AND MBMInvoiceNo = '" & RemStr(Trim(tmpInvNo)) & "'"
            End If
        Else
            If Len(Trim(TxtInvoiceNo1)) Then
                StrAppend = StrAppend + " AND MBMInvoiceNo >= '" & RemStr(Trim(TxtInvoiceNo1)) & "'"
            End If
            
            If Len(Trim(TxtInvoiceNo2)) Then
                StrAppend = StrAppend + " AND MBMInvoiceNo <= '" & RemStr(Trim(TxtInvoiceNo2)) & "'"
            End If
        End If
        If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
            StrAppend = StrAppend + " And MBMInvoiceDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtINVDate2)) > 0 And IsDate(txtINVDate2) = True Then
            StrAppend = StrAppend + " And MBMInvoiceDate <= '" & Format(txtINVDate2, "mm/dd/yyyy") & "'"
        End If
        
    End If
    
    If Len(Trim(txtRecNoFr)) Then
        StrAppend = StrAppend + " AND ReceiptNo >= '" & RemStr(Trim(txtRecNoFr)) & "'"
    End If
    
    If Len(Trim(txtRecNoTo)) Then
        StrAppend = StrAppend + " AND ReceiptNo <= '" & RemStr(Trim(txtRecNoTo)) & "'"
    End If
    
    If ChkRecpNo.Value = 1 Then
        sql = ""
        sql = " AND (ReceiptNo Is null OR ReceiptNo = '')" & ""
        StrAppend = StrAppend + sql
    End If
    
    If Len(Trim(txtRecDateFr)) > 0 And IsDate(txtRecDateFr) = True Then
        StrAppend = StrAppend + " And ReceiptDate >= '" & Format(txtRecDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRecDateTo)) > 0 And IsDate(txtRecDateTo) = True Then
        StrAppend = StrAppend + " And ReceiptDate <= '" & Format(txtRecDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If chkRecpDate.Value = 1 Then
        sql = ""
        sql = " AND (ReceiptDate Is null OR ReceiptDate = '')" & ""
        StrAppend = StrAppend + sql
    End If
    
    If chkCRNote.Value = 0 Then
        If Len(Trim(txtCRNoFr)) > 0 And IsDate(txtCRNoFr) = True Then
            StrAppend = StrAppend + " And MBMCRNote >= '" & txtCRNoFr & "'"
        End If
        
        If Len(Trim(txtCRNoTo)) > 0 And IsDate(txtCRNoTo) = True Then
            StrAppend = StrAppend + " And MBMCRNote <= '" & txtCRNoTo & "'"
        End If
    End If
    
    If CboRecType = "M - MCO FEES" Then
        Call ReceiptMCOListing(StrAppend)
        If chkCRNote.Value = 1 Or txtRecDateTo <> "__/__/____" Or txtRecDateTo <> "__/__/____" Then
            Call SearchCRListing
        End If
        
    Else
        If Len(CboRecType) Then
            StrAppend = StrAppend + " And ReceiptType = '" & Mid(CboRecType, 1, 1) & "'"
        End If
        Call ReceiptOthersListing(StrAppend)
    End If
   
End Function

Private Sub ReceiptMCOListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlREC As String
Dim Total As String
Dim Current As String
Dim TotalAmt As String

    Total = 0
    Current = 0
    TotalAmt = 0
    If ChkScan.Value = "0" Then
        lstReceiptMCO.ListItems.Clear
    End If
    sql = " Select * From VIEW_Receipt_MCO Where 0 = 0 "
        
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If

    rst2.Open sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
   
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
    Else
    
    Do
        Do Until rst2.EOF
                    
            With rst2
                Set Record = lstReceiptMCO.ListItems.Add(, , "")
                                
                If Len(!MBMInvoiceNo) Then
                    Record.SubItems(1) = Trim(!MBMInvoiceNo)
                End If
                
                If Len(!MBMInvoiceDate) Then
                    Record.SubItems(2) = Format(!MBMInvoiceDate, "dd/mm/yyyy")
                End If
                
                If Len(!TotalMCOFee) Then
                    Record.SubItems(3) = Format(!TotalMCOFee, "#,##0.00")
                End If
                
                If Len(!PAYCode) Then
                    Record.SubItems(4) = Trim(!PAYCode)
                End If

                
                If Len(!ReceiptNo) Then
                    Record.SubItems(5) = Trim(!ReceiptNo)
                End If
                
                If Len(!ReceiptDate) Then
                    Record.SubItems(6) = Format(!ReceiptDate, "dd/mm/yyyy")
                End If
                
                If Len(!ReceiptAmt) Then
                    Record.SubItems(7) = Format(!ReceiptAmt, "#,##0.00")
                End If
                
                If Len(!INVRecAmt) Then
                    Record.SubItems(8) = Format(!INVRecAmt, "#,##0.00")
                End If
                
                If Len(!ReceiptType) Then
                    Record.SubItems(9) = Trim(!ReceiptType)
                End If
                
                If Len(!RecMCORemarks) Then
                    Record.SubItems(10) = Trim(!RecMCORemarks)
                End If
                
                If Len(!RecMCOIndex) Then
                    Record.SubItems(11) = Trim(!RecMCOIndex)
                End If
                
                If Len(!RecChqNo) Then
                    Record.SubItems(12) = Trim(!RecChqNo)
                End If
                
                If Len(!RecChqDate) Then
                    Record.SubItems(13) = Trim(!RecChqDate)
                End If
                                            
                If Len(!ReceiptAmt) Then
                    TotalAmt = CDbl(TotalAmt) + CDbl(!ReceiptAmt)
                End If

                txtTotalAmt = Format(TotalAmt, "#,##0.00")
                
                Current = Current + 1
                lblCurRecord = Current
            End With
            rst2.MoveNext
            
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
    Loop
   
    Loop Until check = False
    
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    
    CmdSearch.Enabled = True
    
End Sub

Private Sub SearchCRListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlREC As String
Dim Total As String
Dim Current As String
Dim TotalAmt As String
Dim StrAppendCR As String
Dim strsqlREF As String
Dim MCORef As New ADODB.Recordset

    If Len(Trim(txtRecDateFr)) > 0 And IsDate(txtRecDateFr) = True Then
        StrAppendCR = StrAppendCR + " And ReceiptDate >= '" & Format(txtRecDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRecDateTo)) > 0 And IsDate(txtRecDateTo) = True Then
        StrAppendCR = StrAppendCR + " And ReceiptDate <= '" & Format(txtRecDateTo, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(txtCRNoFr)) Then
        StrAppendCR = StrAppendCR + " And MBMCRNote >= '" & txtCRNoFr & "'"
    End If
    
    If Len(Trim(txtCRNoTo)) Then
        StrAppendCR = StrAppendCR + " And MBMCRNote <= '" & txtCRNoTo & "'"
    End If

    sql = " Select * From VIEW_SearchMBM_Refund Where 0 = 0 "
        'sql = ""
    If Len(Trim(StrAppendCR)) Then
        sql = sql + StrAppendCR
    End If
    If StrAppendCR <> "" Then
        rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
       
        If rst2.EOF = True Then
            'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
            CmdSearch.Enabled = True
        Else
        
        Do
            Do Until rst2.EOF
                        
                With rst2
                
                    Set Record = lstReceiptMCO.ListItems.Add(, , "")
                                    
                    If Len(!MBMCRNote) Then
                        Record.SubItems(1) = Trim(!MBMCRNote)
                    End If
                    
                    If Len(!MBMCRNoteDate) Then
                        Record.SubItems(2) = Format(!MBMCRNoteDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMMCORefund) Then
                        Record.SubItems(3) = Format(!MBMMCORefund, "#,##0.00")
                    End If
                    
                    If Len(!PAYCode) Then
                        Record.SubItems(4) = Trim(!PAYCode)
                    End If
    
                    
                    If Len(!ReceiptNo) Then
                        Record.SubItems(5) = Trim(!ReceiptNo)
                    End If
                    
                    If Len(!ReceiptDate) Then
                        Record.SubItems(6) = Format(!ReceiptDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!ReceiptAmt) Then
                        Record.SubItems(7) = Format(!ReceiptAmt, "#,##0.00")
                    End If
                    
                    If Len(!INVRecAmt) Then
                        Record.SubItems(8) = Format(!INVRecAmt, "#,##0.00")
                    End If
                    
                    If Len(!ReceiptType) Then
                        Record.SubItems(9) = Trim(!ReceiptType)
                    End If
                    
                    If Len(!RecMCORemarks) Then
                        Record.SubItems(10) = Trim(!RecMCORemarks)
                    End If
                    
                    If Len(!RecMCOIndex) Then
                        Record.SubItems(11) = Trim(!RecMCOIndex)
                    End If
                    
                    If Len(!RecChqNo) Then
                        Record.SubItems(12) = Trim(!RecChqNo)
                    End If
                    
                    If Len(!RecChqDate) Then
                        Record.SubItems(13) = Trim(!RecChqDate)
                    End If
                                                
                    'If Len(!ReceiptAmt) Then
                    '    TotalAmt = CDbl(TotalAmt) + CDbl(!ReceiptAmt)
                    'End If
    
                    'txtTotalAmt = Format(TotalAmt, "#,##0.00")
                    
                    'Current = Current + 1
                    'lblCurRecord = Current
                End With
                rst2.MoveNext
                
                DoEvents
                    If intExit = 1 Then
                       check = False
                       Exit Do
                    End If
        Loop
       
        Loop Until check = False
        
        End If
        intExit = 0
        rst2.Close
        Set rst2 = Nothing
        
        CmdSearch.Enabled = True
    End If
End Sub

Private Sub ReceiptOthersListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim REC As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlREC As String
Dim Total As String
Dim Current As String
Dim TotalAmt As String

    Total = 0
    Current = 0
    TotalAmt = 0
    
    lstReceiptOther.ListItems.Clear
    
    sql = " Select * From RECOthers Where 0 = 0 "
        
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If

    rst2.Open sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
   
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdSearch.Enabled = True
    Else
    
    Do
        Do Until rst2.EOF
                    
            With rst2
                If Len(!ReceiptNo) Then
                    Set Record = lstReceiptOther.ListItems.Add(, , !ReceiptNo)
                Else
                    Set Record = lstReceiptOther.ListItems.Add(, , "")
                End If
                                
                If Len(!ReceiptDate) Then
                    Record.SubItems(1) = Format(!ReceiptDate, "dd/mm/yyyy")
                End If
                
                If Len(!ReceiptAmt) Then
                    Record.SubItems(2) = Format(!ReceiptAmt, "#,##0.00")
                End If
                
                If Len(!ReceiptType) Then
                    Record.SubItems(3) = Trim(!ReceiptType)
                End If
                
                If Len(!RecChqNo) Then
                    Record.SubItems(4) = !RecChqNo
                End If
                
                If Len(!RecChqDate) Then
                    Record.SubItems(5) = Format(!RecChqDate, "dd/mm/yyyy")
                End If
                
                Record.SubItems(6) = IIf(Len(!RECAccCode), !RECAccCode, "")
                
                If Len(!RecOthersRemarks) Then
                    Record.SubItems(7) = Trim(!RecOthersRemarks)
                End If
                
                If Len(!RecOthersIndex) Then
                    Record.SubItems(8) = Trim(!RecOthersIndex)
                End If
                
                If Len(!ReceiptAmt) Then
                    TotalAmt = TotalAmt + !ReceiptAmt
                End If
                txtTotalAmt = Format(TotalAmt, "#,##0.00")
                
                Current = Current + 1
                lblCurRecord = Current
            End With
            rst2.MoveNext
            
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
    Loop
   
    Loop Until check = False
    
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    
    CmdSearch.Enabled = True
End Sub

Private Sub cmdEXPtoExcel_Click()
    'Screen.MousePointer = vbHourglass
        'If CboRecType = "M - MCO FEES" Then
            'Call ExportListViewtoExcel(lstReceiptMCO)
        'Else
            'Call ExportListViewtoExcel(lstReceiptOther)
        'End If
    'Screen.MousePointer = vbDefault
    
    Screen.MousePointer = vbHourglass
        
    If CboRecType = "M - MCO FEES" Then
        If lstReceiptMCO.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(lstReceiptMCO)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
        
    Else
    
        If lstReceiptOther.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(lstReceiptOther)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
    
    End If
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdSave_Click()
Dim RecordSaved
Dim SaveOthers As New ADODB.Recordset
Dim SaveMCO As New ADODB.Recordset
Dim ItemChecked As Boolean
Dim i As Integer
Dim INVNo As String
Dim InvDate As String
Dim tmpRecInvAmt As String
Dim ReceiptNumber As String
Dim strsqlREC As String

    ItemChecked = False
    CmdSave.Enabled = False
    strcriteria = ""
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

    If CboRecType = "" Then
       MsgBox "Please Enter Rec. Type", vbOKOnly + vbCritical, DialogBoxTitle
       CboRecType.SetFocus
    ElseIf TxtReceiptNo = "" Then
        MsgBox "Please Enter Receipt No", vbOKOnly + vbCritical, DialogBoxTitle
       TxtReceiptNo.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    ElseIf txtAmt = "" Then
        MsgBox "Please Enter Receipt Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtAmt.SetFocus
    Else
        
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            'ReceiptNumber = GenerateReceiptMCONo
            'txtReceiptNo = ""
            'txtReceiptNo = ReceiptNumber

            If CboRecType = "M - MCO FEES" Then
                For i = 1 To lstReceiptMCO.ListItems.Count
                    If lstReceiptMCO.ListItems(i).Checked = True Then
                        ItemChecked = True
                        Exit For
                    End If
                Next i
                
                SaveMCO.Open "RECMCO", medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                If ItemChecked = True Then
                    
                    
                    
                    For i = 1 To lstReceiptMCO.ListItems.Count
                        If lstReceiptMCO.ListItems(i).Checked = True Then
                            
                            INVNo = Trim(lstReceiptMCO.ListItems(i).ListSubItems(1).Text)
                            InvDate = Trim(lstReceiptMCO.ListItems(i).ListSubItems(2).Text)
                            tmpRecInvAmt = Trim(lstReceiptMCO.ListItems(i).ListSubItems(3).Text)
                           ' strsqlREC = " Select * from RECMCO where 0=0"
                           ' SaveMCO.Open strsqlREC, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                            
                            'SaveMCO.Open "RECMCO", medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            
                            With SaveMCO
                                !InvoiceNo = Trim(INVNo)
                                !InvoiceDate = Format(InvDate, "dd/mm/yyyy")
                                !ReceiptNo = ChkStr(TxtReceiptNo)
                                !ReceiptDate = Format(txtReceiptDate, "dd/mm/yyyy")
                                !RecChqNo = txtChequeNo
                                !RecChqDate = Format(txtChequeDate, "dd/mm/yyyy")
                                !ReceiptAmt = txtAmt
                                If Mid(INVNo, 1, 1) = "S" Then
                                    !INVRecAmt = tmpRecInvAmt
                                Else
                                    !INVRecAmt = tmpRecInvAmt * -1
                                End If
                                    
                                !ReceiptType = Mid(CboRecType, 1, 1)
                                !RecMCORemarks = ChkStr(Trim(TxtRecRemarks))
                                !RecMCOLastUpdateUser = Logon
                                !RecMCOLastUpdatedate = Date
                            End With
                            SaveMCO.Update
                            
                        End If
                    Next i
                        
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    lstReceiptMCO.ListItems.Clear
                    'Call ClearForm(Me)
                    strcriteria = strcriteria + " AND ReceiptNo = '" & RemStr(Trim(TxtReceiptNo)) & "'"
                    Call ReceiptMCOListing(strcriteria)
                    
                End If
            SaveMCO.Close
            Set SaveMCO = Nothing
            Else
            
                'ReceiptNumber = GenerateReceiptMCONo
                'txtReceiptNo = ""
                'txtReceiptNo = ReceiptNumber
                    
                SaveOthers.Open "RECOthers", medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                SaveOthers.AddNew
                    
                With SaveOthers
                    !RECAccCode = txtAccCode
                    !ReceiptNo = ChkStr(TxtReceiptNo)
                    !ReceiptDate = Format(txtReceiptDate, "dd/mm/yyyy")
                    !ReceiptAmt = txtAmt
                    !RecChqNo = txtChequeNo
                    !RecChqDate = Format(txtChequeDate, "dd/mm/yyyy")
                    !ReceiptType = Mid(CboRecType, 1, 1)
                    !RecOthersRemarks = ChkStr(Trim(TxtRecRemarks))
                    !RecLastUpdateUser = Logon
                    !RecLastUpdatedate = Date
                End With
                SaveOthers.Update
                'Call ClearForm(Me)
            
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                strcriteria = strcriteria + " AND ReceiptNo = '" & RemStr(Trim(TxtReceiptNo)) & "'"
                Call ReceiptOthersListing(strcriteria)

                SaveOthers.Close
                Set SaveOthers = Nothing
            End If
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    CmdSave.Enabled = True
End Sub

Private Sub CmdSearch_Click()
Dim listrecord As Integer

    Screen.MousePointer = vbHourglass
    
    CmdSearch.Enabled = False
    listloop = 1
    check = False
    lblchkAmt = 0
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If Trim(CboRecType) = "" Then
        MsgBox "Please select the Receipt Type", vbOKOnly + vbCritical
    Else
        If ChkScan.Value = 1 And Mid(CboRecType, 1, 1) = "M" Then
    
            If TxtInvoiceNo1 = "" Then
                MsgBox "Please Enter Invoice No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtInvoiceNo1.SetFocus
            Else
                listrecord = lstReceiptMCO.ListItems.Count
                If listrecord <> 0 Then
                    For listloop = 1 To listrecord
                        Call InvNoCheck
                    Next listloop
                End If
                    If check = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                        TxtInvoiceNo1.SetFocus
                    Else
                        Call SearchRecord
                    End If
            End If
        Else
            Call SearchRecord
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    CmdSearch.Enabled = True

    Screen.MousePointer = vbDefault
        
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
End Sub


Private Sub CmdUpdate_Click()
Dim RecordSaved
Dim RptMBM As New ADODB.Recordset
Dim UpdateRptMBM As New ADODB.Recordset
Dim strsql As String
    
    CmdUpdate.Enabled = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If EditFlag = True Then
    
        If CboRecType = "M - MCO FEES" Then
            If lstReceiptMCO.ListItems.Count = 0 Then
                MsgBox "Please select the record", vbOKOnly + vbCritical
            ElseIf IsDate(txtReceiptDate) = False Then
                MsgBox "Invalid Receipt Date!", vbCritical
                txtReceiptDate.SetFocus
            ElseIf txtReceiptDate = "__/__/____" Then
                MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
                txtReceiptDate.SetFocus
            ElseIf txtAmt = "" Then
                MsgBox "Please Enter Cheque Amt!", vbCritical
                txtAmt.SetFocus
            Else
                strsql = "Select * From RECMCO Where RecMCOIndex = '" & Trim(lstReceiptMCO.SelectedItem.SubItems(11)) & "'"
                RptMBM.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                
                RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                        
                    If RptMBM.recordCount = 0 Then
                        RptMBM.Close
                        Set RptMBM = Nothing
                        Exit Sub
                    Else
                        
                        With RptMBM
                            !ReceiptNo = TxtReceiptNo
                            !ReceiptDate = txtReceiptDate
                            !ReceiptAmt = txtAmt
                            !INVRecAmt = Trim(lstReceiptMCO.SelectedItem.SubItems(8))
                            !InvoiceNo = Trim(lstReceiptMCO.SelectedItem.SubItems(1))
                            !InvoiceDate = Trim(lstReceiptMCO.SelectedItem.SubItems(2))
                            !ReceiptType = Mid(CboRecType, 1, 1)
                            !RecChqNo = txtChequeNo
                            !RecChqDate = Format(txtChequeDate, "dd/mm/yyyy")
                            !RecMCORemarks = Trim(TxtRecRemarks)
                            !RecMCOLastUpdatedate = Date
                            !RecMCOLastUpdateUser = Logon
                            .Update
                        End With
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
                  
                    Call SearchRecord
                    EditFlag = False
                    End If
                End If
                RptMBM.Close
                Set RptMBM = Nothing
                Call ClearReceipt
                CmdSave.Enabled = True
                CmdUpdate.Enabled = True
            End If
        
        Else
            If lstReceiptOther.ListItems.Count = 0 Then
                MsgBox "Please select the record", vbOKOnly + vbCritical
            ElseIf IsDate(txtReceiptDate) = False Then
                MsgBox "Invalid Receipt Date!", vbCritical
                txtReceiptDate.SetFocus
            ElseIf txtReceiptDate = "__/__/____" Then
                MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
                txtReceiptDate.SetFocus
            ElseIf txtAmt = "" Then
                MsgBox "Please Enter Cheque Amt!", vbCritical
                txtAmt.SetFocus
            Else
            strsql = ""
                strsql = "Select * From RECOthers Where RecOthersIndex = '" & Trim(lstReceiptOther.SelectedItem.SubItems(8)) & "'"
                RptMBM.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                    
                RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                        
                    If RptMBM.recordCount = 0 Then
                        RptMBM.AddNew
                    End If
                        
                    With RptMBM
                        !ReceiptNo = TxtReceiptNo
                        !ReceiptDate = txtReceiptDate
                        !RecChqNo = txtChequeNo
                        !RecChqDate = Format(txtChequeDate, "dd/mm/yyyy")
                        !ReceiptAmt = txtAmt
                        !ReceiptType = Mid(CboRecType, 1, 1)
                        !RecOthersRemarks = Trim(TxtRecRemarks)
                        !RecLastUpdatedate = Date
                        !RecLastUpdateUser = Logon
                        .Update
                    End With
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
                  
                    Call SearchRecord
                    EditFlag = False
                End If
                RptMBM.Close
                Set RptMBM = Nothing
                    
                Call ClearReceipt
                CmdSave.Enabled = True
                CmdUpdate.Enabled = False
            End If
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    CmdUpdate.Enabled = True
End Sub

Private Sub Form_Load()
    EditFlag = False
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    Call ComboListing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    
    txtTotalAmt = 0
    intExit = 0
End Sub

Private Sub lstReceiptMCO_Click()
    ItemCheck = False
    Call DisplayData
End Sub
Private Sub DisplayData(Optional Item As ListItem, Optional ItemCheck As Boolean)
Dim lstRec As ListItem
    'If ItemCheck = True Then
        Call ClearReceipt
        CmdUpdate.Enabled = False
        If Mid(CboRecType, 1, 1) = "M" Then
        
            If lstReceiptMCO.ListItems.Count Then
                If tmpItemCheck = False Then
                    If Trim(lstReceiptMCO.SelectedItem.SubItems(5)) <> "" Then
                        Set lstRec = lstReceiptMCO.SelectedItem
                        TxtReceiptNo = IIf(Len(Trim(lstRec.SubItems(5))), lstRec.SubItems(5), "")
                        If Len(lstRec.SubItems(6)) Then
                            txtReceiptDate = Format(lstRec.SubItems(6), "DD/MM/YYYY")
                        End If
                        If Len(lstRec.SubItems(7)) Then
                            txtAmt = lstRec.SubItems(7)
                        End If
                        
                        tmpRecType = IIf(Len(lstRec.SubItems(9)), lstRec.SubItems(9), "")
                        
                        TxtRecRemarks = IIf(Len(lstRec.SubItems(10)), lstRec.SubItems(10), "")
                        
                        lblIndex = IIf(Len(lstRec.SubItems(11)), lstRec.SubItems(11), "")
                        
                        If Len(lstRec.SubItems(12)) Then
                            txtChequeNo = lstRec.SubItems(12)
                        End If
            
                        If Len(lstRec.SubItems(13)) Then
                            txtChequeDate = Format(lstRec.SubItems(13), "DD/MM/YYYY")
                        End If
                        txtINVNo = lstRec.SubItems(1)
                        CmdSave.Enabled = False
                    'Else
                    '    MsgBox "Receipt has been updated", vbOKOnly + vbCritical
                    End If
                End If
            End If
        Else
            If lstReceiptOther.ListItems.Count Then
                'If Len(Trim(lstReceiptOther.SelectedItem.Text)) Then
                    Set lstRec = lstReceiptOther.SelectedItem
                        TxtReceiptNo = IIf(Len(lstRec.Text), lstRec.Text, "")
                        If Len(lstRec.SubItems(1)) Then
                            txtReceiptDate = Format(lstRec.SubItems(1), "DD/MM/YYYY")
                        End If
                       If Len(lstRec.SubItems(2)) Then
                           txtAmt = lstRec.SubItems(2)
                       End If
                       
                       tmpRecType = IIf(Len(lstRec.SubItems(3)), lstRec.SubItems(3), "")
        
                       If Len(lstRec.SubItems(4)) Then
                           txtChequeNo = lstRec.SubItems(4)
                       End If
           
                       If Len(lstRec.SubItems(5)) Then
                           txtChequeDate = Format(lstRec.SubItems(5), "DD/MM/YYYY")
                       End If
           
                       If Len(lstRec.SubItems(6)) Then
                           txtAccCode = lstRec.SubItems(6)
                       End If
                       
                       TxtRecRemarks = IIf(Len(lstRec.SubItems(7)), lstRec.SubItems(7), "")
                       
                       lblIndex = IIf(Len(lstRec.SubItems(8)), lstRec.SubItems(8), "")
                       
                       CmdSave.Enabled = False
                'Else
                '    MsgBox "Receipt has been updated", vbOKOnly + vbCritical
                'End If
            End If
        End If
        tmpItemCheck = False
    'End If
End Sub

Private Sub lstReceiptMCO_GotFocus()

End Sub

Private Sub lstReceiptMCO_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
Dim tmpAmt As Double
Dim tmpChkAmt As Double
    tmpAmt = 0
    tmpChkAmt = 0
    If Item.Checked = True Then
        'If Item.SubItems(5) <> "" Then
        '    MsgBox "Receipt has been updated!", vbCritical
        '    Item.Checked = False
        'Else
        If Mid(Item.SubItems(1), 1, 1) = "S" Then
            tmpAmt = IIf(IsNumeric(txtAmt), txtAmt, 0)
            txtAmt = IIf(Len(Item.SubItems(7)), Item.SubItems(7), 0)
            tmpChkAmt = IIf(IsNumeric(lblchkAmt), lblchkAmt, 0)
            lblchkAmt = tmpChkAmt + CDbl(Item.SubItems(3))
        ElseIf Mid(Item.SubItems(1), 1, 1) = "C" Then
            tmpAmt = IIf(IsNumeric(Item.SubItems(3)), Item.SubItems(3), 0)
            txtAmt = IIf(Len(Item.SubItems(7)), Item.SubItems(7), 0)
            tmpChkAmt = IIf(IsNumeric(lblchkAmt), lblchkAmt, 0)
            lblchkAmt = tmpChkAmt - CDbl(Item.SubItems(3))
        End If
        
        'End If
    Else
        If Mid(Item.SubItems(1), 1, 1) = "S" Then
            tmpAmt = IIf(IsNumeric(txtAmt), txtAmt, 0)
            txtAmt = IIf(Len(Item.SubItems(7)), Item.SubItems(7), 0)
            tmpChkAmt = IIf(IsNumeric(lblchkAmt), lblchkAmt, 0)
            lblchkAmt = tmpChkAmt - CDbl(Item.SubItems(3))
        ElseIf Mid(Item.SubItems(1), 1, 1) = "C" Then
            tmpAmt = IIf(IsNumeric(Item.SubItems(3)), Item.SubItems(3), 0)
            txtAmt = IIf(Len(Item.SubItems(7)), Item.SubItems(7), 0)
            tmpChkAmt = IIf(IsNumeric(lblchkAmt), lblchkAmt, 0)
            lblchkAmt = tmpChkAmt + CDbl(Item.SubItems(3))
        End If
    End If
End Sub

Private Sub lstReceiptMCO_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DisplayData
End Sub

Private Sub lstReceiptMCO_KeyUp(KeyCode As Integer, Shift As Integer)
    Call DisplayData
End Sub

Private Sub lstReceiptOther_Click()
    Call DisplayData
End Sub

Private Sub lstReceiptOther_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DisplayData
End Sub

Private Sub lstReceiptOther_KeyUp(KeyCode As Integer, Shift As Integer)
    Call DisplayData
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub txtInvDate1_LostFocus()
    If IsDate(txtInvDate1) Then
    Else
        If txtInvDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtInvDate2_LostFocus()
    If IsDate(txtINVDate2) Then
    Else
        If txtINVDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtINVDate2.SetFocus
        End If
    End If
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

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
    Else
        If txtReceiptDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDate.SetFocus
        End If
    End If
End Sub

Private Sub lstReceiptMCO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstReceiptMCO, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstReceiptMCO, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstReceiptMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstReceiptMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstReceiptMCO, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstReceiptMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstReceiptMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstReceiptMCO, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstReceiptMCO, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstReceiptMCO, ColumnHeader.Index, ldtString, m_blnDirection(10)
    End Select
End Sub

Private Sub lstReceiptOther_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstReceiptOther, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstReceiptOther, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
    Case 3
        SortListView lstReceiptOther, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstReceiptOther, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstReceiptOther, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstReceiptOther, ColumnHeader.Index, ldtString, m_blnDirection(6)
    End Select
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

Private Sub CboRecType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRecType.Text, CboRecType.SelStart) + Chr(KeyAscii) + Mid(CboRecType.Text, CboRecType.SelStart + CboRecType.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboRecType.ListCount - 1
       
        If NewText = LCase(Left(CboRecType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboRecType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboRecType.Text = ValidValue
               CboRecType.SelStart = 0
               CboRecType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub ClearReceipt()
    TxtReceiptNo.Text = ""
    txtReceiptDate.Text = "__/__/____"
    txtAmt.Text = ""
    txtChequeNo.Text = ""
    txtChequeDate.Text = "__/__/____"
    TxtRecRemarks.Text = ""
End Sub

Private Function InvNoCheck()
Dim INVNo As String

    INVNo = Trim(lstReceiptMCO.ListItems(listloop).SubItems(1))
    If Trim(ChkStr(TxtInvoiceNo1)) = ChkStr(INVNo) Then
        check = True
    Else
         check = False
    End If
End Function


