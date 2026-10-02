VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEnquiryPayMember 
   Caption         =   "Enquiry Payment to Member"
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
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   120
      TabIndex        =   57
      Top             =   6960
      Width           =   11655
      Begin VB.TextBox lblAmtPaid 
         Height          =   285
         Left            =   6600
         TabIndex        =   85
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox lblSelectedAmt 
         Height          =   285
         Left            =   3840
         TabIndex        =   84
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox lblCurRecord 
         Height          =   285
         Left            =   120
         TabIndex        =   83
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtPVStatus2 
         Height          =   375
         Left            =   5160
         TabIndex        =   72
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10320
         TabIndex        =   63
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   62
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   330
         Left            =   8040
         TabIndex        =   61
         Top             =   150
         Width           =   735
      End
      Begin VB.TextBox txtPVStatus 
         Height          =   375
         Left            =   2760
         TabIndex        =   60
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   330
         Left            =   9360
         TabIndex        =   59
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   8760
         TabIndex        =   58
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Total Amt Paid"
         Height          =   255
         Left            =   5280
         TabIndex        =   76
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   65
         Top             =   600
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   1440
         TabIndex        =   67
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   66
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Pay Member Amt"
         Height          =   255
         Left            =   2400
         TabIndex        =   64
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   120
      TabIndex        =   33
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox ChkSunNo 
         Height          =   255
         Left            =   11040
         TabIndex        =   80
         Top             =   1540
         Width           =   255
      End
      Begin VB.CheckBox ChkSunDate 
         Height          =   255
         Left            =   11040
         TabIndex        =   79
         Top             =   1820
         Width           =   375
      End
      Begin VB.CheckBox ChkPVDate 
         Height          =   255
         Left            =   11040
         TabIndex        =   23
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox ChkPVNo 
         Height          =   255
         Left            =   11040
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
         Height          =   255
         Left            =   11040
         TabIndex        =   31
         Top             =   1000
         Width           =   375
      End
      Begin VB.CheckBox chkCRAmt 
         Height          =   255
         Left            =   6000
         TabIndex        =   30
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkCRNo 
         Height          =   255
         Left            =   11040
         TabIndex        =   27
         Top             =   740
         Width           =   375
      End
      Begin VB.CheckBox chkChequeNo 
         Height          =   255
         Left            =   6000
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkPeriodYr 
         Height          =   255
         Left            =   11040
         TabIndex        =   32
         Top             =   1280
         Width           =   375
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtVerNo 
         Height          =   285
         Left            =   4080
         MaxLength       =   15
         TabIndex        =   1
         Top             =   200
         Width           =   615
      End
      Begin VB.TextBox txtCASNo 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   0
         Top             =   200
         Width           =   2175
      End
      Begin VB.TextBox txtPVNoFr 
         Height          =   285
         Left            =   7920
         MaxLength       =   20
         TabIndex        =   8
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox txtPVNoTo 
         Height          =   285
         Left            =   9600
         MaxLength       =   20
         TabIndex        =   9
         Top             =   210
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDateFr 
         Height          =   285
         Left            =   7920
         TabIndex        =   11
         Top             =   460
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCHENoFr 
         Height          =   285
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCRNoFr 
         Height          =   285
         Left            =   7920
         MaxLength       =   20
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtCRAmtFr 
         Height          =   285
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   28
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmEnquiryPayMember.frx":0000
         Left            =   7920
         List            =   "FrmEnquiryPayMember.frx":0028
         TabIndex        =   15
         Top             =   975
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmEnquiryPayMember.frx":0053
         Left            =   7920
         List            =   "FrmEnquiryPayMember.frx":00DB
         TabIndex        =   17
         Top             =   1260
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtPVDateTo 
         Height          =   285
         Left            =   9600
         TabIndex        =   12
         Top             =   460
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCHENoTo 
         Height          =   285
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCRNoTo 
         Height          =   285
         Left            =   9600
         MaxLength       =   20
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtCRAmtTo 
         Height          =   285
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmEnquiryPayMember.frx":01E7
         Left            =   9600
         List            =   "FrmEnquiryPayMember.frx":020F
         TabIndex        =   16
         Top             =   975
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmEnquiryPayMember.frx":023A
         Left            =   9600
         List            =   "FrmEnquiryPayMember.frx":02C2
         TabIndex        =   18
         Top             =   1260
         Width           =   1335
      End
      Begin VB.ComboBox cboWSStatus 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   77
         Top             =   450
         Width           =   3015
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         ItemData        =   "FrmEnquiryPayMember.frx":03CE
         Left            =   1680
         List            =   "FrmEnquiryPayMember.frx":03D0
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox TxtMBMNo 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1005
         Width           =   3015
      End
      Begin VB.TextBox TxtMBMName 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1260
         Width           =   3015
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   1515
         Width           =   3015
      End
      Begin VB.ComboBox CboPayment 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtSunNo1 
         Height          =   285
         Left            =   7920
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1550
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtSUNDate1 
         Height          =   285
         Left            =   7920
         TabIndex        =   21
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtSunNo2 
         Height          =   285
         Left            =   9600
         MaxLength       =   20
         TabIndex        =   20
         Top             =   1550
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtSUNDate2 
         Height          =   285
         Left            =   9600
         TabIndex        =   22
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Caption         =   "SUN PV No Fr"
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
         TabIndex        =   82
         Top             =   1605
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "SUN PV Date Fr"
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
         TabIndex        =   81
         Top             =   1845
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "WS Claimability"
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
         TabIndex        =   78
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Mem No"
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
         TabIndex        =   55
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Mem Name"
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
         TabIndex        =   54
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Payment Status"
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
         TabIndex        =   53
         Top             =   1800
         Width           =   1335
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
         Left            =   120
         TabIndex        =   52
         Top             =   1515
         Width           =   1215
      End
      Begin VB.Label Label17 
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
         Left            =   9240
         TabIndex        =   51
         Top             =   510
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "PV Date from"
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
         TabIndex        =   50
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "PV No from"
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
         TabIndex        =   49
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label22 
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
         Left            =   9240
         TabIndex        =   48
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label25 
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
         Left            =   9240
         TabIndex        =   47
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label26 
         Caption         =   "CR No From"
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
         TabIndex        =   46
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label Label27 
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
         Left            =   9240
         TabIndex        =   45
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label29 
         Caption         =   "Cheq. No from"
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
         Left            =   4800
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label30 
         Caption         =   "A/C Yr Fr"
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
         TabIndex        =   43
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label31 
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
         Left            =   9240
         TabIndex        =   42
         Top             =   1830
         Width           =   375
      End
      Begin VB.Label Label32 
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
         Left            =   9240
         TabIndex        =   41
         Top             =   1530
         Width           =   375
      End
      Begin VB.Label Label34 
         Caption         =   "A/C Mth Fr"
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
         TabIndex        =   40
         Top             =   1035
         Width           =   1815
      End
      Begin VB.Label Label16 
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
         Left            =   9240
         TabIndex        =   39
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "CR Amt From"
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
         Left            =   4800
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label14 
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
         Left            =   11040
         TabIndex        =   37
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label36 
         Caption         =   "Claim No"
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
         TabIndex        =   36
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label37 
         Caption         =   "-"
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
         Left            =   3960
         TabIndex        =   35
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label38 
         Caption         =   "Health Type"
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
         TabIndex        =   34
         Top             =   720
         Width           =   1815
      End
   End
   Begin MSComctlLib.ListView lstPaymentMBM 
      Height          =   1335
      Left            =   120
      TabIndex        =   56
      Top             =   2520
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2355
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
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ref. No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Pymt Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "MBM Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Pay To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Claim No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Amt Paid"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "OS Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "AR Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "PV No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "PV Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Cheq. No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Cheq. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Bank In"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Account No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Remarks"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "WS Claimability"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "SUN PV"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "SUN PVDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstReceiptMBM 
      Height          =   1380
      Left            =   120
      TabIndex        =   71
      Top             =   4080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2434
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CaseNo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "MBM No."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Rept Fr"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   " "
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amt Rec'd"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   " "
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Rept No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Rept Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Acc Period"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Remarks"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "CasNo"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lstPaymentMBMReissue 
      Height          =   1215
      Left            =   120
      TabIndex        =   73
      Top             =   5760
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   2143
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Patient"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Claim No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Amt Paid"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Prev Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Reissue PV"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Reissue Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Reissue ChqNo"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Reissue ChqDate"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Reissue ChqAmt"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "PayTo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ReissueBank"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ReissueAccount"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "BankCharges"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "AccPeriod"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "IssueBy"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "ReissuePVRemark"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Received From MBM"
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
      TabIndex        =   75
      Top             =   3860
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Reissue Cheque"
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
      TabIndex        =   74
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Enquiry Payment to Member"
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
      TabIndex        =   70
      Top             =   0
      Width           =   11805
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
      TabIndex        =   69
      Top             =   7515
      Width           =   1575
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
      Left            =   2160
      TabIndex        =   68
      Top             =   7515
      Width           =   2295
   End
End
Attribute VB_Name = "FrmEnquiryPayMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Integer
Dim ScanCode As String
Dim listloop1 As Integer
Dim Check As Boolean
Dim RecordSelected As Integer
Private m_blnDirection(1 To 39) As Boolean

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

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        lstPaymentMBM.ListItems.Clear
        txtCASNo.SetFocus
    Else
        lstPaymentMBM.ListItems.Clear
        txtCASNo.SetFocus
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

Private Sub Form_Load()
    Call ComboListing
    CboPayment.ListIndex = 0
    intExit = 0
    lblCurRecord = 0
    lblSelectedAmt = 0
    lblAmtPaid = 0
End Sub

'=============== Click Command Button ============================
Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        'Call ExportToExcelAnalysis1(lstPaymentMBM)
    If lstPaymentMBM.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstPaymentMBM)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub cmdSearch_Click()
Dim listrecord As Integer
    
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
    ScanCode = txtCASNo
    If Mid(ScanCode, 1, 2) = "LO" Then
        txtCASNo = ""
        Call SearchForm(ScanCode)
    Else
    lblSelectedAmt = 0
    lblAmtPaid = 0
        
        If ChkScan.Value = 1 Then
            If txtCASNo = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                txtCASNo.SetFocus
            Else
            
                Screen.MousePointer = vbHourglass
                
                listrecord = lstPaymentMBM.ListItems.Count
                If listrecord <> 0 Then
                    
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                    If Check = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    ElseIf Check = False Then
                        Call SearchRecord
                    End If
                    Check = False
                Screen.MousePointer = Default
                txtCASNo = ""
                txtCASNo.SetFocus
            End If
            txtCASNo.SetFocus
        Else
            Call SearchRecord
        End If
    End If
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    CboPayment.ListIndex = 0
    lstReceiptMBM.ListItems.Clear
    lstPaymentMBM.ListItems.Clear
    lstPaymentMBMReissue.ListItems.Clear
    lblCurRecord = 0
    lblSelectedAmt = 0
    lblAmtPaid = 0
    ChkScan.Value = 1
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
End Sub

Private Sub cmdExit_Click()
        Unload Me
End Sub
'=============== Click Command Button ============================

'************************Start ComboListing**************************

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim monthstart
Dim yearstart
Dim HLT As New ADODB.Recordset
Dim cmd As New ADODB.Command

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHLTType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HLT.CursorLocation = adUseClient
    HLT.CursorType = adOpenForwardOnly
    HLT.CacheSize = 100
    Set HLT = cmd.Execute
    
    Do Until HLT.EOF
        With HLT
            cboHLTCode.AddItem ChkStr(!HLTCode) & " - " & ChkStr(!HLTDescription)
            HLT.MoveNext
        End With
    Loop
    HLT.Close
    Set HLT = Nothing
    
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
            CboPayor.AddItem !PAYCode & " - " & Trim(!PAYCompanyName)
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
    
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
    With CboPayment
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
        .AddItem "3 - Partial Paid"
        .AddItem "4 - Paid & Partial Paid"
        .AddItem "5 - Unpaid & Partial Paid"
    End With
    
    For monthstart = 1 To 12
        'With CboACPeriodMth
            '.AddItem monthstart
        'End With
        With cboACPeriodMthFr
            .AddItem monthstart
        End With
        With cboACPeriodMthTo
            .AddItem monthstart
        End With
    Next monthstart
    
    For yearstart = 1999 To 3000
        'With CboACPeriodYr
            '.AddItem yearstart
        'End With
        With cboACPeriodYRFr
            .AddItem yearstart
        End With
        With cboACPeriodYRTo
            .AddItem yearstart
        End With
    Next yearstart
    
    cboWSStatus.Clear
    cboWSStatus.AddItem "AA" & " - " & "Accepted"
    cboWSStatus.AddItem "RR" & " - " & "Repudiated"
    cboWSStatus.AddItem "CC" & " - " & "Cancelled"
    cboWSStatus.AddItem "SA" & " - " & "Special Approval"
    cboWSStatus.AddItem "PP" & " - " & "Pay & Claim"
    cboWSStatus.AddItem "NN" & " - " & "Not Decided"
    cboWSStatus.AddItem "NA" & " - " & "Not Decided/Accepted"
    cboWSStatus.AddItem "NR" & " - " & "Not Decided/Repudiated"
    cboWSStatus.AddItem "PA" & " - " & "Pay&Claim/Accepted"
    cboWSStatus.AddItem "PR" & " - " & "Pay&Claim/Rejected"
    cboWSStatus.AddItem "AR" & " - " & "Accepted/Rejected"
    cboWSStatus.AddItem "RA" & " - " & "Rejected/Accepted"
    cboWSStatus.AddItem "BR" & " - " & "BTBG/Rejected"

    
End Sub
'************************Start ComboListing**************************


'**********************Start Search Record*******************************
Private Function SearchRecord()
Dim strSql As String
Dim strAppend As String
Dim sql As String
    
    strAppend = ""
    
    If Len(Trim(txtCASNo)) Then
        strAppend = strAppend & " And CASNumber = '" & ChkStr(txtCASNo) & "'"
    End If
    
    If Len(Trim(txtVerNo)) Then
        strAppend = strAppend & " And ClaimVersion = '" & ChkStr(txtVerNo) & "'"
    End If
    If Len(Trim(cboWSStatus)) Then
        strAppend = strAppend & " and CLMWSStatus = '" & Mid(cboWSStatus, 1, 2) & "'"
    End If
    If Trim(cboHLTCode) <> "" Then
        strAppend = strAppend & " and HLTCode = '" & Mid(cboHLTCode, 1, 1) & "'"
    End If

    If Len(Trim(TxtMBMNo)) Then
        strAppend = strAppend & " And MBMNumber = '" & ChkStr(TxtMBMNo) & "'"
    End If
    
    If Len(Trim(TxtMBMName)) Then
        strAppend = strAppend & " And MBMName like '%" & ChkStr(TxtMBMName) & "%'"
    End If
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend & " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Mid(CboPayment, 1, 1) = 1 Then
      strAppend = strAppend & " and PVPayMBMStatus = 'F'"
    End If
    
    If Mid(CboPayment, 1, 1) = 2 Then
      strAppend = strAppend & " and (PVPayMBMStatus = '' or PVPayMBMStatus is null) "
    End If
    
    If Mid(CboPayment, 1, 1) = 3 Then
      strAppend = strAppend & " and PVPayMBMStatus = 'P'"
    End If
    
    If Mid(CboPayment, 1, 1) = 4 Then
      strAppend = strAppend & " and (PVPayMBMStatus = 'P' or PVPayMBMStatus = 'F') "
    End If
    
    If Mid(CboPayment, 1, 1) = 5 Then
      strAppend = strAppend & " and (PVPayMBMStatus = '' or PVPayMBMStatus is null or PVPayMBMStatus = 'P') "
    End If

    If Len(txtPVNoFr) Then
        strAppend = strAppend & " and PVNo >= '" & Trim(txtPVNoFr) & "'"
    End If
        
    If Len(txtPVNoTo) Then
        strAppend = strAppend & " and PVNo <= '" & Trim(txtPVNoTo) & "'"
    End If
    
    If ChkPVNo.Value = 1 Then
        sql = ""
        sql = "AND (PVNo Is null OR PVNo = '') " & ""
        strAppend = strAppend & sql
    End If
    
    If Len(Trim(txtPVDateFr)) > 0 And IsDate(txtPVDateFr) = True Then
        strAppend = strAppend & " And PVDate >= '" & Format(txtPVDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDateTo)) > 0 And IsDate(txtPVDateTo) = True Then
        strAppend = strAppend & " And PVDate <= '" & Format(txtPVDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If ChkPVDate.Value = 1 Then
        sql = ""
        sql = "AND (PVDate Is null OR PVDate = '') " & ""
        strAppend = strAppend & sql
    End If
    
'    If Len(txtCHENoFr) Then
'        strAppend = strAppend & " and PVChequeNo >= '" & Trim(txtCHENoFr) & "'"
'    End If
'
'    If Len(txtCHENoTo) Then
'        strAppend = strAppend & " and PVChequeNo <= '" & Trim(txtCHENoTo) & "'"
'    End If
    
'    If chkChequeNo.value = 1 Then
'        sql = ""
'        sql = "AND (PVChequeNo Is null OR PVChequeNo = '') " & ""
'        strAppend = strAppend & sql
'    End If
    
    If Len(txtCRNoFr) Then
        strAppend = strAppend & " and DebitCreditRefNo >= '" & Trim(txtCRNoFr) & "'"
    End If
        
    If Len(txtCRNoTo) Then
        strAppend = strAppend & " and DebitCreditRefNo <= '" & Trim(txtCRNoTo) & "'"
    End If
    
    If chkCRNo.Value = 1 Then
        sql = ""
        sql = "AND (DebitCreditRefNo Is null OR DebitCreditRefNo = '') " & ""
        strAppend = strAppend & sql
    End If
    
'    If Len(txtCRAmtFr) Then
'        strAppend = strAppend & " and CLMPayableByMedix2M >= " & Trim(txtCRAmtFr) & ""
'    End If
'    If Len(txtCRAmtTo) Then
'        strAppend = strAppend & " and CLMPayableByMedix2M <= " & Trim(txtCRAmtTo) & ""
'    End If
    
'    If chkCRAmt.value = 1 Then
'        sql = ""
'        sql = "AND (CLMPayableByMedix2M Is null OR CLMPayableByMedix2M = '') " & ""
'        strAppend = strAppend & sql
'    End If
    
    If Len(cboACPeriodMthFr) Then
        strAppend = strAppend & " and PAYAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
        
    If Len(cboACPeriodMthTo) Then
        strAppend = strAppend & " and PAYAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If chkPeriodMth.Value = 1 Then
        sql = ""
        sql = " AND PAYAccPeriodMth Is null " & ""
        strAppend = strAppend & sql
    End If
    
    If Len(cboACPeriodYRTo) Then
        strAppend = strAppend & " and PAYAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
        
    If Len(cboACPeriodYRTo) Then
        strAppend = strAppend & " and PAYAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkPeriodYr.Value = 1 Then
        sql = ""
        sql = " AND PAYAccPeriodYr Is null " & ""
        strAppend = strAppend & sql
    End If
        If Len(txtPVNoFr) Then
        strAppend = strAppend & " and PVNo = '" & Trim(txtPVNoFr) & "'"
    End If
        
    If Len(txtSunNo1) Then
        strAppend = strAppend & " and PVSunNo >= '" & Trim(txtSunNo1) & "'"
    End If
    If Len(txtSunNo2) Then
        strAppend = strAppend & " and PVSunNo <= '" & Trim(txtSunNo2) & "'"
    End If
    If ChkSunNo.Value = 1 Then
        sql = ""
        sql = "AND (PVSunNo Is null OR PVSunNo = '') " & ""
        strAppend = strAppend & sql
    End If
    
    If Len(Trim(txtSUNDate1)) > 0 And IsDate(txtSUNDate1) = True Then
        strAppend = strAppend & " And PVSUNDate >= '" & Format(txtSUNDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtSUNDate2)) > 0 And IsDate(txtSUNDate2) = True Then
        strAppend = strAppend & " And PVSUNDate <= '" & Format(txtSUNDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkSunDate.Value = 1 Then
        sql = ""
        sql = "AND (PVSUNDate Is null OR PVSUNDate = '') " & ""
        strAppend = strAppend & sql
    End If
    
  '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

    Call Payment2MBMListing(strAppend)
    
    Call Payment2MBMReissueListing(strAppend & " And PVStatus = 'R'")
   ' medixmasterconnPayment.Close
   ' Set medixmasterconnPayment = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing

End Function
'**********************End Search Record*******************************

'**********************Start List Record*******************************

Private Sub Payment2MBMListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim INVAmt As Variant
Dim CASNumber As String
Dim WSVersion As String
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim Check As Boolean
Dim TotalAmt As Variant
Dim TotalAmt2 As Variant
Dim cmd1 As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim RECFound As Boolean
Dim tmpPayableAmt As Double
   total = 0
   Current = 0
    If ChkScan.Value = 0 Then
        lstPaymentMBM.ListItems.Clear
        lstReceiptMBM.ListItems.Clear
    End If
'    sql = " Select  DebitCreditRefNo, MBMName, HLTCode," & _
'           " CASNumber, ClaimVersion, CLMPayableByMedix2M, PVBankCharges, " & _
'           " MBMPaidAmt,  PVNo, PVMBMIndex, PVDate, PVChequeNo, PVPayMBMStatus, PVChequeDate," & _
'           " Pay2MBMLastUpdateUser, PAYAccPeriodMth, PAYAccPeriodYr, ClaimCloseDate, " & _
'           " PVPayTo, PVPayeeBank, PVPayeeACNo, Remark, CASARStatus, PVStatus " & _
'           " From VIEW_Payment2MBMNew where 0 = 0 "'

'    If Len(Trim(strAppend)) Then
'        sql = sql + strAppend
'    End If
    
'    rst2.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    RECFound = False
    Set cmd1.ActiveConnection = medixmasterconnPayment
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_Payment"
    Set prm1 = cmd1.CreateParameter("PaymentType", adVarChar, adParamInput, 2000, "Pay2MBM")
    cmd1.Parameters.Append prm1
    Set prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, strAppend)
    cmd1.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd1.Execute
    
    Do
        Do Until rst2.EOF
        RECFound = True
        txtPVStatus = ""
        If Len(rst2!PVPayMBMStatus) Then
            txtPVStatus = rst2!PVPayMBMStatus
        End If
        With rst2
            Set Record = lstPaymentMBM.ListItems.Add(, , IIf(Len(rst2!DebitCreditRefNo), rst2!DebitCreditRefNo, ""))
            If Trim(txtPVStatus) = "F" Then
                Record.SubItems(1) = "Full"
            ElseIf Trim(txtPVStatus) = "P" Then
                Record.SubItems(1) = "Partial"
            Else
                Record.SubItems(1) = "Unpaid"
            End If
            Record.SubItems(2) = IIf(Len(rst2!MBMName), Trim(rst2!MBMName), "")
            Record.SubItems(3) = IIf(Len(rst2!CLMPayto), Trim(rst2!CLMPayto), "")
            If Len(rst2!CASNumber & "-" & rst2!ClaimVersion) Then
                Record.SubItems(4) = (rst2!CASNumber & "-" & rst2!ClaimVersion)
            End If
            CASNumber = rst2!CASNumber
            WSVersion = rst2!ClaimVersion
            
            If Len(rst2!CLMPayableByMedix2M) Then
                Record.SubItems(5) = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
            End If
            If Len(rst2!MBMPaidAmt) Then
                Record.SubItems(6) = Format(rst2!MBMPaidAmt, "#,##0.00")
            End If
                            
            Record.SubItems(8) = IIf(Len(Trim(rst2!CASARStatus)), Trim(rst2!CASARStatus), "")
                
            If Len(rst2!PVNo) Then
                Record.SubItems(9) = Trim(rst2!PVNo)
            End If
            
            If Len(rst2!PVStatus) Then
                Record.SubItems(10) = rst2!PVStatus
            End If
            
            If Len(rst2!PVChequeNo) Then
                Record.SubItems(11) = Trim(rst2!PVChequeNo)
            End If
                            
            If Len(!PVChequeDate) Then
                Record.SubItems(12) = Trim(rst2!PVChequeDate)
            End If
            
            If Len(rst2!PVPayeeBank) Then
                Record.SubItems(13) = Trim(rst2!PVPayeeBank)
            End If

            If Len(rst2!PVPayeeACNo) Then
                Record.SubItems(14) = Trim(rst2!PVPayeeACNo)
            End If
            
            If Len(rst2!Remark) Then
                Record.SubItems(15) = Trim(rst2!Remark)
            End If
            Record.SubItems(16) = IIf(Trim(rst2!CLMWSStatus) <> "", Trim(rst2!CLMWSStatus), "")
            Record.SubItems(17) = IIf(Len(!PVSUNNo), Trim(!PVSUNNo), "")
            Record.SubItems(18) = IIf(Len(!PVSUNDate), Trim(!PVSUNDate), "")

            strsqlPAY = "Select * from Payment2MBM where CASNumber = '" & Trim(CASNumber) & "' and ClaimVersion = '" & Trim(WSVersion) & "'"
            PAY.Open strsqlPAY, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            INVAmt = 0
                
            If Trim(txtPVStatus) = "P" Then
                Do Until PAY.EOF
                    INVAmt = INVAmt + PAY!MBMPaidAmt
                PAY.MoveNext
                Loop
                If Len(INVAmt) Then
                    Record.SubItems(6) = INVAmt
                    Record.SubItems(7) = rst2!CLMPayableByMedix2M - INVAmt
                End If
            ElseIf txtPVStatus = "" Then
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(7) = rst2!CLMPayableByMedix2M
                End If
            End If
            PAY.Close
            Set PAY = Nothing
            Current = Current + 1
            lblCurRecord = Current
            
            TotalAmt = TotalAmt + rst2!CLMPayableByMedix2M
            lblSelectedAmt = Format(TotalAmt, "#,##0.00")
            
            If Len(rst2!MBMPaidAmt) Then
                TotalAmt2 = TotalAmt2 + rst2!MBMPaidAmt
                lblAmtPaid = Format(TotalAmt2, "#,##0.00")
            End If
            tmpPayableAmt = IIf(IsNumeric(!CLMPayableByMedix2M), !CLMPayableByMedix2M, 0)
            Call SelectReceipt(CASNumber, WSVersion, tmpPayableAmt, !MBMNumber)
        End With
        
        rst2.MoveNext
        
        DoEvents
            If intExit = 1 Then
               Check = False
               Exit Do
            End If
        Loop
    Loop Until Check = False

    If RECFound = False Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    End If

    intExit = 0
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub SelectReceipt(tmpCasNo As String, tmpVEr As String, tmpPayableAmt As Double, tmpMBM As String)
Dim rst2 As New ADODB.Recordset
Dim strsqlPAY As String
Dim INVAmt As Variant
Dim CASNumber As String
Dim WSVersion As String
Dim cmd1 As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim tmpAmt As Double
Dim tmpReceiptNo As String
Dim tmpReceiptDate As String
Dim tmpReceiptRemarks As String
Dim tmpReceiptAccMth As String
Dim tmpReceiptAccYr As String

    If ReceiptNoCheck(tmpCasNo & "-" & tmpVEr) = True Then
        Exit Sub
    End If
    strsqlPAY = "And CASNumber = '" & Trim(tmpCasNo) & "' and ClaimVersion='" & Trim(tmpVEr) & "'"
    Set cmd1.ActiveConnection = medixmasterconnPayment
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_Payment"
    Set prm1 = cmd1.CreateParameter("PaymentType", adVarChar, adParamInput, 2000, "Receipt")
    cmd1.Parameters.Append prm1
    Set prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, strsqlPAY)
    cmd1.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd1.Execute
    Do While Not rst2.EOF
        With rst2
            If Trim(!ReceiptNoMNRB) <> "" Then
                tmpAmt = IIf(IsNumeric(!CheckAmtMNRB), !CheckAmtMNRB, 0)
                tmpReceiptNo = IIf(Trim(!ReceiptNoMNRB) <> "", !ReceiptNoMNRB, "")
                tmpReceiptDate = IIf(Trim(!ReceiptDateMNRB) <> "", !ReceiptDateMNRB, "")
                tmpReceiptRemarks = IIf(Trim(!RemarkMNRB) <> "", !RemarkMNRB, "")
                tmpReceiptAccMth = IIf(Trim(!ReceiptAccMthMNRB) <> "", !ReceiptAccMthMNRB, "")
                tmpReceiptAccYr = IIf(Trim(!ReceiptAccYRMNRB) <> "", !ReceiptAccYRMNRB, "")
                Call ReceivedFrMBMListing("MNRB", tmpCasNo, tmpVEr, tmpPayableAmt, tmpMBM, _
                     tmpReceiptNo, tmpReceiptDate, tmpReceiptRemarks, tmpAmt, tmpReceiptAccMth, tmpReceiptAccYr)
            End If
            If Trim(!ReceiptNoPAY) <> "" Then
                tmpAmt = IIf(IsNumeric(!CheckAmtPAY), !CheckAmtPAY, 0)
                tmpReceiptNo = IIf(Trim(!ReceiptNoPAY) <> "", !ReceiptNoPAY, "")
                tmpReceiptDate = IIf(Trim(!ReceiptDatePAY) <> "", !ReceiptDatePAY, "")
                tmpReceiptRemarks = IIf(Trim(!RemarkPAY) <> "", !RemarkPAY, "")
                tmpReceiptAccMth = IIf(Trim(!ReceiptAccMthPAY) <> "", !ReceiptAccMthPAY, "")
                tmpReceiptAccYr = IIf(Trim(!ReceiptAccYRPAY) <> "", !ReceiptAccYRPAY, "")
                Call ReceivedFrMBMListing("PAY", tmpCasNo, tmpVEr, tmpPayableAmt, tmpMBM, _
                     tmpReceiptNo, tmpReceiptDate, tmpReceiptRemarks, tmpAmt, tmpReceiptAccMth, tmpReceiptAccYr)
            End If
            If Trim(!ReceiptNoMBM) <> "" Then
                tmpAmt = IIf(IsNumeric(!ReceiptAmtMBM), !ReceiptAmtMBM, 0)
                tmpReceiptNo = IIf(Trim(!ReceiptNoMBM) <> "", !ReceiptNoMBM, "")
                tmpReceiptDate = IIf(Trim(!ReceiptDateMBM) <> "", !ReceiptDateMBM, "")
                tmpReceiptRemarks = IIf(Trim(!ReceiptRemarksMBM) <> "", !ReceiptRemarksMBM, "")
                tmpReceiptAccMth = IIf(Trim(!ReceiptAccMthMBM) <> "", !ReceiptAccMthMBM, "")
                tmpReceiptAccYr = IIf(Trim(!ReceiptMBMAccYrMBM) <> "", !ReceiptMBMAccYrMBM, "")
                Call ReceivedFrMBMListing("MBM", tmpCasNo, tmpVEr, tmpPayableAmt, tmpMBM, _
                     tmpReceiptNo, tmpReceiptDate, tmpReceiptRemarks, tmpAmt, tmpReceiptAccMth, tmpReceiptAccYr)
            End If
        End With
        rst2.MoveNext
    Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub ReceivedFrMBMListing(tmpType As String, tmpCasNo As String, tmpVEr As String, tmpPayableAmt As Double, tmpMBM As String, _
ReceiptNo As String, ReceiptDate As String, ReceiptRemarks As String, tmpAmt As Double, ReceiptAccMth As String, ReceiptAccYr As String)
Dim Record As ListItem
    Set Record = lstReceiptMBM.ListItems.Add(, , tmpCasNo & "-" & tmpVEr)
    Record.SubItems(1) = IIf(Trim(tmpMBM) <> "", Trim(tmpMBM), "")
    Record.SubItems(2) = tmpType
    Record.SubItems(4) = Format(tmpPayableAmt, "#,##0.00")
    Record.SubItems(5) = Format(tmpAmt, "#,##0.00")
    Record.SubItems(7) = IIf(Trim(ReceiptNo) <> "", Trim(ReceiptNo), "")
    If Len(ReceiptDate) Then
        Record.SubItems(8) = Format(ReceiptDate, "dd/mm/yyyy")
    End If
                
    If Trim(ReceiptAccMth) <> "" Or Trim(ReceiptAccYr) <> "" Then
        Record.SubItems(9) = ReceiptAccMth & "/" & ReceiptAccYr
    End If
    Record.SubItems(10) = IIf(Trim(ReceiptRemarks) <> "", Trim(ReceiptRemarks), "")
End Sub

Private Sub Payment2MBMReissueListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim Check As Boolean
Dim CasXRec As New ADODB.Recordset
Dim CASXSql As String
Dim CASNumber As String
Dim ClmVer As String
Dim PatientName As String
Dim Add1 As String
Dim Add2 As String
Dim Add3 As String
Dim PostCode As String
Dim CityCode As String
Dim StateCode As String
Dim WsDate As String
Dim BordxRef As String
Dim cmd1 As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
   
    Set cmd1.ActiveConnection = medixmasterconnPayment
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_Payment"
    Set prm1 = cmd1.CreateParameter("PaymentType", adVarChar, adParamInput, 2000, "Pay2MBM")
    cmd1.Parameters.Append prm1
    Set prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, strAppend)
    cmd1.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd1.Execute

    Do
        Do Until rst2.EOF
            With rst2
                Set Record = lstPaymentMBMReissue.ListItems.Add(, , IIf(Len(PatientName), PatientName, ""))
                If Len(!CASNumber & "-" & !ClaimVersion) Then
                    Record.SubItems(1) = (!CASNumber & "-" & !ClaimVersion)
                End If
                
                Record.SubItems(2) = IIf(IsNumeric(!MBMPaidAmt), !MBMPaidAmt, 0)
                Record.SubItems(3) = !PVMBMIndex
                If Len(!PVNo) Then
                    Record.SubItems(4) = Trim(!PVNo)
                End If
                If IsDate(!PVDate) Then
                    Record.SubItems(5) = !PVDate
                End If
                
                Record.SubItems(6) = IIf(Len(Trim(!PVChequeNo)), Trim(!PVChequeNo), "")
                If Len(!PVChequeDate) Then
                    Record.SubItems(7) = Trim(!PVChequeDate)
                End If
                Record.SubItems(8) = IIf(IsNumeric(!PVChequeAmt), !PVChequeAmt, 0)
                Record.SubItems(9) = IIf(Len(Trim(!PVPayTo)), Trim(!PVPayTo), "")
                Record.SubItems(10) = IIf(Len(Trim(!PVPayeeBank)), Trim(!PVPayeeBank), "")
                Record.SubItems(11) = IIf(Len(Trim(!PVPayeeACNo)), Trim(!PVPayeeACNo), "")
                Record.SubItems(12) = IIf(IsNumeric(!PVBankCharges), Trim(!PVBankCharges), 0)
                If Len(!PAYAccPeriodMth & !PAYAccPeriodYr) Then
                    Record.SubItems(13) = Format((!PAYAccPeriodMth & "/" & !PAYAccPeriodYr), "mm/yyyy")
                End If
                Record.SubItems(14) = IIf(Len(Trim(!Pay2MBMLastUpdateUser)), Trim(!Pay2MBMLastUpdateUser), "")
                Record.SubItems(15) = IIf(Len(Trim(!Remark)), Trim(!Remark), "")
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
End Sub

Private Function GetReceiptNO(CASNumber, ClmVer, ReceiptNo As String)
Dim REceiptRec As New ADODB.Recordset
Dim REceiptSql As String

    REceiptSql = "select RECEIPTNO from ReceiptfrMNRB where CASNumber = '" & Trim(CASNumber) & "' And ClaimVersion =  '" & Trim(ClmVer) & "'"
    REceiptRec.Open REceiptSql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    If REceiptRec.recordCount Then
        ReceiptNo = REceiptRec!ReceiptNo
    End If
    REceiptRec.Close
    Set REceiptRec = Nothing
End Function

'**********************End List Record*******************************


'Private Sub lstPaymentMBM_Click()
'Dim rst3 As New ADODB.Recordset
'Dim strsql As String

'If one = True Then

'CmdSave.Enabled = False


'If lstPaymentMBM.listitems.count Then

    'strsql = " SELECT * From Payment2MBM Where PVMBMIndex = '" & lstPaymentMBM.SelectedItem.SubItems(12) & "'"
    'rst3.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
    'If rst3.recordCount = 0 Then
        'MsgBox "Payment hasn't been made!", vbOKOnly + vbCritical, DialogBoxTitle
    'Else
        'TxtPayTo.SetFocus
        'With rst3
            'If Len(!PVPayTo) Then
                'TxtPayTo = !PVPayTo
            'End If
            'If Len(!PVPayeeBank) Then
                'TxtBank = !PVPayeeBank
            'End If
            'If Len(!PVPayeeACNo) Then
                'TxtAccountNo = !PVPayeeACNo
            'End If
            'If Len(!PVNo) Then
                'txtPVNo = !PVNo
            'End If
            'If Len(!PVDate) Then
                'txtPVDate = !PVDate
            'End If
            'If Len(!PVChequeDate) Then
                'txtChequeDate = !PVChequeDate
            'End If
            'If Len(!PVChequeNo) Then
                'txtChequeNo = !PVChequeNo
            'End If
            'If Len(!MBMPaidAmt) Then
                'txtChequeAmt = !MBMPaidAmt
            'End If
            'If Len(!Remark) Then
                'TxtRemark = !Remark
            'End If
            'If Len(!PAYAccPeriodMth) Then
                'CboACPeriodMth = !PAYAccPeriodMth
            'End If
            'If Len(!PAYAccPeriodYr) Then
                'CboACPeriodYr = !PAYAccPeriodYr
            'End If
            'If Len(!PVBankCharges) Then
                'txtBankCharges = !PVBankCharges
            'Else
                'txtBankCharges = ""
            'End If
        'End With
        'CmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
'End If
'End If
'End Sub

Private Sub txtCASNo_LostFocus()
txtCASNo = ChkStr(txtCASNo)
End Sub

'*************************** Start Diable Non Integer Value

Private Sub lstPaymentMBM_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtDatetime, m_blnDirection(13)
    Case 14
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstPaymentMBM, ColumnHeader.Index, ldtString, m_blnDirection(16)
    End Select
End Sub

Private Sub lstReceiptMBM_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtDatetime, m_blnDirection(9)
    Case 10
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(11)
    End Select
End Sub

Private Sub lstPaymentMBMReissue_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtDatetime, m_blnDirection(6)
    Case 7
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtDatetime, m_blnDirection(8)
    Case 9
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    Case 15
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtDatetime, m_blnDirection(16)
    Case 17
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtDatetime, m_blnDirection(18)
    Case 19
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(21)
    Case 22
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(22)
    Case 23
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
    Case 24
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
    Case 25
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
    Case 26
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(26)
    Case 27
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(27)
    Case 28
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtDatetime, m_blnDirection(28)
    Case 29
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(29)
    Case 30
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(30)
    Case 31
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(31)
    Case 32
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(32)
    Case 33
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(33)
    Case 34
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(34)
    Case 35
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(35)
    Case 36
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(36)
    Case 37
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtDatetime, m_blnDirection(37)
    Case 38
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(38)
    Case 39
        SortListView lstPaymentMBMReissue, ColumnHeader.Index, ldtString, m_blnDirection(39)
    End Select
End Sub

'***********************Start EntriesEnable***************************
'Private Sub EntriesEnable(status As Boolean)

'Dim ctl As Control

    'For Each ctl In Me.Controls
        'If TypeOf ctl Is TextBox Then
            'ctl.Enabled = status
        'ElseIf TypeOf ctl Is ComboBox Then
            'ctl.Enabled = status
        'End If
    'Next ctl
    
'End Sub
'***********************Start EntriesEnable***************************

Private Sub cboPayor_Change()
If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If
End Sub

Private Sub cboPayor_KeyPress(KeyAscii As Integer)
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

Private Sub txtCRAmtFr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtCRAmtTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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

Private Function CaseNoCheck()
    If Trim(ChkStr(txtCASNo)) = Trim(Mid(lstPaymentMBM.ListItems(listloop1).SubItems(4), 1, IIf(InStr(1, lstPaymentMBM.ListItems(listloop1).SubItems(4), "-") = 0, 10, InStr(1, lstPaymentMBM.ListItems(listloop1).SubItems(4), "-") - 1))) Then
        Check = True
    End If
End Function

Private Function ReceiptNoCheck(tmpCasNo As String) As Boolean
Dim AA As Integer
    ReceiptNoCheck = False
    If lstReceiptMBM.ListItems.Count > 0 Then
        For AA = 1 To lstReceiptMBM.ListItems.Count
            If Trim(tmpCasNo) = lstReceiptMBM.ListItems(AA).Text Then
                ReceiptNoCheck = True
                Exit For
            End If
        Next AA
    End If
End Function


