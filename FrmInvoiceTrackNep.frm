VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInvoiceTrackNep 
   Caption         =   "Invoice Tracking (Jupiter)"
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
      Height          =   3345
      Left            =   0
      TabIndex        =   58
      Top             =   720
      Width           =   11895
      Begin VB.CheckBox chkBordxDate 
         Height          =   255
         Left            =   11280
         TabIndex        =   40
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox ChkGrpCom 
         Height          =   255
         Left            =   5115
         TabIndex        =   14
         Top             =   1650
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom2 
         Height          =   255
         Left            =   5400
         TabIndex        =   60
         Top             =   1650
         Width           =   255
      End
      Begin VB.CheckBox ChkHos2 
         Height          =   255
         Left            =   5400
         TabIndex        =   59
         Top             =   1100
         Width           =   255
      End
      Begin VB.CheckBox ChkVersion 
         Height          =   255
         Left            =   5115
         TabIndex        =   6
         Top             =   495
         Width           =   375
      End
      Begin VB.TextBox TxtWSNo2 
         Height          =   310
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   2
         Top             =   225
         Width           =   1695
      End
      Begin VB.TextBox TxtWSNo1 
         Height          =   310
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   1
         Top             =   225
         Width           =   1725
      End
      Begin VB.CheckBox ChkBordxNo 
         Height          =   255
         Left            =   5115
         TabIndex        =   23
         Top             =   2480
         Width           =   375
      End
      Begin VB.CheckBox ChkCaseNo 
         Height          =   255
         Left            =   5115
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkInvNo 
         Height          =   255
         Left            =   5115
         TabIndex        =   17
         Top             =   1920
         Width           =   375
      End
      Begin VB.CheckBox ChkPayor 
         Height          =   255
         Left            =   5115
         TabIndex        =   12
         Top             =   1370
         Width           =   375
      End
      Begin VB.CheckBox ChkHos 
         Height          =   255
         Left            =   5115
         TabIndex        =   10
         Top             =   1100
         Width           =   255
      End
      Begin VB.CheckBox ChkPayTo 
         Height          =   255
         Left            =   5115
         TabIndex        =   8
         Top             =   825
         Width           =   375
      End
      Begin VB.CheckBox ChkInvAmt 
         Height          =   255
         Left            =   5115
         TabIndex        =   20
         Top             =   2200
         Width           =   375
      End
      Begin VB.CheckBox chkDOA 
         Height          =   255
         Left            =   11280
         TabIndex        =   28
         Top             =   255
         Width           =   375
      End
      Begin VB.CheckBox chkDOD 
         Height          =   255
         Left            =   11280
         TabIndex        =   31
         Top             =   555
         Width           =   375
      End
      Begin VB.CheckBox chkInvDate 
         Height          =   255
         Left            =   11280
         TabIndex        =   34
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox chkInvReg 
         Height          =   255
         Left            =   11280
         TabIndex        =   37
         Top             =   1120
         Width           =   375
      End
      Begin VB.TextBox TxtWSVersion1 
         Height          =   310
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   4
         Top             =   480
         Width           =   1725
      End
      Begin VB.TextBox TxtWSVersion2 
         Height          =   310
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox CboPayTo 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Text            =   "H - Hospital"
         Top             =   770
         Width           =   3855
      End
      Begin VB.ComboBox CboHos 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1050
         Width           =   3855
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   1330
         Width           =   3855
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   1620
         Width           =   3855
      End
      Begin VB.TextBox TxtInvNo1 
         Height          =   310
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1900
         Width           =   1815
      End
      Begin VB.TextBox TxtInvNo2 
         Height          =   310
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   16
         Top             =   1900
         Width           =   1575
      End
      Begin VB.TextBox txtINVAmtFr 
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   2190
         Width           =   1815
      End
      Begin VB.TextBox txtINVAmtTo 
         Height          =   285
         Left            =   3480
         TabIndex        =   19
         Top             =   2190
         Width           =   1575
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   310
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   21
         Top             =   2450
         Width           =   1815
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   310
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   22
         Top             =   2450
         Width           =   1575
      End
      Begin VB.ComboBox CboStayType 
         Height          =   315
         Left            =   1200
         TabIndex        =   24
         Top             =   2730
         Width           =   3855
      End
      Begin VB.ComboBox CboPayType 
         Height          =   315
         Left            =   1200
         TabIndex        =   25
         Top             =   3000
         Width           =   3855
      End
      Begin MSMask.MaskEdBox txtDOADate1 
         Height          =   315
         Left            =   7560
         TabIndex        =   26
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate2 
         Height          =   315
         Left            =   9720
         TabIndex        =   27
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate1 
         Height          =   315
         Left            =   7560
         TabIndex        =   29
         Top             =   525
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate2 
         Height          =   315
         Left            =   9720
         TabIndex        =   30
         Top             =   525
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   7560
         TabIndex        =   32
         Top             =   810
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
         Left            =   9720
         TabIndex        =   33
         Top             =   810
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtIRDDate2 
         Height          =   315
         Left            =   9720
         TabIndex        =   36
         Top             =   1100
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtIRDDate1 
         Height          =   315
         Left            =   7560
         TabIndex        =   35
         Top             =   1100
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   315
         Left            =   9720
         TabIndex        =   39
         Top             =   1380
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   315
         Left            =   7560
         TabIndex        =   38
         Top             =   1380
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label24 
         Caption         =   "Bordx Date"
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   92
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label23 
         Caption         =   "To"
         Height          =   255
         Index           =   1
         Left            =   9240
         TabIndex        =   91
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label49 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label33 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   135
         Left            =   5400
         TabIndex        =   86
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label52 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   135
         Left            =   5160
         TabIndex        =   85
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   3000
         TabIndex        =   84
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label31 
         Caption         =   "Version"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label47 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   82
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   81
         Top             =   1980
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "PayType"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "Inv Reg"
         Height          =   255
         Index           =   0
         Left            =   6600
         TabIndex        =   79
         Top             =   1155
         Width           =   900
      End
      Begin VB.Label Label23 
         Caption         =   "To"
         Height          =   255
         Index           =   0
         Left            =   9240
         TabIndex        =   78
         Top             =   1150
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "DOA"
         Height          =   255
         Left            =   6600
         TabIndex        =   75
         ToolTipText     =   "Date of Admission"
         Top             =   320
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "DOD"
         Height          =   255
         Left            =   6600
         TabIndex        =   74
         ToolTipText     =   "Date of Discharged"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Inv Stay Type"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   2800
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   72
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   71
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "To"
         Height          =   255
         Left            =   9240
         TabIndex        =   70
         Top             =   850
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Inv Date"
         Height          =   255
         Left            =   6600
         TabIndex        =   69
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Pay To"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Clm No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   67
         Top             =   280
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Inv No"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   3000
         TabIndex        =   64
         Top             =   285
         Width           =   255
      End
      Begin VB.Label Label40 
         Caption         =   "Inv Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   2250
         Width           =   855
      End
      Begin VB.Label Label41 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   62
         Top             =   2250
         Width           =   375
      End
      Begin VB.Label Label18 
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
         Left            =   11280
         TabIndex        =   61
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Height          =   550
      Left            =   0
      TabIndex        =   56
      Top             =   240
      Width           =   11895
      Begin VB.ComboBox cboRPTSelection 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   150
         Width           =   4095
      End
      Begin VB.Label Label25 
         Caption         =   "Rpt Selection"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   57
         Top             =   195
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   43
      Top             =   7080
      Width           =   11775
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3960
         TabIndex        =   94
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   93
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   7080
         TabIndex        =   51
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10200
         TabIndex        =   50
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   49
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9240
         TabIndex        =   48
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   330
         Left            =   8640
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   6360
         TabIndex        =   41
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   330
         Left            =   7800
         TabIndex        =   46
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   330
         Left            =   8520
         TabIndex        =   45
         Top             =   200
         Width           =   735
      End
      Begin VB.TextBox txtNoofDay 
         Height          =   285
         Left            =   2520
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1560
         TabIndex        =   55
         Top             =   195
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1440
         TabIndex        =   54
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3120
         TabIndex        =   53
         Top             =   195
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1200
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lstInvoiceTrackList 
      Height          =   2985
      Left            =   0
      TabIndex        =   42
      Top             =   4080
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5265
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
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
         Alignment       =   2
         SubItemIndex    =   11
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label48 
      Caption         =   "* - Empty Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   90
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "* - Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   89
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Invoice  Tracking (Jupiter)"
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
      Height          =   255
      Left            =   0
      TabIndex        =   88
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmInvoiceTrackNep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Public searchCriteria As String
Private HOspitals As String
Private GrpCompany As String
Dim ScanCode As String
Private AA As Integer
Private m_blnDirection(1 To 15) As Boolean

Private Sub cboRPTSelection_Click()
    Call ListViewHeader
End Sub

Private Sub ListViewHeader()
    lstInvoiceTrackList.ListItems.Clear
    If Mid(CboRptSelection, 1, 1) = "0" Then
        lstInvoiceTrackList.ColumnHeaders(1).Text = "InvMonth"
        lstInvoiceTrackList.ColumnHeaders(1).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(2).Text = "Total Cases"
        lstInvoiceTrackList.ColumnHeaders(2).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(3).Text = "Payable2Hosp"
        lstInvoiceTrackList.ColumnHeaders(3).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(4).Text = "Payable2MBM"
        lstInvoiceTrackList.ColumnHeaders(4).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(4).Width = 0
        lstInvoiceTrackList.ColumnHeaders(5).Text = "Inv Amount"
        lstInvoiceTrackList.ColumnHeaders(5).Alignment = lvwColumnRight
        For AA = 6 To 15
            lstInvoiceTrackList.ColumnHeaders(AA).Text = ""
        Next AA
    ElseIf Mid(CboRptSelection, 1, 1) = "1" Then
        lstInvoiceTrackList.ColumnHeaders(1).Text = "Hosp Code"
        lstInvoiceTrackList.ColumnHeaders(1).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(2).Text = "Hospital"
        lstInvoiceTrackList.ColumnHeaders(2).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(3).Text = "InvMonth"
        lstInvoiceTrackList.ColumnHeaders(3).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(4).Text = "Total Cases"
        lstInvoiceTrackList.ColumnHeaders(4).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(4).Width = 1000
        lstInvoiceTrackList.ColumnHeaders(5).Text = "Payable2Hosp"
        lstInvoiceTrackList.ColumnHeaders(5).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(6).Text = "Payable2MBM"
        lstInvoiceTrackList.ColumnHeaders(6).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(6).Width = 0
        lstInvoiceTrackList.ColumnHeaders(7).Text = "Inv Amount"
        lstInvoiceTrackList.ColumnHeaders(7).Alignment = lvwColumnRight
        For AA = 8 To 15
            lstInvoiceTrackList.ColumnHeaders(AA).Text = ""
        Next AA
    
    ElseIf Mid(CboRptSelection, 1, 1) = "2" Then
        lstInvoiceTrackList.ColumnHeaders(1).Text = "Hosp Code"
        lstInvoiceTrackList.ColumnHeaders(1).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(2).Text = "Hospital"
        lstInvoiceTrackList.ColumnHeaders(2).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(3).Text = "Inv. No"
        lstInvoiceTrackList.ColumnHeaders(3).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(4).Text = "Inv. Date"
        lstInvoiceTrackList.ColumnHeaders(4).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(4).Width = 1000
        lstInvoiceTrackList.ColumnHeaders(5).Text = "Payable2Hosp"
        lstInvoiceTrackList.ColumnHeaders(5).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(6).Text = "Payable2MBM"
        lstInvoiceTrackList.ColumnHeaders(6).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(6).Width = 0
        lstInvoiceTrackList.ColumnHeaders(7).Text = "Inv Amount"
        lstInvoiceTrackList.ColumnHeaders(7).Alignment = lvwColumnRight
        For AA = 8 To 15
            lstInvoiceTrackList.ColumnHeaders(AA).Text = ""
        Next AA
    ElseIf Mid(CboRptSelection, 1, 1) = "3" Then
        lstInvoiceTrackList.ColumnHeaders(1).Text = "Worksheet No"
        lstInvoiceTrackList.ColumnHeaders(1).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(2).Text = "Bordx No"
        lstInvoiceTrackList.ColumnHeaders(2).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(3).Text = "Bordx Date"
        lstInvoiceTrackList.ColumnHeaders(3).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(4).Text = "Inv. No"
        lstInvoiceTrackList.ColumnHeaders(4).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(4).Width = 1000
        lstInvoiceTrackList.ColumnHeaders(5).Text = "Inv. Date"
        lstInvoiceTrackList.ColumnHeaders(5).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(6).Text = "Inv Amt"
        lstInvoiceTrackList.ColumnHeaders(6).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(6).Width = 1000
        lstInvoiceTrackList.ColumnHeaders(7).Text = "DOR"
        lstInvoiceTrackList.ColumnHeaders(7).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(8).Text = "DOA"
        lstInvoiceTrackList.ColumnHeaders(8).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(9).Text = "DOD"
        lstInvoiceTrackList.ColumnHeaders(9).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(10).Text = "Payee"
        lstInvoiceTrackList.ColumnHeaders(10).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(11).Text = "Group Company"
        lstInvoiceTrackList.ColumnHeaders(11).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(12).Text = "Inv. Reg"
        lstInvoiceTrackList.ColumnHeaders(12).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(13).Text = "Inv. Entry By"
        lstInvoiceTrackList.ColumnHeaders(13).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(14).Text = "Stay Type"
        lstInvoiceTrackList.ColumnHeaders(14).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(15).Text = "Pay Type"
        lstInvoiceTrackList.ColumnHeaders(15).Alignment = lvwColumnLeft
    ElseIf Mid(CboRptSelection, 1, 1) = "4" Then
        lstInvoiceTrackList.ColumnHeaders(1).Text = "Hosp Code"
        lstInvoiceTrackList.ColumnHeaders(1).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(2).Text = "Hospital"
        lstInvoiceTrackList.ColumnHeaders(2).Alignment = lvwColumnLeft
        lstInvoiceTrackList.ColumnHeaders(3).Text = "Payable2Hosp"
        lstInvoiceTrackList.ColumnHeaders(3).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(4).Text = "Payable2MBM"
        lstInvoiceTrackList.ColumnHeaders(4).Alignment = lvwColumnRight
        lstInvoiceTrackList.ColumnHeaders(4).Width = 0
        lstInvoiceTrackList.ColumnHeaders(5).Text = "Inv Amount"
        lstInvoiceTrackList.ColumnHeaders(5).Alignment = lvwColumnRight
        For AA = 8 To 15
            lstInvoiceTrackList.ColumnHeaders(AA).Text = ""
        Next AA
        
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
    intExit = 0
    lblCurrent = 0
    'lblTotal = 0
    LblTotalAmt = 0
    medixneptuneconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
    medixneptuneconn.Close
    Set medixneptuneconn = Nothing
End Sub

'************************Start Command Button**************************
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lstInvoiceTrackList.ListItems.Clear
    lblCurrent = 0
    LblTotalAmt = 0
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
End Sub

Public Sub CmdSearch_Click()
On Error GoTo errRun
    ScanCode = TxtWSNo1
    CmdSearch.Enabled = False
    If Mid(ScanCode, 1, 2) = "LO" Then
        TxtWSNo1 = ""
        Call SearchForm(ScanCode)
    Else
        CmdClear.Enabled = False
        CmdPrintFile.Enabled = False
        CmdExit.Enabled = False
        Screen.MousePointer = vbHourglass
        'lstInvoiceTrackList.Height = 6705
        'lstInvoiceTrackList.Left = 0
        'lstInvoiceTrackList.Top = 360
        'lstInvoiceTrackList.Width = 11775
        'Frame1.Visible = False
        medixneptuneconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        medixneptuneconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

        Call SearchRecord
        
        medixneptuneconn.Close
        Set medixneptuneconn = Nothing
        medixneptuneconnWS.Close
        Set medixneptuneconnWS = Nothing

        Screen.MousePointer = Default
    End If
    CmdSearch.Enabled = True
Exit Sub
errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstInvoiceTrackList.ListItems.Clear
    lblCurrent = ""
    CmdPrintFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    medixneptuneconn.Close
    Set medixneptuneconn = Nothing
    medixneptuneconnWS.Close
    Set medixneptuneconnWS = Nothing
    
End Sub

Private Sub CmdBack_Click()
    lstInvoiceTrackList.Height = 2865
    lstInvoiceTrackList.Left = 0
    lstInvoiceTrackList.Top = 4200
    lstInvoiceTrackList.Width = 11775
    Frame1.Visible = True
End Sub

'Private Sub CmdPrint_Click()
    'Screen.MousePointer = vbHourglass
        
    'If lstInvoiceTrackList.listitems.count <> 0 Then
        'RptInvoiceTrack.show
    'Else
        'MsgBox "Report cannot be printed without any record.", vbInformation
    'End If
    
    'Screen.MousePointer = Default
'End Sub

Private Sub CmdPrintFile_Click()
    Screen.MousePointer = vbHourglass
    
    If lstInvoiceTrackList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstInvoiceTrackList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = Default
End Sub

Private Sub CmdView_Click()
    lstInvoiceTrackList.Height = 6705
    lstInvoiceTrackList.Left = 0
    lstInvoiceTrackList.Top = 360
    lstInvoiceTrackList.Width = 11775
    Frame1.Visible = False
End Sub
'************************End Command Button**************************


'************************Start sorted lstInvoiceTrackList**************************

Private Sub lstInvoiceTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    If Mid(CboRptSelection, 1, 1) = 0 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtnumeric, m_blnDirection(1)
        Case 2
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
    ElseIf Mid(CboRptSelection, 1, 1) = 1 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
    ElseIf Mid(CboRptSelection, 1, 1) = 2 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
    ElseIf Mid(CboRptSelection, 1, 1) = 3 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
        Case 11
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
        Case 12
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(12)
        Case 13
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(15)
        End Select
        
    End If
End Sub
'************************End sorted lstInvoiceTrackList**************************

'************************Start ComboListing**************************
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cnn As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim sql As String
Dim yearstart
Dim monthstart

'PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixneptuneconn, adOpenForwardOnly, adLockBatchOptimistic
'INS.Open "Select INSCode, INSDescription from INS", medixneptuneconn, adOpenForwardOnly, adLockBatchOptimistic
        
    Set cmd.ActiveConnection = medixneptuneconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute

    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
     
    With cboPayTo
        .AddItem "H - Hospital"
        .AddItem "M - Member"
    End With
    
    With CboStayType
        .AddItem "0 - In-Patient"
        .AddItem "1 - Out-Patient"
        .AddItem "2 - Follow-up"
        .AddItem "3 - Pre-Hospitalisation"
        .AddItem "4 - Follow-up & Pre-Hospitalisation"
    End With
    
    With cboPayType
        .AddItem "C - Cashless"
        .AddItem "R - Reimbursement"
    End With
        
   
    CboRptSelection.AddItem "0 - Invoive Month(All Hospitals)"
    CboRptSelection.AddItem "1 - Invoive Month(By Hospitals)"
    CboRptSelection.AddItem "2 - General Listing"
    CboRptSelection.AddItem "3 - General Listing(Detail)"
    CboRptSelection.AddItem "4 - Summary(By Hospitals)"
End Sub
'************************Start ComboListing**************************

'************************Start Search Record**************************

Private Function SearchRecord()
Dim HOS As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim StrSql As String
Dim StrAppend As String
Dim strsqlHOS As String
Dim sql As String
Dim sql2 As String
Dim Bordx As Integer
    
    StrAppend = ""
    HOspitals = ""
    GrpCompany = ""
    
    If CboRptSelection = "" Then
        MsgBox "Please select a Report.", vbOKOnly + vbCritical, DialogBoxTitle
        CboRptSelection.SetFocus
        cmdPrint.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
        Exit Function
    End If
    

     If Len(Trim(txtInvNo1)) Then
         StrAppend = StrAppend + " AND InvNo >= '" & Trim(txtInvNo1) & "'"
     End If
 
     If Len(Trim(txtInvNo2)) Then
         StrAppend = StrAppend + " AND InvNo <= '" & Trim(txtInvNo2) & "'"
     End If
 
     If ChkInvNo.Value = 1 Then
         sql = ""
         sql = "AND (INVNo Is null OR INVNo = '')" & ""
         StrAppend = StrAppend + sql
     End If
     
     If Mid(CboStayType, 1, 1) = "0" Then
         StrAppend = StrAppend + " AND INVStayType = 'I'"
     End If
     
     If Mid(CboStayType, 1, 1) = "1" Then
         StrAppend = StrAppend + " AND INVStayType = 'O'"
     End If
     
     If Mid(CboStayType, 1, 1) = "2" Then
         StrAppend = StrAppend + " AND INVStayType = 'F'"
     End If
     
     If Mid(CboStayType, 1, 1) = "3" Then
         StrAppend = StrAppend + " AND INVStayType = 'P'"
     End If
     
     If Mid(CboStayType, 1, 1) = "4" Then
         StrAppend = StrAppend + " AND (INVStayType = 'P' OR INVStayType = 'F'"
     End If
     
     If Len(Trim(Mid(cboPayType, 1, 1))) Then
         StrAppend = StrAppend + " AND INVPAYType = '" & Mid(cboPayType, 1, 1) & "'"
     End If
     
     If Len(Trim(TxtWSVersion1)) Then
         StrAppend = StrAppend + " AND INVWSVersion >= " & CDbl(TxtWSVersion1)
     End If
     
     If Len(Trim(TxtWSVersion2)) Then
         StrAppend = StrAppend + " AND INVWSVersion <= " & CDbl(TxtWSVersion2)
     End If
     
     If ChkVersion.Value = 1 Then
         sql = ""
         sql = "AND (INVWSVersion Is null OR INVWSVersion = '')" & ""
         StrAppend = StrAppend + sql
     End If
         
     If Len(Trim(cboPayTo)) Then
         StrAppend = StrAppend + " AND INVPayTo = '" & Trim(Mid(cboPayTo, 1, IIf(InStr(1, cboPayTo, "-") = 0, 4, InStr(1, cboPayTo, "-") - 1))) & "'"
     End If
     
     If ChkPayTo.Value = 1 Then
         sql = ""
         sql = "AND (INVPayTo Is null OR INVPayTo = '')" & ""
         StrAppend = StrAppend + sql
     End If
     
     If Len(Trim(TxtBordx1)) Then
         StrAppend = StrAppend + " AND INVBordxRefNo >= '" & Trim(TxtBordx1) & "'"
     End If
     
     If Len(Trim(TxtBordx2)) Then
         StrAppend = StrAppend + " AND INVBordxRefNo <= '" & Trim(TxtBordx2) & "'"
     End If
            
     If ChkBordxNo.Value = 1 Then
         sql = ""
         sql = "AND (INVBordxRefNo Is null OR INVBordxRefNo = '')" & ""
         StrAppend = StrAppend + sql
     End If
     
     If ChkInvDate.Value = 1 Then
         sql = ""
         sql = "AND (InvDate Is null OR InvDate = '')" & ""
         StrAppend = StrAppend + sql
     End If
     
     If chkInvReg.Value = 1 Then
         sql = ""
         sql = "AND (INVEnteredDate Is null OR INVEnteredDate = '')" & ""
         StrAppend = StrAppend + sql
     End If
     
    
     If Len(Trim(txtIRDDate1)) > 0 And IsDate(txtIRDDate1) = True Then
         StrAppend = StrAppend + " And INVEnteredDate >= '" & Format(txtIRDDate1, "mm/dd/yyyy") & "'"
     End If
     
     If Len(Trim(txtIRDDate2)) > 0 And IsDate(txtIRDDate2) = True Then
         StrAppend = StrAppend + " And INVEnteredDate <= '" & Format(txtIRDDate2, "mm/dd/yyyy") & "'"
     End If
     
     If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
         StrAppend = StrAppend + " And INVDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
     End If
     
     If Len(Trim(txtINVDate2)) > 0 And IsDate(txtINVDate2) = True Then
         StrAppend = StrAppend + " And INVDate <= '" & Format(txtINVDate2, "mm/dd/yyyy") & "'"
     End If
    
    
    If Len(Trim(CboHos)) Then
        StrAppend = StrAppend + " AND HOSName = '" & Trim(CboHos) & "'"
    End If
    
    If Len(Trim(CboGrpCom)) Then
       GrpCompany = Replace(CboGrpCom, "'", "''")
       StrAppend = StrAppend + " AND MBMGroupCompany = '" & Trim(GrpCompany) & "'"
    End If
    If ChkGrpCom.Value = 1 Then
       StrAppend = StrAppend + " AND (MBMGroupCompany IS NULL OR MBMGroupCompany = '')"
    End If
    
    If Len(Trim(txtINVAmtFr)) Then
        StrAppend = StrAppend + " AND INVAmount >= " & Trim(txtINVAmtFr) & ""
    End If
    
    If Len(Trim(txtINVAmtTo)) Then
        StrAppend = StrAppend + " AND INVAmount <= " & Trim(txtINVAmtTo) & ""
    End If
    
    If ChkInvAmt.Value = 1 Then
        sql = ""
        sql = "AND (INVAmount Is null)" & ""
        StrAppend = StrAppend + sql
    End If
           
    If Len(Trim(CboPayor)) Then
        StrAppend = StrAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If chkPayor.Value = 1 Then
        sql = ""
        sql = "AND (PAYCode Is null OR PAYCode = '')" & ""
        StrAppend = StrAppend + sql
    End If
    
    If Len(Trim(TxtWSNo1)) Then
        StrAppend = StrAppend + " AND CASNumber >= '" & Trim(TxtWSNo1) & "'"
    End If
    
    If Len(Trim(TxtWSNo2)) Then
        StrAppend = StrAppend + " AND CASNumber <= '" & Trim(TxtWSNo2) & "'"
    End If
    
    If ChkCaseNo.Value = 1 Then
        sql = ""
        sql = "AND (CASNumber Is null OR CASNumber = '')" & ""
        StrAppend = StrAppend + sql
    End If
    
    If ChkDOA.Value = 1 Then
        sql = ""
        sql = "AND (CasDateAdmitted Is null OR CasDateAdmitted = '')" & ""
        StrAppend = StrAppend + sql
    End If
    
    If ChkDOD.Value = 1 Then
        sql = ""
        sql = "AND (CasDischargeDate Is null OR CasDischargeDate = '')" & ""
        StrAppend = StrAppend + sql
    End If
        
    If Len(Trim(txtDOADate1)) > 0 And IsDate(txtDOADate1) = True Then
        StrAppend = StrAppend + " And CasDateAdmitted >= '" & Format(txtDOADate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOADate2)) > 0 And IsDate(txtDOADate2) = True Then
        StrAppend = StrAppend + " And CasDateAdmitted <= '" & Format(txtDOADate2, "mm/dd/yyyy") & "'"
    End If
        
    If Len(Trim(txtDODDate1)) > 0 And IsDate(txtDODDate1) = True Then
        StrAppend = StrAppend + " And CasDischargeDate >= '" & Format(txtDODDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDODDate2)) > 0 And IsDate(txtDODDate2) = True Then
        StrAppend = StrAppend + " And CasDischargeDate <= '" & Format(txtDODDate2, "mm/dd/yyyy") & "'"
    End If
    
    If chkBordxDate.Value = 1 Then
        sql = ""
        sql = "AND (BordxDate Is null OR BordxDate = '')" & ""
        StrAppend = StrAppend + sql
    End If
        
    If Len(Trim(txtBordxDateFr)) > 0 And IsDate(txtBordxDateFr) = True Then
        StrAppend = StrAppend + " And BordxDate >= '" & Format(txtBordxDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtBordxDateTo)) > 0 And IsDate(txtBordxDateTo) = True Then
        StrAppend = StrAppend + " And BordxDate <= '" & Format(txtBordxDateTo, "mm/dd/yyyy") & "'"
    End If
    

  
    searchCriteria = StrAppend
    If Mid(CboRptSelection, 1, 1) = "0" Then
        Call InvByMthCLM(StrAppend)
    ElseIf Mid(CboRptSelection, 1, 1) = "1" Then
        Call InvByMthHospCLM(StrAppend)
    ElseIf Mid(CboRptSelection, 1, 1) = "2" Then
        Call InvGenListingCLM(StrAppend)
    ElseIf Mid(CboRptSelection, 1, 1) = "3" Then
        Call InvTrackListingCLM(StrAppend)
    ElseIf Mid(CboRptSelection, 1, 1) = "4" Then
        Call InvByHosp(StrAppend)

    End If
End Function
'************************End Search Record**************************

'************************Start List Record**************************
Private Sub InvTrackListingCLM(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
Dim TotAmt As Variant
Dim TotalInvNo As String
    
    Current = 0
    TotAmt = 0
    LblTotalAmt = ""
    lstInvoiceTrackList.ListItems.Clear
            
    'Sql = " Select CASNumber, InvWSVersion, MBMNumber, InvNo, InvDate, " & _
          " InvAmount, CASDischargeDate, CaseRegDate, CASDRD, " & _
          " CASDateAdmitted, CASTokenDate, MBMPatientCovID, FinalDiag, " & _
          " FirstDiag, MBMCName, INVPayTo, INVPayee, InvEnteredDate, " & _
          " InvEnteredBy, CASStatus, InvWSed, ClaimStatus, MBMGroupCompany, " & _
          " CASAdmissionReason, ClaimBordxStatus, INVBordxRefNo, MBMName, " & _
          " INVfolUpDate, CASClaimability, INVPaidbyMem, AccPeriodMonth, " & _
          " AccPeriodYear, NoofDay, INVStayType, INVPAYType, BordxDate, PLNCode, INSCode, CLMTotalApproved " & _
          " From VIEW_Invoice_Tracking where 0 = 0 "
          
    
    'If Len(Trim(strAppend)) Then
    '    Sql = Sql + strAppend
    'End If
    
    'rst2.Open Sql, medixneptuneconnWS, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "InvoiceTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF
            
            lbltotal = rst2.recordCount
            
            With rst2
                
                Set Record = lstInvoiceTrackList.ListItems.Add(, , rst2!CASNumber & "-" & rst2!ClaimVersion)
                
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(1) = Trim(rst2!BordxRefNo)
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(2) = Trim(rst2!BordxDate)
                End If
                If Len(rst2!INVNo) Then
                    Record.SubItems(3) = Trim(rst2!INVNo)
                End If
                If Len(rst2!InvDate) Then
                   Record.SubItems(4) = Format(rst2!InvDate, "dd/mm/yyyy")
                End If
                If Len(rst2!INVAmount) Then
                    Record.SubItems(5) = Format(rst2!INVAmount, "#,##0.00")
                Else
                    Record.SubItems(5) = "0.00"
                End If
                If Len(rst2!CaseRegDate) Then
                    Record.SubItems(6) = Format(rst2!CaseRegDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(7) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(8) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                If Len(rst2!INVPayee) Then
                    Record.SubItems(9) = rst2!INVPayee
                End If
                If Len(rst2!MBMGroupCompany) Then
                    Record.SubItems(10) = Trim(rst2!MBMGroupCompany)
                End If
                If Len(rst2!INVEnteredDate) Then
                    Record.SubItems(11) = Format(rst2!INVEnteredDate, "dd/mm/yyyy")
                End If
                'If Len(rst2!INVEnteredBy) Then
                '    Record.SubItems(12) = Trim(rst2!INVEnteredBy)
                'End If
                If Len(rst2!INVStayType) Then
                    Record.SubItems(13) = Trim(rst2!INVStayType)
                End If
                If Len(rst2!INVPAYType) Then
                    Record.SubItems(14) = Trim(rst2!INVPAYType)
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!INVAmount) Then
                    TotAmt = TotAmt + rst2!INVAmount
                End If
                LblTotalAmt = Format(TotAmt, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
        Loop
    Loop Until check = False
    
    If lstInvoiceTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdPrint.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    
End Sub

Private Sub InvByMthCLM(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
Dim TotAmt As Variant
Dim TotalInvNo As String
Dim INVMonth As String
Dim INVYear As String

    Current = 0
    TotAmt = 0
    LblTotalAmt = ""
    lstInvoiceTrackList.ListItems.Clear
            
    strAppendCase = "2"
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "InvoiceTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF
            
            lbltotal = rst2.recordCount
            
            With rst2
                If Len(rst2!INVMonth) Then
                    If rst2!INVMonth > 9 Then
                        INVMonth = rst2!INVMonth
                    Else
                        INVMonth = "0" & rst2!INVMonth
                    End If
                End If
                Set Record = lstInvoiceTrackList.ListItems.Add(, , rst2!INVYear & "/" & INVMonth)
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(1) = Trim(rst2!TotalCases)
                End If
                If Len(rst2!CLMPayableByMedix2H) Then
                    Record.SubItems(2) = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(3) = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                End If
                
                If Len(rst2!INVAmount) Then
                    Record.SubItems(4) = Format(rst2!INVAmount, "#,##0.00")
                End If
                
                Current = Current + rst2!TotalCases
                
                lblCurrent = Current
                If Len(rst2!INVAmount) Then
                    TotAmt = TotAmt + rst2!INVAmount
                End If
                LblTotalAmt = Format(TotAmt, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
        Loop
    Loop Until check = False
    
    If lstInvoiceTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdPrint.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    
End Sub

Private Sub InvByMthHospCLM(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
Dim TotAmt As Variant
Dim TotalInvNo As String
Dim INVMonth As String
    Current = 0
    TotAmt = 0
    LblTotalAmt = ""
    lstInvoiceTrackList.ListItems.Clear
            
    strAppendCase = "3"
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "InvoiceTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF
            
            lbltotal = rst2.recordCount
            
            With rst2
                
                Set Record = lstInvoiceTrackList.ListItems.Add(, , rst2!CASHOSCode)
                If Len(rst2!HosName) Then
                    Record.SubItems(1) = Trim(rst2!HosName)
                End If
                If Len(rst2!INVMonth) Then
                    INVMonth = IIf(rst2!INVMonth > 9, rst2!INVMonth, "0" & rst2!INVMonth)
                    Record.SubItems(2) = Trim(rst2!INVYear) & "/" & Trim(INVMonth)
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = Trim(rst2!TotalCases)
                End If
                If Len(rst2!CLMPayableByMedix2H) Then
                    Record.SubItems(4) = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(5) = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                End If
                
                If Len(rst2!INVAmount) Then
                    Record.SubItems(6) = Format(rst2!INVAmount, "#,##0.00")
                End If
                
                Current = Current + rst2!TotalCases
                
                lblCurrent = Current
                If Len(rst2!INVAmount) Then
                    TotAmt = TotAmt + rst2!INVAmount
                End If
                LblTotalAmt = Format(TotAmt, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
        Loop
    Loop Until check = False
    
    If lstInvoiceTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdPrint.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    
End Sub

Private Sub InvByHosp(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
Dim TotAmt As Variant
Dim TotalInvNo As String
Dim INVMonth As String
    Current = 0
    TotAmt = 0
    LblTotalAmt = ""
    lstInvoiceTrackList.ListItems.Clear
            
    strAppendCase = "4"
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "InvoiceTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF
            
            lbltotal = rst2.recordCount
            
            With rst2
                
                Set Record = lstInvoiceTrackList.ListItems.Add(, , rst2!CASHOSCode)
                If Len(rst2!HosName) Then
                    Record.SubItems(1) = Trim(rst2!HosName)
                End If
                If Len(rst2!INVMonth) Then
                    INVMonth = IIf(rst2!INVMonth > 9, rst2!INVMonth, "0" & rst2!INVMonth)
                    Record.SubItems(2) = Trim(rst2!INVYear) & "/" & Trim(INVMonth)
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = Trim(rst2!TotalCases)
                End If
                If Len(rst2!CLMPayableByMedix2H) Then
                    Record.SubItems(4) = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(5) = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                End If
                
                If Len(rst2!INVAmount) Then
                    Record.SubItems(6) = Format(rst2!INVAmount, "#,##0.00")
                End If
                
                Current = Current + rst2!TotalCases
                
                lblCurrent = Current
                If Len(rst2!INVAmount) Then
                    TotAmt = TotAmt + rst2!INVAmount
                End If
                LblTotalAmt = Format(TotAmt, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
        Loop
    Loop Until check = False
    
    If lstInvoiceTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdPrint.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    
End Sub

Private Sub InvGenListingCLM(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
Dim TotAmt As Variant
Dim TotalInvNo As String
    
    Current = 0
    TotAmt = 0
    LblTotalAmt = ""
    lstInvoiceTrackList.ListItems.Clear
            
    strAppendCase = "4"
    Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "InvoiceTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    Do
        Do Until rst2.EOF
            
            'lblTotal = rst2.recordCount
            
            With rst2
                
                Set Record = lstInvoiceTrackList.ListItems.Add(, , rst2!CASHOSCode)
                If Len(rst2!HosName) Then
                    Record.SubItems(1) = Trim(rst2!HosName)
                End If
                If Len(rst2!INVNo) Then
                    Record.SubItems(2) = Trim(rst2!INVNo)
                End If
                
                If Len(rst2!InvDate) Then
                    Record.SubItems(3) = Trim(rst2!InvDate)
                End If
                If Len(rst2!CLMPayableByMedix2H) Then
                    Record.SubItems(4) = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(5) = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                End If
                
                If Len(rst2!INVAmount) Then
                    Record.SubItems(6) = Format(rst2!INVAmount, "#,##0.00")
                End If
                
                Current = Current + 1
                
                lblCurrent = Current
                If Len(rst2!INVAmount) Then
                    TotAmt = TotAmt + rst2!INVAmount
                End If
                LblTotalAmt = Format(TotAmt, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
        Loop
    Loop Until check = False
    
    If lstInvoiceTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdPrint.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    
End Sub


'************************Start ComboBox Change/KeyPress**************************

Private Sub CboPayTo_Click()
    If InStr(cboPayTo.Text, "null") Then
       cboPayTo.Enabled = False
    End If
End Sub

Private Sub CboPayor_Change()
If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If
End Sub

Private Sub cboPayor_Click()
Dim StrSql As String
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend As String

    If ChkGrpCom2.Value = 1 Then
        CboGrpCom.Clear
        medixneptuneconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        'strsql = "select MBMGroupCompany from View_PAYGrpCom where PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
        'PAYGRP.Open strsql, medixneptuneconn, adOpenForwardOnly, adLockBatchOptimistic
        strAppendCase = "2"
        StrAppend = " AND PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
        
        Set cmd.ActiveConnection = medixneptuneconn
        cmd.CommandText = "InvoiceTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        PAYGRP.CursorLocation = adUseClient
        PAYGRP.CursorType = adOpenForwardOnly
        PAYGRP.CacheSize = 100
        Set PAYGRP = cmd.Execute
        
        If PAYGRP.recordCount Then
            Do Until PAYGRP.EOF
               CboGrpCom.AddItem Trim(PAYGRP!MBMGroupCompany)
               PAYGRP.MoveNext
            Loop
        End If
        PAYGRP.Close
        Set PAYGRP = Nothing
        medixneptuneconn.Close
        Set medixneptuneconn = Nothing
    
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

Private Sub CboClaimability_Change()
    If InStr(CboClaimability.Text, "null") Then
       CboClaimability.Enabled = False
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

Private Sub CboInsType_Change()
    If InStr(CboInsType.Text, "null") Then
       CboInsType.Enabled = False
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


Private Sub CboPayTo_Change()
    If InStr(cboPayTo.Text, "null") Then
       cboPayTo.Enabled = False
    End If
End Sub

Private Sub cboPayTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayTo.Text, cboPayTo.SelStart) + Chr(KeyAscii) + Mid(cboPayTo.Text, cboPayTo.SelStart + cboPayTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayTo.ListCount - 1
 
        If NewText = LCase(Left(cboPayTo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayTo.Text = ValidValue
                cboPayTo.SelStart = 0
                cboPayTo.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboStayType_Change()
    If InStr(CboStayType.Text, "null") Then
       CboStayType.Enabled = False
    End If
End Sub

Private Sub CboStayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboStayType.Text, CboStayType.SelStart) + Chr(KeyAscii) + Mid(CboStayType.Text, CboStayType.SelStart + CboStayType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboStayType.ListCount - 1
 
        If NewText = LCase(Left(CboStayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboStayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboStayType.Text = ValidValue
                CboStayType.SelStart = 0
                CboStayType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPayType_Change()
    If InStr(cboPayType.Text, "null") Then
       cboPayType.Enabled = False
    End If
End Sub

Private Sub CboPayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayType.Text, cboPayType.SelStart) + Chr(KeyAscii) + Mid(cboPayType.Text, cboPayType.SelStart + cboPayType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayType.ListCount - 1
 
        If NewText = LCase(Left(cboPayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayType.Text = ValidValue
                cboPayType.SelStart = 0
                cboPayType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboACPeriodMthFr_Change()
If cboACPeriodMthFr.Text = "" Then
    If InStr(cboACPeriodMthFr.Text, "null") Then
       cboACPeriodMthFr.Enabled = False
    End If
End If
End Sub


Private Sub cboACPeriodMthFr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If cboACPeriodMthFr.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboACPeriodMthFr.Text, cboACPeriodMthFr.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodMthFr.Text, cboACPeriodMthFr.SelStart + cboACPeriodMthFr.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodMthFr.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodMthFr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodMthFr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodMthFr.Text = ValidValue
                cboACPeriodMthFr.SelStart = 0
                cboACPeriodMthFr.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub

Private Sub cboACPeriodMthTo_Change()
If cboACPeriodMthTo.Text = "" Then
    If InStr(cboACPeriodMthTo.Text, "null") Then
       cboACPeriodMthTo.Enabled = False
    End If
End If
End Sub


Private Sub cboACPeriodMthTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If cboACPeriodMthTo.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboACPeriodMthTo.Text, cboACPeriodMthTo.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodMthTo.Text, cboACPeriodMthTo.SelStart + cboACPeriodMthTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodMthTo.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodMthTo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodMthTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodMthTo.Text = ValidValue
                cboACPeriodMthTo.SelStart = 0
                cboACPeriodMthTo.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub

Private Sub cboACPeriodYRFr_Change()
If cboACPeriodYRFr.Text = "" Then
    If InStr(cboACPeriodYRFr.Text, "null") Then
       cboACPeriodYRFr.Enabled = False
    End If
End If
End Sub


Private Sub cboACPeriodYRFr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If cboACPeriodYRFr.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboACPeriodYRFr.Text, cboACPeriodYRFr.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodYRFr.Text, cboACPeriodYRFr.SelStart + cboACPeriodYRFr.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodYRFr.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodYRFr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodYRFr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodYRFr.Text = ValidValue
                cboACPeriodYRFr.SelStart = 0
                cboACPeriodYRFr.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub

Private Sub cboACPeriodYRTo_Change()
If cboACPeriodYRTo.Text = "" Then
    If InStr(cboACPeriodYRTo.Text, "null") Then
       cboACPeriodYRTo.Enabled = False
    End If
End If
End Sub


Private Sub cboACPeriodYRTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If cboACPeriodYRTo.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboACPeriodYRTo.Text, cboACPeriodYRTo.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodYRTo.Text, cboACPeriodYRTo.SelStart + cboACPeriodYRTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodYRTo.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodYRTo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodYRTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodYRTo.Text = ValidValue
                cboACPeriodYRTo.SelStart = 0
                cboACPeriodYRTo.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub
'************************End ComboBox Change/KeyPress**************************




Private Sub txtDOADate1_LostFocus()
    If IsDate(txtDOADate1) Then
    Else
        If txtDOADate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOADate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDOADate2_LostFocus()
    If IsDate(txtDOADate2) Then
    Else
        If txtDOADate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOADate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDODDate1_LostFocus()
    If IsDate(txtDODDate1) Then
    Else
        If txtDODDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDODDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDODDate2_LostFocus()
    If IsDate(txtDODDate2) Then
    Else
        If txtDODDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDODDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtINVAmtFr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtINVAmtTo_KeyPress(KeyAscii As Integer)
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

Private Sub txtIRDDate1_LostFocus()
    If IsDate(txtIRDDate1) Then
    Else
        If txtIRDDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtIRDDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtIRDDate2_LostFocus()
    If IsDate(txtIRDDate2) Then
    Else
        If txtIRDDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtIRDDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtWSNo1_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtWSNo1_LostFocus()
    TxtWSNo1 = ChkStr(TxtWSNo1)
    CmdSearch.Default = False
End Sub

Private Sub TxtWSVersion1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtWSVersion2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub ChkHos2_Click()
Dim HOS As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
    
    If ChkHos2.Value = 1 Then
        CboHos.Clear
        medixneptuneconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        'HOS.Open "Select HOSName from HOS", medixneptuneconn, adOpenForwardOnly, adLockBatchOptimistic

        Set cmd.ActiveConnection = medixneptuneconn
        cmd.CommandText = "SearchHOS"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        HOS.CursorLocation = adUseClient
        HOS.CursorType = adOpenForwardOnly
        HOS.CacheSize = 100
        Set HOS = cmd.Execute
        
        Do Until HOS.EOF
            CboHos.AddItem Trim(HOS!HosName)
            HOS.MoveNext
        Loop
        HOS.Close
        Set HOS = Nothing
        medixneptuneconn.Close
        Set medixneptuneconn = Nothing
    
    Else
        CboHos.Clear
    End If

End Sub

Private Sub ChkGrpCom2_Click()
Dim StrSql As String
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend As String

    If ChkGrpCom2.Value = 1 Then
        CboGrpCom.Clear
        medixneptuneconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        'strsql = "select MBMGroupCompany from View_PAYGrpCom where PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
        'PAYGRP.Open strsql, medixneptuneconn, adOpenForwardOnly, adLockBatchOptimistic
        strAppendCase = "2"
        StrAppend = " AND PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
        Set cmd.ActiveConnection = medixneptuneconn
        cmd.CommandText = "InvoiceTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        PAYGRP.CursorLocation = adUseClient
        PAYGRP.CursorType = adOpenForwardOnly
        PAYGRP.CacheSize = 100
        Set PAYGRP = cmd.Execute
        
        If PAYGRP.recordCount Then
            Do Until PAYGRP.EOF
               CboGrpCom.AddItem Trim(PAYGRP!MBMGroupCompany)
               PAYGRP.MoveNext
            Loop
        End If
        PAYGRP.Close
        Set PAYGRP = Nothing
        medixneptuneconn.Close
        Set medixneptuneconn = Nothing
    
    Else
        CboGrpCom.Clear
    End If
End Sub

Private Sub CboGrpCom_Change()
    If InStr(CboGrpCom.Text, "null") Then
       CboGrpCom.Enabled = False
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



