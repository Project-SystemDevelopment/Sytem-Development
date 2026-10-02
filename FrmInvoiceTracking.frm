VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInvoiceTracking 
   Caption         =   "Invoice Tracking"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "For Update A/C Period"
      Height          =   615
      Left            =   0
      TabIndex        =   109
      Top             =   6720
      Width           =   5775
      Begin VB.ComboBox CboPeriodMth 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":0000
         Left            =   1320
         List            =   "FrmInvoiceTracking.frx":0002
         TabIndex        =   48
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CboPeriodYr 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":0004
         Left            =   4080
         List            =   "FrmInvoiceTracking.frx":0006
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   2880
         TabIndex        =   110
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Invoice  Tracking"
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
         TabIndex        =   59
         Top             =   0
         Width           =   3285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   73
      Top             =   7320
      Width           =   11895
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   330
         Left            =   7920
         TabIndex        =   54
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   330
         Left            =   7200
         TabIndex        =   53
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   6480
         TabIndex        =   52
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   330
         Left            =   5040
         TabIndex        =   50
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   330
         Left            =   8640
         TabIndex        =   55
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9360
         TabIndex        =   56
         Top             =   200
         Width           =   975
      End
      Begin VB.TextBox txtNoofDay 
         Height          =   285
         Left            =   120
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   11040
         TabIndex        =   58
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10320
         TabIndex        =   57
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   5760
         TabIndex        =   51
         Top             =   200
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1200
         TabIndex        =   74
         Top             =   200
         Width           =   735
      End
      Begin VB.Label LblTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3600
         TabIndex        =   86
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2760
         TabIndex        =   85
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   2040
         TabIndex        =   77
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   960
         TabIndex        =   75
         Top             =   200
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lstInvoiceTrackList 
      Height          =   2700
      Left            =   0
      TabIndex        =   72
      Top             =   3960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4763
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   28
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "A/C Period Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Case Number"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Version"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Worksheet No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Bordx"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Inv. No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Inv. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Inv Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Excess Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "DOR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "DOA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "DOD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "DPO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "No of Days (DPO - DOR)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "Follow Up Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Text            =   "Inv. Reg"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Inv. Entry By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   19
         Text            =   "Claim Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   20
         Text            =   "Case Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   21
         Text            =   "Wks Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   22
         Text            =   "Bordx Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "First Diag."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Final Diag."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Reason for Admission"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Text            =   "A/C Period"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "InvPayTo"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3585
      Left            =   0
      TabIndex        =   60
      Top             =   300
      Width           =   11775
      Begin VB.CheckBox ChkYear 
         Height          =   255
         Left            =   11040
         TabIndex        =   47
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox ChkMonth 
         Height          =   255
         Left            =   11040
         TabIndex        =   44
         Top             =   2960
         Width           =   375
      End
      Begin VB.CheckBox ChkFirst 
         Height          =   255
         Left            =   11040
         TabIndex        =   24
         Top             =   970
         Width           =   375
      End
      Begin VB.CheckBox ChkFinal 
         Height          =   255
         Left            =   11040
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkInvReg 
         Height          =   255
         Left            =   11040
         TabIndex        =   39
         Top             =   2400
         Width           =   375
      End
      Begin VB.CheckBox chkInvDate 
         Height          =   255
         Left            =   11040
         TabIndex        =   36
         Top             =   2130
         Width           =   375
      End
      Begin VB.CheckBox chkDPO 
         Height          =   255
         Left            =   11040
         TabIndex        =   33
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox chkDOD 
         Height          =   255
         Left            =   11040
         TabIndex        =   30
         Top             =   1560
         Width           =   375
      End
      Begin VB.ComboBox CboPayTo 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Text            =   "H - Hospital"
         Top             =   130
         Width           =   4320
      End
      Begin VB.CheckBox chkDOA 
         Height          =   255
         Left            =   11040
         TabIndex        =   27
         Top             =   1260
         Width           =   375
      End
      Begin VB.ComboBox CboWksStatus 
         Height          =   315
         Left            =   9480
         TabIndex        =   16
         Top             =   130
         Width           =   1455
      End
      Begin VB.ComboBox CboBordxStatus 
         Height          =   315
         Left            =   7080
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
      Begin VB.ComboBox CboHos 
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   4320
      End
      Begin VB.TextBox TxtAgeband2 
         Height          =   285
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   18
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox TxtFinal2 
         Height          =   310
         Left            =   9480
         MaxLength       =   3
         TabIndex        =   20
         Top             =   680
         Width           =   1455
      End
      Begin VB.TextBox TxtFirst2 
         Height          =   285
         Left            =   9480
         MaxLength       =   3
         TabIndex        =   23
         Top             =   960
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtDOADate2 
         Height          =   315
         Left            =   9480
         TabIndex        =   26
         Top             =   1220
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtWSNo 
         Height          =   310
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   3
         Top             =   700
         Width           =   3000
      End
      Begin VB.TextBox TxtWSVersion 
         Height          =   310
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   700
         Width           =   840
      End
      Begin VB.TextBox TxtInvNo 
         Height          =   310
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   5
         Top             =   980
         Width           =   4320
      End
      Begin VB.TextBox TxtMBMNo 
         Height          =   310
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1260
         Width           =   4320
      End
      Begin VB.TextBox TxtBordx 
         Height          =   310
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1550
         Width           =   4320
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   1830
         Width           =   4320
      End
      Begin VB.TextBox TxtPatientName 
         Height          =   310
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2110
         Width           =   4320
      End
      Begin VB.ComboBox CboClaimType 
         Height          =   315
         Left            =   4200
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox cboINSType 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   4200
         TabIndex        =   13
         Top             =   2680
         Width           =   1455
      End
      Begin VB.ComboBox CboCaseStatus 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   2680
         Width           =   1935
      End
      Begin VB.ComboBox CboDept 
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   2970
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtDODDate2 
         Height          =   315
         Left            =   9480
         TabIndex        =   29
         Top             =   1500
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDPODate2 
         Height          =   315
         Left            =   9480
         TabIndex        =   32
         Top             =   1790
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
         Left            =   9480
         TabIndex        =   35
         Top             =   2070
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
         Left            =   9480
         TabIndex        =   38
         Top             =   2360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtINVAmtTo 
         Height          =   285
         Left            =   9480
         TabIndex        =   41
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":0008
         Left            =   9480
         List            =   "FrmInvoiceTracking.frx":000A
         TabIndex        =   43
         Top             =   2900
         Width           =   1455
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":000C
         Left            =   9480
         List            =   "FrmInvoiceTracking.frx":000E
         TabIndex        =   46
         Top             =   3180
         Width           =   1455
      End
      Begin VB.TextBox TxtAgeband1 
         Height          =   285
         Left            =   7080
         MaxLength       =   2
         TabIndex        =   17
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox TxtFinal1 
         Height          =   310
         Left            =   7080
         MaxLength       =   3
         TabIndex        =   19
         Top             =   680
         Width           =   1335
      End
      Begin VB.TextBox TxtFirst1 
         Height          =   285
         Left            =   7080
         MaxLength       =   3
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtDOADate1 
         Height          =   315
         Left            =   7080
         TabIndex        =   25
         Top             =   1220
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate1 
         Height          =   315
         Left            =   7080
         TabIndex        =   28
         Top             =   1500
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDPODate1 
         Height          =   315
         Left            =   7080
         TabIndex        =   31
         Top             =   1790
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   7080
         TabIndex        =   34
         Top             =   2070
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtIRDDate1 
         Height          =   315
         Left            =   7080
         TabIndex        =   37
         Top             =   2360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtINVAmtFr 
         Height          =   285
         Left            =   7080
         TabIndex        =   40
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":0010
         Left            =   7080
         List            =   "FrmInvoiceTracking.frx":0012
         TabIndex        =   42
         Top             =   2900
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmInvoiceTracking.frx":0014
         Left            =   7080
         List            =   "FrmInvoiceTracking.frx":0016
         TabIndex        =   45
         Top             =   3180
         Width           =   1335
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
         Left            =   11040
         TabIndex        =   108
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label45 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   5880
         TabIndex        =   107
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label44 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   106
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label43 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   105
         Top             =   3260
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   5880
         TabIndex        =   104
         Top             =   3260
         Width           =   1215
      End
      Begin VB.Label Label41 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   103
         Top             =   2750
         Width           =   375
      End
      Begin VB.Label Label40 
         Caption         =   "Inv Amount"
         Height          =   255
         Left            =   5880
         TabIndex        =   102
         Top             =   2750
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "Ins Type"
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   2445
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   3360
         TabIndex        =   100
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label34 
         Caption         =   "Age Band"
         Height          =   255
         Left            =   5880
         TabIndex        =   99
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Case Status"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   97
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label38 
         Caption         =   "First Diag."
         Height          =   255
         Left            =   5880
         TabIndex        =   96
         Top             =   1040
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   95
         Top             =   1040
         Width           =   375
      End
      Begin VB.Label Label36 
         Caption         =   "Final Diag."
         Height          =   255
         Left            =   5880
         TabIndex        =   94
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   93
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Member No"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "Wks Status"
         Height          =   255
         Left            =   8520
         TabIndex        =   91
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Bordx Status"
         Height          =   255
         Left            =   5880
         TabIndex        =   90
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "--"
         Height          =   255
         Left            =   4440
         TabIndex        =   88
         Top             =   780
         Width           =   255
      End
      Begin VB.Label Label27 
         Caption         =   "Bordx Ref"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1660
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Inv No"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   1040
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Clm No"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Pay To"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Inv Date"
         Height          =   255
         Left            =   5880
         TabIndex        =   80
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   79
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Dept Stage"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   3075
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   71
         Top             =   1900
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   70
         Top             =   1600
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   69
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Claim Type"
         Height          =   255
         Left            =   3360
         TabIndex        =   68
         Top             =   2445
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "DPO"
         Height          =   255
         Left            =   5880
         TabIndex        =   67
         ToolTipText     =   "Date Pass On"
         Top             =   1900
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "DOD"
         Height          =   255
         Left            =   5880
         TabIndex        =   66
         ToolTipText     =   "Date of Discharged"
         Top             =   1600
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "DOA"
         Height          =   255
         Left            =   5880
         TabIndex        =   65
         ToolTipText     =   "Date of Admission"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "To"
         Height          =   255
         Left            =   8520
         TabIndex        =   62
         Top             =   2440
         Width           =   495
      End
      Begin VB.Label Label24 
         Caption         =   "Inv Reg"
         Height          =   255
         Left            =   5880
         TabIndex        =   61
         Top             =   2440
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmInvoiceTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public intExit As Integer
Dim CaseNoArray() As String
Public searchCriteria As String
Dim CASNo As String
Dim CASVer As String
Dim PAYTo As String
Dim INVNO As String
'Sort Direction for each column in the ListView
Private m_blnDirection(1 To 28) As Boolean

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
    lblTotal = 0
    LblTotalAmt = 0
    ReDim CaseNoArray(0)
    Call ComboListing
End Sub

'************************Start Command Button**************************
Private Sub CmdExit_Click()
Dim RecordExit
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
    End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    ReDim CaseNoArray(0)
    lstInvoiceTrackList.listitems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Public Sub cmdSearch_Click()
    CmdClear.Enabled = False
    Screen.MousePointer = vbHourglass
        Arraycnt = 0
        lstInvoiceTrackList.Height = 6300
        lstInvoiceTrackList.Left = 0
        lstInvoiceTrackList.Top = 360
        lstInvoiceTrackList.Width = 11775
        Frame1.Visible = False
        Call SearchRecord
    Screen.MousePointer = Default
End Sub

Private Sub CmdBack_Click()
    lstInvoiceTrackList.Height = 2700
    lstInvoiceTrackList.Left = 0
    lstInvoiceTrackList.Top = 3960
    lstInvoiceTrackList.Width = 11775
    Frame1.Visible = True
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        
    If lstInvoiceTrackList.listitems.count <> 0 Then
        RptInvoiceTrack.show
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = Default
End Sub

Private Sub CmdPrintFile_Click()
    
    Call ExportToExcelInvoice(lstInvoiceTrackList)

End Sub

Private Sub CmdView_Click()
    lstInvoiceTrackList.Height = 6300
    lstInvoiceTrackList.Left = 0
    lstInvoiceTrackList.Top = 360
    lstInvoiceTrackList.Width = 11775
    Frame1.Visible = False
End Sub
'************************End Command Button**************************


'************************Start sorted ListView**************************

Private Sub lstInvoiceTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    'Case 1
        'SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    'Case 2
        'SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    'Case 3
        'SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
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
    Case 16
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(18)
    Case 19
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(19)
    Case 20
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(21)
    Case 22
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(22)
    Case 23
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(23)
    Case 24
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(24)
    Case 25
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(25)
    Case 26
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(26)
    Case 27
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
    Case 28
        SortListView lstInvoiceTrackList, ColumnHeader.Index, ldtString, m_blnDirection(29)
    End Select
End Sub
'************************End sorted ListView**************************


'************************Start ComboListing**************************

Private Sub ComboListing()
Dim HOS As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim PAYTYPE As New ADODB.Recordset
Dim sql As String
Dim YearStart
Dim MonthStart

HOS.Open "Select HOSName from HOS", medixmasterconn, adOpenKeyset, adLockOptimistic
PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
PAYTYPE.Open "Select PayTypeCode from PAYTYPE", medixmasterconn, adOpenKeyset, adLockOptimistic
INS.Open "Select INSCode, INSDescription from INS", medixmasterconn, adOpenKeyset, adLockOptimistic

    Do Until HOS.EOF
        CboHos.AddItem Trim(HOS!HOSName)
        HOS.MoveNext
    Loop
    HOS.Close
    Set HOS = Nothing
    
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
        
    Do Until PAYTYPE.EOF
        If Len(PAYTYPE!PayTypeCode) Then
            CboClaimType.AddItem PAYTYPE!PayTypeCode
        End If
        PAYTYPE.MoveNext
    Loop
    PAYTYPE.Close
    Set PAYTYPE = Nothing
    
    Do Until INS.EOF
        cboINSType.AddItem ChkStr(INS!INSCode) & " - " & ChkStr(INS!INSDescription)
        INS.MoveNext
    Loop
    INS.Close
    Set INS = Nothing
    
    With CboCaseStatus
        .AddItem "O - Open"
        .AddItem "C - Close"
        .AddItem "D - Delete"
    End With
    
    With CboPayTo
        .AddItem "H - Hospital"
        .AddItem "M - Member"
    End With
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    With CboBordxStatus
        .AddItem "0 - Not yet Bordx"
        .AddItem "1 - Bordx"
        .AddItem "2 - Both"
    End With
    
    With CboWksStatus
        .AddItem "O - Open"
        .AddItem "C - Close"
    End With
    
    
    With CboDept
        .AddItem "0 - Admission"
        .AddItem "1 - Claim"
        .AddItem "2 - Both"
    End With
    
    For MonthStart = 1 To 12
        With CboPeriodMth
            .AddItem MonthStart
        End With
        With cboACPeriodMthFr
            .AddItem MonthStart
        End With
        With cboACPeriodMthTo
            .AddItem MonthStart
        End With
    Next MonthStart
    
    For YearStart = 1999 To 3000
        With CboPeriodYr
            .AddItem YearStart
        End With
        With cboACPeriodYRFr
            .AddItem YearStart
        End With
        With cboACPeriodYRTo
            .AddItem YearStart
        End With
    Next YearStart
End Sub
'************************Start ComboListing**************************

'************************Start Search Record**************************

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim sql As String
Dim sql2 As String
Dim Bordx As Integer
Dim CheckDate As New ADODB.Recordset
    
    
    strAppend = ""
    
    If Len(Trim(TxtInvNo)) Then
        strAppend = strAppend + " AND InvNo like '" & ChkStr(TxtInvNo) & "%'"
    End If
    
    If Len(Trim(TxtAgeband1)) Then
        strAppend = strAppend + " AND AGEBand >= '" & Trim(TxtAgeband1) & "'"
    End If
    
    If Len(Trim(TxtAgeband2)) Then
        strAppend = strAppend + " AND AGEBand <= '" & Trim(TxtAgeband2) & "'"
    End If
    
       
    If Len(Trim(TxtFinal1)) Then
        strAppend = strAppend + " AND FinalDiag >= '" & Trim(TxtFinal1) & "'"
    End If
    
    If Len(Trim(TxtFinal2)) Then
        strAppend = strAppend + " AND FinalDiag <= '" & Trim(TxtFinal2) & "'"
    End If
    
    If ChkFinal.value = 1 Then
        sql = ""
        sql = "AND (FinalDiag Is null OR FinalDiag = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtFirst1)) Then
        strAppend = strAppend + " AND FirstDiag >= '" & Trim(TxtFirst1) & "'"
    End If
    
    If Len(Trim(TxtFirst2)) Then
        strAppend = strAppend + " AND FirstDiag <= '" & Trim(TxtFirst2) & "'"
    End If
    
    If ChkFirst.value = 1 Then
        sql = ""
        sql = "AND (FirstDiag Is null OR FirstDiag = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtPatientName)) Then
       strAppend = strAppend + " And (MBMName like '%" & ChkStr(TxtPatientName) & "%' " & _
       " OR MBMCName like '%" & ChkStr(TxtPatientName) & "%')"
    End If
    
    If Len(Trim(TxtMBMNo)) Then
        strAppend = strAppend + " AND MBMNumber like '" & ChkStr(TxtMBMNo) & "%' "
    End If
    
       
    If Len(Trim(CboClaimType)) Then
        strAppend = strAppend + " AND ClaimType like '" & RemStr(Mid(CboClaimType, 1, IIf(InStr(1, CboClaimType, "-") = 0, 4, InStr(1, CboClaimType, "-") - 1))) & "%'"
    End If
    
           
    If Len(Trim(TxtWSNo)) Then
        strAppend = strAppend + " AND CASNumber like '" & ChkStr(TxtWSNo) & "%' "
    End If
    If Len(Trim(TxtWSVersion)) Then
        strAppend = strAppend + " AND INVWSVersion = " & Trim(TxtWSVersion)
    End If
    
    
    If Len(Trim(CboPayTo)) Then
        strAppend = strAppend + " AND INVPayTo = '" & Trim(Mid(CboPayTo, 1, IIf(InStr(1, CboPayTo, "-") = 0, 4, InStr(1, CboPayTo, "-") - 1))) & "'"
    End If
    
    If Len(Trim(CboHos)) Then
    
        strsqlHOS = "select HosCode, HosName from HOS where HosName = '" & Trim(CboHos) & "'"
        HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        strAppend = strAppend + " AND INVPayee = '" & Trim(HOS!HosCode) & "'"
        
        HOS.Close
        Set HOS = Nothing
    End If
    
    If Mid(CboCaseStatus, 1, 1) = "C" Then
        strAppend = strAppend + " AND CASStatus = 'C'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "D" Then
        strAppend = strAppend + " AND CASStatus = 'D'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "O" Then
        strAppend = strAppend + " AND CASStatus = 'O'"
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
          
    If Len(Trim(TxtBordx)) Then
        strAppend = strAppend + " AND BordxRefNo like '%" & ChkStr(TxtBordx) & "%'"
    End If
           
    If Mid(CboDept, 1, 1) = "0" Or Mid(CboDept, 1, 1) = "1" Then
       sql = ""
       sql = " AND  CASClaimAdmin = '" & Trim(Mid(CboDept, 1, 1)) & "'"
       strAppend = strAppend + sql
    End If
           
    If chkDOA.value = 1 Then
        sql = ""
        sql = "AND (CasDateAdmitted Is null OR CasDateAdmitted = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If chkDOD.value = 1 Then
        sql = ""
        sql = "AND (CasDischargeDate Is null OR CasDischargeDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If chkDPO.value = 1 Then
        sql = ""
        sql = "AND (CASTokenDate Is null OR CASTokenDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If chkInvDate.value = 1 Then
        sql = ""
        sql = "AND (InvDate Is null OR InvDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If chkInvReg.value = 1 Then
        sql = ""
        sql = "AND (INVEnteredDate Is null OR INVEnteredDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If ChkMonth.value = 1 Then
        sql = ""
        sql = "AND AccPeriodMonth Is null " & ""
        strAppend = strAppend + sql
    End If
    
    If ChkYear.value = 1 Then
        sql = ""
        sql = "AND AccPeriodYear Is null " & ""
        strAppend = strAppend + sql
    End If

    
    If (txtDPODate1) <> "__/__/____" And IsDate(txtDPODate1) = True _
        And (txtDPODate2) <> "__/__/____" And IsDate(txtDPODate2) = True Then
        sql = ""
        sql = " AND CASTokenDate >= '" & Format(Trim(txtDPODate1), "mm/dd/yyyy") & _
              "' And CASTokenDate <= '" & Format(Trim(txtDPODate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDPODate1)) > 0 And IsDate(txtDPODate1) = True Then
        strAppend = strAppend + " And CASTokenDate >= '" & Format(txtDPODate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDPODate2)) > 0 And IsDate(txtDPODate2) = True Then
        strAppend = strAppend + " And CASTokenDate <= '" & Format(txtDPODate2, "mm/dd/yyyy") & "'"
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
    
    If (txtIRDDate1) <> "__/__/____" And IsDate(txtIRDDate1) = True _
        And (txtIRDDate2) <> "__/__/____" And IsDate(txtIRDDate2) = True Then
        sql = ""
        sql = " AND INVEnteredDate >= '" & Format(Trim(txtIRDDate1), "mm/dd/yyyy") & _
              "' And INVEnteredDate <= '" & Format(Trim(txtIRDDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtIRDDate1)) > 0 And IsDate(txtIRDDate1) = True Then
        strAppend = strAppend + " And INVEnteredDate >= '" & Format(txtIRDDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtIRDDate2)) > 0 And IsDate(txtIRDDate2) = True Then
        strAppend = strAppend + " And INVEnteredDate <= '" & Format(txtIRDDate2, "mm/dd/yyyy") & "'"
    End If
    
    
    If (txtInvDate1) <> "__/__/____" And IsDate(txtInvDate1) = True _
        And (txtInvDate2) <> "__/__/____" And IsDate(txtInvDate2) = True Then
        sql = ""
        sql = " AND INVDate >= '" & Format(Trim(txtInvDate1), "mm/dd/yyyy") & _
              "' And INVDate <= '" & Format(Trim(txtInvDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
        strAppend = strAppend + " And INVDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDate2)) > 0 And IsDate(txtInvDate2) = True Then
        strAppend = strAppend + " And INVDate <= '" & Format(txtInvDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Mid(CboBordxStatus, 1, 1) = "0" Or Mid(CboBordxStatus, 1, 1) = "1" Then
       sql = ""
       sql = " AND ClaimBordxStatus = '" & Trim(Mid(CboBordxStatus, 1, 1)) & "'"
       strAppend = strAppend + sql
    End If
           
    If Len(Trim(CboWksStatus)) Then
        strAppend = strAppend + " And ClaimStatus = '" & Trim(Mid(CboWksStatus, 1, IIf(InStr(1, CboWksStatus, "-") = 0, 4, InStr(1, CboWksStatus, "-") - 1))) & "'"
    End If
    
    If Len(Trim(cboACPeriodMthFr)) Then
        strAppend = strAppend + " AND AccPeriodMonth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
    
    If Len(Trim(cboACPeriodMthTo)) Then
        strAppend = strAppend + " AND AccPeriodMonth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If Len(Trim(cboACPeriodYRFr)) Then
        strAppend = strAppend + " AND AccPeriodYear >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
    
    If Len(Trim(cboACPeriodYRTo)) Then
        strAppend = strAppend + " AND AccPeriodYear <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    searchCriteria = strAppend
    Call InvoiceTrackingListing(strAppend)

End Function
'************************End Search Record**************************


'************************Start List Record**************************

Private Sub InvoiceTrackingListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlHOS As String
Dim Total As String
Dim Current As String
Dim TotAmt As Variant
Dim TotalInvNo As String
    
    Total = 0
    Current = 0
    TotAmt = 0
    LblTotalAmt = ""
    lstInvoiceTrackList.listitems.Clear
            
    sql = " Select CASNumber, InvWSVersion, MBMNumber, InvNo, InvDate, " & _
          " InvAmount, CASClaimAdmin, CASDischargeDate, CaseRegDate, " & _
          " CASDateAdmitted, CASTokenDate, MBMPatientCovID, FinalDiag, " & _
          " FirstDiag, AGEBand, MBMCName, INVPayTo, INVPayee, InvEnteredDate, " & _
          " InvEnteredBy, ClaimType, CASStatus, InvWSed, ClaimStatus, " & _
          " CASAdmissionReason, ClaimBordxStatus, INVBordxRefNo, MBMName, " & _
          " INVfolUpDate, CASClaimability, INVPaidbyMem, AccPeriodMonth, AccPeriodYear, " & _
          " AccPeriodStatus From VIEW_Invoice_Tracking where 0=0 "
          
          
    strCount = "select count(*) as totalRecord From VIEW_Invoice_Tracking where 0=0 "
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
        strCount = strCount + strAppend
    End If

    Set rstCount = medixmasterconnWS.Execute(strCount)
 
    
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
       
    If rstCount!totalrecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
         
    Do
        Do Until rst2.EOF
        Total = rstCount!totalrecord
        lblTotal = Total
        txtNoofDay = ""
        If Len(rst2!CASTokenDate) And Len(rst2!CaseRegDate) Then
            txtNoofDay = DateValue(rst2!CASTokenDate) - DateValue((rst2!CaseRegDate))
        End If
            
        
        If txtNoofDay = "0" Then
                txtNoofDay = txtNoofDay + 1
        End If
            
            With rst2
                                            
                                
                If !AccPeriodStatus = True Then
                    Set Record = lstInvoiceTrackList.listitems.Add(, , "Updated")
                Else
                    Set Record = lstInvoiceTrackList.listitems.Add(, , "Not")
                End If
                
                If Len(rst2!CASNumber) Then
                    Record.SubItems(1) = Trim(rst2!CASNumber)
                End If
                
                If Len(rst2!INVWSVersion) Then
                    Record.SubItems(2) = Trim(rst2!INVWSVersion)
                End If
                
                If Len(rst2!CASNumber & "-" & rst2!INVWSVersion) Then
                    Record.SubItems(3) = rst2!CASNumber & "-" & rst2!INVWSVersion
                End If
                If Len(rst2!INVBordxRefNo) Then
                    Record.SubItems(4) = Trim(rst2!INVBordxRefNo)
                End If
                If Len(rst2!INVNO) Then
                    Record.SubItems(5) = Trim(rst2!INVNO)
                End If
                If Len(rst2!InvDate) Then
                   Record.SubItems(6) = Format(rst2!InvDate, "dd/mm/yyyy")
                End If
                If Len(rst2!INVAmount) Then
                    Record.SubItems(7) = Format(rst2!INVAmount, "#,##0.00")
                End If
                If Len(rst2!INVpaidbyMem) Then
                    Record.SubItems(8) = Format(rst2!INVpaidbyMem, "#,##0.00;(#,##0.00)")
                End If
                If Len(rst2!CaseRegDate) Then
                    Record.SubItems(9) = Format(rst2!CaseRegDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(10) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(11) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CASTokenDate) Then
                    Record.SubItems(12) = Format(rst2!CASTokenDate, "dd/mm/yyyy")
                End If
                If Len(txtNoofDay) Then
                    Record.SubItems(13) = txtNoofDay
                End If
                If Len(rst2!INVFolUpDate) Then
                    Record.SubItems(14) = Format(rst2!INVFolUpDate, "dd/mm/yyyy")
                End If
                If Len(rst2!INVPAYEE) Then
                    Record.SubItems(15) = rst2!INVPAYEE
                End If
                                                                       
                strsqlHOS = "select HosCode, HosName from HOS where HosCode = '" & Trim(Record.SubItems(15)) & "'"
                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                If HOS.recordCount Then
                    Record.SubItems(15) = Trim(HOS!HOSName)
                Else
                    Record.SubItems(15) = "MEMBER"
                End If
                
                If rst2!MBMPatientCovID = "00" Then
                    If Len(rst2!MBMName) Then
                        Record.SubItems(16) = Trim(rst2!MBMName)
                    End If
                Else
                    If Len(rst2!MBMCName) Then
                        Record.SubItems(16) = Trim(rst2!MBMCName)
                    End If
                End If
                
                If Len(rst2!INVEnteredDate) Then
                    Record.SubItems(17) = Format(rst2!INVEnteredDate, "dd/mm/yyyy")
                End If
                If Len(rst2!InvEnteredBy) Then
                    Record.SubItems(18) = Trim(rst2!InvEnteredBy)
                End If
                If Len(rst2!ClaimType) Then
                    Record.SubItems(19) = Trim(rst2!ClaimType)
                End If
                If Len(rst2!CASStatus) Then
                    Record.SubItems(20) = Trim(rst2!CASStatus)
                End If
                If Len(rst2!ClaimStatus) Then
                    Record.SubItems(21) = Trim(rst2!ClaimStatus)
                End If
                If Len(rst2!ClaimBordxStatus) Then
                    Record.SubItems(22) = Trim(rst2!ClaimBordxStatus)
                End If
                If Len(rst2!FirstDiag) Then
                    Record.SubItems(23) = Trim(rst2!FirstDiag)
                End If
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(24) = Trim(rst2!FinalDiag)
                End If
                If Len(rst2!CasAdmissionReason) Then
                    Record.SubItems(25) = Trim(rst2!CasAdmissionReason)
                End If
                If Len(rst2!AccPeriodMonth & "/" & rst2!AccPeriodYear) Then
                    Record.SubItems(26) = rst2!AccPeriodMonth & "/" & rst2!AccPeriodYear
                End If
                If Len(rst2!INVPayTo) Then
                    Record.SubItems(27) = rst2!INVPayTo
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
            HOS.Close
            Set HOS = Nothing
        Loop
    Loop Until check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    
End Sub
'************************End List Record**************************

'************************Start ComboBox Change/KeyPress**************************

Private Sub CboPayTo_Click()
Dim PAYTYPE As New ADODB.Recordset
Dim strsql As String
    If InStr(CboPayTo.Text, "null") Then
       CboPayTo.Enabled = False
    End If
    
    If Mid(ChkStr(CboPayTo), 1, 1) = "M" Then
           strsql = "select PAYTypeCode , PAYTypeName   from  PAYType where PAYTypeCode like '%R' "
        
        PAYTYPE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        CboClaimType.Clear
        Do Until PAYTYPE.EOF
            CboClaimType.AddItem PAYTYPE!PayTypeCode & " - " & Trim(PAYTYPE!PayTypeName)
            PAYTYPE.MoveNext
        Loop
        PAYTYPE.Close
        Set PAYTYPE = Nothing
    Else
        
            strsql = "select PAYTypeCode, PAYTypeName from PAYType where PAYTypeCode like '%C' "
        
        PAYTYPE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        CboClaimType.Clear
        Do Until PAYTYPE.EOF
            CboClaimType.AddItem PAYTYPE!PayTypeCode & " - " & Trim(PAYTYPE!PayTypeName)
            PAYTYPE.MoveNext
        Loop
        PAYTYPE.Close
        Set PAYTYPE = Nothing
        End If
End Sub

Private Sub CboPayor_Change()
If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If
End Sub


Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If CboPayor.Text = "" Then
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

Private Sub CboCaseStatus_Change()
    If InStr(CboCaseStatus.Text, "null") Then
       CboCaseStatus.Enabled = False
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

Private Sub cboINSType_Change()
    If InStr(cboINSType.Text, "null") Then
       cboINSType.Enabled = False
    End If
End Sub

Private Sub cboINSType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboINSType.Text, cboINSType.SelStart) + Chr(KeyAscii) + Mid(cboINSType.Text, cboINSType.SelStart + cboINSType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboINSType.ListCount - 1
 
        If NewText = LCase(Left(cboINSType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINSType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboINSType.Text = ValidValue
                cboINSType.SelStart = 0
                cboINSType.SelLength = Len(ValidValue)
            End If
        End If
End Sub


Private Sub CboPayTo_Change()
Dim PAYTYPE As New ADODB.Recordset
Dim strsql As String
    If InStr(CboPayTo.Text, "null") Then
       CboPayTo.Enabled = False
    End If
    
    If Mid(ChkStr(CboPayTo), 1, 1) = "M" Then
           strsql = "select PAYTypeCode , PAYTypeName   from  PAYType where PAYTypeCode like '%R' "
        
        PAYTYPE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        CboClaimType.Clear
        Do Until PAYTYPE.EOF
            CboClaimType.AddItem PAYTYPE!PayTypeCode & " - " & Trim(PAYTYPE!PayTypeName)
            PAYTYPE.MoveNext
        Loop
        PAYTYPE.Close
        Set PAYTYPE = Nothing
    Else
        
            strsql = "select PAYTypeCode, PAYTypeName from PAYType where PAYTypeCode like '%C' "
        
        PAYTYPE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        CboClaimType.Clear
        Do Until PAYTYPE.EOF
            CboClaimType.AddItem PAYTYPE!PayTypeCode & " - " & Trim(PAYTYPE!PayTypeName)
            PAYTYPE.MoveNext
        Loop
        PAYTYPE.Close
        Set PAYTYPE = Nothing
        End If
End Sub

Private Sub CboPayTo_KeyPress(KeyAscii As Integer)
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

Private Sub CboWksStatus_Change()
    If InStr(CboWksStatus.Text, "null") Then
       CboWksStatus.Enabled = False
    End If
End Sub

Private Sub CboWksStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboWksStatus.Text, CboWksStatus.SelStart) + Chr(KeyAscii) + Mid(CboWksStatus.Text, CboWksStatus.SelStart + CboWksStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboWksStatus.ListCount - 1
 
        If NewText = LCase(Left(CboWksStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboWksStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboWksStatus.Text = ValidValue
                CboWksStatus.SelStart = 0
                CboWksStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboBordxStatus_Change()
    If InStr(CboBordxStatus.Text, "null") Then
       CboBordxStatus.Enabled = False
    End If
End Sub

Private Sub CboBordxStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboBordxStatus.Text, CboBordxStatus.SelStart) + Chr(KeyAscii) + Mid(CboBordxStatus.Text, CboBordxStatus.SelStart + CboBordxStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboBordxStatus.ListCount - 1
 
        If NewText = LCase(Left(CboBordxStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboBordxStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboBordxStatus.Text = ValidValue
                CboBordxStatus.SelStart = 0
                CboBordxStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboDept_Change()
    If InStr(CboDept.Text, "null") Then
       CboDept.Enabled = False
    End If
End Sub

Private Sub CboDept_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboDept.Text, CboDept.SelStart) + Chr(KeyAscii) + Mid(CboDept.Text, CboDept.SelStart + CboDept.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboDept.ListCount - 1
 
        If NewText = LCase(Left(CboDept.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboDept.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboDept.Text = ValidValue
                CboDept.SelStart = 0
                CboDept.SelLength = Len(ValidValue)
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

Private Sub CboPeriodMth_Change()
If CboPeriodMth.Text = "" Then
    If InStr(CboPeriodMth.Text, "null") Then
       CboPeriodMth.Enabled = False
    End If
End If
End Sub


Private Sub CboPeriodMth_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If CboPeriodMth.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPeriodMth.Text, CboPeriodMth.SelStart) + Chr(KeyAscii) + Mid(CboPeriodMth.Text, CboPeriodMth.SelStart + CboPeriodMth.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPeriodMth.ListCount - 1
 
        If NewText = LCase(Left(CboPeriodMth.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPeriodMth.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPeriodMth.Text = ValidValue
                CboPeriodMth.SelStart = 0
                CboPeriodMth.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub

Private Sub CboPeriodYr_Change()
If CboPeriodYr.Text = "" Then
    If InStr(CboPeriodYr.Text, "null") Then
       CboPeriodYr.Enabled = False
    End If
End If
End Sub


Private Sub CboPeriodYr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If CboPeriodYr.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPeriodYr.Text, CboPeriodYr.SelStart) + Chr(KeyAscii) + Mid(CboPeriodYr.Text, CboPeriodYr.SelStart + CboPeriodYr.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPeriodYr.ListCount - 1
 
        If NewText = LCase(Left(CboPeriodYr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPeriodYr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPeriodYr.Text = ValidValue
                CboPeriodYr.SelStart = 0
                CboPeriodYr.SelLength = Len(ValidValue)
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

Private Sub txtDPODate1_LostFocus()
    If IsDate(txtDPODate1) Then
    Else
        If txtDPODate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDPODate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDPODate2_LostFocus()
    If IsDate(txtDPODate2) Then
    Else
        If txtDPODate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDPODate2.SetFocus
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
    If IsDate(txtInvDate2) Then
    Else
        If txtInvDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDate2.SetFocus
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


Private Sub CmdUpdate_Click()
Dim ii As Integer
Dim cnt As Integer
Dim recordFound As Boolean
Dim delimiter1, delimiter2, delimiter3 As Integer
    
        
        If CboPeriodMth = "" Then
            MsgBox "Month cannot be Empty.", vbCritical
        ElseIf CboPeriodYr = "" Then
            MsgBox "Year cannot be empty!", vbCritical
        ElseIf lstInvoiceTrackList.listitems.count = 0 Then
            MsgBox "Please select a record(s) before proceeding!", vbOKOnly + vbCritical
        ElseIf UBound(CaseNoArray) = 0 Then
            MsgBox "Please select at least one record to be updated! ", vbCritical
        Else
    
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                cnt = 0
           
                recordFound = False
                If lstInvoiceTrackList.listitems.count <> 0 Then
                    For ii = 0 To UBound(CaseNoArray)
                        If Trim(CaseNoArray(ii)) <> "" Then
                                                                                
                            delimiter1 = InStr(1, CaseNoArray(ii), "-")
                            CASNo = Trim(Mid(CaseNoArray(ii), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                            delimiter2 = InStr(1, CaseNoArray(ii), "~")
                            CASVer = Trim(Mid(CaseNoArray(ii), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                            delimiter3 = InStr(delimiter2 + 1, CaseNoArray(ii), "~")
                            PAYTo = Trim(Mid(CaseNoArray(ii), delimiter2 + 1, delimiter3 - delimiter2 - 1))
                            INVNO = Trim(Mid(CaseNoArray(ii), delimiter3 + 1, delimiter3 - 1))
        
                            Call UpdateRecord
                            recordFound = True
                        
                        End If
                    Next ii
                    Erase CaseNoArray
                    Arraycnt = 0
                    ReDim CaseNoArray(0)
                    MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    Call ClearForm(Me)
                    Call InvoiceTrackingListing(searchCriteria)
                End If
            End If
            
        End If
End Sub

'**********************Start Update Record*******************************
Private Sub UpdateRecord()
Dim Update
Dim strsql As String
Dim AccPeriod As New ADODB.Recordset

         
    strsql = "Select CASNumber, INVWSVersion, AccPeriodStatus, " & _
    " AccPeriodMonth, AccPeriodYear, LastAccPeriodUser, LastAccPeriodDate " & _
    " From ClaimInvoice where CASNumber = '" & Trim(CASNo) & "' " & _
    " AND INVWSVersion = '" & Trim(CASVer) & "' AND INVPayTo = '" & Trim(PAYTo) & "' " & _
    " AND INVNo = '" & Trim(INVNO) & "'"
    AccPeriod.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
      
    With AccPeriod
        !LastAccPeriodUser = Logon
        !LastAccPeriodDate = Date
        !AccPeriodStatus = True
        !AccPeriodMonth = Trim(CboPeriodMth)
        !AccPeriodYear = Trim(CboPeriodYr)
    End With
    AccPeriod.Update
    AccPeriod.Close
    Set AccPeriod = Nothing
End Sub
'**********************End Update Record*******************************

'****************** Start lstInvoiceTrackList ItemCheck *****************************

Private Sub lstInvoiceTrackList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Static Arraycnt As Integer
Dim temp

       

    If Item.Checked = True Then
        
            'If CboPeriodMth = "" Then
                'MsgBox "Month cannot be Empty.", vbCritical
                'Item.Checked = False
            'ElseIf CboPeriodYr = "" Then
                'MsgBox "Year cannot be empty!", vbCritical
                'Item.Checked = False
            'Else
                Arraycnt = IIf(UBound(CaseNoArray) = 0, 1, UBound(CaseNoArray) + 1)
                ReDim Preserve CaseNoArray(Arraycnt)
                CaseNoArray(Arraycnt) = Item.SubItems(3) & "~" & Item.SubItems(27) & "~" & Item.SubItems(5)
               
            'End If
    Else
    
        ' Remove from the array
        newbound = 0
    
        If UBound(CaseNoArray) <> 0 Then
            For i = 0 To UBound(CaseNoArray)
                If (Item.SubItems(3) & "~" & Item.SubItems(27) & "~" & Item.SubItems(5)) <> CaseNoArray(i) Then
                'If (Item.SubItems(1) & "-" & Item.SubItems(2) & "+" & Item.SubItems(27)) <> CaseNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = CaseNoArray(i)
                    newbound = newbound + 1
                End If
            Next
            Erase CaseNoArray
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim CaseNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                CaseNoArray(i) = tempArray(i)
            Next
            Erase tempArray
        Else
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            Arraycnt = 0
        End If
    End If
    
End Sub
'****************** End lstInvoiceTrackList ItemCheck *****************************

