VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReportMain 
   Caption         =   "Report"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Leaflet"
      TabPicture(0)   =   "frmReportMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameLeaflet"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdLLPrint"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "MCO Listing"
      TabPicture(1)   =   "frmReportMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdMCOPrint"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(2)=   "cmdMCOPrintPreview"
      Tab(1).Control(3)=   "FrameMCO"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Invoice"
      TabPicture(2)   =   "frmReportMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1"
      Tab(2).Control(1)=   "cmdInvPrint"
      Tab(2).Control(2)=   "Frame7"
      Tab(2).Control(3)=   "txtInvoiceNo"
      Tab(2).Control(4)=   "cmdINVPrintPreview"
      Tab(2).Control(5)=   "TxtInvoiceYear"
      Tab(2).Control(6)=   "FrameInv"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Card Export"
      TabPicture(3)   =   "frmReportMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame11"
      Tab(3).Control(1)=   "cmdCRDPrint"
      Tab(3).Control(2)=   "cmdCRDPrintPreview"
      Tab(3).Control(3)=   "FrameExp"
      Tab(3).ControlCount=   4
      Begin VB.Frame FrameExp 
         Caption         =   "Enter Criteria For Searching"
         Height          =   3135
         Left            =   -74760
         TabIndex        =   95
         Top             =   480
         Width           =   11055
         Begin VB.ComboBox cboExpPrintSeq 
            Height          =   315
            ItemData        =   "frmReportMain.frx":0070
            Left            =   6600
            List            =   "frmReportMain.frx":0080
            TabIndex        =   106
            Top             =   2310
            Width           =   2175
         End
         Begin VB.ComboBox cboExpPrintOpt 
            Height          =   315
            ItemData        =   "frmReportMain.frx":00B3
            Left            =   2160
            List            =   "frmReportMain.frx":00C9
            TabIndex        =   105
            Top             =   2310
            Width           =   2415
         End
         Begin VB.ComboBox cboExpCardVer 
            Height          =   315
            ItemData        =   "frmReportMain.frx":010F
            Left            =   2160
            List            =   "frmReportMain.frx":011C
            TabIndex        =   104
            Top             =   2700
            Width           =   2415
         End
         Begin VB.TextBox txtExpBatchNo 
            Height          =   315
            Left            =   6600
            TabIndex        =   103
            Text            =   "1"
            Top             =   1140
            Width           =   1335
         End
         Begin VB.TextBox txtExpMemberNoTo 
            Height          =   315
            Left            =   6600
            TabIndex        =   102
            Top             =   750
            Width           =   2175
         End
         Begin VB.TextBox txtExpMemberNoFr 
            Height          =   315
            Left            =   3240
            TabIndex        =   101
            Top             =   750
            Width           =   2175
         End
         Begin VB.ComboBox cboExpPayor 
            Height          =   315
            Left            =   2160
            TabIndex        =   100
            Top             =   1920
            Width           =   6615
         End
         Begin VB.TextBox txtExpMemberNo 
            Height          =   315
            Left            =   3240
            TabIndex        =   99
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton cmdExpSearch 
            Caption         =   "&Search"
            Height          =   375
            Left            =   9480
            TabIndex        =   98
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdExpClose 
            Caption         =   "&Exit"
            Height          =   375
            Left            =   9480
            TabIndex        =   97
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdExpClear 
            Caption         =   "&Clear"
            Height          =   375
            Left            =   9480
            TabIndex        =   96
            Top             =   1200
            Width           =   1215
         End
         Begin Crystal.CrystalReport cRptExp 
            Left            =   6120
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin MSMask.MaskEdBox txtExpBordxDate 
            Height          =   315
            Left            =   3240
            TabIndex        =   107
            Top             =   1140
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtExpBordxDateFr 
            Height          =   315
            Left            =   3240
            TabIndex        =   108
            Top             =   1530
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtExpBordxDateTo 
            Height          =   315
            Left            =   6600
            TabIndex        =   109
            Top             =   1530
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl51 
            Caption         =   "Sequence"
            Height          =   255
            Left            =   5760
            TabIndex        =   124
            Top             =   2310
            Width           =   855
         End
         Begin VB.Label lbl50 
            Caption         =   "Print Option"
            Height          =   255
            Left            =   240
            TabIndex        =   123
            Top             =   2310
            Width           =   855
         End
         Begin VB.Label lbl48 
            Caption         =   "To"
            Height          =   255
            Left            =   5760
            TabIndex        =   122
            Top             =   1530
            Width           =   495
         End
         Begin VB.Label lbl47 
            Caption         =   "Bordx From"
            Height          =   255
            Left            =   2160
            TabIndex        =   121
            Top             =   1530
            Width           =   855
         End
         Begin VB.Label lbl46 
            Caption         =   "Bordereaux Range :"
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Top             =   1530
            Width           =   1455
         End
         Begin VB.Label lbl45 
            Caption         =   "Batch No"
            Height          =   255
            Left            =   5760
            TabIndex        =   119
            Top             =   1140
            Width           =   855
         End
         Begin VB.Label lbl43 
            Caption         =   "Bordereaux Batch :"
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   1140
            Width           =   1575
         End
         Begin VB.Label lbl42 
            Caption         =   "To"
            Height          =   255
            Left            =   5760
            TabIndex        =   117
            Top             =   750
            Width           =   375
         End
         Begin VB.Label lbl41 
            Caption         =   "From"
            Height          =   255
            Left            =   2160
            TabIndex        =   116
            Top             =   750
            Width           =   855
         End
         Begin VB.Label lbl39 
            Caption         =   "Member No."
            Height          =   255
            Left            =   2160
            TabIndex        =   115
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lbl40 
            Caption         =   "Member No. Range :"
            Height          =   255
            Left            =   240
            TabIndex        =   114
            Top             =   750
            Width           =   1575
         End
         Begin VB.Label lbl52 
            Caption         =   "Card Version"
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   2700
            Width           =   1335
         End
         Begin VB.Label lbl44 
            Caption         =   "Bordx Date "
            Height          =   255
            Left            =   2160
            TabIndex        =   112
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label lbl49 
            Caption         =   "Payor"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label lbl38 
            Caption         =   "Single Member :"
            Height          =   255
            Left            =   240
            TabIndex        =   110
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame FrameInv 
         Caption         =   "Enter Criteria For Searching"
         Height          =   3135
         Left            =   -74760
         TabIndex        =   73
         Top             =   480
         Width           =   11055
         Begin VB.ComboBox cboInvPayor 
            Height          =   315
            Left            =   2160
            TabIndex        =   80
            Top             =   1860
            Width           =   6615
         End
         Begin VB.ComboBox cboInvListingVer 
            Height          =   315
            ItemData        =   "frmReportMain.frx":0156
            Left            =   2160
            List            =   "frmReportMain.frx":0160
            TabIndex        =   79
            Top             =   2640
            Width           =   2415
         End
         Begin VB.ComboBox cboInvPrintOpt 
            Height          =   315
            ItemData        =   "frmReportMain.frx":018A
            Left            =   2160
            List            =   "frmReportMain.frx":0197
            TabIndex        =   78
            Top             =   2250
            Width           =   2415
         End
         Begin VB.CommandButton cmdInvClear 
            Caption         =   "&Clear"
            Height          =   375
            Left            =   9480
            TabIndex        =   77
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdInvClose 
            Caption         =   "&Exit"
            Height          =   375
            Left            =   9480
            TabIndex        =   76
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdInvSearch 
            Caption         =   "&Search"
            Height          =   375
            Left            =   9480
            TabIndex        =   75
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtInvBatchNo 
            Height          =   315
            Left            =   6600
            TabIndex        =   74
            Text            =   "1"
            Top             =   600
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtInvBordxDate 
            Height          =   315
            Left            =   3240
            TabIndex        =   81
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Crystal.CrystalReport cRptInv 
            Left            =   8760
            Top             =   720
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin MSMask.MaskEdBox txtInvBordxDateTo 
            Height          =   315
            Left            =   6600
            TabIndex        =   82
            Top             =   990
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtInvBordxDateFr 
            Height          =   315
            Left            =   3240
            TabIndex        =   83
            Top             =   990
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl37 
            Caption         =   "MCO / Premium Listing"
            Height          =   255
            Left            =   4080
            TabIndex        =   84
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label lbl36 
            Caption         =   "Bordereaux Options"
            Height          =   255
            Left            =   4080
            TabIndex        =   85
            Top             =   240
            Width           =   1455
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   9120
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   9120
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label lbl33 
            Caption         =   "Payor"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1860
            Width           =   1095
         End
         Begin VB.Label lbl35 
            Caption         =   "Listing Versions"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lbl34 
            Caption         =   "Print Option"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   2250
            Width           =   855
         End
         Begin VB.Label lbl32 
            Caption         =   "To"
            Height          =   255
            Left            =   5760
            TabIndex        =   91
            Top             =   990
            Width           =   495
         End
         Begin VB.Label lbl31 
            Caption         =   "Bordx From"
            Height          =   255
            Left            =   2160
            TabIndex        =   90
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lbl30 
            Caption         =   "Bordereaux Range :"
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   990
            Width           =   1455
         End
         Begin VB.Label lbl27 
            Caption         =   "Bordereaux Batch :"
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lbl29 
            Caption         =   "Batch No"
            Height          =   255
            Left            =   5760
            TabIndex        =   87
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lbl28 
            Caption         =   "Bordx Date "
            Height          =   255
            Left            =   2160
            TabIndex        =   86
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame FrameMCO 
         Caption         =   "Enter Criteria For Searching"
         Height          =   3135
         Left            =   -74760
         TabIndex        =   60
         Top             =   480
         Width           =   11055
         Begin VB.ComboBox cboMCOPayor 
            Height          =   315
            Left            =   2160
            TabIndex        =   5
            Top             =   1860
            Width           =   6615
         End
         Begin VB.ComboBox cboMCOListingVer 
            Height          =   315
            ItemData        =   "frmReportMain.frx":01BA
            Left            =   2160
            List            =   "frmReportMain.frx":01D0
            TabIndex        =   8
            Top             =   2250
            Width           =   6615
         End
         Begin VB.ComboBox cboMCOPrintOpt 
            Height          =   315
            ItemData        =   "frmReportMain.frx":0292
            Left            =   2160
            List            =   "frmReportMain.frx":029F
            TabIndex        =   6
            Top             =   2640
            Width           =   2415
         End
         Begin VB.ComboBox cboMCOPrintSeq 
            Height          =   315
            ItemData        =   "frmReportMain.frx":02C2
            Left            =   6600
            List            =   "frmReportMain.frx":02D2
            TabIndex        =   7
            Top             =   2640
            Width           =   2175
         End
         Begin VB.CommandButton cmdMCOClear 
            Caption         =   "&Clear"
            Height          =   375
            Left            =   9480
            TabIndex        =   10
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdMCOClose 
            Caption         =   "&Exit"
            Height          =   375
            Left            =   9480
            TabIndex        =   11
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdMCOSearch 
            Caption         =   "&Search"
            Height          =   375
            Left            =   9480
            TabIndex        =   9
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtMCOBatchNo 
            Height          =   315
            Left            =   6600
            TabIndex        =   2
            Text            =   "1"
            Top             =   600
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtMCOBordxDate 
            Height          =   315
            Left            =   3240
            TabIndex        =   1
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Crystal.CrystalReport cRptMCO 
            Left            =   8760
            Top             =   720
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin MSMask.MaskEdBox txtMCOBordxDateTo 
            Height          =   315
            Left            =   6600
            TabIndex        =   4
            Top             =   990
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMCOBordxDateFr 
            Height          =   315
            Left            =   3240
            TabIndex        =   3
            Top             =   990
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl26 
            Caption         =   "Sequence"
            Height          =   255
            Left            =   5760
            TabIndex        =   72
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label lbl24 
            Caption         =   "MCO / Premium Listing"
            Height          =   255
            Left            =   3840
            TabIndex        =   71
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   9120
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lbl23 
            Caption         =   "Bordereaux Options"
            Height          =   255
            Left            =   3840
            TabIndex        =   70
            Top             =   240
            Width           =   1455
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   9120
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label lbl20 
            Caption         =   "Payor"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   1860
            Width           =   1095
         End
         Begin VB.Label lbl22 
            Caption         =   "Listing Versions"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   2250
            Width           =   1335
         End
         Begin VB.Label lbl21 
            Caption         =   "Print Option"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label lbl19 
            Caption         =   "To"
            Height          =   255
            Left            =   5760
            TabIndex        =   66
            Top             =   990
            Width           =   495
         End
         Begin VB.Label lbl18 
            Caption         =   "Bordx From"
            Height          =   255
            Left            =   2160
            TabIndex        =   65
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lbl17 
            Caption         =   "Bordereaux Range :"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   990
            Width           =   1455
         End
         Begin VB.Label lbl15 
            Caption         =   "Bordereaux Batch :"
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Batch No"
            Height          =   255
            Left            =   5760
            TabIndex        =   62
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lbl16 
            Caption         =   "Bordx Date "
            Height          =   255
            Left            =   2160
            TabIndex        =   61
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdCRDPrintPreview 
         Caption         =   "Print Pr&eview"
         Height          =   495
         Left            =   -73440
         TabIndex        =   44
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton cmdCRDPrint 
         Caption         =   "Pri&nt"
         Height          =   495
         Left            =   -74760
         TabIndex        =   43
         Top             =   7800
         Width           =   1215
      End
      Begin VB.Frame Frame11 
         Caption         =   "Listing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74760
         TabIndex        =   41
         Top             =   3720
         Width           =   11055
         Begin MSComctlLib.ListView lvwCRD 
            Height          =   3615
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   6376
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Payor Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Membership No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Bordereaux Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Policy No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Health Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Batch No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Printed"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.TextBox TxtInvoiceYear 
         Height          =   375
         Left            =   -70920
         TabIndex        =   40
         Top             =   7800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdINVPrintPreview 
         Caption         =   "Print Pr&eview"
         Height          =   495
         Left            =   -73440
         TabIndex        =   39
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton cmdMCOPrintPreview 
         Caption         =   "Print Pr&eview"
         Height          =   495
         Left            =   -73440
         TabIndex        =   14
         Top             =   7800
         Width           =   1215
      End
      Begin VB.TextBox txtInvoiceNo 
         Height          =   375
         Left            =   -71760
         TabIndex        =   38
         Top             =   7800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdLLPrint 
         Caption         =   "Pri&nt"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   7800
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Listing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   240
         TabIndex        =   37
         Top             =   3360
         Width           =   11055
         Begin MSComctlLib.ListView lvwLeaflet 
            Height          =   3975
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   7011
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Payor Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Membership No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Bordereaux Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Policy No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Health Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Batch No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Printed"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Listing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74760
         TabIndex        =   36
         Top             =   3720
         Width           =   11055
         Begin MSComctlLib.ListView lvwMCO 
            Height          =   3615
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   6376
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Payor Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Membership No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Bordereaux Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Policy No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Health Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Batch No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Printed"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CommandButton cmdMCOPrint 
         Caption         =   "Pri&nt"
         Height          =   495
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7800
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Listing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74760
         TabIndex        =   34
         Top             =   3720
         Width           =   11055
         Begin MSComctlLib.ListView lvwInvoice 
            Height          =   3615
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   6376
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Payor Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Bordereaux Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Health Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Batch No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Printed"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CommandButton cmdInvPrint 
         Caption         =   "Pri&nt"
         Height          =   495
         Left            =   -74760
         TabIndex        =   33
         Top             =   7800
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -70080
         TabIndex        =   32
         Top             =   7800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrameLeaflet 
         Caption         =   "Enter Criteria For Searching"
         Height          =   2775
         Left            =   240
         TabIndex        =   45
         Top             =   480
         Width           =   11055
         Begin VB.ComboBox cboLLCriteriaSel 
            Height          =   315
            ItemData        =   "frmReportMain.frx":030C
            Left            =   2160
            List            =   "frmReportMain.frx":031C
            TabIndex        =   0
            Top             =   360
            Width           =   2415
         End
         Begin VB.ComboBox cboLLPrintSeq 
            Height          =   315
            ItemData        =   "frmReportMain.frx":0365
            Left            =   6600
            List            =   "frmReportMain.frx":0375
            TabIndex        =   24
            Top             =   1830
            Width           =   2175
         End
         Begin VB.ComboBox cboLLPrintOpt 
            Height          =   315
            ItemData        =   "frmReportMain.frx":03A8
            Left            =   2160
            List            =   "frmReportMain.frx":03B5
            TabIndex        =   23
            Top             =   1830
            Width           =   2415
         End
         Begin VB.ComboBox cboLeafletVer 
            Height          =   315
            ItemData        =   "frmReportMain.frx":03D8
            Left            =   2160
            List            =   "frmReportMain.frx":03E5
            TabIndex        =   25
            Top             =   2340
            Width           =   2415
         End
         Begin VB.TextBox txtLLBatchNo 
            Height          =   315
            Left            =   6600
            TabIndex        =   19
            Top             =   870
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtLLMemberNoTo 
            Height          =   315
            Left            =   6600
            TabIndex        =   17
            Top             =   870
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtLLMemberNoFr 
            Height          =   315
            Left            =   3240
            TabIndex        =   16
            Top             =   870
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox cboLLPayor 
            Height          =   315
            Left            =   2160
            TabIndex        =   22
            Top             =   1320
            Width           =   6615
         End
         Begin VB.TextBox txtLLMemberNo 
            Height          =   315
            Left            =   3240
            TabIndex        =   15
            Top             =   870
            Width           =   2175
         End
         Begin VB.CommandButton cmdLLSearch 
            Caption         =   "&Search"
            Height          =   375
            Left            =   9480
            TabIndex        =   26
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdLLClose 
            Caption         =   "&Exit"
            Height          =   375
            Left            =   9480
            TabIndex        =   28
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdLLClear 
            Caption         =   "&Clear"
            Height          =   375
            Left            =   9480
            TabIndex        =   27
            Top             =   1200
            Width           =   1215
         End
         Begin Crystal.CrystalReport cRptLL 
            Left            =   6120
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin MSMask.MaskEdBox txtLLBordxDate 
            Height          =   315
            Left            =   3240
            TabIndex        =   18
            Top             =   870
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtLLBordxDateFr 
            Height          =   315
            Left            =   3240
            TabIndex        =   20
            Top             =   870
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtLLBordxDateTo 
            Height          =   315
            Left            =   6600
            TabIndex        =   21
            Top             =   870
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "This range only allow to 'Print All'"
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
            Left            =   5520
            TabIndex        =   127
            Top             =   2400
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Label lbl4 
            Caption         =   "From"
            Height          =   255
            Left            =   2160
            TabIndex        =   126
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lbl53 
            Caption         =   "Criteria Seletion"
            Height          =   255
            Left            =   240
            TabIndex        =   125
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lbl13 
            Caption         =   "Sequence"
            Height          =   255
            Left            =   5760
            TabIndex        =   59
            Top             =   1830
            Width           =   855
         End
         Begin VB.Label lbl12 
            Caption         =   "Print Option"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   1830
            Width           =   855
         End
         Begin VB.Label lbl11 
            Caption         =   "To"
            Height          =   255
            Left            =   5760
            TabIndex        =   57
            Top             =   870
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl10 
            Caption         =   "Bordx From"
            Height          =   255
            Left            =   2160
            TabIndex        =   56
            Top             =   870
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl9 
            Caption         =   "Bordereaux Range :"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   870
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lbl8 
            Caption         =   "Batch No"
            Height          =   255
            Left            =   5760
            TabIndex        =   54
            Top             =   870
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl6 
            Caption         =   "Bordereaux Batch :"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   870
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lbl5 
            Caption         =   "To"
            Height          =   255
            Left            =   5760
            TabIndex        =   52
            Top             =   870
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lbl3 
            Caption         =   "Members No. Range :"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   870
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lbl14 
            Caption         =   "Leaflet Version"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   2340
            Width           =   1335
         End
         Begin VB.Label lbl25 
            Caption         =   "Payor"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lbl1 
            Caption         =   "Single Member :"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   870
            Width           =   1455
         End
         Begin VB.Label lbl7 
            Caption         =   "Bordx Date "
            Height          =   255
            Left            =   2160
            TabIndex        =   48
            Top             =   870
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl2 
            Caption         =   "Member No."
            Height          =   255
            Left            =   2160
            TabIndex        =   51
            Top             =   870
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmReportMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const LEAF_PAYCODE = 0
Const LEAF_MEMBERNO = 1
Const LEAF_BORDXDATE = 2
Const LEAF_POLICYNO = 3
Const LEAF_HLTCODE = 4
Const LEAF_BATCHNO = 5
Const LEAF_NAME = 6
Const LEAF_PRINTED = 7

Const MCO_PAYCODE = 0
Const MCO_MEMBERNO = 1
Const MCO_BORDXDATE = 2
Const MCO_POLICYNO = 3
Const MCO_HLTCODE = 4
Const MCO_BATCHNO = 5
Const MCO_NAME = 6
Const MCO_PRINTED = 7

Const INV_PAYCODE = 0
Const INV_BORDXDATE = 1
Const INV_HLTCODE = 2
Const INV_BATCHNO = 3
Const INV_PRINTED = 4

Const CRD_PAYCODE = 0
Const CRD_MEMBERNO = 1
Const CRD_BORDXDATE = 2
Const CRD_POLICYNO = 3
Const CRD_HLTCODE = 4
Const CRD_BATCHNO = 5
Const CRD_NAME = 6
Const CRD_PRINTED = 7

Dim SDate As String
Dim DateFrom As String
Dim DateTo As String
Dim Exportsql As String

Private Sub RecordSearching_LeafLet(objListView As Object)
On Error Resume Next
Dim rsSearch As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim ColumnItem As ListItem
Dim sql As String
Dim itemToPrint As Integer
Dim isPrinted As Boolean
Dim CoverQty As Integer
Dim PrintM As String
Dim i As Integer
Dim A As Integer
Dim B As Integer

Screen.MousePointer = vbHourglass
If cboLLCriteriaSel.ListIndex = 0 Then
    objListView.listitems.clear
     sql = "SELECT * FROM MBM "
     sql = sql & " WHERE MBMNumber = '" & txtLLMemberNo & "'"
           
     sql = sql & " ORDER BY MBMNumber"
     rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
        
    If rsSearch.recordCount = 0 Then
       MsgBox "Record not found !", vbInformation + vbOKOnly
       Screen.MousePointer = Default
       Exit Sub
    End If
    
    Do While Not rsSearch.EOF
        Set ColumnItem = objListView.listitems.Add(, , rsSearch!PayCode)
        ColumnItem.SubItems(1) = rsSearch!MBMNumber
        ColumnItem.SubItems(2) = rsSearch!MBMBordxDate
        ColumnItem.SubItems(3) = rsSearch!MBMPolicyNo
        ColumnItem.SubItems(4) = Trim(rsSearch!HLTCode)
        ColumnItem.SubItems(5) = rsSearch!MBMBatchNo
        ColumnItem.SubItems(6) = rsSearch!MBMName
        ColumnItem.SubItems(3) = rsSearch!MBMBatchNo
       
        sql = "SELECT * FROM MBMPrinting"
        SDate = CStr(Month(rsSearch!MBMBordxDate)) & "/" & CStr(Day(rsSearch!MBMBordxDate)) & "/" & CStr(Year(rsSearch!MBMBordxDate))
        sql = sql & " WHERE MBMBordxDate = '" & SDate & "'"
        sql = sql & " AND PAYCode = '" & rsSearch!PayCode & "'"
        sql = sql & " AND HLTCode = '" & rsSearch!HLTCode & "'"
        sql = sql & " AND MBMBatchno = '" & rsSearch!MBMBatchNo & "'"
        sql = sql & " AND MBMNumber = '" & rsSearch!MBMNumber & "'"
        
        rsPrint.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        Select Case SSTab1.Tab
        Case 0 'Leaflet
            itemToPrint = LEAF_PRINTED
            cmdLLPrint.Enabled = True
        Case 1 'MCO Listing
            itemToPrint = MCO_PRINTED
            cmdMCOPrint.Enabled = True
            cmdMCOPrintPreview.Enabled = True
        Case 2 'Invoice
            itemToPrint = INV_PRINTED
            cmdInvPrint.Enabled = True
            cmdINVPrintPreview.Enabled = True
        Case 3 'Card Printing
            itemToPrint = CRD_PRINTED
            cmdCRDPrint.Enabled = True
            cmdCRDPrintPreview.Enabled = True
        End Select
        
        If rsPrint.recordCount > 0 Then
            Select Case SSTab1.Tab
            Case 0 'Leaflet
                isPrinted = rsPrint!MBMLeafletPrinted
            Case 1 'MCO Listing
                isPrinted = rsPrint!MBMMCOListPrinted
            Case 2 'Invoice
                isPrinted = rsPrint!MBMInvoiced
            Case 3 'Card Printing
                isPrinted = rsPrint!MBMExpListed
            End Select
            If isPrinted = True Then
                ColumnItem.SubItems(itemToPrint) = "YES"
            Else
                ColumnItem.SubItems(itemToPrint) = "NO"
            End If
        Else
            ColumnItem.SubItems(itemToPrint) = "NO"
        End If
        rsPrint.Close
        rsSearch.MoveNext
    Loop
    
    rsSearch.Close
   
    Set rsPrint = Nothing
    Set rsSearch = Nothing
Else
    If cboLLCriteriaSel.ListIndex = 1 Then
         objListView.listitems.clear
        If cboLLPrintSeq = "Membership No" Then
            sql = "SELECT * FROM MBM "
            sql = sql & " WHERE MBMNumber >= '" & txtLLMemberNoFr & "'"
            sql = sql & " AND MBMNumber <= '" & txtLLMemberNoTo & "'"
            
            sql = sql & " ORDER BY MBMNumber"
            rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
        Else
            If cboLLPrintSeq = "Policy No" Then
                sql = "SELECT * FROM MBM "
                sql = sql & " WHERE MBMNumber >= '" & txtLLMemberNoFr & "'"
                sql = sql & " AND MBMNumber <= '" & txtLLMemberNoTo & "'"
        
                sql = sql & " ORDER BY MBMPolicyNo"
                rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
            Else
                If cboLLPrintSeq = "Name" Then
                    sql = "SELECT * FROM MBM "
                    sql = sql & " WHERE MBMNumber >= '" & txtLLMemberNoFr & "'"
                    sql = sql & " AND MBMNumber <= '" & txtLLMemberNoTo & "'"
        
                    sql = sql & " ORDER BY MBMName"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                Else
                    If cboLLPrintSeq = "Group Company" Then
                        sql = "SELECT * FROM VIEW_LEAFLET_SORTBY_GRPCO "
                        sql = sql & " WHERE MBMNumber >= '" & txtLLMemberNoFr & "'"
                        sql = sql & " AND MBMNumber <= '" & txtLLMemberNoTo & "'"
                    
                        sql = sql & " ORDER BY MBMGroupCompany"
                        rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                    End If
                End If
            End If
        End If
        
         If rsSearch.recordCount = 0 Then
            MsgBox "Record not found !", vbInformation + vbOKOnly
            Screen.MousePointer = Default
            Exit Sub
         End If
         
         Do While Not rsSearch.EOF
             Set ColumnItem = objListView.listitems.Add(, , rsSearch!PayCode)
             ColumnItem.SubItems(1) = rsSearch!MBMNumber
             ColumnItem.SubItems(2) = rsSearch!MBMBordxDate
             ColumnItem.SubItems(3) = rsSearch!MBMPolicyNo
             ColumnItem.SubItems(4) = Trim(rsSearch!HLTCode)
             ColumnItem.SubItems(5) = rsSearch!MBMBatchNo
             ColumnItem.SubItems(6) = rsSearch!MBMName
             
             sql = "SELECT * FROM MBMPrinting"
             SDate = CStr(Month(rsSearch!MBMBordxDate)) & "/" & CStr(Day(rsSearch!MBMBordxDate)) & "/" & CStr(Year(rsSearch!MBMBordxDate))
             sql = sql & " WHERE MBMBordxDate = '" & SDate & "'"
             sql = sql & " AND PAYCode = '" & rsSearch!PayCode & "'"
             sql = sql & " AND HLTCode = '" & rsSearch!HLTCode & "'"
             sql = sql & " AND MBMBatchno = '" & rsSearch!MBMBatchNo & "'"
             sql = sql & " AND MBMNumber = '" & rsSearch!MBMNumber & "'"
             
             rsPrint.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
             Select Case SSTab1.Tab
             Case 0 'Leaflet
                 itemToPrint = LEAF_PRINTED
                 cmdLLPrint.Enabled = True
             Case 1 'MCO Listing
                 itemToPrint = MCO_PRINTED
                 cmdMCOPrint.Enabled = True
                 cmdMCOPrintPreview.Enabled = True
             Case 2 'Invoice
                 itemToPrint = INV_PRINTED
                 cmdInvPrint.Enabled = True
                 cmdINVPrintPreview.Enabled = True
             Case 3 'Card Printing
                 itemToPrint = CRD_PRINTED
                 cmdCRDPrint.Enabled = True
                 cmdCRDPrintPreview.Enabled = True
             End Select
             
             If rsPrint.recordCount > 0 Then
                 Select Case SSTab1.Tab
                 Case 0 'Leaflet
                     isPrinted = rsPrint!MBMLeafletPrinted
                 Case 1 'MCO Listing
                     isPrinted = rsPrint!MBMMCOListPrinted
                 Case 2 'Invoice
                     isPrinted = rsPrint!MBMInvoiced
                 Case 3 'Card Printing
                     isPrinted = rsPrint!MBMExpListed
                 End Select
                 If isPrinted = True Then
                     ColumnItem.SubItems(itemToPrint) = "YES"
                 Else
                     ColumnItem.SubItems(itemToPrint) = "NO"
                 End If
             Else
                 ColumnItem.SubItems(itemToPrint) = "NO"
             End If
             rsPrint.Close
             rsSearch.MoveNext
         Loop
         
         rsSearch.Close
        
         Set rsPrint = Nothing
         Set rsSearch = Nothing
    Else
        If cboLLCriteriaSel.ListIndex = 2 Then
            SDate = CStr(Month(txtLLBordxDate)) & "/" & CStr(Day(txtLLBordxDate)) & "/" & CStr(Year(txtLLBordxDate))
            
            objListView.listitems.clear
            If cboLLPrintSeq = "Membership No" Then
                sql = "SELECT * FROM MBM "
                sql = sql & " WHERE MBMBordxDate = '" & SDate & "'"
                sql = sql & " AND MBMBatchNo = '" & txtLLBatchNo & "'"
                sql = sql & " AND PAYCode = '" & Mid(cboLLPayor, 1, 2) & "'"
                
                sql = sql & " ORDER BY MBMNumber"
                rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
            Else
                If cboLLPrintSeq = "Policy No" Then
                    sql = "SELECT * FROM MBM "
                    sql = sql & " WHERE MBMBordxDate = '" & SDate & "'"
                    sql = sql & " AND MBMBatchNo = '" & txtLLBatchNo & "'"
                    sql = sql & " AND PAYCode = '" & Mid(cboLLPayor, 1, 2) & "'"
                    
                    sql = sql & " ORDER BY MBMPolicyNo"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                Else
                    If cboLLPrintSeq = "Name" Then
                        sql = "SELECT * FROM MBM "
                        sql = sql & " WHERE MBMBordxDate = '" & SDate & "'"
                        sql = sql & " AND MBMBatchNo = '" & txtLLBatchNo & "'"
                        sql = sql & " AND PAYCode = '" & Mid(cboLLPayor, 1, 2) & "'"
            
                        sql = sql & " ORDER BY MBMName"
                        rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                    Else
                        If cboLLPrintSeq = "Group Company" Then
                            sql = "SELECT * FROM VIEW_LEAFLET_SORTBY_GRPCO "
                            sql = sql & " WHERE MBMBordxDate = '" & SDate & "'"
                            sql = sql & " AND MBMBatchNo = '" & txtLLBatchNo & "'"
                            sql = sql & " AND PAYCode = '" & Mid(cboLLPayor, 1, 2) & "'"
                        
                            sql = sql & " ORDER BY MBMGroupCompany"
                            rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                        End If
                    End If
                End If
            End If
            
             If rsSearch.recordCount = 0 Then
                MsgBox "Record not found !", vbInformation + vbOKOnly
                Screen.MousePointer = Default
                Exit Sub
             End If
             
             Do While Not rsSearch.EOF
                 Set ColumnItem = objListView.listitems.Add(, , rsSearch!PayCode)
                 ColumnItem.SubItems(1) = rsSearch!MBMNumber
                 ColumnItem.SubItems(2) = rsSearch!MBMBordxDate
                 ColumnItem.SubItems(3) = rsSearch!MBMPolicyNo
                 ColumnItem.SubItems(4) = Trim(rsSearch!HLTCode)
                 ColumnItem.SubItems(5) = rsSearch!MBMBatchNo
                 ColumnItem.SubItems(6) = rsSearch!MBMName
                 
                 sql = "SELECT * FROM BordxPrinting"
                 SDate = CStr(Month(rsSearch!MBMBordxDate)) & "/" & CStr(Day(rsSearch!MBMBordxDate)) & "/" & CStr(Year(rsSearch!MBMBordxDate))
                 sql = sql & " WHERE MBMBordxDate = '" & SDate & "'"
                 sql = sql & " AND PAYCode = '" & rsSearch!PayCode & "'"
                 sql = sql & " AND HLTCode = '" & rsSearch!HLTCode & "'"
                 sql = sql & " AND MBMBatchno = '" & rsSearch!MBMBatchNo & "'"
                 sql = sql & " AND MBMNumber = '" & rsSearch!MBMNumber & "'"
                 
                 rsPrint.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
                 Select Case SSTab1.Tab
                 Case 0 'Leaflet
                     itemToPrint = LEAF_PRINTED
                     cmdLLPrint.Enabled = True
                 Case 1 'MCO Listing
                     itemToPrint = MCO_PRINTED
                     cmdMCOPrint.Enabled = True
                     cmdMCOPrintPreview.Enabled = True
                 Case 2 'Invoice
                     itemToPrint = INV_PRINTED
                     cmdInvPrint.Enabled = True
                     cmdINVPrintPreview.Enabled = True
                 Case 3 'Card Printing
                     itemToPrint = CRD_PRINTED
                     cmdCRDPrint.Enabled = True
                     cmdCRDPrintPreview.Enabled = True
                 End Select
                 
                 If rsPrint.recordCount > 0 Then
                     Select Case SSTab1.Tab
                     Case 0 'Leaflet
                         isPrinted = rsPrint!MBMLeafletPrinted
                     Case 1 'MCO Listing
                         isPrinted = rsPrint!MBMMCOListPrinted
                     Case 2 'Invoice
                         isPrinted = rsPrint!MBMInvoiced
                     Case 3 'Card Printing
                         isPrinted = rsPrint!MBMExpListed
                     End Select
                     If isPrinted = True Then
                         ColumnItem.SubItems(itemToPrint) = "YES"
                     Else
                         ColumnItem.SubItems(itemToPrint) = "NO"
                     End If
                 Else
                     ColumnItem.SubItems(itemToPrint) = "NO"
                 End If
                 rsPrint.Close
                 rsSearch.MoveNext
             Loop
             
             rsSearch.Close
            
             Set rsPrint = Nothing
             Set rsSearch = Nothing
        Else
            If cboLLCriteriaSel.ListIndex = 3 Then
                DateFrom = CStr(Month(txtLLBordxDateFr)) & "/" & CStr(Day(txtLLBordxDateFr)) & "/" & CStr(Year(txtLLBordxDateFr))
                DateTo = CStr(Month(txtLLBordxDateTo)) & "/" & CStr(Day(txtLLBordxDateTo)) & "/" & CStr(Year(txtLLBordxDateTo))
                
                objListView.listitems.clear
                If cboLLPrintSeq = "Membership No" Then
                    sql = "SELECT * FROM MBM "
                    sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                    sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
                    sql = sql & " AND PAYCode = '" & Mid(cboLLPayor, 1, 2) & "'"
                    
                    sql = sql & " ORDER BY MBMNumber"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                Else
                    If cboLLPrintSeq = "Policy No" Then
                        sql = "SELECT * FROM MBM "
                        sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                        sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
                        sql = sql & " AND PAYCode = '" & Mid(cboLLPayor, 1, 2) & "'"
                        
                        sql = sql & " ORDER BY MBMPolicyNo"
                        rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                    Else
                        If cboLLPrintSeq = "Name" Then
                            sql = "SELECT * FROM MBM "
                            sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                            sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
                            sql = sql & " AND PAYCode = '" & Mid(cboLLPayor, 1, 2) & "'"
                
                            sql = sql & " ORDER BY MBMName"
                            rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                        Else
                            If cboLLPrintSeq = "Group Company" Then
                                sql = "SELECT * FROM VIEW_LEAFLET_SORTBY_GRPCO "
                                sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                                sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
                                sql = sql & " AND PAYCode = '" & Mid(cboLLPayor, 1, 2) & "'"
                            
                                sql = sql & " ORDER BY MBMGroupCompany"
                                rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                            End If
                        End If
                    End If
                End If
                
                 If rsSearch.recordCount = 0 Then
                    MsgBox "Record not found !", vbInformation + vbOKOnly
                    Screen.MousePointer = Default
                    Exit Sub
                 End If
                 
                 Do While Not rsSearch.EOF
                     Set ColumnItem = objListView.listitems.Add(, , rsSearch!PayCode)
                     ColumnItem.SubItems(1) = rsSearch!MBMNumber
                     ColumnItem.SubItems(2) = rsSearch!MBMBordxDate
                     ColumnItem.SubItems(3) = rsSearch!MBMPolicyNo
                     ColumnItem.SubItems(4) = Trim(rsSearch!HLTCode)
                     ColumnItem.SubItems(5) = rsSearch!MBMBatchNo
                     ColumnItem.SubItems(6) = rsSearch!MBMName
                     
                     sql = "SELECT * FROM BordxPrinting"
                     SDate = CStr(Month(rsSearch!MBMBordxDate)) & "/" & CStr(Day(rsSearch!MBMBordxDate)) & "/" & CStr(Year(rsSearch!MBMBordxDate))
                     sql = sql & " WHERE MBMBordxDate = '" & SDate & "'"
                     sql = sql & " AND PAYCode = '" & rsSearch!PayCode & "'"
                     sql = sql & " AND HLTCode = '" & rsSearch!HLTCode & "'"
                     sql = sql & " AND MBMBatchno = '" & rsSearch!MBMBatchNo & "'"
                     sql = sql & " AND MBMNumber = '" & rsSearch!MBMNumber & "'"
                     
                     rsPrint.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
                     Select Case SSTab1.Tab
                     Case 0 'Leaflet
                         itemToPrint = LEAF_PRINTED
                         cmdLLPrint.Enabled = True
                     Case 1 'MCO Listing
                         itemToPrint = MCO_PRINTED
                         cmdMCOPrint.Enabled = True
                         cmdMCOPrintPreview.Enabled = True
                     Case 2 'Invoice
                         itemToPrint = INV_PRINTED
                         cmdInvPrint.Enabled = True
                         cmdINVPrintPreview.Enabled = True
                     Case 3 'Card Printing
                         itemToPrint = CRD_PRINTED
                         cmdCRDPrint.Enabled = True
                         cmdCRDPrintPreview.Enabled = True
                     End Select
                     
                     If rsPrint.recordCount > 0 Then
                         Select Case SSTab1.Tab
                         Case 0 'Leaflet
                             isPrinted = rsPrint!MBMLeafletPrinted
                         Case 1 'MCO Listing
                             isPrinted = rsPrint!MBMMCOListPrinted
                         Case 2 'Invoice
                             isPrinted = rsPrint!MBMInvoiced
                         Case 3 'Card Printing
                             isPrinted = rsPrint!MBMExpListed
                         End Select
                         If isPrinted = True Then
                             ColumnItem.SubItems(itemToPrint) = "YES"
                         Else
                             ColumnItem.SubItems(itemToPrint) = "NO"
                         End If
                     Else
                         ColumnItem.SubItems(itemToPrint) = "NO"
                     End If
                     rsPrint.Close
                     rsSearch.MoveNext
                 Loop
                 
                 rsSearch.Close
                
                 Set rsPrint = Nothing
                 Set rsSearch = Nothing
            End If
        End If
    End If
End If
    
    
    A = 0
    B = 0
    For i = 1 To lvwLeaflet.listitems.Count
        Set lvwLeaflet.SelectedItem = lvwLeaflet.listitems(i)   'lvwLeaflet.ListItems(ii)
            PrintM = lvwLeaflet.SelectedItem.SubItems(7)
            If PrintM = "YES" Then
                A = A + 1
            Else
                B = B + 1
            End If
        'Loop
    Next i
    
    If (A <> 0 And B <> 0) And cboLLPrintOpt.ListIndex = 0 Then
        Label1.Visible = True
        cboLLPrintOpt.ListIndex = 2
        cboLLPrintOpt.Enabled = False
    ElseIf (A <> 0 And B <> 0) And cboLLPrintOpt.ListIndex = 1 Then
        Label1.Visible = True
        cboLLPrintOpt.ListIndex = 2
        cboLLPrintOpt.Enabled = False
    ElseIf (A <> 0 And B <> 0) And cboLLPrintOpt.ListIndex = 2 Then
        LeafLetPrint "ALL"
    ElseIf (A = 0 And B <> 0) And cboLLPrintOpt.ListIndex = 1 Then
        LeafLetPrint "NO"
    ElseIf (A <> 0 And B = 0) And cboLLPrintOpt.ListIndex = 2 Then
        LeafLetPrint "YES"
    End If
Screen.MousePointer = Default

End Sub

Private Sub LeafLetPrint(PrintMode As String)
Dim ii As Integer
Dim MemberNo As String

Dim tmpBDate, tmpBatchNo, tmpPayCode, tmpHLTCode, tmpMemberNo As String
Dim tmpPrintedBDate, tmpPrintedBatchNo, tmpPrintedPaycode, tempPrintedHLTcode, tmpPrintedMemberNo As String
Dim BDateArray() As String
Dim BatchNoArray() As String
Dim PaycodeArray() As String
Dim HLTcodeArray() As String
Dim MemberNoArray() As String
Dim cnt As Integer

 Screen.MousePointer = vbHourglass
    cnt = 0
    If PrintMode = "ALL" Then
        For ii = 1 To lvwLeaflet.listitems.Count
         'If lvwLeaflet.ListItems.Count > 0 Then
            Set lvwLeaflet.SelectedItem = lvwLeaflet.listitems(ii)
            tmpMemberNo = lvwLeaflet.SelectedItem.SubItems(LEAF_MEMBERNO)
            tmpBDate = lvwLeaflet.SelectedItem.SubItems(LEAF_BORDXDATE)
            tmpHLTCode = lvwLeaflet.SelectedItem.SubItems(LEAF_HLTCODE)
            tmpBatchNo = lvwLeaflet.SelectedItem.SubItems(LEAF_BATCHNO)
            tmpPayCode = lvwLeaflet.SelectedItem
            
            cRptLL.ReportFileName = "C:\My Documents\RPT\Leaflet.rpt"
            cRptLL.Destination = crptToPrinter
            cRptLL.SelectionFormula = "{VIEW_LEAFLET.MBMNumber} = '" & tmpMemberNo & "'"
            'If cboLLPrintSeq.ListIndex = 0 Then
            '    ascending cRpt.SelectionFormula = "{VIEW_LEAFLET.MBMNumber} = '" & tmpMemberNo & "'"
            'End If
            cRpt.Action = 1
            
            If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And tmpPayCode = tmpPrintedPaycode And _
               tmpHLTCode = tempPrintedHLTcode And tmpMemberNo = tmpPrintedMemberNo Then
            Else
               tmpPrintedBDate = tmpBDate
               tmpPrintedBatchNo = tmpBatchNo
               tmpPrintedPaycode = tmpPayCode
               tempPrintedHLTcode = tmpHLTCode
               tmpPrintedMemberNo = tmpMemberNo
               cnt = cnt + 1
               ReDim Preserve BDateArray(cnt)
               ReDim Preserve BatchNoArray(cnt)
               ReDim Preserve PaycodeArray(cnt)
               ReDim Preserve HLTcodeArray(cnt)
               ReDim Preserve MemberNoArray(cnt)
               BDateArray(cnt) = tmpBDate
               BatchNoArray(cnt) = tmpBatchNo
               PaycodeArray(cnt) = tmpPayCode
               HLTcodeArray(cnt) = tmpHLTCode
               MemberNoArray(cnt) = tmpMemberNo
            End If
        'End If
        Next ii
    Else
        'For ii = 1 To lvwLeaflet.ListItems.Count
         'If lvwLeaflet.ListItems.Count > 0 Then
            For ii = 1 To lvwLeaflet.listitems.Count
                Set lvwLeaflet.SelectedItem = lvwLeaflet.listitems(ii)   'lvwLeaflet.ListItems(ii)
                If lvwLeaflet.SelectedItem.SubItems(LEAF_PRINTED) = PrintMode Then
                    tmpMemberNo = lvwLeaflet.SelectedItem.SubItems(LEAF_MEMBERNO)
                    tmpBDate = lvwLeaflet.SelectedItem.SubItems(LEAF_BORDXDATE)
                    tmpHLTCode = lvwLeaflet.SelectedItem.SubItems(LEAF_HLTCODE)
                    tmpBatchNo = lvwLeaflet.SelectedItem.SubItems(LEAF_BATCHNO)
                    tmpPayCode = lvwLeaflet.SelectedItem
                    
                    cRptLL.ReportFileName = "C:\My Documents\RPT\Leaflet.rpt"
                    cRptLL.Destination = crptToPrinter
                    cRptLL.SelectionFormula = "{VIEW_LEAFLET.MBMNumber} = '" & tmpMemberNo & "'"
                    cRptLL.Action = 1
                    
                
                    If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And tmpPayCode = tmpPrintedPaycode And _
                        tmpHLTCode = tempPrintedHLTcode And tmpMemberNo = tmpPrintedMemberNo Then
                    Else
                        tmpPrintedBDate = tmpBDate
                        tmpPrintedBatchNo = tmpBatchNo
                        tmpPrintedPaycode = tmpPayCode
                        tempPrintedHLTcode = tmpHLTCode
                        tmpPrintedMemberNo = tmpMemberNo
                        cnt = cnt + 1
                        ReDim Preserve BDateArray(cnt)
                        ReDim Preserve BatchNoArray(cnt)
                        ReDim Preserve PaycodeArray(cnt)
                        ReDim Preserve HLTcodeArray(cnt)
                        ReDim Preserve MemberNoArray(cnt)
                        BDateArray(cnt) = tmpBDate
                        BatchNoArray(cnt) = tmpBatchNo
                        PaycodeArray(cnt) = tmpPayCode
                        HLTcodeArray(cnt) = tmpHLTCode
                        MemberNoArray(cnt) = tmpMemberNo
                   End If
                End If
            Next ii
        'End If
   End If

    'Update MBMPrinting
    If cboLLPrintOpt.ListIndex = 0 And lvwLeaflet.SelectedItem.SubItems(LEAF_PRINTED) = "YES" Then
       MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        If cboLLPrintOpt.ListIndex = 1 And lvwLeaflet.SelectedItem.SubItems(LEAF_PRINTED) = "NO" Then
            MsgBox "Please select 'Print' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            For ii = 1 To UBound(MemberNoArray)
                UpdatePrintingFile BDateArray(ii), BatchNoArray(ii), PaycodeArray(ii), HLTcodeArray(ii), MemberNoArray(ii), "LEAFLET"
            Next ii
        End If
    End If
Screen.MousePointer = Default
End Sub

Private Sub UpdatePrintingFile(BordxDate As String, BatchNo As String, PayCode As String, HLTCode As String, MemberNo As String, WhichField As String)
'On Error Resume Next
'Dim BordxDate As String
'Dim BatchNo As String
'Dim PayCode As String
'Dim HLTCode As String
'Dim WhichField As String
Dim rsUpdate As New ADODB.Recordset
Dim sql As String
    
'Update file to MBMPrinting for Single Member
'If cboLLCriteriaSel.ListIndex = 0 Then
   'Update MBMPrinting
    sql = "SELECT * FROM MBMPrinting"
    sql = sql & " WHERE MBMBordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'"
    sql = sql & " AND MBMBatchNo = " & BatchNo & ""
    sql = sql & " AND PAYCode = '" & PayCode & "'"
    sql = sql & " AND HLTCode = '" & Trim(HLTCode) & "'"
    sql = sql & " AND MBMNumber = '" & MemberNo & "'"
    
    rsUpdate.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
    If rsUpdate.recordCount = 0 Then 'AddNew
        rsUpdate.AddNew
        rsUpdate!MBMBordxDate = BordxDate
        rsUpdate!MBMBatchNo = BatchNo
        rsUpdate!PayCode = PayCode
        rsUpdate!HLTCode = Trim(HLTCode)
        rsUpdate!MBMNumber = MemberNo
        
        Select Case WhichField
        Case "LEAFLET"
            rsUpdate!MBMLeafletPrinted = "1"
            rsUpdate!MBMLeafletPrintedDate = Date
        Case "MCO"
            rsUpdate!MBMMCOListPrinted = "1"
            rsUpdate!MBMMCOListDate = Date
        Case "INV"
            rsUpdate!MBMInvoiced = "1"
            rsUpdate!MBMInvoiceDate = Date
            rsUpdate!MBMInvoiceNo = TxtInvoiceYear & txtInvoiceNo
        Case "CRD"
            rsUpdate!MBMExpListed = "1"
            rsUpdate!MBMExpListDate = Date
        End Select
        rsUpdate.Update
    Else 'Update
        Select Case WhichField
        Case "LEAFLET"
            rsUpdate!MBMLeafletPrinted = "1"
            rsUpdate!MBMLeafletPrintedDate = Date
        Case "MCO"
            rsUpdate!MBMMCOListPrinted = "1"
            rsUpdate!MBMMCOListDate = Date
        Case "INV"
            rsUpdate!MBMInvoiced = "1"
            rsUpdate!MBMInvoiceDate = Date
            rsUpdate!MBMInvoiceNo = TxtInvoiceYear & txtInvoiceNo
        Case "CRD"
            rsUpdate!MBMExpListed = "1"
            rsUpdate!MBMExpListDate = Date
        End Select
        rsUpdate.Update
    End If
    rsUpdate.Close

End Sub

Sub ComboListing()
    Dim PAY As New ADODB.Recordset
    Dim sqlstr As String
   
    sqlstr = "select PAYCode , PAYCompanyName from PAY "
    
    PAY.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do While Not PAY.EOF
        cboLLPayor.AddItem (PAY!PayCode) & " - " & (PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub
Private Sub cboLLCriteriaSel_Click()
Select Case cboLLCriteriaSel.ListIndex
    Case 0
        lbl1.Visible = True
        lbl2.Visible = True
        txtLLMemberNo.Visible = True
        lbl3.Visible = False
        lbl4.Visible = False
        lbl5.Visible = False
        lbl6.Visible = False
        lbl7.Visible = False
        lbl8.Visible = False
        lbl9.Visible = False
        lbl10.Visible = False
        lbl11.Visible = False
        txtLLMemberNoFr.Visible = False
        txtLLMemberNoTo.Visible = False
        txtLLBordxDate.Visible = False
        txtLLBatchNo.Visible = False
        txtLLBordxDateFr.Visible = False
        txtLLBordxDateTo.Visible = False
        cboLLPayor.Enabled = False
    
    Case 1
        lbl1.Visible = False
        lbl2.Visible = False
        lbl3.Visible = True
        lbl4.Visible = True
        lbl5.Visible = True
        txtLLMemberNoFr.Visible = True
        txtLLMemberNoTo.Visible = True
        lbl6.Visible = False
        lbl7.Visible = False
        lbl8.Visible = False
        lbl9.Visible = False
        lbl10.Visible = False
        lbl11.Visible = False
        txtLLMemberNo.Visible = False
        txtLLBordxDate.Visible = False
        txtLLBatchNo.Visible = False
        txtLLBordxDateFr.Visible = False
        txtLLBordxDateTo.Visible = False
        cboLLPayor.Enabled = False
        
    Case 2
        lbl1.Visible = False
        lbl2.Visible = False
        lbl3.Visible = False
        lbl4.Visible = False
        lbl5.Visible = False
        lbl6.Visible = True
        lbl7.Visible = True
        lbl8.Visible = True
        txtLLBordxDate.Visible = True
        txtLLBatchNo.Visible = True
        lbl9.Visible = False
        lbl10.Visible = False
        lbl11.Visible = False
        txtLLMemberNo.Visible = False
        txtLLMemberNoFr.Visible = False
        txtLLMemberNoTo.Visible = False
        txtLLBordxDateFr.Visible = False
        txtLLBordxDateTo.Visible = False
        cboLLPayor.Enabled = True
    
    Case 3
        lbl1.Visible = False
        lbl2.Visible = False
        lbl3.Visible = False
        lbl4.Visible = False
        lbl5.Visible = False
        lbl6.Visible = False
        lbl7.Visible = False
        lbl8.Visible = False
        lbl9.Visible = True
        lbl10.Visible = True
        lbl11.Visible = True
        txtLLBordxDateFr.Visible = True
        txtLLBordxDateTo.Visible = True
        txtLLMemberNo.Visible = False
        txtLLMemberNoFr.Visible = False
        txtLLMemberNoTo.Visible = False
        txtLLBordxDate.Visible = False
        txtLLBatchNo.Visible = False
        cboLLPayor.Enabled = True
End Select
        
End Sub

Private Sub cmdLLClear_Click()
    txtLLMemberNoFr = ""
    txtLLMemberNoTo = ""
    txtLLMemberNo = ""
    txtLLBatchNo = ""
    cboLLPayor = ""
    txtLLBordxDate = "__/__/____"
    txtLLBordxDateFr = "__/__/____"
    txtLLBordxDateTo = "__/__/____"
    cboLLPrintOpt.ListIndex = 0
    cboLLPrintSeq.ListIndex = 0
    cboLeafletVer.ListIndex = 0
    cboLLPrintOpt.Enabled = True
    Label1.Visible = False
    lvwLeaflet.listitems.clear
       
End Sub

Private Sub cmdLLPrint_Click()

    If cboLLPrintOpt.ListIndex = 0 Then
        LeafLetPrint "NO"
    ElseIf cboLLPrintOpt.ListIndex = 1 Then
        LeafLetPrint "YES"
    ElseIf cboLLPrintOpt.ListIndex = 2 Then
        LeafLetPrint "ALL"
        cboLLPrintOpt.Enabled = True
    End If
    
End Sub

Private Sub cmdLLSearch_Click()
If cboLLCriteriaSel = "Single Member" And txtLLMemberNo = "" Then
    MsgBox "Please Enter Member No !", vbOKOnly + vbCritical, DialogBoxTitle
    Exit Sub
Else
    If cboLLCriteriaSel = "Members No Range" And txtLLMemberNoFr = "" And txtLLMemberNoTo = "" Then
        MsgBox "Please Enter Members No range !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If cboLLCriteriaSel = "Bordereaux Batch" And cboLLPayor = "" And txtLLBordxDate = "__/__/____" And txtLLBatchNo = "" Then
            MsgBox "Please Enter Bordereaux Date, Batch No and Payor !", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        Else
            If cboLLCriteriaSel = "Bordereaux Range" And cboLLPayor = "" And txtLLBordxDateFr = "__/__/____" And txtLLBordxDateTo = "__/__/____" Then
                MsgBox "Please Enter Bordereaux Range and Payor!", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            End If
        End If
    End If
End If

If cboLLPrintOpt = "" Then
    MsgBox "Please Enter Print Option !", vbOKOnly + vbCritical, DialogBoxTitle
    Exit Sub
Else
    If cboLLPrintSeq = "" Then
        MsgBox "Please Enter Sequence !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If cboLeafletVer = "" Then
            MsgBox "Please Enter Leaflet Version !", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    End If
End If

Select Case SSTab1.Tab
    Case 0  'Leaflet
        RecordSearching_LeafLet lvwLeaflet
    'Case 1  'MCO Listing
    '    RecordSearching lvwMCO
    'Case 2  'Invoice
    '    RecordSearching lvwInvoice
    'Case 3  'Card Printing
    '    RecordSearching lvwCRD
    End Select
End Sub

Private Sub Form_Load()
Call ComboListing
cboLLCriteriaSel.ListIndex = 0
cboLLPrintOpt.ListIndex = 0
cboLLPrintSeq.ListIndex = 0
cboLeafletVer.ListIndex = 0
cboLLPayor.Enabled = False
End Sub
