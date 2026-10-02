VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMBMCardExpSTMB 
   Caption         =   "Card Export (STMB)"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstMemberListing 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IC/No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Plan No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Effective Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Expiry Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Department"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Employee No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Branch"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Subsidary"
         Object.Width           =   2540
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Leaflet"
      TabPicture(0)   =   "frmMBMCardExpSTMB.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameLeaflet"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Card Export"
      TabPicture(1)   =   "frmMBMCardExpSTMB.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FrameExp"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Card &Export (Clinical)"
      TabPicture(2)   =   "frmMBMCardExpSTMB.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Update Card Sent Date"
      TabPicture(3)   =   "frmMBMCardExpSTMB.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameCardSent"
      Tab(3).Control(1)=   "Frame6"
      Tab(3).ControlCount=   2
      Begin VB.Frame FrameCardSent 
         Caption         =   "Enter Criteria For Searching"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   118
         Top             =   480
         Width           =   11415
         Begin VB.OptionButton OptCLN 
            Caption         =   "Clinical"
            Height          =   375
            Left            =   1560
            TabIndex        =   126
            Top             =   200
            Width           =   855
         End
         Begin VB.OptionButton OptHS 
            Caption         =   "H && S"
            Height          =   375
            Left            =   240
            TabIndex        =   125
            Top             =   200
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.ComboBox cboCriteriaCardSent 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":0070
            Left            =   1560
            List            =   "frmMBMCardExpSTMB.frx":007D
            TabIndex        =   124
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtCardBatchNo 
            Height          =   315
            Left            =   7800
            TabIndex        =   123
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtCardMemberNoTo 
            Height          =   315
            Left            =   7800
            TabIndex        =   122
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtCardMemberNoFr 
            Height          =   315
            Left            =   5520
            TabIndex        =   121
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboCardSentPayor 
            Height          =   315
            Left            =   1560
            TabIndex        =   120
            Top             =   950
            Width           =   3015
         End
         Begin VB.TextBox txtCardMemberNo 
            Height          =   315
            Left            =   5520
            TabIndex        =   119
            Top             =   600
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtCardBordxDate 
            Height          =   315
            Left            =   5520
            TabIndex        =   127
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCardBordxDateFr 
            Height          =   315
            Left            =   5520
            TabIndex        =   128
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCardBordxDateTo 
            Height          =   315
            Left            =   7800
            TabIndex        =   129
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDateSent 
            Height          =   285
            Left            =   5520
            TabIndex        =   130
            Top             =   950
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label15 
            Caption         =   "Date Sent"
            Height          =   255
            Left            =   4680
            TabIndex        =   139
            Top             =   1000
            Width           =   855
         End
         Begin VB.Label LblSel 
            Caption         =   "Criteria Selection"
            Height          =   255
            Left            =   240
            TabIndex        =   138
            Top             =   650
            Width           =   1335
         End
         Begin VB.Label lblTo 
            Caption         =   "To"
            Height          =   255
            Left            =   7080
            TabIndex        =   137
            Top             =   650
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblBordxFr 
            Caption         =   "Bordx From"
            Height          =   255
            Left            =   4680
            TabIndex        =   136
            Top             =   650
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblBatchNo 
            Caption         =   "Batch No"
            Height          =   255
            Left            =   7080
            TabIndex        =   135
            Top             =   650
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblFrom 
            Caption         =   "From"
            Height          =   255
            Left            =   4680
            TabIndex        =   134
            Top             =   650
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label LblMBMNo 
            Caption         =   "Member No."
            Height          =   255
            Left            =   4680
            TabIndex        =   133
            Top             =   650
            Width           =   975
         End
         Begin VB.Label lblBordxDate 
            Caption         =   "Bordx Date "
            Height          =   255
            Left            =   4680
            TabIndex        =   132
            Top             =   650
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblPayor 
            Caption         =   "Payor"
            Height          =   255
            Left            =   240
            TabIndex        =   131
            Top             =   1000
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Height          =   615
         Left            =   -74880
         TabIndex        =   114
         Top             =   1740
         Width           =   11415
         Begin VB.CommandButton cmdClearCard 
            Caption         =   "&Clear"
            Default         =   -1  'True
            Height          =   300
            Left            =   9600
            TabIndex        =   117
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdcardExit 
            Caption         =   "&Exit"
            Height          =   300
            Left            =   10440
            TabIndex        =   116
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update"
            Height          =   300
            Left            =   8760
            TabIndex        =   115
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   -74880
         TabIndex        =   106
         Top             =   1860
         Width           =   11415
         Begin VB.CommandButton cmdEXPClose2 
            Cancel          =   -1  'True
            Caption         =   "&Exit"
            Height          =   300
            Left            =   10440
            TabIndex        =   109
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdEXPClear2 
            Caption         =   "&Clear"
            Height          =   300
            Left            =   9600
            TabIndex        =   108
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdCRDPrint2 
            Caption         =   "E&xport"
            Height          =   300
            Left            =   8760
            TabIndex        =   107
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "Total Principal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   120
            TabIndex        =   113
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label txtEXPTotalPrincipal2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            TabIndex        =   112
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label txtEXPTotalSupplementary2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3720
            TabIndex        =   111
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Total Supp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2760
            TabIndex        =   110
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Enter Criteria For Searching"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   82
         Top             =   600
         Width           =   11415
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   8950
            TabIndex        =   91
            Top             =   760
            Width           =   255
         End
         Begin VB.ComboBox TxtCLNPlanCode2 
            Height          =   315
            Left            =   7800
            Sorted          =   -1  'True
            TabIndex        =   90
            Top             =   720
            Width           =   1095
         End
         Begin VB.ComboBox TxtCLNPlanCode1 
            Height          =   315
            Left            =   5520
            Sorted          =   -1  'True
            TabIndex        =   89
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox cboEXPCriteriaSel2 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":00B4
            Left            =   1560
            List            =   "frmMBMCardExpSTMB.frx":00C1
            TabIndex        =   88
            Top             =   360
            Width           =   3015
         End
         Begin VB.TextBox txtExpMemberNoTo2 
            Height          =   315
            Left            =   7800
            TabIndex        =   87
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtExpBatchNo2 
            Height          =   315
            Left            =   7800
            TabIndex        =   86
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtExpMemberNo2 
            Height          =   315
            Left            =   5520
            TabIndex        =   85
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox CboCLNPayor 
            Height          =   315
            Left            =   1560
            TabIndex        =   84
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtExpMemberNoFr2 
            Height          =   315
            Left            =   5520
            TabIndex        =   83
            Top             =   360
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtExpBordxDate2 
            Height          =   315
            Left            =   5520
            TabIndex        =   92
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtExpBordxDateFr2 
            Height          =   315
            Left            =   5520
            TabIndex        =   93
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtExpBordxDateTo2 
            Height          =   315
            Left            =   7800
            TabIndex        =   94
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            Caption         =   "To"
            Height          =   255
            Left            =   7080
            TabIndex        =   105
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Plan Code "
            Height          =   255
            Left            =   4680
            TabIndex        =   104
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Criteria Selection"
            Height          =   255
            Left            =   240
            TabIndex        =   103
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "To"
            Height          =   255
            Left            =   7080
            TabIndex        =   102
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Bordx From"
            Height          =   255
            Left            =   4680
            TabIndex        =   101
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Batch No"
            Height          =   255
            Left            =   7080
            TabIndex        =   100
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "To"
            Height          =   255
            Left            =   7560
            TabIndex        =   99
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "From"
            Height          =   255
            Left            =   4680
            TabIndex        =   98
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Member No."
            Height          =   255
            Left            =   4680
            TabIndex        =   97
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Bordx Date "
            Height          =   255
            Left            =   4680
            TabIndex        =   96
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Payor"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame FrameExp 
         Caption         =   "Enter Criteria For Searching"
         Height          =   1215
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   11415
         Begin VB.ComboBox cboCardSequence 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":00F8
            Left            =   5640
            List            =   "frmMBMCardExpSTMB.frx":0108
            TabIndex        =   67
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtExpMemberNo 
            Height          =   315
            Left            =   5640
            TabIndex        =   66
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboExpPayor 
            Height          =   315
            Left            =   1560
            TabIndex        =   65
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtExpMemberNoFr 
            Height          =   315
            Left            =   5640
            TabIndex        =   64
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtExpMemberNoTo 
            Height          =   315
            Left            =   7800
            TabIndex        =   63
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtExpBatchNo 
            Height          =   315
            Left            =   7800
            TabIndex        =   62
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboExpPrintOpt 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":0140
            Left            =   10920
            List            =   "frmMBMCardExpSTMB.frx":014D
            TabIndex        =   61
            Top             =   1200
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.ComboBox cboEXPCriteriaSel 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":0174
            Left            =   1560
            List            =   "frmMBMCardExpSTMB.frx":0184
            TabIndex        =   60
            Top             =   360
            Width           =   3015
         End
         Begin VB.PictureBox cRptExp 
            Height          =   480
            Left            =   4800
            ScaleHeight     =   420
            ScaleWidth      =   1140
            TabIndex        =   59
            Top             =   960
            Width           =   1200
         End
         Begin MSMask.MaskEdBox txtExpBordxDate 
            Height          =   315
            Left            =   5640
            TabIndex        =   68
            Top             =   360
            Visible         =   0   'False
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
            Left            =   5640
            TabIndex        =   69
            Top             =   360
            Visible         =   0   'False
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
            Left            =   7800
            TabIndex        =   70
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl13 
            Caption         =   "Sequence"
            Height          =   255
            Index           =   1
            Left            =   4680
            TabIndex        =   81
            Top             =   765
            Width           =   855
         End
         Begin VB.Label lbl49 
            Caption         =   "Payor"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lbl44 
            Caption         =   "Bordx Date "
            Height          =   255
            Left            =   4680
            TabIndex        =   79
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl39 
            Caption         =   "Member No."
            Height          =   255
            Index           =   0
            Left            =   4680
            TabIndex        =   78
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lbl41 
            Caption         =   "From"
            Height          =   255
            Left            =   4680
            TabIndex        =   77
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl42 
            Caption         =   "To"
            Height          =   255
            Left            =   7200
            TabIndex        =   76
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lbl45 
            Caption         =   "Batch No"
            Height          =   255
            Left            =   7080
            TabIndex        =   75
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl47 
            Caption         =   "Bordx From"
            Height          =   255
            Left            =   4680
            TabIndex        =   74
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl48 
            Caption         =   "To"
            Height          =   255
            Left            =   7080
            TabIndex        =   73
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl50 
            Caption         =   "Print Option"
            Height          =   255
            Left            =   9600
            TabIndex        =   72
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Criteria Selection"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   71
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame FrameLeaflet 
         Caption         =   "Enter Criteria For Searching"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   11295
         Begin VB.ComboBox cmbCMCName 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":01CB
            Left            =   8640
            List            =   "frmMBMCardExpSTMB.frx":01CD
            TabIndex        =   39
            Top             =   1440
            Width           =   2535
         End
         Begin VB.ComboBox CmbPlncode 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":01CF
            Left            =   5280
            List            =   "frmMBMCardExpSTMB.frx":01D1
            TabIndex        =   38
            Top             =   1440
            Width           =   2295
         End
         Begin VB.OptionButton optEdymium 
            Caption         =   "Edymium Leaflet"
            Height          =   255
            Left            =   3720
            TabIndex        =   37
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton OPTSpeLeaflet 
            Caption         =   "Clinical Leaflet"
            Height          =   255
            Left            =   2160
            TabIndex        =   36
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton OPTNorLeaflet 
            Caption         =   "Normal Leaflet"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.CheckBox chkREPCard 
            Caption         =   "Replacement Card"
            Height          =   255
            Left            =   6240
            TabIndex        =   34
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtLLHLTcode 
            Height          =   285
            Left            =   10320
            TabIndex        =   33
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cboLLCriteriaSel 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":01D3
            Left            =   2160
            List            =   "frmMBMCardExpSTMB.frx":01E3
            TabIndex        =   32
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtLLBatchNo 
            Height          =   315
            Left            =   6960
            TabIndex        =   31
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtLLMemberNoTo 
            Height          =   315
            Left            =   7080
            TabIndex        =   30
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtLLMemberNoFr 
            Height          =   315
            Left            =   4560
            TabIndex        =   29
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cboLLPayor 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":022C
            Left            =   2160
            List            =   "frmMBMCardExpSTMB.frx":022E
            TabIndex        =   28
            Top             =   885
            Width           =   3975
         End
         Begin VB.TextBox txtLLMemberNo 
            Height          =   315
            Left            =   4560
            TabIndex        =   27
            Top             =   600
            Width           =   1575
         End
         Begin VB.PictureBox cRptLL 
            Height          =   480
            Left            =   8640
            ScaleHeight     =   420
            ScaleWidth      =   1140
            TabIndex        =   26
            Top             =   240
            Width           =   1200
         End
         Begin VB.ComboBox cboLLPrintOpt 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":0230
            Left            =   7320
            List            =   "frmMBMCardExpSTMB.frx":023D
            TabIndex        =   25
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cboLeafletVer 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":0260
            Left            =   2160
            List            =   "frmMBMCardExpSTMB.frx":0279
            TabIndex        =   24
            Top             =   1155
            Width           =   1815
         End
         Begin VB.ComboBox cboLLPrintSeq 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":02D8
            Left            =   4800
            List            =   "frmMBMCardExpSTMB.frx":02F4
            TabIndex        =   23
            Top             =   1155
            Width           =   1335
         End
         Begin VB.ComboBox cboPrintMode 
            Height          =   315
            ItemData        =   "frmMBMCardExpSTMB.frx":0351
            Left            =   2160
            List            =   "frmMBMCardExpSTMB.frx":0353
            TabIndex        =   22
            Top             =   1440
            Width           =   1815
         End
         Begin MSMask.MaskEdBox txtLLBordxDate 
            Height          =   315
            Left            =   4560
            TabIndex        =   40
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtLLBordxDateFr 
            Height          =   315
            Left            =   4560
            TabIndex        =   41
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtLLBordxDateTo 
            Height          =   315
            Left            =   6960
            TabIndex        =   42
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSComDlg.CommonDialog dlgCommonDialog 
            Left            =   8640
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label20 
            Caption         =   "CMC Name"
            Height          =   255
            Left            =   7680
            TabIndex        =   57
            Top             =   1470
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   4080
            TabIndex        =   56
            Top             =   1470
            Width           =   1095
         End
         Begin VB.Label lbl14 
            Caption         =   "Printing Mode"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   55
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lbl4 
            Caption         =   "From"
            Height          =   255
            Left            =   4080
            TabIndex        =   54
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lbl53 
            Caption         =   "Single/Batch/Range"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lbl13 
            Caption         =   "Sequence"
            Height          =   255
            Index           =   0
            Left            =   4050
            TabIndex        =   52
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lbl12 
            Caption         =   "Print Option"
            Height          =   255
            Left            =   7560
            TabIndex        =   51
            Top             =   840
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl11 
            Caption         =   "To"
            Height          =   255
            Left            =   6360
            TabIndex        =   50
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lbl10 
            Caption         =   "From"
            Height          =   255
            Left            =   4200
            TabIndex        =   49
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl8 
            Caption         =   "Batch No"
            Height          =   255
            Left            =   6240
            TabIndex        =   48
            Top             =   600
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl5 
            Caption         =   "To"
            Height          =   255
            Left            =   6480
            TabIndex        =   47
            Top             =   600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lbl14 
            Caption         =   "Leaflet Version"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   46
            Top             =   1155
            Width           =   1335
         End
         Begin VB.Label lbl25 
            Caption         =   "Payor"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   915
            Width           =   1095
         End
         Begin VB.Label lbl7 
            Caption         =   "Date "
            Height          =   255
            Left            =   4200
            TabIndex        =   44
            Top             =   600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lbl2 
            Caption         =   "From"
            Height          =   255
            Left            =   4200
            TabIndex        =   43
            Top             =   600
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   -74880
         TabIndex        =   12
         Top             =   2340
         Width           =   11295
         Begin VB.CommandButton cmdLLSearch 
            Caption         =   "&Search"
            Height          =   300
            Left            =   7800
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdLLPrint 
            Caption         =   "&Print"
            Height          =   300
            Left            =   8640
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdLLClear 
            Caption         =   "&Clear"
            Height          =   300
            Left            =   9480
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdLLClose 
            Caption         =   "&Exit"
            Height          =   300
            Left            =   10320
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.Label txtLLTotalPrincipal 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1200
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label txtLLTotalSupplementary1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3600
            TabIndex        =   19
            Top             =   240
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Total Supp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2640
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Total Principal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   1860
         Width           =   11415
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   300
            Left            =   7920
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdEXPSearch 
            Caption         =   "&Search"
            Enabled         =   0   'False
            Height          =   300
            Left            =   7920
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdCRDPrint 
            Caption         =   "E&xport"
            Height          =   300
            Left            =   8760
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdEXPClear 
            Caption         =   "&Clear"
            Height          =   300
            Left            =   9600
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdEXPClose 
            Caption         =   "&Exit"
            Height          =   300
            Left            =   10440
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label208 
            Alignment       =   2  'Center
            Caption         =   "Total Supp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2760
            TabIndex        =   11
            Top             =   240
            Width           =   960
         End
         Begin VB.Label txtEXPTotalSupplementary1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3720
            TabIndex        =   10
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label txtEXPTotalPrincipal 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            TabIndex        =   9
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label207 
            Alignment       =   2  'Center
            Caption         =   "Total Principal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1200
         End
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Card Export (STMB)"
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
      TabIndex        =   141
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Member Listing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   140
      Top             =   3240
      Width           =   1200
   End
End
Attribute VB_Name = "frmMBMCardExpSTMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Dim SDateTo As String
Dim sDate As String
Dim DateFrom As String
Dim DateTo As String
Dim Exportsql As String
Dim MBMNo As String
Dim PolNo As String
Dim HLTCode As String
Dim GroupCompany As String
Dim GrpCode As String
Dim EmployeeNo As String
Dim Subsidary As String
Dim PLNPayorCode As String
Dim Joindt As String
Dim AgentCode As String
Dim AgentName As String
Dim POLSubNo1 As String
Dim Department As String
Dim MBMICNo As String
Dim MBMICNo1 As String
Dim MBMICNo2 As String
Dim PLNCode As String
Dim MBMName As String
Dim ExpDate As String
Dim MBMCName As String
Dim MBMCCoverID As String
Dim MBMCICNo  As String
Dim MBMBranch  As String
Dim Printed As String
Dim CancelPrinting As Boolean
Dim bPrintPass As Boolean
Dim TotalSupplementary As Integer
Dim INVNo As String
Dim Search As Boolean
Dim row As Integer
Dim Rcounter As Integer
Dim Check As Boolean
Dim tmpBDate, tmpBatchNo, tmpPayCode, tmpHltCode, tmpMemberNo As String

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Private Sub cboCardSentPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCardSentPayor.Text, cboCardSentPayor.SelStart) + Chr(KeyAscii) + Mid(cboCardSentPayor.Text, cboCardSentPayor.SelStart + cboCardSentPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCardSentPayor.ListCount - 1
 
        If NewText = LCase(Left(cboCardSentPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCardSentPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCardSentPayor.Text = ValidValue
                cboCardSentPayor.SelStart = 0
                cboCardSentPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboLLPayor_LostFocus()
    Dim PAY As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim SqlQry As String
    SqlQry = "Select PLnCode, PLNDescription  from PLN  where PAYCode ='" + Mid(cboLLPayor, 1, 2) + "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = SqlQry
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    CmbPlncode.Clear
    Do While Not PAY.EOF
        With PAY
            CmbPlncode.AddItem !PLNCode & " - " & !PLNDescription
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
       
    '
   ' Dim StrQry As String
   ' Dim Grpcompany() As String
   ' Grpcompany = Split(cboLLPayor, "-")
   ' StrQry = "  select  Subsidiary, CMCName, MBMGroupCompany   from MBMOthers  Where LTrim(CMCName) Is Not Null And LTrim(CMCName) <>  N''   Group by Subsidiary, CMCName, MBMGroupCompany  "
   ' PAY.Open StrQry, medixmasterconn, adOpenKeyset, adLockOptimistic
   ' cmbCMCName.Clear
   '     Do While Not PAY.EOF
   '         cmbCMCName.AddItem (PAY!Subsidiary)
   '     PAY.MoveNext
   ' Loop
   ' PAY.Close
   ' Set PAY = Nothing
    
End Sub

Private Sub Check1_Click()
Dim Plan As New ADODB.Recordset
Dim sqlstr As String
   
    If Check1.Value = 1 Then
        sqlstr = " Select PLNCode From PLN "
        Plan.Open sqlstr, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
        Do While Not Plan.EOF
            TxtCLNPlanCode1.AddItem Plan!PLNCode
            TxtCLNPlanCode2.AddItem Plan!PLNCode
            Plan.MoveNext
        Loop
    
        Plan.Close
        Set Plan = Nothing
    Else
        TxtCLNPlanCode1.Clear
        TxtCLNPlanCode2.Clear
    End If
End Sub

Private Sub chkREPCard_Click()
    If chkREPCard.Value = 1 Then
        With cboLeafletVer
            cboLeafletVer.Clear
            .AddItem "Sihat Malaysia(Replacement Card)"
            .AddItem "Group Company(Replacement Card)"
            .AddItem "Dept and Emp No(Replacement Card)"
            .AddItem "Dept(Replacement Card)"
            .AddItem "Branch(Replacement Card)"
            .AddItem "By IC"
        End With
    Else
        With cboLeafletVer
            cboLeafletVer.Clear
            .AddItem "Sihat Malaysia"
            .AddItem "Group Company"
            .AddItem "Dept and Emp No"
            .AddItem "Dept"
            .AddItem "Branch"
            .AddItem "Without Address"
            .AddItem "By IC"
        End With
    End If
End Sub

Private Sub cmdcardExit_Click()
    Unload Me
End Sub

Private Sub cmdClearCard_Click()
    txtCardMemberNoFr = ""
    txtCardMemberNoTo = ""
    txtCardMemberNo = ""
    txtCardBatchNo = ""
    cboCardSentPayor = ""
    txtCardBordxDate = "__/__/____"
    txtCardBordxDateFr = "__/__/____"
    txtCardBordxDateTo = "__/__/____"
    txtDateSent = "__/__/____"
End Sub

Private Sub cmdCRDPrint2_Click()
Dim RecordExit

    cmdCRDPrint2.Enabled = False
    If cboEXPCriteriaSel2.ListIndex = 0 And txtExpMemberNo2 = "" Then
        MsgBox "Please Enter Member No !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If cboEXPCriteriaSel2.ListIndex = 1 And (txtExpMemberNoFr2 = "" Or txtExpMemberNoTo2 = "") Then
            MsgBox "Please Enter Members No range !", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        Else
            If cboEXPCriteriaSel2.ListIndex = 2 And (txtExpBordxDate2 = "__/__/____" Or txtExpBatchNo2 = "") Then
                MsgBox "Please Enter Bordereaux Date, Batch No !", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            ElseIf cboEXPCriteriaSel2.ListIndex = 2 And CboCLNPayor = "" Then
                MsgBox "Please Enter Payor Code !", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            End If
        End If
    End If

    RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        Screen.MousePointer = vbHourglass
           ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            
                Call EXPPrincipalListing2
            
          '  medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        Screen.MousePointer = Default
    End If
    cmdCRDPrint2.Enabled = True
End Sub

Private Sub cmdEXPClear2_Click()
    txtExpMemberNoFr2 = ""
    txtExpMemberNoTo2 = ""
    txtExpMemberNo2 = ""
    txtExpBatchNo2 = ""
    cboExpPayor2 = ""
    txtExpBordxDate2 = "__/__/____"
    txtExpBordxDateFr2 = "__/__/____"
    txtExpBordxDateTo2 = "__/__/____"
    txtEXPTotalPrincipal2 = 0
    txtEXPTotalSupplementary2 = 0
End Sub

Private Sub cmdEXPClose2_Click()
    Unload Me
End Sub

Private Sub cboCriteriaCardSent_Click()
Select Case cboCriteriaCardSent.ListIndex
    Case 0
        LblMBMNo.Visible = True
        txtCardMemberNo.Visible = True
        lblFrom.Visible = False
        lblBordxDate.Visible = False
        lblBatchNo.Visible = False
        lblBordxFr.Visible = False
        lblTo.Visible = False
        txtCardMemberNoFr.Visible = False
        txtCardMemberNoTo.Visible = False
        txtCardBordxDate.Visible = False
        txtCardBatchNo.Visible = False
        txtCardBordxDateFr.Visible = False
        txtCardBordxDateTo.Visible = False
        cboCardSentPayor.Enabled = False
        txtCardMemberNoFr = ""
        txtCardMemberNoTo = ""
        txtCardMemberNo = ""
        txtCardBatchNo = ""
        cboCardSentPayor = ""
        txtCardBordxDate = "__/__/____"
        txtCardBordxDateFr = "__/__/____"
        txtCardBordxDateTo = "__/__/____"
            
    Case 1
        LblMBMNo.Visible = False
        lblFrom.Visible = True
        txtCardMemberNoFr.Visible = True
        txtCardMemberNoTo.Visible = True
        lblBordxDate.Visible = False
        lblBatchNo.Visible = False
        lblBordxFr.Visible = False
        lblTo.Visible = True
        txtCardMemberNo.Visible = False
        txtCardBordxDate.Visible = False
        txtCardBatchNo.Visible = False
        txtCardBordxDate.Visible = False
        txtCardBordxDateFr.Visible = False
        txtCardBordxDateTo.Visible = False
        cboCardSentPayor.Enabled = False
        txtCardMemberNoTo = ""
        txtCardMemberNo = ""
        txtCardBatchNo = ""
        cboCardSentPayor = ""
        txtCardBordxDate = "__/__/____"
        txtCardBordxDateFr = "__/__/____"
        txtCardBordxDateTo = "__/__/____"
                
    Case 2
        LblMBMNo.Visible = False
        lblFrom.Visible = False
        lblBordxDate.Visible = True
        lblBatchNo.Visible = True
        txtCardBordxDate.Visible = True
        txtCardBatchNo.Visible = True
        lblBordxFr.Visible = False
        lblTo.Visible = False
        txtCardMemberNo.Visible = False
        txtCardMemberNoFr.Visible = False
        txtCardMemberNoTo.Visible = False
        txtCardBordxDateFr.Visible = False
        cboCardSentPayor.Enabled = True
        txtCardMemberNoTo = ""
        txtCardMemberNo = ""
        txtCardBatchNo = ""
        cboCardSentPayor = ""
        txtCardBordxDate = "__/__/____"
        txtCardBordxDateFr = "__/__/____"
        txtCardBordxDateTo = "__/__/____"
        
    
End Select
End Sub

Private Sub cboCriteriaCardSent_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboCriteriaCardSent.Text, cboCriteriaCardSent.SelStart) + Chr(KeyAscii) + Mid(cboCriteriaCardSent.Text, cboCriteriaCardSent.SelStart + cboCriteriaCardSent.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboCriteriaCardSent.ListCount - 1
        
        If NewText = LCase(Left(cboCriteriaCardSent.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCriteriaCardSent.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboCriteriaCardSent.Text = ValidValue
               cboCriteriaCardSent.SelStart = 0
               cboCriteriaCardSent.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        If lstMemberListing.ListItems.Count > 0 Then
            Call ExportListViewtoExcel(lstMemberListing)
        End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdUpdate_Click()
    
    cmdUpdate.Enabled = False
    If cboCriteriaCardSent.ListIndex = 0 And txtCardMemberNo = "" Then
        MsgBox "Please Enter Member No!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdUpdate.Enabled = True
        Exit Sub
    ElseIf txtDateSent = "__/__/____" Then
        MsgBox "Please Enter Date Sent!", vbOKOnly + vbCritical, DialogBoxTitle
        txtDateSent.SetFocus
        cmdUpdate.Enabled = True
        Exit Sub
    Else
        If cboCriteriaCardSent.ListIndex = 1 And (txtCardMemberNoFr = "" Or txtCardMemberNoTo = "") Then
            MsgBox "Please Enter Members No range!", vbOKOnly + vbCritical, DialogBoxTitle
            cmdUpdate.Enabled = True
            Exit Sub
        ElseIf txtDateSent = "__/__/____" Then
            MsgBox "Please Enter Date Sent!", vbOKOnly + vbCritical, DialogBoxTitle
            txtDateSent.SetFocus
            cmdUpdate.Enabled = True
            Exit Sub
        Else
            If cboCriteriaCardSent.ListIndex = 2 And (txtCardBordxDate = "__/__/____" Or txtCardBatchNo = "" Or cboCardSentPayor = "") Then
                MsgBox "Please Enter Bordereaux Date, Batch No and Payor!", vbOKOnly + vbCritical, DialogBoxTitle
                cmdUpdate.Enabled = True
                Exit Sub
            ElseIf txtDateSent = "__/__/____" Then
                MsgBox "Please Enter Date Sent!", vbOKOnly + vbCritical, DialogBoxTitle
                txtDateSent.SetFocus
                cmdUpdate.Enabled = True
                Exit Sub
            End If
        End If
    End If
    
    Screen.MousePointer = vbHourglass
        If OptHS.Value = True Then
            Call CardSentListing
        ElseIf OptCLN.Value = True Then
            Call CardSentListing2
        End If
    Screen.MousePointer = Default
    cmdUpdate.Enabled = True
End Sub

Private Sub CardSentListing()
Dim strsql As String
Dim strsqls As String
Dim sql As String
Dim MBM As New ADODB.Recordset
Dim SentDate As String

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Select Case cboCriteriaCardSent.ListIndex
        Case 0 'Card Sent for Single Member
            sql = " SELECT MBMNumber, INSCode From MBM WHERE MBMNumber = '" & Trim(txtCardMemberNo) & "'" & _
                     " AND MBMDataStatus <> 'D' AND MBMStatus <> 'C' AND HSCLNInd <> 'C' "
            
            MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
        Case 1 'Card Sent for Member's No Range
            sql = "SELECT MBMNumber, INSCode From MBM WHERE MBMNumber >= '" & txtCardMemberNoFr & "'"
            sql = sql & " AND MBMNumber <= '" & txtCardMemberNoTo & "'"
            sql = sql & " AND MBMDataStatus <> 'D'"
            sql = sql & " AND MBMStatus <> 'C'"
            sql = sql & " AND HSCLNInd <> 'C'"
            
            MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        Case 2
            sDate = CStr(Month(txtCardBordxDate)) & "/" & CStr(Day(txtCardBordxDate)) & "/" & CStr(Year(txtCardBordxDate))
            
            sql = "SELECT MBMNumber, INSCode From MBM "
            sql = sql & " WHERE MBMBordxDate = '" & sDate & "'"
            sql = sql & " AND MBMBatchNo = '" & txtCardBatchNo & "'"
            sql = sql & " AND PAYCode = '" & Mid(cboCardSentPayor, 1, 2) & "'"
            sql = sql & " AND MBMDataStatus <> 'D'"
            sql = sql & " AND MBMStatus <> 'C'"
            sql = sql & " AND HSCLNInd <> 'C'"
            
            MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
    End Select
    
                
    If MBM.EOF = True Then
        MsgBox "Record not found !", vbInformation + vbOKOnly
        Screen.MousePointer = Default
        
        MBM.Close
        Set MBM = Nothing
        
     '   medixmasterconn.Close
     '   Set medixmasterconn = Nothing
        
        Exit Sub
    End If
        
    Do Until MBM.EOF
                
        SentDate = Format(txtDateSent, "YYYY/MM/DD")
        
        strsqls = " update MBMTwo set " & _
                  " MBMCardSentDate = '" & SentDate & "', " & _
                  " MBMCardSentUser = '" & Logon & "' " & _
                  " Where MBMNumber = '" & Trim(MBM!MBMNumber) & "'"
    
        medixmasterconn.Execute strsqls
        
        If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
            
            SentDate = Format(txtDateSent, "YYYY/MM/DD")
        
            strsql = " update MBMCoveredPersonsTwo set " & _
                      " MBMCCardSentDate = '" & SentDate & "', " & _
                      " MBMCCardSentUser = '" & Logon & "' " & _
                      " Where MBMCNumber = '" & Trim(MBM!MBMNumber) & "'"
    
            medixmasterconn.Execute strsql
            
        End If
    
    MBM.MoveNext
    Loop
    
    MsgBox "Update successful!", vbOKOnly + vbInformation, DialogBoxTitle
            
    MBM.Close
    Set MBM = Nothing
    
    medixmasterconn.Close
    Set medixmasterconn = Nothing

End Sub

Private Sub CardSentListing2()
Dim strsql As String
Dim strsqls As String
Dim sql As String
Dim MBMCLN As New ADODB.Recordset
Dim SentDate As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter

  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Select Case cboCriteriaCardSent.ListIndex
        Case 0 'Card Sent for Single Member
            
            sql = " AND MBMNumber = '" & Trim(txtCardMemberNo) & "'"
            
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "CardSentList"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
            cmd.Parameters.Append Prm1
            MBMCLN.CursorLocation = adUseClient
            MBMCLN.CursorType = adOpenForwardOnly
            MBMCLN.CacheSize = 100
            Set MBMCLN = cmd.Execute
            
        Case 1 'Card Sent for Member's No Range
                       
            sql = " AND MBMNumber >= '" & txtCardMemberNoFr & "' AND MBMNumber <= '" & txtCardMemberNoTo & "'"
            
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "CardSentList"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
            cmd.Parameters.Append Prm1
            MBMCLN.CursorLocation = adUseClient
            MBMCLN.CursorType = adOpenForwardOnly
            MBMCLN.CacheSize = 100
            Set MBMCLN = cmd.Execute
        
        Case 2
            sDate = CStr(Month(txtCardBordxDate)) & "/" & CStr(Day(txtCardBordxDate)) & "/" & CStr(Year(txtCardBordxDate))
            
            sql = " AND MBMBordxDate = '" & sDate & "' AND MBMBatchNo = '" & txtCardBatchNo & "'" & _
                  " AND PAYCode = '" & Mid(cboCardSentPayor, 1, 2) & "'"
            
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "CardSentList"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
            cmd.Parameters.Append Prm1
            MBMCLN.CursorLocation = adUseClient
            MBMCLN.CursorType = adOpenForwardOnly
            MBMCLN.CacheSize = 100
            Set MBMCLN = cmd.Execute
            
    End Select
    
                
    If MBMCLN.EOF = True Then
        MsgBox "Record not found !", vbInformation + vbOKOnly
        Screen.MousePointer = Default
        
        MBMCLN.Close
        Set MBMCLN = Nothing
        
    '    medixmasterconn.Close
     '   Set medixmasterconn = Nothing
        
        Exit Sub
    End If
        
    Do Until MBMCLN.EOF
                
        SentDate = Format(txtDateSent, "YYYY/MM/DD")
        
        strsqls = " update MBMCLNOthers set " & _
                  " MBMCLNCardSentDate = '" & SentDate & "', " & _
                  " MBMCLNCardSentUser = '" & Logon & "' " & _
                  " Where MBMNumber = '" & Trim(MBMCLN!MBMNumber) & "'"
    
        medixmasterconn.Execute strsqls
        
        If Trim(MBMCLN!InsType) = "F" Or Trim(MBMCLN!InsType) = "H" Then
            
            SentDate = Format(txtDateSent, "YYYY/MM/DD")
        
            strsql = " update MBMCoveredPersonsCLN set " & _
                      " MBMCCLNCardSentDate = '" & SentDate & "', " & _
                      " MBMCCLNCardSentUser = '" & Logon & "' " & _
                      " Where MBMCNumber = '" & Trim(MBMCLN!MBMNumber) & "'"
    
            medixmasterconn.Execute strsql
            
        End If
    
    MBMCLN.MoveNext
    Loop
    
    MsgBox "Update successful!", vbOKOnly + vbInformation, DialogBoxTitle
            
    MBMCLN.Close
    Set MBMCLN = Nothing
    
    medixmasterconn.Close
    Set medixmasterconn = Nothing

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

Public Function CreatePath(NewPath) As Boolean
    'Add a trailing slash if none
    If Right(NewPath, 1) <> "\" Then
        NewPath = NewPath & "\"
    End If
    
    'Call API
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        'No errors, return True
        CreatePath = True
    End If

End Function
 
'Private Function PrintFunction() As Boolean
'On Error GoTo CancelPrint
'Select Case SSTab1.Tab
'    Case 0
'        If cRptLL Is Nothing Then Exit Function
'
'        With dlgCommonDialog
'            .DialogTitle = "Print Crystal Report "
'            .CancelError = True
'            .Flags = cdlPDNoSelection + cdlPDNoPageNums
'            .ShowPrinter
'
'            If Err <> MSComDlg.cdlCancel Then
'                cRptLL.PrintReport
 '           End If
'        End With
'        PrintFunction = True
'        Exit Function
'CancelPrint:
'        PrintFunction = False
'        CancelPrinting = True
'
'    Case 1
 '       If cRptMCO Is Nothing Then Exit Function
 '
 '           With dlgCommonDialog_MCO
 '               .DialogTitle = "Print Crystal Report "
 '               .CancelError = True
 '               .Flags = cdlPDNoSelection + cdlPDNoPageNums
 '               .ShowPrinter
 '
 '               If Err <> MSComDlg.cdlCancel Then
 '                   cRptMCO.PrintReport
 '               End If
 '           End With
 '           PrintFunction = True
 '           Exit Function
'cancelPrint1:
'                PrintFunction = False
'                CancelPrinting = True
'                Exit Function'

    'Case 2
    '    If cRptINV Is Nothing Then Exit Function
   '
   '         With dlgCommonDialog_INV
   '             .DialogTitle = "Print Crystal Report "
   '             .CancelError = True
   '              'For ii = 1 To UBound(MemberNoArray)
   '                 .Flags = cdlPDNoSelection + cdlPDNoPageNums
   '                 'If ActiveForm.rtfText.SelLength = 0 Then
   '                  '   .Flags = .Flags + cdlPDAllPages
   '                 'Else
   '                  '   .Flags = .Flags + cdlPDSelection
   '                 'End If
   '
   '             .ShowPrinter
   '
   '             If Err <> MSComDlg.cdlCancel Then
   '                 'If cboINVPrintOpt.ListIndex = 0 Then
   '                 '    Call IncInvoiceNo
   '                 '    cRptINV.ParameterFields(0) = "MCOInvoiceNo;" & (TxtInvoiceYear & txtInvoiceNo) & ";true"
   '                 '    cRptINV.PrintReport
   '                 'Else
   '                 '    If cboINVPrintOpt.ListIndex = 1 Then
   '                 '        cRptINV.PrintReport
   '                 '    End If
   '                 'End If
   '             End If
   '         ' Next ii
   '         End With
   '         PrintFunction = True
   '         Exit Function
'can celPrint2:
'                PrintFunction = False
'                CancelPrinting = True
'                Exit Function
'
'    Case 3
'        If cRptLL Is Nothing Then Exit Function
'
'            With dlgCommonDialog
'                .DialogTitle = "Print Crystal Report "
'                .CancelError = True
'                 'For ii = 1 To UBound(MemberNoArray)
'                    .Flags = cdlPDNoSelection + cdlPDNoPageNums
'                    'If ActiveForm.rtfText.SelLength = 0 Then
'                     '   .Flags = .Flags + cdlPDAllPages
'                    'Else
'                     '   .Flags = .Flags + cdlPDSelection
'                    'End If
'
  '              .ShowPrinter
  '
  '              If Err <> MSComDlg.cdlCancel Then
  '                  cRptLL.PrintReport
  '              End If
  '          ' Next ii
  '          End With
  '          PrintFunction = True
  '          Exit Function
'cancelPrint3:
'                PrintFunction = False
'                CancelPrinting = True
'                Exit Function''

'End Select
    
'End Function

Private Sub FormTabView()
Select Case SSTab1.Tab
    Case 0
        SSTab1.TabEnabled(0) = True
        cboLLCriteriaSel.ListIndex = 0
        cboLLPrintOpt.ListIndex = 0
        cboLLPrintSeq.ListIndex = 0
        cboLeafletVer.ListIndex = 0
        cboLLPayor.Enabled = False
    Case 1
        SSTab1.TabEnabled(1) = True
        cboEXPCriteriaSel.ListIndex = 0
        cboExpPrintOpt.ListIndex = 0
        cboCardSequence.ListIndex = 0
    Case 2
        SSTab1.TabEnabled(2) = True
        cboEXPCriteriaSel2.ListIndex = 0
    Case 3
        SSTab1.TabEnabled(3) = True
        cboCriteriaCardSent.ListIndex = 0
    End Select
    lstMemberListing.Visible = False
    Label16.Visible = False
    cmdPrint.Visible = False
End Sub
Private Sub LeafLetPrint()
'On Error Resume Next
Dim strsql As String
Dim MBM As New ADODB.Recordset
Dim ii As Integer
Dim MemberNo As String
Dim tmpPrintedBDate, tmpPrintedBatchNo, tmpPrintedPaycode, tempPrintedHLTcode, tmpPrintedMemberNo As String
Dim Cnt As Integer
Dim PrintM As String
 
 bPrintPass = False
 CancelPrinting = False
 Screen.MousePointer = vbHourglass
 Frame1.Enabled = False
    Cnt = 0 ' Sihat Malaysia
    'If PrintMode = "ALL" Then
    If cboLLCriteriaSel.ListIndex = 0 Then
        If Len(txtLLMemberNo) = "8" Then
            strsql = "Select MBMNumber from MBMCrossReference where substring(MBMNumber,1,8) = '" & Trim(txtLLMemberNo) & "' And substring(MBMNumber,10,2) = '01' "
            MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        ElseIf Len(txtLLMemberNo) = "9" Then
            strsql = "Select MBMNumber from MBMCrossReference where substring(MBMNumber,1,9) = '" & Trim(txtLLMemberNo) & "' And substring(MBMNumber,11,2) = '01' "
            MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        ElseIf Len(Trim(txtLLMemberNo)) = 11 And Mid(Trim(txtLLMemberNo), 9, 1) <> "*" And ChkStr(Mid(Trim(txtLLMemberNo), 1, 2)) <> "DC" Then
            strsql = "Select MBMNumber from MBMCrossReference where substring(MBMNumber,1,11) = '" & Trim(txtLLMemberNo) & "' And substring(MBMNumber,13,2) = '01' "
            MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        Else
            strsql = "Select MBMNumber from MBMCrossReference where MBMNumber = '" & Trim(txtLLMemberNo) & "'"
            MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        End If
    ElseIf cboLLCriteriaSel.ListIndex = 1 Then
        strsql = "SELECT MBMNumber FROM MBM "
        If Len(txtLLMemberNoFr) = "8" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,8)  >= '" & Trim(txtLLMemberNoFr) & "'"
            strsql = strsql & " AND substring(MBMNumber,1,8) <= '" & Trim(txtLLMemberNoTo) & "'"
            strsql = strsql & " AND substring(MBMNumber,10,2)  = '01'"
            strsql = strsql & " Order By MBMDepartment, MBMNumber"
        ElseIf Len(txtLLMemberNoFr) = "9" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,9)  >= '" & Trim(txtLLMemberNoFr) & "'"
            strsql = strsql & " AND substring(MBMNumber,1,9) <= '" & Trim(txtLLMemberNoTo) & "'"
            strsql = strsql & " AND substring(MBMNumber,11,2)  = '01'"
            strsql = strsql & " Order By MBMDepartment, MBMNumber"
        ElseIf Len(Trim(txtLLMemberNo)) = 11 And Mid(Trim(txtLLMemberNo), 9, 1) <> "*" And ChkStr(Mid(Trim(txtLLMemberNo), 1, 2)) <> "DC" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,11)  >= '" & Trim(txtLLMemberNoFr) & "'"
            strsql = strsql & " AND substring(MBMNumber,1,11) <= '" & Trim(txtLLMemberNoTo) & "'"
            strsql = strsql & " AND substring(MBMNumber,13,2)  = '01'"
            strsql = strsql & " Order By MBMDepartment, MBMNumber"
        Else
            strsql = strsql & " WHERE MBMNumber >= '" & Trim(txtLLMemberNoFr) & "'"
            strsql = strsql & " AND MBMNumber <= '" & Trim(txtLLMemberNoTo) & "'"
            strsql = strsql & " Order By MBMNumber"
        End If
        
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    ElseIf cboLLCriteriaSel.ListIndex = 2 Then
        strsql = ""
        sDate = CStr(Month(txtLLBordxDate)) & "/" & CStr(Day(txtLLBordxDate)) & "/" & CStr(Year(txtLLBordxDate))
        strsql = "SELECT MBMNumber FROM View_MBMCrossReference "
        strsql = strsql & " WHERE MBMBordxDate = '" & sDate & "' "
        strsql = strsql & " AND MBMBatchNo = '" & Trim(txtLLBatchNo) & "'"
        
        If Len(cboLLPayor) Then
            strsql = strsql & " AND PAYCode = '" & Mid(cboLLPayor, 1, 2) & "'"
        End If
        
        If Len(CmbPlncode) Then
            strsql = strsql & " AND PLNCode='" & Mid(CmbPlncode, 1, 4) & "'"
        End If
        If Len(cmbCMCName) Then
            strsql = strsql & " AND Subsidiary like '%" & cmbCMCName & "%'"
        End If
        If cboLLPrintSeq.ListIndex = 0 Then
            strsql = strsql & " Order By MBMNumber "
        ElseIf cboLLPrintSeq.ListIndex = 1 Then
            strsql = strsql & " Order By MBMPolicyNo "
        ElseIf cboLLPrintSeq.ListIndex = 2 Then
            strsql = strsql & " Order By MBMName "
        ElseIf cboLLPrintSeq.ListIndex = 3 Then
            strsql = strsql & " Order By MBMDepartment, MBMNumber"
        ElseIf cboLLPrintSeq.ListIndex = 4 Then
            strsql = strsql & " Order By MBMDepartment, MBMName"
        ElseIf cboLLPrintSeq.ListIndex = 5 Then
            strsql = strsql & " Order By MBMDepartment, MBMEmployeeNo"
        ElseIf cboLLPrintSeq.ListIndex = 6 Then
            strsql = strsql & " Order By MBMEmployeeNo"
        End If
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    ElseIf cboLLCriteriaSel.ListIndex = 3 Then
        sDate = CStr(Month(txtLLBordxDateFr)) & "/" & CStr(Day(txtLLBordxDateFr)) & "/" & CStr(Year(txtLLBordxDateFr))
        SDateTo = CStr(Month(txtLLBordxDateTo)) & "/" & CStr(Day(txtLLBordxDateTo)) & "/" & CStr(Year(txtLLBordxDateTo))
        strsql = "SELECT MBMNumber FROM View_MBMCrossReference "
        strsql = strsql & " WHERE MBMBordxDate >= '" & sDate & "'"
        strsql = strsql & " AND MBMBordxDate <= '" & Trim(SDateTo) & "'"
        
        If Len(cboLLPayor) Then
            strsql = strsql & " AND PAYCode = '" & Mid(cboLLPayor, 1, 2) & "'"
        End If
        If Len(CmbPlncode) Then
            strsql = strsql & " AND PLNCode='" & Mid(CmbPlncode, 1, 4) & "'"
        End If
        
        If Len(cmbCMCName) Then
            strsql = strsql & " AND Subsidiary like '%" & (cmbCMCName) & "%'"
        End If
        
        If cboLLPrintSeq.ListIndex = 0 Then
            strsql = strsql & " Order By MBMNumber "
        ElseIf cboLLPrintSeq.ListIndex = 1 Then
            strsql = strsql & " Order By MBMPolicyNo "
        ElseIf cboLLPrintSeq.ListIndex = 2 Then
            strsql = strsql & " Order By MBMName "
        ElseIf cboLLPrintSeq.ListIndex = 3 Then
            strsql = strsql & " Order By MBMDepartment, MBMNumber"
        ElseIf cboLLPrintSeq.ListIndex = 4 Then
            strsql = strsql & " Order By MBMDepartment, MBMName"
        ElseIf cboLLPrintSeq.ListIndex = 5 Then
            strsql = strsql & " Order By MBMDepartment, MBMEmployeeNo"
        ElseIf cboLLPrintSeq.ListIndex = 6 Then
            strsql = strsql & " Order By MBMEmployeeNo"
        End If
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    End If
    If MBM.recordCount = 0 Then
        MsgBox "Record not found !", vbInformation + vbOKOnly
        Screen.MousePointer = Default
        Exit Sub
    End If
    
    ' Sihat Malaysia
    If cboLeafletVer.ListIndex = 0 Then
        Do Until MBM.EOF
            tmpMemberNo = MBM!MBMNumber
            Call InsertLealet
            Call MakeWordDoc(tmpMemberNo)
        MBM.MoveNext
        Loop

        'Group Company
        ElseIf cboLeafletVer.ListIndex = 1 Then
            Do Until MBM.EOF
                tmpMemberNo = MBM!MBMNumber
                Call InsertLealet
                Call MakeWordDoc(tmpMemberNo)
            MBM.MoveNext
            Loop
                    
             ' Dept and Employee No
           ElseIf cboLeafletVer.ListIndex = 2 Then
                    Do Until MBM.EOF
                        tmpMemberNo = MBM!MBMNumber
                        Call InsertLealet
                        Call MakeWordDoc(tmpMemberNo)
                    MBM.MoveNext
                    Loop
             
            ElseIf cboLeafletVer.ListIndex = 3 Then
                    Do Until MBM.EOF
                        tmpMemberNo = MBM!MBMNumber
                        Call InsertLealet
                        Call MakeWordDoc(tmpMemberNo)
                    MBM.MoveNext
                    Loop
             'Branch
                ElseIf cboLeafletVer.ListIndex = 4 Then
                    Do Until MBM.EOF
                        tmpMemberNo = MBM!MBMNumber
                        Call InsertLealet
                        Call MakeWordDoc(tmpMemberNo)
                    MBM.MoveNext
                    Loop
                
             'Without Address
                ElseIf cboLeafletVer.ListIndex = 5 Then
                    Do Until MBM.EOF
                        tmpMemberNo = MBM!MBMNumber
                        Call InsertLealet
                        Call MakeWordDoc(tmpMemberNo)
                    MBM.MoveNext
                    Loop
               ' End If
                    'By IC
                ElseIf cboLeafletVer.ListIndex = 6 Then
                    Do Until MBM.EOF
                        tmpMemberNo = MBM!MBMNumber
                        Call InsertLealet
                        Call MakeWordDoc(tmpMemberNo)
                    MBM.MoveNext
                    Loop
                    End If
    MBM.Close
    Set MBM = Nothing
    Frame1.Enabled = True
    Screen.MousePointer = Default
End Sub
'New sub for Print
Private Sub MakeWordDoc(MBMNumber As String)
'Get member details
        Dim PRSet As New ADODB.Recordset
        If OPTNorLeaflet.Value = True Then
            sqlbor = "Select * from VIEW_Leaflet where MBMNumber = '" & Trim(MBMNumber) & "'"
            PRSet.Open sqlbor, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        ElseIf OPTSpeLeaflet.Value = True Then
            sqlbor = "Select * from View_Leaflet_Clinical where MBMNumber = '" & Trim(MBMNumber) & "'"
            PRSet.Open sqlbor, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        End If
        'Template Path
        Dim fname As String
         If Mid(cboPrintMode, 1, 2) = "RI" Then
            fname = crdPath & "Leflet\RITemplate.dot"
         ElseIf Mid(cboPrintMode, 1, 2) = "R8" Then
            fname = crdPath & "Leflet\R8Template.dot"
         ElseIf Mid(cboPrintMode, 1, 2) = "HP" Then
            fname = crdPath & "Leflet\PrintTemplate.dot"
         End If
        
        Dim word_app As Word.Application
        Dim word_doc As Word.Document
        Dim objSelection
        
        Dim SuppCount As Integer
        Dim strArray(20) As String
        SuppCount = 0
        Dim isFirst As Boolean
        isFirst = True
        
        Dim MBMNo As String
       
      
        Do While Not PRSet.EOF
         MBMNo = PRSet!MBMNumber
            
        If InStr(1, MBMNo, "*") Then
            Dim sArr() As String
            sArr = Split(MBMNo, "*")
            MBMNo = sArr(0)
        End If
        
            If isFirst Then
            strArray(6) = PRSet!MBMPolicyNo & " " & IIf(Len(PRSet!MBMAgentCode), (", " & PRSet!MBMAgentCode), "")
            strArray(7) = PRSet!MBMName
            strArray(8) = PRSet!MBMAddress1
            strArray(9) = PRSet!MBMAddress2
            strArray(10) = IIf(Len(PRSet!MBMAddress3), (PRSet!MBMAddress3 & ", "), "") & PRSet!MBMPostCode & ", " & PRSet!CTYCode
            strArray(11) = PRSet!STACode
            strArray(12) = "MembershipNo: " & MBMNo
            strArray(13) = PRSet!BRNCode
            isFirst = False
            End If
            
            If SuppCount < 6 Then
                strArray(SuppCount) = strArray(SuppCount) + "                                                                                             " & PRSet!MBMCName
            ElseIf SuppCount = 6 And SuppCount < 14 Then
                strArray(SuppCount) = strArray(SuppCount) + "                                                                        " & PRSet!MBMCName
            Else
                strArray(SuppCount) = strArray(SuppCount) + "                                                                                            " & PRSet!MBMCName
            End If
            
            SuppCount = SuppCount + 1
            PRSet.MoveNext
        Loop
        
        Dim arrcount As Integer
        arrcount = 0
        
        Dim strText As String
        
        For arrcount = 0 To 20
            'If Len(StrArray(arrcount)) Then
                strText = strText & strArray(arrcount) & Chr(11)
            'Else
            '    strText = strText & StrArray(arrcount) & Chr(11)
            'End If
        Next
        
         'Print details with word
            Set word_app = New Word.Application
'            Set word_doc = word_app.Documents.Add(DocumentType:=wdNewBlankDocument)
            Set word_doc = word_app.Documents.Open(fname)
            word_app.Selection.TypeParagraph
            word_app.Selection.TypeParagraph
            word_app.Selection.TypeText strText
            Set objSelection = word_app.Selection
            objSelection.WholeStory
            objSelection.Font.Name = "Times New Roman"
            objSelection.Font.Size = 10
            objSelection.Font.Bold = False
            
            word_doc.SaveAs filename:="C:\Worddoc.doc"
            word_doc.PrintOut Range:=wdPrintCurrentPage
            Sleep 1000
            word_doc.Close False
            word_app.Quit False
End Sub



Sub ComboListing()
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
    
    Do While Not PAY.EOF
        With PAY
            cboLLPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            cboExpPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            CboCLNPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            cboCardSentPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    
    PAY.Close
    Set PAY = Nothing
    
    cboPrintMode.AddItem "RI - Ricoh"
    cboPrintMode.AddItem "HP - Hewlett Packard"
    cboPrintMode.AddItem "R8 - Ricoh Pro 8100"
    cboPrintMode.Text = "RI - Ricoh"
End Sub

Private Sub cboEXPCriteriaSel_Click()
Select Case cboEXPCriteriaSel.ListIndex
    Case 0
        'lbl38.Visible = True
        lbl39(0).Visible = True
        txtExpMemberNo.Visible = True
        'lbl40.Visible = False
        lbl41.Visible = False
        lbl42.Visible = False
        'lbl43.Visible = False
        lbl44.Visible = False
        lbl45.Visible = False
        'lbl46.Visible = False
        lbl47.Visible = False
        lbl48.Visible = False
        txtExpMemberNoFr.Visible = False
        txtExpMemberNoTo.Visible = False
        txtExpBordxDate.Visible = False
        txtExpBatchNo.Visible = False
        txtExpBordxDateFr.Visible = False
        txtExpBordxDateTo.Visible = False
        cboExpPayor.Enabled = False
        txtExpMemberNoFr = ""
        txtExpMemberNoTo = ""
        txtExpMemberNo = ""
        txtExpBatchNo = ""
        cboExpPayor = ""
        txtExpBordxDate = "__/__/____"
        txtExpBordxDateFr = "__/__/____"
        txtExpBordxDateTo = "__/__/____"
        cboExpPrintOpt.ListIndex = 0
        cboExpPrintOpt.Enabled = True
        'lvwCRD.listitems.Clear
        'cmdCRDPrint.Enabled = False
        txtEXPTotalPrincipal = 0
        txtEXPTotalSupplementary1 = 0
        cmdCRDPrint.Caption = "Export"
        Label16.Visible = False
        lstMemberListing.Visible = False
        cmdPrint.Visible = False
    Case 1
        'lbl38.Visible = False
        lbl39(0).Visible = False
        'lbl40.Visible = True
        lbl41.Visible = True
        lbl42.Visible = True
        'lbl43.Visible = False
        txtExpMemberNoFr.Visible = True
        txtExpMemberNoTo.Visible = True
        lbl44.Visible = False
        lbl45.Visible = False
        'lbl46.Visible = False
        lbl47.Visible = False
        lbl48.Visible = False
        txtExpMemberNo.Visible = False
        txtExpBordxDate.Visible = False
        txtExpBatchNo.Visible = False
        txtExpBordxDateFr.Visible = False
        txtExpBordxDateTo.Visible = False
        cboExpPayor.Enabled = False
        txtExpMemberNoTo = ""
        txtExpMemberNo = ""
        txtExpBatchNo = ""
        cboExpPayor = ""
        txtExpBordxDate = "__/__/____"
        txtExpBordxDateFr = "__/__/____"
        txtExpBordxDateTo = "__/__/____"
        cboExpPrintOpt.ListIndex = 0
        cboExpPrintOpt.Enabled = True
        'lvwCRD.listitems.Clear
        'cmdCRDPrint.Enabled = False
        txtEXPTotalPrincipal = 0
        txtEXPTotalSupplementary1 = 0
        cmdCRDPrint.Caption = "Export"
        Label16.Visible = False
        lstMemberListing.Visible = False
        cmdPrint.Visible = False
    Case 2
        'lbl38.Visible = False
        lbl39(0).Visible = False
        'lbl40.Visible = False
        lbl41.Visible = False
        lbl42.Visible = False
        'lbl43.Visible = True
        lbl44.Visible = True
        lbl45.Visible = True
        txtExpBordxDate.Visible = True
        txtExpBatchNo.Visible = True
        'lbl46.Visible = False
        lbl47.Visible = False
        lbl48.Visible = False
        txtExpMemberNo.Visible = False
        txtExpMemberNoFr.Visible = False
        txtExpMemberNoTo.Visible = False
        txtExpBordxDateFr.Visible = False
        txtExpBordxDateTo.Visible = False
        cboExpPayor.Enabled = True
        txtExpMemberNoTo = ""
        txtExpMemberNo = ""
        txtExpBatchNo = ""
        cboExpPayor = ""
        txtExpBordxDate = "__/__/____"
        txtExpBordxDateFr = "__/__/____"
        txtExpBordxDateTo = "__/__/____"
        cboExpPrintOpt.ListIndex = 0
        cboExpPrintOpt.Enabled = True
        'lvwCRD.listitems.Clear
        'cmdCRDPrint.Enabled = False
        txtEXPTotalPrincipal = 0
        txtEXPTotalSupplementary1 = 0
        cmdCRDPrint.Caption = "Export"
        Label16.Visible = False
        lstMemberListing.Visible = False
        cmdPrint.Visible = False
    Case 3
            'lbl38.Visible = True
        lbl39(0).Visible = True
        txtExpMemberNo.Visible = True
        txtExpMemberNo.SetFocus
        'lbl40.Visible = False
        lbl41.Visible = False
        lbl42.Visible = False
        'lbl43.Visible = False
        lbl44.Visible = False
        lbl45.Visible = False
        'lbl46.Visible = False
        lbl47.Visible = False
        lbl48.Visible = False
        txtExpMemberNoFr.Visible = False
        txtExpMemberNoTo.Visible = False
        txtExpBordxDate.Visible = False
        txtExpBatchNo.Visible = False
        txtExpBordxDateFr.Visible = False
        txtExpBordxDateTo.Visible = False
        cboExpPayor.Enabled = False
        txtExpMemberNoFr = ""
        txtExpMemberNoTo = ""
        txtExpMemberNo = ""
        txtExpBatchNo = ""
        cboExpPayor = ""
        txtExpBordxDate = "__/__/____"
        txtExpBordxDateFr = "__/__/____"
        txtExpBordxDateTo = "__/__/____"
        cboExpPrintOpt.ListIndex = 0
        cboExpPrintOpt.Enabled = True
        txtEXPTotalPrincipal = 0
        txtEXPTotalSupplementary1 = 0
        cmdCRDPrint.Caption = "Search"
        cmdCRDPrint.Default = True
        Label16.Visible = True
        lstMemberListing.Visible = True
        cmdPrint.Visible = True
        cmdCRDPrint.Enabled = True
End Select
End Sub

Private Sub cboEXPCriteriaSel_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboEXPCriteriaSel.Text, cboEXPCriteriaSel.SelStart) + Chr(KeyAscii) + Mid(cboEXPCriteriaSel.Text, cboEXPCriteriaSel.SelStart + cboEXPCriteriaSel.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboEXPCriteriaSel.ListCount - 1
        
        If NewText = LCase(Left(cboEXPCriteriaSel.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboEXPCriteriaSel.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboEXPCriteriaSel.Text = ValidValue
               cboEXPCriteriaSel.SelStart = 0
               cboEXPCriteriaSel.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboExpPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboExpPayor.Text, cboExpPayor.SelStart) + Chr(KeyAscii) + Mid(cboExpPayor.Text, cboExpPayor.SelStart + cboExpPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboExpPayor.ListCount - 1
 
        If NewText = LCase(Left(cboExpPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboExpPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboExpPayor.Text = ValidValue
                cboExpPayor.SelStart = 0
                cboExpPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboExpPrintOpt_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboExpPrintOpt.Text, cboExpPrintOpt.SelStart) + Chr(KeyAscii) + Mid(cboExpPrintOpt.Text, cboExpPrintOpt.SelStart + cboExpPrintOpt.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboExpPrintOpt.ListCount - 1
        
        If NewText = LCase(Left(cboExpPrintOpt.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboExpPrintOpt.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboExpPrintOpt.Text = ValidValue
               cboExpPrintOpt.SelStart = 0
               cboExpPrintOpt.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboLeafletVer_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboLeafletVer.Text, cboLeafletVer.SelStart) + Chr(KeyAscii) + Mid(cboLeafletVer.Text, cboLeafletVer.SelStart + cboLeafletVer.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboLeafletVer.ListCount - 1
        
        If NewText = LCase(Left(cboLeafletVer.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboLeafletVer.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboLeafletVer.Text = ValidValue
               cboLeafletVer.SelStart = 0
               cboLeafletVer.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboLLCriteriaSel_Click()
Select Case cboLLCriteriaSel.ListIndex
    Case 0
        'lbl1.Visible = True
        lbl2.Visible = True
        txtLLMemberNo.Visible = True
        'lbl3.Visible = False
        lbl4.Visible = False
        lbl5.Visible = False
        'lbl6.Visible = False
        lbl7.Visible = False
        lbl8.Visible = False
        'lbl9.Visible = False
        lbl10.Visible = False
        lbl11.Visible = False
        txtLLMemberNoFr.Visible = False
        txtLLMemberNoTo.Visible = False
        txtLLBordxDate.Visible = False
        txtLLBatchNo.Visible = False
        txtLLBordxDateFr.Visible = False
        txtLLBordxDateTo.Visible = False
        cboLLPayor.Enabled = False
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
        'lvwLeaflet.listitems.Clear
        'cmdLLPrint.Enabled = False
        txtLLTotalPrincipal = 0
        txtLLTotalSupplementary1 = 0
    
    Case 1
        'lbl1.Visible = False
        lbl2.Visible = False
        'lbl3.Visible = True
        lbl4.Visible = True
        lbl5.Visible = True
        txtLLMemberNoFr.Visible = True
        txtLLMemberNoTo.Visible = True
        'lbl6.Visible = False
        lbl7.Visible = False
        lbl8.Visible = False
        'lbl9.Visible = False
        lbl10.Visible = False
        lbl11.Visible = False
        txtLLMemberNo.Visible = False
        txtLLBordxDate.Visible = False
        txtLLBatchNo.Visible = False
        txtLLBordxDateFr.Visible = False
        txtLLBordxDateTo.Visible = False
        cboLLPayor.Enabled = False
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
        'lvwLeaflet.listitems.Clear
        'cmdLLPrint.Enabled = False
        txtLLTotalPrincipal = 0
        txtLLTotalSupplementary1 = 0
        
    Case 2
        'lbl1.Visible = False
        lbl2.Visible = False
        'lbl3.Visible = False
        lbl4.Visible = False
        lbl5.Visible = False
        'lbl6.Visible = True
        lbl7.Visible = True
        lbl8.Visible = True
        txtLLBordxDate.Visible = True
        txtLLBatchNo.Visible = True
        'lbl9.Visible = False
        lbl10.Visible = False
        lbl11.Visible = False
        txtLLMemberNo.Visible = False
        txtLLMemberNoFr.Visible = False
        txtLLMemberNoTo.Visible = False
        txtLLBordxDateFr.Visible = False
        txtLLBordxDateTo.Visible = False
        cboLLPayor.Enabled = True
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
        'lvwLeaflet.listitems.Clear
        'cmdLLPrint.Enabled = False
        txtLLTotalPrincipal = 0
        txtLLTotalSupplementary1 = 0
    
    Case 3
        'lbl1.Visible = False
        lbl2.Visible = False
        'lbl3.Visible = False
        lbl4.Visible = False
        lbl5.Visible = False
        'lbl6.Visible = False
        lbl7.Visible = False
        lbl8.Visible = False
        'lbl9.Visible = True
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
        'lvwLeaflet.listitems.Clear
        'cmdLLPrint.Enabled = False
        txtLLTotalPrincipal = 0
        txtLLTotalSupplementary1 = 0
End Select
        
End Sub

Private Sub cboLLCriteriaSel_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboLLCriteriaSel.Text, cboLLCriteriaSel.SelStart) + Chr(KeyAscii) + Mid(cboLLCriteriaSel.Text, cboLLCriteriaSel.SelStart + cboLLCriteriaSel.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboLLCriteriaSel.ListCount - 1
        
        If NewText = LCase(Left(cboLLCriteriaSel.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboLLCriteriaSel.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboLLCriteriaSel.Text = ValidValue
               cboLLCriteriaSel.SelStart = 0
               cboLLCriteriaSel.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboLLPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboLLPayor.Text, cboLLPayor.SelStart) + Chr(KeyAscii) + Mid(cboLLPayor.Text, cboLLPayor.SelStart + cboLLPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboLLPayor.ListCount - 1
 
        If NewText = LCase(Left(cboLLPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboLLPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboLLPayor.Text = ValidValue
                cboLLPayor.SelStart = 0
                cboLLPayor.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboLLPrintOpt_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboLLPrintOpt.Text, cboLLPrintOpt.SelStart) + Chr(KeyAscii) + Mid(cboLLPrintOpt.Text, cboLLPrintOpt.SelStart + cboLLPrintOpt.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboLLPrintOpt.ListCount - 1
        
        If NewText = LCase(Left(cboLLPrintOpt.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboLLPrintOpt.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboLLPrintOpt.Text = ValidValue
               cboLLPrintOpt.SelStart = 0
               cboLLPrintOpt.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboLLPrintSeq_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboLLPrintSeq.Text, cboLLPrintSeq.SelStart) + Chr(KeyAscii) + Mid(cboLLPrintSeq.Text, cboLLPrintSeq.SelStart + cboLLPrintSeq.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboLLPrintSeq.ListCount - 1
        
        If NewText = LCase(Left(cboLLPrintSeq.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboLLPrintSeq.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboLLPrintSeq.Text = ValidValue
               cboLLPrintSeq.SelStart = 0
               cboLLPrintSeq.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cmdCRDPrint_Click()
Dim RecordExit
    cmdCRDPrint.Enabled = False
    If cboEXPCriteriaSel.ListIndex = 0 And txtExpMemberNo = "" Then
        MsgBox "Please Enter Member No !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If cboEXPCriteriaSel.ListIndex = 1 And (txtExpMemberNoFr = "" Or txtExpMemberNoTo = "") Then
            MsgBox "Please Enter Members No range !", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        Else
            If cboEXPCriteriaSel.ListIndex = 2 And (txtExpBordxDate = "__/__/____" Or txtExpBatchNo = "" Or cboExpPayor = "") Then
                MsgBox "Please Enter Bordereaux Date, Batch No and Payor !", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            
            End If
        End If
    End If

    RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        Screen.MousePointer = vbHourglass
            Call EXPPrincipalListing
            Call EXPPrincipalAck
            Call EXPPrincipalEnvelope
        Screen.MousePointer = Default
    End If
    cmdCRDPrint.Enabled = True
End Sub

Private Sub cmdExpClear_Click()
    txtExpMemberNoFr = ""
    txtExpMemberNoTo = ""
    txtExpMemberNo = ""
    txtExpBatchNo = ""
    cboExpPayor = ""
    txtExpBordxDate = "__/__/____"
    txtExpBordxDateFr = "__/__/____"
    txtExpBordxDateTo = "__/__/____"
    cboExpPrintOpt.ListIndex = 0
    cboExpPrintOpt.Enabled = True
    'lvwCRD.listitems.Clear
    'cmdCRDPrint.Enabled = False
    txtEXPTotalPrincipal = 0
    txtEXPTotalSupplementary1 = 0
    lstMemberListing.ListItems.Clear
End Sub

Private Sub cmdExpClose_Click()
    Unload Me
End Sub

Private Sub cmdExpSearch_Click()
    cmdEXPSearch.Enabled = False
    If (cboEXPCriteriaSel.ListIndex = 0 Or cboEXPCriteriaSel.ListIndex = 3) And txtExpMemberNo = "" Then
        MsgBox "Please Enter Member No !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If cboEXPCriteriaSel.ListIndex = 1 And txtExpMemberNoFr = "" And txtExpMemberNoTo = "" Then
            MsgBox "Please Enter Members No range !", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        Else
            If cboEXPCriteriaSel.ListIndex = 2 And txtExpBordxDate = "__/__/____" And txtExpBatchNo = "" And cboExpPayor = "" Then
                MsgBox "Please Enter Bordereaux Date, Batch No and Payor !", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            
            End If
        End If
    End If
    
    If cboExpPrintOpt = "" Then
        MsgBox "Please Enter Print Option !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
        Call EXPPrincipalListing
    Screen.MousePointer = Default
    cmdEXPSearch.Enabled = True
    
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
    'lvwLeaflet.listitems.Clear
    'cmdLLPrint.Enabled = False
    txtLLTotalPrincipal = 0
    txtLLTotalSupplementary1 = 0
End Sub

Private Sub cmdLLClose_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdLLPrint_Click()
Dim RecordExit
Dim tmpPath As String

    cmdLLPrint.Enabled = False
    If cboLLCriteriaSel.ListIndex = 0 And txtLLMemberNo = "" Then
        MsgBox "Please Enter Member No !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If cboLLCriteriaSel.ListIndex = 1 And (txtLLMemberNoFr = "" Or txtLLMemberNoTo = "") Then
            MsgBox "Please Enter Members No range !", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        Else
            'If cboLLCriteriaSel.ListIndex = 2 And (txtLLBordxDate = "__/__/____" Or txtLLBatchNo = "" Or cboLLPayor = "") Then
            If cboLLCriteriaSel.ListIndex = 2 And (txtLLBordxDate = "__/__/____" Or txtLLBatchNo = "") Then
                MsgBox "Please Enter Bordereaux Date, Batch No and Payor !", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            Else
                If cboLLCriteriaSel.ListIndex = 3 And (cboLLPayor = "" Or txtLLBordxDateFr = "__/__/____" Or txtLLBordxDateTo = "__/__/____") Then
                    MsgBox "Please Enter Bordereaux Range and Payor!", vbOKOnly + vbCritical, DialogBoxTitle
                    Exit Sub
                End If
            End If
        End If
    End If
    If cboPrintMode = "" Then
        MsgBox "Please select printer!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdLLPrint.Enabled = True
        Exit Sub
    End If
    RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        tempPath = ""
        tempPath = "\\PLUTO\HISSaturn\Conn\mediconnectPluto.UDL"
        medixmasterconnPluto.Open "File Name=" & tempPath
        Call LeafLetPrint
        medixmasterconnPluto.Close
        Set medixmasterconnPluto = Nothing
    Frame1.Enabled = True
    End If
    cmdLLPrint.Enabled = True
End Sub

Private Sub cmdLLSearch_Click()
    cmdLLSearch.Enabled = False
    If cboLLCriteriaSel.ListIndex = 0 And txtLLMemberNo = "" Then
        MsgBox "Please Enter Member No !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If cboLLCriteriaSel.ListIndex = 1 And txtLLMemberNoFr = "" And txtLLMemberNoTo = "" Then
            MsgBox "Please Enter Members No range !", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        Else
            If cboLLCriteriaSel.ListIndex = 2 And txtLLBordxDate = "__/__/____" And txtLLBatchNo = "" And cboLLPayor = "" Then
                MsgBox "Please Enter Bordereaux Date, Batch No and Payor !", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            Else
                If cboLLCriteriaSel.ListIndex = 3 And cboLLPayor = "" And txtLLBordxDateFr = "__/__/____" And txtLLBordxDateTo = "__/__/____" Then
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
    cmdLLSearch.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim X As Boolean
    Call EnterToTab(KeyAscii)
End Sub

Private Sub Form_Load()

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
    Call FormTabView
    X = CreatePath("c:\cardPrinting")
    If X = False Then
        MsgBox "Error Creating Folder c:\CardPrinting !" & vbNewLine & _
                " Please create that folder manually to store the output of leaflet in sofe copy!"
    End If
    
    
    'Dim PAY As New ADODB.Recordset
    'Dim StrQry As String
    'Dim GrpCompany() As String
    'GrpCompany = Split(cboLLPayor, "-")
    'StrQry = "  select  Subsidiary, CMCName, MBMGroupCompany   from MBMOthers  Where LTrim(CMCName) Is Not Null And LTrim(CMCName) <>  N''   Group by Subsidiary, CMCName, MBMGroupCompany  Order by Subsidiary "
    'PAY.Open StrQry, medixmasterconn, adOpenKeyset, adLockOptimistic
    'cmbCMCName.Clear
    '    Do While Not PAY.EOF
    '        If Len(PAY!Subsidiary) Then
    '            cmbCMCName.AddItem (PAY!Subsidiary)
    '        End If
    '    PAY.MoveNext
    'Loop
    'PAY.Close
    'Set PAY = Nothing
    
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    Call FormTabView
End Sub

Private Sub EXPPrincipalListing()
Dim strsql As String
Dim strsqls As String
Dim sql As String
Dim MBMOther As New ADODB.Recordset
Dim MBM As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim ListMaster As ListItem
Dim isPrinted As Boolean
Dim rsSearch As New ADODB.Recordset
Dim ColumnItem As ListItem
Dim itemToPrint As Integer
Dim CoverQty As Integer
Dim TotalPrincipal As Integer
Dim strsqlCOV As String
Dim MBMCov As New ADODB.Recordset
Dim MBMCPLNCode  As String
Dim tempInsPLN As String
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim tmpSuppExpDate As String
Dim Effdate As String
Dim tmpSuppEffDate As String
Dim tmpMBMPrevMemNo As String
Dim tmpMBMNumber As String
Dim tmpEncoding As String
Dim tmpEncoding1 As String

TotalSupplementary = 0
TotalPrincipal = 0
MBMCPLNCode = ""
tmpEncoding = ""

Select Case cboEXPCriteriaSel.ListIndex
    Case 0 'Card Export Criteria Selection for Single Member
        'lvwCRD.listitems.Clear
        If Len(txtExpMemberNo) = "8" Then
        'strsql = ""
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate from View_MBM_CardExport where substring(MBMNumber,1,8) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,10,2) = '01'  "
        ElseIf Len(txtExpMemberNo) = "9" Then
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate from View_MBM_CardExport where substring(MBMNumber,1,9) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,11,2) = '01'  "
        ElseIf Len(txtExpMemberNo) = "11" Then
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate from View_MBM_CardExport where substring(MBMNumber,1,11) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,13,2) = '01'  "
            
        Else
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate from View_MBM_CardExport where MBMNumber = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D'  "
        End If
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Case 1 'Card Export Criteria Selection for Member's No Range
        strsql = "SELECT INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate from View_MBM_CardExport "
        
        If Len(txtExpMemberNoFr) = "8" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,8) >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND substring(MBMNumber,1,8) <= '" & txtExpMemberNoTo & "'"
            strsql = strsql & " AND substring(MBMNumber,10,2) <= '01'"
        ElseIf Len(txtExpMemberNoFr) = "9" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,9) >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND substring(MBMNumber,1,9) <= '" & txtExpMemberNoTo & "'"
            strsql = strsql & " AND substring(MBMNumber,11,2) <= '01'"
        ElseIf Len(Trim(txtExpMemberNoFr)) = 11 And Mid(Trim(txtExpMemberNoFr), 9, 1) <> "*" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,11) >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND substring(MBMNumber,1,11) <= '" & txtExpMemberNoTo & "'"
            strsql = strsql & " AND substring(MBMNumber,13,2) <= '01'"
        Else
            strsql = strsql & " WHERE MBMNumber >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND MBMNumber <= '" & txtExpMemberNoTo & "'"
        End If
        strsql = strsql & " AND MBMDataStatus <> 'D' "
        strsql = strsql & " AND MBMStatus <> 'C' "
        
      If cboCardSequence.ListIndex = 0 Then
            strsql = strsql & " Order By MBMNumber "
        ElseIf cboCardSequence.ListIndex = 1 Then
            strsql = strsql & " Order By MBMPrevMemNo "
        ElseIf cboCardSequence.ListIndex = 2 Then
            strsql = strsql & " Order By MBMName "
        'ElseIf cboLLPrintSeq.ListIndex = 3 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMNumber"
        ElseIf cboCardSequence.ListIndex = 3 Then
            strsql = strsql & " Order By MBMDepartment, MBMName"
        'ElseIf cboLLPrintSeq.ListIndex = 5 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMEmployeeNo"
        'ElseIf cboLLPrintSeq.ListIndex = 6 Then
        '    strsql = strsql & " Order By MBMEmployeeNo"
        End If
        
        'strsql = strsql & "order by MBMNumber"
        
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Case 2
        sDate = CStr(Month(txtExpBordxDate)) & "/" & CStr(Day(txtExpBordxDate)) & "/" & CStr(Year(txtExpBordxDate))
        strsql = ""
        strsql = "SELECT INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorCode,JoinDate from View_MBM_CardExport "
        strsql = strsql & " WHERE MBMBordxDate = '" & Format(txtExpBordxDate, "mm/dd/yyyy") & "'"
        strsql = strsql & " AND MBMBatchNo = '" & txtExpBatchNo & "'"
        strsql = strsql & " AND PAYCode = '" & Mid(cboExpPayor, 1, 2) & "'"
        strsql = strsql & " AND MBMDataStatus <> 'D' "
        strsql = strsql & " AND MBMStatus <> 'C' "
        
        If cboCardSequence.ListIndex = 0 Then
            strsql = strsql & " Order By MBMNumber "
        ElseIf cboCardSequence.ListIndex = 1 Then
            strsql = strsql & " Order By MBMPrevMemNo "
        ElseIf cboCardSequence.ListIndex = 2 Then
            strsql = strsql & " Order By MBMName "
        'ElseIf cboLLPrintSeq.ListIndex = 3 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMNumber"
        ElseIf cboCardSequence.ListIndex = 3 Then
            strsql = strsql & " Order By MBMDepartment, MBMName"
        'ElseIf cboLLPrintSeq.ListIndex = 5 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMEmployeeNo"
        'ElseIf cboLLPrintSeq.ListIndex = 6 Then
        '    strsql = strsql & " Order By MBMEmployeeNo"
        End If
        'strsql = strsql & "order by MBMNumber"
        'strsql = strsql & "order by MBMNumber"
        'strsql = strsql & " Order By MBMNumber"
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
     Case 3
        Call MemberListing
   End Select
    
        If cboEXPCriteriaSel.ListIndex <> "3" Then
            If MBM.EOF = True Then
                MsgBox "Record not found !", vbInformation + vbOKOnly
                'cmdCRDPrint.Enabled = False
                Screen.MousePointer = Default
                Exit Sub
            End If
        
            Frame2.Enabled = False
            txtSVRPath = crdPath & "CardPrinting\"
            txtPath = "C:\CardPrinting\"
            If cboEXPCriteriaSel.ListIndex = 0 Then
                txtFilName = txtExpMemberNo & "-" & "Print Card"
            ElseIf cboEXPCriteriaSel.ListIndex = 1 Then
                txtFilName = txtExpMemberNoFr & "-" & "Print Card"
            Else
                txtFilName = Mid(cboExpPayor, 1, 2) & Mid(txtExpBordxDate, 1, 2) & Mid(txtExpBordxDate, 4, 2) & Mid(txtExpBordxDate, 7, 4) & "-" & "Print Card"
            End If
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set txtSVRfile = fso.CreateTextFile((Trim(txtSVRPath) & Trim(txtFilName & ".xls")), True)
            Set txtfile = fso.CreateTextFile((Trim(txtPath) & Trim(txtFilName & ".xls")), True)
            txtfile = (Trim(txtPath) & Trim(txtFilName & ".xls"))
            txtSVRfile = (Trim(txtSVRPath) & Trim(txtFilName & ".xls"))
                
            Set xlApp = CreateObject("Excel.application")
            Set xlBook = xlApp.Workbooks.Open(txtfile)
            Set xlsheet = xlBook.Sheets(1)
             xlsheet.Cells(1, 1).Value = cboExpPayor
             xlsheet.Cells(2, 1).Value = "DATE :" & Mid(txtExpBordxDate, 1, 2) & "/" & Mid(txtExpBordxDate, 4, 2) & "/" & Mid(txtExpBordxDate, 7, 4)
             xlsheet.Cells(3, 1).Value = "FILE NAME :" & txtFilName
             row = 5
             Rcounter = 0
            
        
        Do Until MBM.EOF
            strsqls = ""
            tmpEncoding = ""
            tmpMBMNumber = Split(MBM!MBMNumber, "*")(0)
            strsqls = "Select MBMNumber,MBMSecurityCode,MBMCardType,MBMClientType from MBMSCSecurity where MBMNumber = '" & Trim(tmpMBMNumber) & "'"
            MBMOther.Open strsqls, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                If MBMOther.recordCount > 0 Then
                    'If Mid(cboExpPayor, 1, 2) = "EM" Or Mid(cboExpPayor, 1, 2) = "ES" Then
                    '    tmpEncoding1 = MBMOther!MBMSecurityCode & MBMOther!MBMCardType & MBMOther!MBMClientType & "  etiqahc.com.my                />" & MBMOther!MBMNumber
                    '    tmpEncoding = tmpEncoding1 & "-00"
                    'Else
                        tmpEncoding1 = MBMOther!MBMSecurityCode & MBMOther!MBMCardType & MBMOther!MBMClientType & "  hcsb.com.my                   />" & MBMOther!MBMNumber
                        tmpEncoding = tmpEncoding1 & "-00"
                    'End If
                    '6001mmed  hcsb.com.my                   />dcga8524393-00
                End If
            MBMOther.Close
            Set MBMOther = Nothing
            
            strsqls = ""
            strsqls = "Select MBMBranch, MBMPOLSubNo1, MBMGroupCompany, MBMEmployeeNo, MBMDepartment,MBMPrevMemNo,Subsidiary,MBMAgentCode,AgentName,PLN.PLNPayorPlan,MBMOthers.MBMFirstJoinedDate AS JoinDate from MBMOthers  inner join MBM on MBM.MBMnumber = MBMOthers.MBMnumber inner join PLN on PLN.PLNCode = MBM.PLNCode where MBMOthers.MBMNumber = '" & Trim(MBM!MBMNumber) & "'"
            MBMOther.Open strsqls, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                                                
            If MBMOther.recordCount Then
                If Len(MBMOther!MBMGroupCompany) Then
                    GroupCompany = MBMOther!MBMGroupCompany
                Else
                    GroupCompany = ""
                End If
                
                If Len(MBMOther!MBMEmployeeNo) Then
                    EmployeeNo = MBMOther!MBMEmployeeNo
                Else
                    EmployeeNo = ""
                End If
                
                If Len(MBMOther!MBMDepartment) Then
                    Department = MBMOther!MBMDepartment
                Else
                    Department = ""
                End If
                
                If Len(MBMOther!MBMPolSubNo1) Then
                    POLSubNo1 = MBMOther!MBMPolSubNo1
                Else
                    POLSubNo1 = ""
                End If
                
                If Len(MBMOther!MBMBranch) Then
                    MBMBranch = MBMOther!MBMBranch
                Else
                    MBMBranch = ""
                End If
                
                If Len(MBMOther!MBMPrevMemNo) Then
                    tmpMBMPrevMemNo = MBMOther!MBMPrevMemNo
                Else
                    tmpMBMPrevMemNo = ""
                End If
                
                If Len(MBMOther!Subsidiary) Then
                    Subsidary = MBMOther!Subsidiary
                Else
                    Subsidary = ""
                End If
                
                If Len(MBMOther!MBMAgentCode) Then
                    AgentCode = MBMOther!MBMAgentCode
                Else
                    AgentCode = ""
                End If
                
                If Len(MBMOther!AgentName) Then
                    AgentName = MBMOther!AgentName
                Else
                    AgentName = ""
                End If
                
                If Len(MBMOther!PLNPayorPlan) Then
                    PLNPayorCode = MBMOther!PLNPayorPlan
                Else
                    PLNPayorCode = ""
                End If
                
                If Len(MBMOther!JoinDate) Then
                    Joindt = Format(MBMOther!JoinDate, "dd/MM/yyyy")
                Else
                    Joindt = ""
                End If
                
                
            End If
            MBMOther.Close
            Set MBMOther = Nothing
            strsqls = ""
           
               ' Display record in List View
               Cnt = 0
               Rcounter = Rcounter + 1
               row = row + 1
               xlApp.DisplayAlerts = False
                 ' Export record to excel format
                 
                    MBMName = MBM!MBMName
                    MBMNo = MBM!MBMNumber
                    MBMICNo = MBM!MBMIcBcPp
                    MBMPOLNO = MBM!MBMPolicyNo
                    If Trim(MBM!PLNCode) = "21" Or Trim(MBM!PLNCode) = "31" Or Trim(MBM!PLNCode) = "41" Then
                        PLNCode = "60"
                    ElseIf Trim(MBM!PLNCode) = "22" Or Trim(MBM!PLNCode) = "32" Or Trim(MBM!PLNCode) = "42" Then
                        PLNCode = "80"
                    ElseIf Trim(MBM!PLNCode) = "23" Or Trim(MBM!PLNCode) = "33" Or Trim(MBM!PLNCode) = "43" Then
                        PLNCode = "120"
                    ElseIf Trim(MBM!PLNCode) = "24" Or Trim(MBM!PLNCode) = "34" Or Trim(MBM!PLNCode) = "44" Then
                        PLNCode = "200"
                    ElseIf Trim(MBM!PLNCode) = "25" Or Trim(MBM!PLNCode) = "35" Or Trim(MBM!PLNCode) = "45" Then
                        PLNCode = "350"
                    ElseIf Trim(MBM!PLNCode) = "26" Or Trim(MBM!PLNCode) = "36" Or Trim(MBM!PLNCode) = "46" Then
                        PLNCode = "450"
                    Else
                        PLNCode = IIf(Len(MBM!PLNCode), MBM!PLNCode, "")
                    End If
                    Effdate = Format(MBM!MBMPayorEffDate, "dd/mm/yyyy")
                    ExpDate = Format(MBM!MBMPayorEXPDate, "dd/mm/yyyy")
                   'strsqls = "Select MBMBranch, MBMPOLSubNo1, MBMGroupCompany, MBMEmployeeNo, MBMDepartment from MBMOthers where MBMNumber = '" & Trim(MBM!MBMNumber) & "'"

                    strsqlPLN = "Select PLNCode,PLNClientPlan,PLNRBType,PLNDescription,PLNCardPrintMode from PLN where PLNCode = '" & MBM!PLNCode & "' And PAYCode = '" & MBM!PAYCode & "'"
                    PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If PLN.recordCount Then
                            If Trim(MBM!INSCode) = "G" Then
                                tempInsPLN = "INDIVIDUAL" & " " & PLN!PLNDescription
                            ElseIf Trim(MBM!INSCode) = "H" Then
                                tempInsPLN = "FAMILY" & " " & PLN!PLNDescription
                            End If
                            If PLN!PLNCardPrintMode <> "" Then
                                tmpPrintMode = PLN!PLNCardPrintMode
                            Else
                                tmpPrintMode = ""
                            End If
                        End If
                    PLN.Close
                    Set PLN = Nothing
                    
                    xlsheet.Cells(row, 1).Value = Rcounter
                    xlsheet.Cells(row, 2).Value = MBMName
                    If IsNumeric(MBMICNo) = True Then
                        xlsheet.Cells(row, 3).Value = Mid(MBMICNo, 1, 6) & "-" & Mid(MBMICNo, 7, 2) & "-" & Mid(MBMICNo, 9, 4)
                    Else
                        xlsheet.Cells(row, 3).Value = MBMICNo
                    End If
                                        
                   ' If Len(POLSubNo1) Then
                   '     xlsheet.Cells(row, 4).Value = Trim(MBMPOLNO) & "*" & POLSubNo1
                   ' Else
                   '     xlsheet.Cells(row, 4).Value = "'" & MBMPOLNO
                   ' End If
                    
                    If Mid(MBMNo, 9, 1) = "*" Then
                        xlsheet.Cells(row, 4).Value = Mid(MBMNo, 1, 8) & "-" & "00"
                    ElseIf Mid(MBMNo, 10, 1) = "*" Then
                        xlsheet.Cells(row, 4).Value = Mid(MBMNo, 1, 9) & "-" & "00"
                    ElseIf Mid(MBMNo, 12, 1) = "*" Then
                        xlsheet.Cells(row, 4).Value = Mid(MBMNo, 1, 11) & "-" & "00"
                    Else
                        xlsheet.Cells(row, 4).Value = MBMNo & "-" & "00"
                    End If
                    'xlsheet.Cells(row, 6).Value = PLNCode
                    'xlsheet.Cells(row, 7).Value = Effdate
                    'xlsheet.Cells(row, 8).Value = ExpDate
                    If tmpPrintMode = "G" Then
                        xlsheet.Cells(row, 5).Value = GroupCompany
                    Else
                        xlsheet.Cells(row, 5).Value = Department
                    End If
                    'xlsheet.Cells(row, 10).Value = EmployeeNo
                    'xlsheet.Cells(row, 11).Value = MBMBranch
                        
                    'xlsheet.Cells(row, 13).Value = tempInsPLN
                    'xlsheet.Cells(row, 14).Value = tmpMBMPrevMemNo
                    xlsheet.Cells(row, 6).Value = tmpEncoding
                    'xlsheet.Cells(row, 16).Value = Subsidary
                    'xlsheet.Cells(row, 17).Value = AgentCode
                    'xlsheet.Cells(row, 18).Value = AgentName
                    'xlsheet.Cells(row, 19).Value = PLNPayorCode
                    'xlsheet.Cells(row, 20).Value = Joindt
                    
                        strsqlCOV = ""
                        strsqlCOV = "Select MBMCNumber,MBMCName,MBMCCoverID,MBMCIcBcPpNo,MBMCPLNCode,MBMCEffDate, MBMCExpDate,MBMAXANumber, HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMAgentCode, HISDB.dbo.MBMOthers.CMCName,HISDB.dbo.MBMOthers.AgentName,PLN.PLNPayorPlan,PLN.PLNCardPrintMode from MBMCoveredPersons " & _
                        "inner join HISDB.dbo.MBMOthers on HISDB.dbo.MBMOthers.MBMNumber = MBMCoveredPersons.MBMCNumber " & _
                        "inner join HISDB.dbo.MBM on HISDB.dbo.MBM.MBMNumber = HISDB.dbo.MBMOthers.MBMNumber " & _
                        "inner join HISDB.dbo.PLN on PLN.PLNCode = HISDB.dbo.MBM.PLNCode " & _
                        "where MBMCNumber = '" & Trim(MBMNo) & "'"
                        MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
                    If MBMCov.recordCount Then
                    Do Until MBMCov.EOF
                    
                        'tmpEncoding1 = MBMOther!MBMSecurityCode & MBMOther!MBMCardType & MBMOther!MBMClientType & "  etiqahc.com.my                />" & MBMOther!MBMNumber
                        tmpEncoding = tmpEncoding1 & "-" & MBMCov!MBMCCoverID

                        If Len(MBMCov!MBMCPLNCode) Then
                            If Trim(MBMCov!MBMCPLNCode) = "21" Or Trim(MBMCov!MBMCPLNCode) = "31" Or Trim(MBMCov!MBMCPLNCode) = "41" Then
                                MBMCPLNCode = "60"
                            ElseIf Trim(MBMCov!MBMCPLNCode) = "22" Or Trim(MBMCov!MBMCPLNCode) = "32" Or Trim(MBMCov!MBMCPLNCode) = "42" Then
                                MBMCPLNCode = "80"
                            ElseIf Trim(MBMCov!MBMCPLNCode) = "23" Or Trim(MBMCov!MBMCPLNCode) = "33" Or Trim(MBMCov!MBMCPLNCode) = "43" Then
                                MBMCPLNCode = "120"
                            ElseIf Trim(MBMCov!MBMCPLNCode) = "24" Or Trim(MBMCov!MBMCPLNCode) = "34" Or Trim(MBMCov!MBMCPLNCode) = "44" Then
                                MBMCPLNCode = "200"
                            ElseIf Trim(MBMCov!MBMCPLNCode) = "25" Or Trim(MBMCov!MBMCPLNCode) = "35" Or Trim(MBMCov!MBMCPLNCode) = "45" Then
                                MBMCPLNCode = "350"
                            ElseIf Trim(MBMCov!MBMCPLNCode) = "26" Or Trim(MBMCov!MBMCPLNCode) = "36" Or Trim(MBMCov!MBMCPLNCode) = "46" Then
                                MBMCPLNCode = "450"
                            Else
                                MBMCPLNCode = MBMCov!MBMCPLNCode
                            End If
                        End If
                        
                            If MBMCov!PLNCardPrintMode <> "" Then
                                tmpPrintMode = MBMCov!PLNCardPrintMode
                            Else
                                tmpPrintMode = ""
                            End If
                            
                        Rcounter = Rcounter + 1
                        row = row + 1
                        MBMCName = MBMCov!MBMCName
                        MBMCCoverID = MBMCov!MBMCCoverID
                        MBMCICNo = MBMCov!MBMCICBCPPNo
                        
                        If Len(MBMCov!MBMCEffDate) Then
                            tmpSuppEffDate = MBMCov!MBMCEffDate
                        Else
                            tmpSuppEffDate = ""
                        End If
                        
                        If Len(MBMCov!MBMCExpDate) Then
                            tmpSuppExpDate = MBMCov!MBMCExpDate
                        Else
                            tmpSuppExpDate = ""
                        End If
                        xlsheet.Cells(row, 1).Value = Rcounter
                        xlsheet.Cells(row, 2).Value = MBMCName
                        If IsNumeric(MBMCICNo) Then
                            xlsheet.Cells(row, 3).Value = Mid(MBMCICNo, 1, 6) & "-" & Mid(MBMCICNo, 7, 2) & "-" & Mid(MBMCICNo, 9, 4)
                        Else
                            xlsheet.Cells(row, 3).Value = MBMCICNo
                        End If
                        'If Len(POLSubNo1) Then
                        '    xlsheet.Cells(row, 4).Value = Trim(MBMPOLNO) & "*" & POLSubNo1
                        'Else
                        '    xlsheet.Cells(row, 4).Value = MBMPOLNO
                        'End If
                        If Mid(MBMNo, 9, 1) = "*" Then
                            xlsheet.Cells(row, 4).Value = Mid(MBMNo, 1, 8) & "-" & MBMCCoverID
                        ElseIf Mid(MBMNo, 10, 1) = "*" Then
                            xlsheet.Cells(row, 4).Value = Mid(MBMNo, 1, 9) & "-" & MBMCCoverID
                        ElseIf Mid(MBMNo, 12, 1) = "*" Then
                            xlsheet.Cells(row, 4).Value = Mid(MBMNo, 1, 11) & "-" & MBMCCoverID
                        Else
                            xlsheet.Cells(row, 4).Value = MBMNo & "-" & MBMCCoverID
                        End If
                        'If Len(MBMCov!MBMCPLNCode) Then
                        '    xlsheet.Cells(row, 6).Value = MBMCPLNCode
                        'Else
                        '    xlsheet.Cells(row, 6).Value = PLNCode
                        'End If
                        'If tmpSuppEffDate <> "" Then
                        '    xlsheet.Cells(row, 7).Value = Format(tmpSuppEffDate, "dd/mm/yyyy")
                        'Else
                        '    xlsheet.Cells(row, 7).Value = Format(Effdate, "dd/mm/yyyy")
                        'End If
                       '
                        'If tmpSuppExpDate <> "" Then
                        '    xlsheet.Cells(row, 8).Value = Format(tmpSuppExpDate, "dd/mm/yyyy")
                        'Else
                        '    xlsheet.Cells(row, 8).Value = Format(ExpDate, "dd/mm/yyyy")
                        'End If
                        If tmpPrintMode = "G" Then
                            xlsheet.Cells(row, 5).Value = GroupCompany
                        Else
                            xlsheet.Cells(row, 5).Value = Department
                        End If
                        'xlsheet.Cells(row, 10).Value = EmployeeNo
                        'xlsheet.Cells(row, 11).Value = MBMBranch
                        
                        'xlsheet.Cells(row, 13).Value = tempInsPLN
                        'xlsheet.Cells(row, 14).Value = MBMCov!MBMAXANumber
                        xlsheet.Cells(row, 6).Value = tmpEncoding
                        'xlsheet.Cells(row, 16).Value = Subsidary
                        'xlsheet.Cells(row, 17).Value = IIf(Len(MBMCov!MBMAgentCode), MBMCov!MBMAgentCode, "") 'Agent Code
                        'xlsheet.Cells(row, 18).Value = IIf(Len(MBMCov!AgentName), MBMCov!AgentName, "")  'Agent Name
                        'xlsheet.Cells(row, 19).Value = IIf(Len(MBMCov!PLNPayorPlan), MBMCov!PLNPayorPlan, "")
                        'xlsheet.Cells(row, 20).Value = IIf(Len(Joindt), Format(Joindt, "dd/MM/YYYY"), "")
                    TotalSupplementary = TotalSupplementary + 1
                    txtEXPTotalSupplementary1 = TotalSupplementary
                    MBMCov.MoveNext
                    Loop
                    MBMCov.Close
                    Set MBMCov = Nothing
                    Else
                        Rcounter = Rcounter + 1
                        row = row + 1
                        If MBMCov.recordCount = 0 Then
                            MBMCov.Close
                            Set MBMCov = Nothing
                        End If
                    End If
        TotalPrincipal = TotalPrincipal + 1
        txtEXPTotalPrincipal = TotalPrincipal
        
        MBM.MoveNext
        Loop
        'End If
        xlBook.Save
        xlBook.SaveAs (Trim(txtSVRPath) & Trim(txtFilName & " - Print" & ".xls"))
        MsgBox "Export successful", vbOKOnly + vbInformation, DialogBoxTitle
        'Call cmdExpSearch_Click
        
        Set xlBook = Nothing
        xlApp.Quit
        Set xlApp = Nothing
        MBM.Close
        Set MBM = Nothing
        Frame2.Enabled = True
    End If
End Sub


Private Sub EXPPrincipalAck()
Dim strsql As String
Dim strsqls As String
Dim sql As String
Dim MBMOther As New ADODB.Recordset
Dim MBM As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim ListMaster As ListItem
Dim isPrinted As Boolean
Dim rsSearch As New ADODB.Recordset
Dim ColumnItem As ListItem
Dim itemToPrint As Integer
Dim CoverQty As Integer
Dim TotalPrincipal As Integer
Dim strsqlCOV As String
Dim MBMCov As New ADODB.Recordset
Dim MBMCPLNCode  As String
Dim tempInsPLN As String
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim tmpSuppExpDate As String
Dim Effdate As String
Dim tmpSuppEffDate As String
Dim tmpMBMPrevMemNo As String
Dim tmpMBMNumber As String
Dim tmpEncoding As String
Dim tmpEncoding1 As String

TotalSupplementary = 0
TotalPrincipal = 0
MBMCPLNCode = ""
tmpEncoding = ""

Select Case cboEXPCriteriaSel.ListIndex
    Case 0 'Card Export Criteria Selection for Single Member
        'lvwCRD.listitems.Clear
        If Len(txtExpMemberNo) = "8" Then
        'strsql = ""
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate from View_MBM_CardExport where substring(MBMNumber,1,8) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,10,2) = '01'  "
        ElseIf Len(txtExpMemberNo) = "9" Then
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate from View_MBM_CardExport where substring(MBMNumber,1,9) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,11,2) = '01'  "
        ElseIf Len(txtExpMemberNo) = "11" Then
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate from View_MBM_CardExport where substring(MBMNumber,1,11) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,13,2) = '01'  "
            
        Else
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate from View_MBM_CardExport where MBMNumber = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D'  "
        End If
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Case 1 'Card Export Criteria Selection for Member's No Range
        strsql = "SELECT INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate from View_MBM_CardExport "
        
        If Len(txtExpMemberNoFr) = "8" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,8) >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND substring(MBMNumber,1,8) <= '" & txtExpMemberNoTo & "'"
            strsql = strsql & " AND substring(MBMNumber,10,2) <= '01'"
        ElseIf Len(txtExpMemberNoFr) = "9" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,9) >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND substring(MBMNumber,1,9) <= '" & txtExpMemberNoTo & "'"
            strsql = strsql & " AND substring(MBMNumber,11,2) <= '01'"
        ElseIf Len(Trim(txtExpMemberNoFr)) = 11 And Mid(Trim(txtExpMemberNoFr), 9, 1) <> "*" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,11) >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND substring(MBMNumber,1,11) <= '" & txtExpMemberNoTo & "'"
            strsql = strsql & " AND substring(MBMNumber,13,2) <= '01'"
        Else
            strsql = strsql & " WHERE MBMNumber >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND MBMNumber <= '" & txtExpMemberNoTo & "'"
        End If
        strsql = strsql & " AND MBMDataStatus <> 'D' "
        strsql = strsql & " AND MBMStatus <> 'C' "
        
      If cboCardSequence.ListIndex = 0 Then
            strsql = strsql & " Order By MBMNumber "
        ElseIf cboCardSequence.ListIndex = 1 Then
            strsql = strsql & " Order By MBMPrevMemNo "
        ElseIf cboCardSequence.ListIndex = 2 Then
            strsql = strsql & " Order By MBMName "
        'ElseIf cboLLPrintSeq.ListIndex = 3 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMNumber"
        ElseIf cboCardSequence.ListIndex = 3 Then
            strsql = strsql & " Order By MBMDepartment, MBMName"
        'ElseIf cboLLPrintSeq.ListIndex = 5 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMEmployeeNo"
        'ElseIf cboLLPrintSeq.ListIndex = 6 Then
        '    strsql = strsql & " Order By MBMEmployeeNo"
        End If
        
        'strsql = strsql & "order by MBMNumber"
        
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Case 2
        sDate = CStr(Month(txtExpBordxDate)) & "/" & CStr(Day(txtExpBordxDate)) & "/" & CStr(Year(txtExpBordxDate))
        strsql = ""
        strsql = "SELECT INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorCode,JoinDate from View_MBM_CardExport "
        strsql = strsql & " WHERE MBMBordxDate = '" & Format(txtExpBordxDate, "mm/dd/yyyy") & "'"
        strsql = strsql & " AND MBMBatchNo = '" & txtExpBatchNo & "'"
        strsql = strsql & " AND PAYCode = '" & Mid(cboExpPayor, 1, 2) & "'"
        strsql = strsql & " AND MBMDataStatus <> 'D' "
        strsql = strsql & " AND MBMStatus <> 'C' "
        
        If cboCardSequence.ListIndex = 0 Then
            strsql = strsql & " Order By MBMNumber "
        ElseIf cboCardSequence.ListIndex = 1 Then
            strsql = strsql & " Order By MBMPrevMemNo "
        ElseIf cboCardSequence.ListIndex = 2 Then
            strsql = strsql & " Order By MBMName "
        'ElseIf cboLLPrintSeq.ListIndex = 3 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMNumber"
        ElseIf cboCardSequence.ListIndex = 3 Then
            strsql = strsql & " Order By MBMDepartment, MBMName"
        'ElseIf cboLLPrintSeq.ListIndex = 5 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMEmployeeNo"
        'ElseIf cboLLPrintSeq.ListIndex = 6 Then
        '    strsql = strsql & " Order By MBMEmployeeNo"
        End If
        'strsql = strsql & "order by MBMNumber"
        'strsql = strsql & "order by MBMNumber"
        'strsql = strsql & " Order By MBMNumber"
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
     Case 3
        Call MemberListing
   End Select
    
        If cboEXPCriteriaSel.ListIndex <> "3" Then
            If MBM.EOF = True Then
                MsgBox "Record not found !", vbInformation + vbOKOnly
                'cmdCRDPrint.Enabled = False
                Screen.MousePointer = Default
                Exit Sub
            End If
        
            Frame2.Enabled = False
            txtSVRPath = crdPath & "CardPrinting\"
            txtPath = "C:\CardPrinting\"
            If cboEXPCriteriaSel.ListIndex = 0 Then
                txtFilName = txtExpMemberNo & " - Acknowledgement"
            ElseIf cboEXPCriteriaSel.ListIndex = 1 Then
                txtFilName = txtExpMemberNoFr & " - Acknowledgement"
            Else
                txtFilName = Mid(cboExpPayor, 1, 2) & Mid(txtExpBordxDate, 1, 2) & Mid(txtExpBordxDate, 4, 2) & Mid(txtExpBordxDate, 7, 4) & " - Acknowledgement"
            End If
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set txtSVRfile = fso.CreateTextFile((Trim(txtSVRPath) & Trim(txtFilName & ".xls")), True)
            Set txtfile = fso.CreateTextFile((Trim(txtPath) & Trim(txtFilName & ".xls")), True)
            txtfile = (Trim(txtPath) & Trim(txtFilName & ".xls"))
            txtSVRfile = (Trim(txtSVRPath) & Trim(txtFilName & ".xls"))
                
            Set xlApp = CreateObject("Excel.application")
            Set xlBook = xlApp.Workbooks.Open(txtfile)
            Set xlsheet = xlBook.Sheets(1)
             xlsheet.Cells(1, 1).Value = cboExpPayor
             xlsheet.Cells(2, 1).Value = "DATE :" & Mid(txtExpBordxDate, 1, 2) & "/" & Mid(txtExpBordxDate, 4, 2) & "/" & Mid(txtExpBordxDate, 7, 4)
             xlsheet.Cells(3, 1).Value = "FILE NAME :" & txtFilName
             row = 5
             Rcounter = 0
            
        
        Do Until MBM.EOF
            strsqls = ""
            tmpEncoding = ""
            tmpMBMNumber = Split(MBM!MBMNumber, "*")(0)
            strsqls = "Select MBMNumber,MBMSecurityCode,MBMCardType,MBMClientType from MBMSCSecurity where MBMNumber = '" & Trim(tmpMBMNumber) & "'"
            MBMOther.Open strsqls, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                If MBMOther.recordCount > 0 Then
                    'If Mid(cboExpPayor, 1, 2) = "EM" Or Mid(cboExpPayor, 1, 2) = "ES" Then
                    '    tmpEncoding1 = MBMOther!MBMSecurityCode & MBMOther!MBMCardType & MBMOther!MBMClientType & "  etiqahc.com.my                />" & MBMOther!MBMNumber
                    '    tmpEncoding = tmpEncoding1 & "-00"
                    'Else
                        tmpEncoding1 = MBMOther!MBMSecurityCode & MBMOther!MBMCardType & MBMOther!MBMClientType & "  hcsb.com.my                   />" & MBMOther!MBMNumber
                        tmpEncoding = tmpEncoding1 & "-00"
                    'End If
                    '6001mmed  hcsb.com.my                   />dcga8524393-00
                End If
            MBMOther.Close
            Set MBMOther = Nothing
            
            strsqls = ""
            strsqls = "Select MBMBranch, MBMPOLSubNo1, MBMGroupCompany, MBMEmployeeNo, MBMDepartment,MBMPrevMemNo,Subsidiary,MBMAgentCode,AgentName,PLN.PLNPayorPlan,MBMOthers.MBMFirstJoinedDate AS JoinDate from MBMOthers  inner join MBM on MBM.MBMnumber = MBMOthers.MBMnumber inner join PLN on PLN.PLNCode = MBM.PLNCode where MBMOthers.MBMNumber = '" & Trim(MBM!MBMNumber) & "'"
            MBMOther.Open strsqls, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                                                
            If MBMOther.recordCount Then
                If Len(MBMOther!MBMGroupCompany) Then
                    GroupCompany = MBMOther!MBMGroupCompany
                Else
                    GroupCompany = ""
                End If
                
                If Len(MBMOther!MBMEmployeeNo) Then
                    EmployeeNo = MBMOther!MBMEmployeeNo
                Else
                    EmployeeNo = ""
                End If
                
                If Len(MBMOther!MBMDepartment) Then
                    Department = MBMOther!MBMDepartment
                Else
                    Department = ""
                End If
                
                If Len(MBMOther!MBMPolSubNo1) Then
                    POLSubNo1 = MBMOther!MBMPolSubNo1
                Else
                    POLSubNo1 = ""
                End If
                
                If Len(MBMOther!MBMBranch) Then
                    MBMBranch = MBMOther!MBMBranch
                Else
                    MBMBranch = ""
                End If
                
                If Len(MBMOther!MBMPrevMemNo) Then
                    tmpMBMPrevMemNo = MBMOther!MBMPrevMemNo
                Else
                    tmpMBMPrevMemNo = ""
                End If
                
                If Len(MBMOther!Subsidiary) Then
                    Subsidary = MBMOther!Subsidiary
                Else
                    Subsidary = ""
                End If
                
                If Len(MBMOther!MBMAgentCode) Then
                    AgentCode = MBMOther!MBMAgentCode
                Else
                    AgentCode = ""
                End If
                
                If Len(MBMOther!AgentName) Then
                    AgentName = MBMOther!AgentName
                Else
                    AgentName = ""
                End If
                
                If Len(MBMOther!PLNPayorPlan) Then
                    PLNPayorCode = MBMOther!PLNPayorPlan
                Else
                    PLNPayorCode = ""
                End If
                
                If Len(MBMOther!JoinDate) Then
                    Joindt = Format(MBMOther!JoinDate, "dd/MM/yyyy")
                Else
                    Joindt = ""
                End If
                
                
            End If
            MBMOther.Close
            Set MBMOther = Nothing
            strsqls = ""
           
               ' Display record in List View
               Cnt = 0
               Rcounter = Rcounter + 1
               row = row + 1
               xlApp.DisplayAlerts = False
                 ' Export record to excel format
                 
                    MBMName = MBM!MBMName
                    MBMNo = MBM!MBMNumber
                    MBMICNo = MBM!MBMIcBcPp
                    MBMPOLNO = MBM!MBMPolicyNo
                    If Trim(MBM!PLNCode) = "21" Or Trim(MBM!PLNCode) = "31" Or Trim(MBM!PLNCode) = "41" Then
                        PLNCode = "60"
                    ElseIf Trim(MBM!PLNCode) = "22" Or Trim(MBM!PLNCode) = "32" Or Trim(MBM!PLNCode) = "42" Then
                        PLNCode = "80"
                    ElseIf Trim(MBM!PLNCode) = "23" Or Trim(MBM!PLNCode) = "33" Or Trim(MBM!PLNCode) = "43" Then
                        PLNCode = "120"
                    ElseIf Trim(MBM!PLNCode) = "24" Or Trim(MBM!PLNCode) = "34" Or Trim(MBM!PLNCode) = "44" Then
                        PLNCode = "200"
                    ElseIf Trim(MBM!PLNCode) = "25" Or Trim(MBM!PLNCode) = "35" Or Trim(MBM!PLNCode) = "45" Then
                        PLNCode = "350"
                    ElseIf Trim(MBM!PLNCode) = "26" Or Trim(MBM!PLNCode) = "36" Or Trim(MBM!PLNCode) = "46" Then
                        PLNCode = "450"
                    Else
                        PLNCode = IIf(Len(MBM!PLNCode), MBM!PLNCode, "")
                    End If
                    Effdate = Format(MBM!MBMPayorEffDate, "dd/mm/yyyy")
                    ExpDate = Format(MBM!MBMPayorEXPDate, "dd/mm/yyyy")
                   'strsqls = "Select MBMBranch, MBMPOLSubNo1, MBMGroupCompany, MBMEmployeeNo, MBMDepartment from MBMOthers where MBMNumber = '" & Trim(MBM!MBMNumber) & "'"

                    strsqlPLN = "Select PLNCode,PLNClientPlan,PLNRBType,PLNDescription,PLNCardPrintMode from PLN where PLNCode = '" & MBM!PLNCode & "' And PAYCode = '" & MBM!PAYCode & "'"
                    PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If PLN.recordCount Then
                            If Trim(MBM!INSCode) = "G" Then
                                tempInsPLN = "INDIVIDUAL" & " " & PLN!PLNDescription
                            ElseIf Trim(MBM!INSCode) = "H" Then
                                tempInsPLN = "FAMILY" & " " & PLN!PLNDescription
                            End If
                            If PLN!PLNCardPrintMode <> "" Then
                                tmpPrintMode = PLN!PLNCardPrintMode
                            Else
                                tmpPrintMode = ""
                            End If
                        End If
                    PLN.Close
                    Set PLN = Nothing
                    
                    xlsheet.Cells(row, 1).Value = Rcounter
                    xlsheet.Cells(row, 2).Value = MBMName
                    xlsheet.Cells(row, 3).Value = MBMName
                    If IsNumeric(MBMICNo) = True Then
                        xlsheet.Cells(row, 4).Value = Mid(MBMICNo, 1, 6) & "-" & Mid(MBMICNo, 7, 2) & "-" & Mid(MBMICNo, 9, 4)
                    Else
                        xlsheet.Cells(row, 4).Value = MBMICNo
                    End If
                                        
                   ' If Len(POLSubNo1) Then
                   '     xlsheet.Cells(row, 4).Value = Trim(MBMPOLNO) & "*" & POLSubNo1
                   ' Else
                   '     xlsheet.Cells(row, 4).Value = "'" & MBMPOLNO
                   ' End If
                    
                    If Mid(MBMNo, 9, 1) = "*" Then
                        xlsheet.Cells(row, 5).Value = Mid(MBMNo, 1, 8) & "-" & "00"
                    ElseIf Mid(MBMNo, 10, 1) = "*" Then
                        xlsheet.Cells(row, 5).Value = Mid(MBMNo, 1, 9) & "-" & "00"
                    ElseIf Mid(MBMNo, 12, 1) = "*" Then
                        xlsheet.Cells(row, 5).Value = Mid(MBMNo, 1, 11) & "-" & "00"
                    Else
                        xlsheet.Cells(row, 5).Value = MBMNo & "-" & "00"
                    End If
                    'xlsheet.Cells(row, 6).Value = PLNCode
                    'xlsheet.Cells(row, 7).Value = Effdate
                    'xlsheet.Cells(row, 8).Value = ExpDate
                    'If tmpPrintMode = "G" Then
                        xlsheet.Cells(row, 6).Value = GroupCompany
                    'Else
                        xlsheet.Cells(row, 7).Value = Department
                    'End If
                    'xlsheet.Cells(row, 10).Value = EmployeeNo
                    xlsheet.Cells(row, 8).Value = MBMBranch
                        
                    'xlsheet.Cells(row, 13).Value = tempInsPLN
                    'xlsheet.Cells(row, 14).Value = tmpMBMPrevMemNo
                    'xlsheet.Cells(row, 6).Value = tmpEncoding
                    xlsheet.Cells(row, 9).Value = Subsidary
                    'xlsheet.Cells(row, 17).Value = AgentCode
                    'xlsheet.Cells(row, 18).Value = AgentName
                    'xlsheet.Cells(row, 19).Value = PLNPayorCode
                    'xlsheet.Cells(row, 20).Value = Joindt
                    
                        strsqlCOV = ""
                        strsqlCOV = "Select MBMCNumber,MBMCName,MBMCCoverID,MBMCIcBcPpNo,MBMCPLNCode,MBMCEffDate, MBMCExpDate,MBMAXANumber, HISDB.dbo.MBMOthers.Subsidiary, HISDB.dbo.MBMOthers.MBMAgentCode, HISDB.dbo.MBMOthers.CMCName,HISDB.dbo.MBMOthers.AgentName,PLN.PLNPayorPlan,PLN.PLNCardPrintMode from MBMCoveredPersons " & _
                        "inner join HISDB.dbo.MBMOthers on HISDB.dbo.MBMOthers.MBMNumber = MBMCoveredPersons.MBMCNumber " & _
                        "inner join HISDB.dbo.MBM on HISDB.dbo.MBM.MBMNumber = HISDB.dbo.MBMOthers.MBMNumber " & _
                        "inner join HISDB.dbo.PLN on PLN.PLNCode = HISDB.dbo.MBM.PLNCode " & _
                        "where MBMCNumber = '" & Trim(MBMNo) & "'"
                        MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
                    If MBMCov.recordCount Then
                    Do Until MBMCov.EOF
                    
                        'tmpEncoding1 = MBMOther!MBMSecurityCode & MBMOther!MBMCardType & MBMOther!MBMClientType & "  etiqahc.com.my                />" & MBMOther!MBMNumber
                        tmpEncoding = tmpEncoding1 & "-" & MBMCov!MBMCCoverID

                        If Len(MBMCov!MBMCPLNCode) Then
                            If Trim(MBMCov!MBMCPLNCode) = "21" Or Trim(MBMCov!MBMCPLNCode) = "31" Or Trim(MBMCov!MBMCPLNCode) = "41" Then
                                MBMCPLNCode = "60"
                            ElseIf Trim(MBMCov!MBMCPLNCode) = "22" Or Trim(MBMCov!MBMCPLNCode) = "32" Or Trim(MBMCov!MBMCPLNCode) = "42" Then
                                MBMCPLNCode = "80"
                            ElseIf Trim(MBMCov!MBMCPLNCode) = "23" Or Trim(MBMCov!MBMCPLNCode) = "33" Or Trim(MBMCov!MBMCPLNCode) = "43" Then
                                MBMCPLNCode = "120"
                            ElseIf Trim(MBMCov!MBMCPLNCode) = "24" Or Trim(MBMCov!MBMCPLNCode) = "34" Or Trim(MBMCov!MBMCPLNCode) = "44" Then
                                MBMCPLNCode = "200"
                            ElseIf Trim(MBMCov!MBMCPLNCode) = "25" Or Trim(MBMCov!MBMCPLNCode) = "35" Or Trim(MBMCov!MBMCPLNCode) = "45" Then
                                MBMCPLNCode = "350"
                            ElseIf Trim(MBMCov!MBMCPLNCode) = "26" Or Trim(MBMCov!MBMCPLNCode) = "36" Or Trim(MBMCov!MBMCPLNCode) = "46" Then
                                MBMCPLNCode = "450"
                            Else
                                MBMCPLNCode = MBMCov!MBMCPLNCode
                            End If
                        End If
                        
                            If MBMCov!PLNCardPrintMode <> "" Then
                                tmpPrintMode = MBMCov!PLNCardPrintMode
                            Else
                                tmpPrintMode = ""
                            End If
                            
                        Rcounter = Rcounter + 1
                        row = row + 1
                        MBMCName = MBMCov!MBMCName
                        MBMCCoverID = MBMCov!MBMCCoverID
                        MBMCICNo = MBMCov!MBMCICBCPPNo
                        
                        If Len(MBMCov!MBMCEffDate) Then
                            tmpSuppEffDate = MBMCov!MBMCEffDate
                        Else
                            tmpSuppEffDate = ""
                        End If
                        
                        If Len(MBMCov!MBMCExpDate) Then
                            tmpSuppExpDate = MBMCov!MBMCExpDate
                        Else
                            tmpSuppExpDate = ""
                        End If
                        xlsheet.Cells(row, 1).Value = Rcounter
                        xlsheet.Cells(row, 2).Value = MBMName
                        xlsheet.Cells(row, 3).Value = MBMCName
                        If IsNumeric(MBMCICNo) Then
                            xlsheet.Cells(row, 4).Value = Mid(MBMCICNo, 1, 6) & "-" & Mid(MBMCICNo, 7, 2) & "-" & Mid(MBMCICNo, 9, 4)
                        Else
                            xlsheet.Cells(row, 4).Value = MBMCICNo
                        End If
                        'If Len(POLSubNo1) Then
                        '    xlsheet.Cells(row, 4).Value = Trim(MBMPOLNO) & "*" & POLSubNo1
                        'Else
                        '    xlsheet.Cells(row, 4).Value = MBMPOLNO
                        'End If
                        If Mid(MBMNo, 9, 1) = "*" Then
                            xlsheet.Cells(row, 5).Value = Mid(MBMNo, 1, 8) & "-" & MBMCCoverID
                        ElseIf Mid(MBMNo, 10, 1) = "*" Then
                            xlsheet.Cells(row, 5).Value = Mid(MBMNo, 1, 9) & "-" & MBMCCoverID
                        ElseIf Mid(MBMNo, 12, 1) = "*" Then
                            xlsheet.Cells(row, 5).Value = Mid(MBMNo, 1, 11) & "-" & MBMCCoverID
                        Else
                            xlsheet.Cells(row, 5).Value = MBMNo & "-" & MBMCCoverID
                        End If
                        'If Len(MBMCov!MBMCPLNCode) Then
                        '    xlsheet.Cells(row, 6).Value = MBMCPLNCode
                        'Else
                        '    xlsheet.Cells(row, 6).Value = PLNCode
                        'End If
                        'If tmpSuppEffDate <> "" Then
                        '    xlsheet.Cells(row, 7).Value = Format(tmpSuppEffDate, "dd/mm/yyyy")
                        'Else
                        '    xlsheet.Cells(row, 7).Value = Format(Effdate, "dd/mm/yyyy")
                        'End If
                       '
                        'If tmpSuppExpDate <> "" Then
                        '    xlsheet.Cells(row, 8).Value = Format(tmpSuppExpDate, "dd/mm/yyyy")
                        'Else
                        '    xlsheet.Cells(row, 8).Value = Format(ExpDate, "dd/mm/yyyy")
                        'End If
                        'If tmpPrintMode = "G" Then
                            xlsheet.Cells(row, 6).Value = GroupCompany
                        'Else
                            xlsheet.Cells(row, 7).Value = Department
                        'End If
                        'xlsheet.Cells(row, 10).Value = EmployeeNo
                        xlsheet.Cells(row, 8).Value = MBMBranch
                        
                        'xlsheet.Cells(row, 13).Value = tempInsPLN
                        'xlsheet.Cells(row, 14).Value = MBMCov!MBMAXANumber
                        'xlsheet.Cells(row, 6).Value = tmpEncoding
                        xlsheet.Cells(row, 9).Value = Subsidary
                        'xlsheet.Cells(row, 17).Value = IIf(Len(MBMCov!MBMAgentCode), MBMCov!MBMAgentCode, "") 'Agent Code
                        'xlsheet.Cells(row, 18).Value = IIf(Len(MBMCov!AgentName), MBMCov!AgentName, "")  'Agent Name
                        'xlsheet.Cells(row, 19).Value = IIf(Len(MBMCov!PLNPayorPlan), MBMCov!PLNPayorPlan, "")
                        'xlsheet.Cells(row, 20).Value = IIf(Len(Joindt), Format(Joindt, "dd/MM/YYYY"), "")
                    TotalSupplementary = TotalSupplementary + 1
                    txtEXPTotalSupplementary1 = TotalSupplementary
                    MBMCov.MoveNext
                    Loop
                    MBMCov.Close
                    Set MBMCov = Nothing
                    Else
                        Rcounter = Rcounter + 1
                        row = row + 1
                        If MBMCov.recordCount = 0 Then
                            MBMCov.Close
                            Set MBMCov = Nothing
                        End If
                    End If
        TotalPrincipal = TotalPrincipal + 1
        txtEXPTotalPrincipal = TotalPrincipal
        
        MBM.MoveNext
        Loop
        'End If
        xlBook.Save
        xlBook.SaveAs (Trim(txtSVRPath) & Trim(txtFilName & ".xls"))
        MsgBox "Export successful", vbOKOnly + vbInformation, DialogBoxTitle
        'Call cmdExpSearch_Click
        
        Set xlBook = Nothing
        xlApp.Quit
        Set xlApp = Nothing
        MBM.Close
        Set MBM = Nothing
        Frame2.Enabled = True
    End If
End Sub

Private Sub EXPPrincipalEnvelope()
Dim strsql As String
Dim strsqls As String
Dim sql As String
Dim MBMOther As New ADODB.Recordset
Dim MBM As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim ListMaster As ListItem
Dim isPrinted As Boolean
Dim rsSearch As New ADODB.Recordset
Dim ColumnItem As ListItem
Dim itemToPrint As Integer
Dim CoverQty As Integer
Dim TotalPrincipal As Integer
Dim strsqlCOV As String
Dim MBMCov As New ADODB.Recordset
Dim MBMCPLNCode  As String
Dim tempInsPLN As String
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim tmpSuppExpDate As String
Dim Effdate As String
Dim tmpSuppEffDate As String
Dim tmpMBMPrevMemNo As String
Dim tmpMBMNumber As String
Dim tmpEncoding As String
Dim tmpEncoding1 As String

TotalSupplementary = 0
TotalPrincipal = 0
MBMCPLNCode = ""
tmpEncoding = ""

Select Case cboEXPCriteriaSel.ListIndex
    Case 0 'Card Export Criteria Selection for Single Member
        'lvwCRD.listitems.Clear
        If Len(txtExpMemberNo) = "8" Then
        'strsql = ""
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate,PLNCardPrintMode from View_MBM_CardExport where substring(MBMNumber,1,8) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,10,2) = '01'  "
        ElseIf Len(txtExpMemberNo) = "9" Then
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate,PLNCardPrintMode from View_MBM_CardExport where substring(MBMNumber,1,9) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,11,2) = '01'  "
        ElseIf Len(txtExpMemberNo) = "11" Then
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate,PLNCardPrintMode from View_MBM_CardExport where substring(MBMNumber,1,11) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,13,2) = '01'  "
            
        Else
            strsql = "Select INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate,PLNCardPrintMode from View_MBM_CardExport where MBMNumber = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D'  "
        End If
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Case 1 'Card Export Criteria Selection for Member's No Range
        strsql = "SELECT INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorPlan,JoinDate,PLNCardPrintMode from View_MBM_CardExport "
        
        If Len(txtExpMemberNoFr) = "8" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,8) >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND substring(MBMNumber,1,8) <= '" & txtExpMemberNoTo & "'"
            strsql = strsql & " AND substring(MBMNumber,10,2) <= '01'"
        ElseIf Len(txtExpMemberNoFr) = "9" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,9) >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND substring(MBMNumber,1,9) <= '" & txtExpMemberNoTo & "'"
            strsql = strsql & " AND substring(MBMNumber,11,2) <= '01'"
        ElseIf Len(Trim(txtExpMemberNoFr)) = 11 And Mid(Trim(txtExpMemberNoFr), 9, 1) <> "*" Then
            strsql = strsql & " WHERE substring(MBMNumber,1,11) >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND substring(MBMNumber,1,11) <= '" & txtExpMemberNoTo & "'"
            strsql = strsql & " AND substring(MBMNumber,13,2) <= '01'"
        Else
            strsql = strsql & " WHERE MBMNumber >= '" & txtExpMemberNoFr & "'"
            strsql = strsql & " AND MBMNumber <= '" & txtExpMemberNoTo & "'"
        End If
        strsql = strsql & " AND MBMDataStatus <> 'D' "
        strsql = strsql & " AND MBMStatus <> 'C' "
        
      If cboCardSequence.ListIndex = 0 Then
            strsql = strsql & " Order By MBMNumber "
        ElseIf cboCardSequence.ListIndex = 1 Then
            strsql = strsql & " Order By MBMPrevMemNo "
        ElseIf cboCardSequence.ListIndex = 2 Then
            strsql = strsql & " Order By MBMName "
        'ElseIf cboLLPrintSeq.ListIndex = 3 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMNumber"
        ElseIf cboCardSequence.ListIndex = 3 Then
            strsql = strsql & " Order By MBMDepartment, MBMName"
        'ElseIf cboLLPrintSeq.ListIndex = 5 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMEmployeeNo"
        'ElseIf cboLLPrintSeq.ListIndex = 6 Then
        '    strsql = strsql & " Order By MBMEmployeeNo"
        End If
        
        'strsql = strsql & "order by MBMNumber"
        
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Case 2
        sDate = CStr(Month(txtExpBordxDate)) & "/" & CStr(Day(txtExpBordxDate)) & "/" & CStr(Year(txtExpBordxDate))
        strsql = ""
        strsql = "SELECT INSCode,PAYCode,MBMBatchNo,MBMBordxDate,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEffDate, MBMPayorExpDate,MBMStatus,MBMDataStatus,Subsidiary,MBMAgentCode,AgentName,PLNPayorCode,JoinDate,PLNCardPrintMode from View_MBM_CardExport "
        strsql = strsql & " WHERE MBMBordxDate = '" & Format(txtExpBordxDate, "mm/dd/yyyy") & "'"
        strsql = strsql & " AND MBMBatchNo = '" & txtExpBatchNo & "'"
        strsql = strsql & " AND PAYCode = '" & Mid(cboExpPayor, 1, 2) & "'"
        strsql = strsql & " AND MBMDataStatus <> 'D' "
        strsql = strsql & " AND MBMStatus <> 'C' "
        
        If cboCardSequence.ListIndex = 0 Then
            strsql = strsql & " Order By MBMNumber "
        ElseIf cboCardSequence.ListIndex = 1 Then
            strsql = strsql & " Order By MBMPrevMemNo "
        ElseIf cboCardSequence.ListIndex = 2 Then
            strsql = strsql & " Order By MBMName "
        'ElseIf cboLLPrintSeq.ListIndex = 3 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMNumber"
        ElseIf cboCardSequence.ListIndex = 3 Then
            strsql = strsql & " Order By MBMDepartment, MBMName"
        'ElseIf cboLLPrintSeq.ListIndex = 5 Then
        '    strsql = strsql & " Order By MBMDepartment, MBMEmployeeNo"
        'ElseIf cboLLPrintSeq.ListIndex = 6 Then
        '    strsql = strsql & " Order By MBMEmployeeNo"
        End If
        'strsql = strsql & "order by MBMNumber"
        'strsql = strsql & "order by MBMNumber"
        'strsql = strsql & " Order By MBMNumber"
        MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
     Case 3
        Call MemberListing
   End Select
    
        If cboEXPCriteriaSel.ListIndex <> "3" Then
            If MBM.EOF = True Then
                MsgBox "Record not found !", vbInformation + vbOKOnly
                'cmdCRDPrint.Enabled = False
                Screen.MousePointer = Default
                Exit Sub
            End If
        
            Frame2.Enabled = False
            txtSVRPath = crdPath & "CardPrinting\"
            txtPath = "C:\CardPrinting\"
            If cboEXPCriteriaSel.ListIndex = 0 Then
                txtFilName = txtExpMemberNo & " - Print Envelope"
            ElseIf cboEXPCriteriaSel.ListIndex = 1 Then
                txtFilName = txtExpMemberNoFr & " - Print Envelope"
            Else
                txtFilName = Mid(cboExpPayor, 1, 2) & Mid(txtExpBordxDate, 1, 2) & Mid(txtExpBordxDate, 4, 2) & Mid(txtExpBordxDate, 7, 4) & " - Print Envelope"
            End If
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set txtSVRfile = fso.CreateTextFile((Trim(txtSVRPath) & Trim(txtFilName & ".xls")), True)
            Set txtfile = fso.CreateTextFile((Trim(txtPath) & Trim(txtFilName & ".xls")), True)
            txtfile = (Trim(txtPath) & Trim(txtFilName & ".xls"))
            txtSVRfile = (Trim(txtSVRPath) & Trim(txtFilName & ".xls"))
                
            Set xlApp = CreateObject("Excel.application")
            Set xlBook = xlApp.Workbooks.Open(txtfile)
            Set xlsheet = xlBook.Sheets(1)
             xlsheet.Cells(1, 1).Value = cboExpPayor
             xlsheet.Cells(2, 1).Value = "DATE :" & Mid(txtExpBordxDate, 1, 2) & "/" & Mid(txtExpBordxDate, 4, 2) & "/" & Mid(txtExpBordxDate, 7, 4)
             xlsheet.Cells(3, 1).Value = "FILE NAME :" & txtFilName
             row = 5
             Rcounter = 0
            
        
        Do Until MBM.EOF
            strsqls = ""
            tmpEncoding = ""
            tmpMBMNumber = Split(MBM!MBMNumber, "*")(0)
            strsqls = "Select MBMNumber,MBMSecurityCode,MBMCardType,MBMClientType from MBMSCSecurity where MBMNumber = '" & Trim(tmpMBMNumber) & "'"
            MBMOther.Open strsqls, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                If MBMOther.recordCount > 0 Then
                    'If Mid(cboExpPayor, 1, 2) = "EM" Or Mid(cboExpPayor, 1, 2) = "ES" Then
                    '    tmpEncoding1 = MBMOther!MBMSecurityCode & MBMOther!MBMCardType & MBMOther!MBMClientType & "  etiqahc.com.my                />" & MBMOther!MBMNumber
                    '    tmpEncoding = tmpEncoding1 & "-00"
                    'Else
                        tmpEncoding1 = MBMOther!MBMSecurityCode & MBMOther!MBMCardType & MBMOther!MBMClientType & "  hcsb.com.my                   />" & MBMOther!MBMNumber
                        tmpEncoding = tmpEncoding1 & "-00"
                    'End If
                    '6001mmed  hcsb.com.my                   />dcga8524393-00
                End If
            MBMOther.Close
            Set MBMOther = Nothing
            
            strsqls = ""
            strsqls = "Select MBMBranch, MBMPOLSubNo1, MBMGroupCompany, MBMEmployeeNo, MBMDepartment,MBMPrevMemNo,Subsidiary,MBMAgentCode,AgentName,PLN.PLNPayorPlan,MBMOthers.MBMFirstJoinedDate AS JoinDate from MBMOthers  inner join MBM on MBM.MBMnumber = MBMOthers.MBMnumber inner join PLN on PLN.PLNCode = MBM.PLNCode where MBMOthers.MBMNumber = '" & Trim(MBM!MBMNumber) & "'"
            MBMOther.Open strsqls, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                                                
            If MBMOther.recordCount Then
                If Len(MBMOther!MBMGroupCompany) Then
                    GroupCompany = MBMOther!MBMGroupCompany
                Else
                    GroupCompany = ""
                End If
                
                If Len(MBMOther!MBMEmployeeNo) Then
                    EmployeeNo = MBMOther!MBMEmployeeNo
                Else
                    EmployeeNo = ""
                End If
                
                If Len(MBMOther!MBMDepartment) Then
                    Department = MBMOther!MBMDepartment
                Else
                    Department = ""
                End If
                
                If Len(MBMOther!MBMPolSubNo1) Then
                    POLSubNo1 = MBMOther!MBMPolSubNo1
                Else
                    POLSubNo1 = ""
                End If
                
                If Len(MBMOther!MBMBranch) Then
                    MBMBranch = MBMOther!MBMBranch
                Else
                    MBMBranch = ""
                End If
                
                If Len(MBMOther!MBMPrevMemNo) Then
                    tmpMBMPrevMemNo = MBMOther!MBMPrevMemNo
                Else
                    tmpMBMPrevMemNo = ""
                End If
                
                If Len(MBMOther!Subsidiary) Then
                    Subsidary = MBMOther!Subsidiary
                Else
                    Subsidary = ""
                End If
                
                If Len(MBMOther!MBMAgentCode) Then
                    AgentCode = MBMOther!MBMAgentCode
                Else
                    AgentCode = ""
                End If
                
                If Len(MBMOther!AgentName) Then
                    AgentName = MBMOther!AgentName
                Else
                    AgentName = ""
                End If
                
                If Len(MBMOther!PLNPayorPlan) Then
                    PLNPayorCode = MBMOther!PLNPayorPlan
                Else
                    PLNPayorCode = ""
                End If
                
                If Len(MBMOther!JoinDate) Then
                    Joindt = Format(MBMOther!JoinDate, "dd/MM/yyyy")
                Else
                    Joindt = ""
                End If
                
                
            End If
            MBMOther.Close
            Set MBMOther = Nothing
            strsqls = ""
           
               ' Display record in List View
               Cnt = 0
               Rcounter = Rcounter + 1
               row = row + 1
               xlApp.DisplayAlerts = False
                 ' Export record to excel format
                 
                    MBMName = MBM!MBMName
                    MBMNo = MBM!MBMNumber
                    MBMICNo = MBM!MBMIcBcPp
                    MBMPOLNO = MBM!MBMPolicyNo
                    If Trim(MBM!PLNCode) = "21" Or Trim(MBM!PLNCode) = "31" Or Trim(MBM!PLNCode) = "41" Then
                        PLNCode = "60"
                    ElseIf Trim(MBM!PLNCode) = "22" Or Trim(MBM!PLNCode) = "32" Or Trim(MBM!PLNCode) = "42" Then
                        PLNCode = "80"
                    ElseIf Trim(MBM!PLNCode) = "23" Or Trim(MBM!PLNCode) = "33" Or Trim(MBM!PLNCode) = "43" Then
                        PLNCode = "120"
                    ElseIf Trim(MBM!PLNCode) = "24" Or Trim(MBM!PLNCode) = "34" Or Trim(MBM!PLNCode) = "44" Then
                        PLNCode = "200"
                    ElseIf Trim(MBM!PLNCode) = "25" Or Trim(MBM!PLNCode) = "35" Or Trim(MBM!PLNCode) = "45" Then
                        PLNCode = "350"
                    ElseIf Trim(MBM!PLNCode) = "26" Or Trim(MBM!PLNCode) = "36" Or Trim(MBM!PLNCode) = "46" Then
                        PLNCode = "450"
                    Else
                        PLNCode = IIf(Len(MBM!PLNCode), MBM!PLNCode, "")
                    End If
                    Effdate = Format(MBM!MBMPayorEffDate, "dd/mm/yyyy")
                    ExpDate = Format(MBM!MBMPayorEXPDate, "dd/mm/yyyy")
                   'strsqls = "Select MBMBranch, MBMPOLSubNo1, MBMGroupCompany, MBMEmployeeNo, MBMDepartment from MBMOthers where MBMNumber = '" & Trim(MBM!MBMNumber) & "'"

                    strsqlPLN = "Select PLNCode,PLNClientPlan,PLNRBType,PLNDescription,PLNCardPrintMode from PLN where PLNCode = '" & MBM!PLNCode & "' And PAYCode = '" & MBM!PAYCode & "'"
                    PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If PLN.recordCount Then
                            If Trim(MBM!INSCode) = "G" Then
                                tempInsPLN = "INDIVIDUAL" & " " & PLN!PLNDescription
                            ElseIf Trim(MBM!INSCode) = "H" Then
                                tempInsPLN = "FAMILY" & " " & PLN!PLNDescription
                            End If
                            If PLN!PLNCardPrintMode <> "" Then
                                tmpPrintMode = PLN!PLNCardPrintMode
                            Else
                                tmpPrintMode = ""
                            End If
                        End If
                    PLN.Close
                    Set PLN = Nothing
                    
                    xlsheet.Cells(row, 1).Value = Rcounter
                    xlsheet.Cells(row, 2).Value = MBMName
                    'xlsheet.Cells(row, 3).Value = MBMName
                    'If IsNumeric(MBMICNo) = True Then
                    '    xlsheet.Cells(row, 4).Value = Mid(MBMICNo, 1, 6) & "-" & Mid(MBMICNo, 7, 2) & "-" & Mid(MBMICNo, 9, 4)
                    'Else
                    '    xlsheet.Cells(row, 4).Value = MBMICNo
                    'End If
                                        
                   ' If Len(POLSubNo1) Then
                   '     xlsheet.Cells(row, 4).Value = Trim(MBMPOLNO) & "*" & POLSubNo1
                   ' Else
                   '     xlsheet.Cells(row, 4).Value = "'" & MBMPOLNO
                   ' End If
                    
                    If Mid(MBMNo, 9, 1) = "*" Then
                        xlsheet.Cells(row, 3).Value = Mid(MBMNo, 1, 8) & "-" & "00"
                    ElseIf Mid(MBMNo, 10, 1) = "*" Then
                        xlsheet.Cells(row, 3).Value = Mid(MBMNo, 1, 9) & "-" & "00"
                    ElseIf Mid(MBMNo, 12, 1) = "*" Then
                        xlsheet.Cells(row, 3).Value = Mid(MBMNo, 1, 11) & "-" & "00"
                    Else
                        xlsheet.Cells(row, 3).Value = MBMNo & "-" & "00"
                    End If
                    'xlsheet.Cells(row, 6).Value = PLNCode
                    'xlsheet.Cells(row, 7).Value = Effdate
                    'xlsheet.Cells(row, 8).Value = ExpDate
                    If tmpPrintMode = "G" Then
                        xlsheet.Cells(row, 4).Value = GroupCompany
                    Else
                        xlsheet.Cells(row, 4).Value = Department
                    End If
                    'xlsheet.Cells(row, 10).Value = EmployeeNo
                    'xlsheet.Cells(row, 8).Value = MBMBranch
                        
                    'xlsheet.Cells(row, 13).Value = tempInsPLN
                    'xlsheet.Cells(row, 14).Value = tmpMBMPrevMemNo
                    'xlsheet.Cells(row, 6).Value = tmpEncoding
                    'xlsheet.Cells(row, 9).Value = Subsidary
                    'xlsheet.Cells(row, 17).Value = AgentCode
                    'xlsheet.Cells(row, 18).Value = AgentName
                    'xlsheet.Cells(row, 19).Value = PLNPayorCode
                    'xlsheet.Cells(row, 20).Value = Joindt
                    
        TotalPrincipal = TotalPrincipal + 1
        txtEXPTotalPrincipal = TotalPrincipal
        
        MBM.MoveNext
        Loop
        'End If
        xlBook.Save
        xlBook.SaveAs (Trim(txtSVRPath) & Trim(txtFilName & ".xls"))
        MsgBox "Export successful", vbOKOnly + vbInformation, DialogBoxTitle
        'Call cmdExpSearch_Click
        
        Set xlBook = Nothing
        xlApp.Quit
        Set xlApp = Nothing
        MBM.Close
        Set MBM = Nothing
        Frame2.Enabled = True
    End If
End Sub

Private Sub cboEXPCriteriaSel2_Click()
Select Case cboEXPCriteriaSel2.ListIndex
    Case 0
        Label5.Visible = True ' mbm no
        Label6.Visible = False 'from
        Label11.Visible = False 'to
        Label4.Visible = False 'brodx date
        Label9.Visible = False ' batch no
        Label10.Visible = False 'brodx from
        Label8.Visible = False 'to
        txtExpMemberNo2.Visible = True
        txtExpMemberNoFr2.Visible = False
        txtExpMemberNoTo2.Visible = False
        txtExpBordxDate2.Visible = False
        txtExpBatchNo2.Visible = False
        txtExpBordxDateFr2.Visible = False
        txtExpBordxDateTo2.Visible = False
        CboCLNPayor.Enabled = False
        txtExpMemberNoFr2 = ""
        txtExpMemberNoTo2 = ""
        txtExpMemberNo2 = ""
        txtExpBatchNo2 = ""
        CboCLNPayor = ""
        txtExpBordxDate2 = "__/__/____"
        txtExpBordxDateFr2 = "__/__/____"
        txtExpBordxDateTo2 = "__/__/____"
        txtEXPTotalPrincipal2 = 0
        txtEXPTotalSupplementary2 = 0
    
    Case 1
        Label5.Visible = False
        Label6.Visible = True
        Label11.Visible = True
        Label4.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        Label8.Visible = False
        txtExpMemberNoFr2.Visible = True
        txtExpMemberNoTo2.Visible = True
        txtExpMemberNo2.Visible = False
        txtExpBordxDate2.Visible = False
        txtExpBatchNo2.Visible = False
        txtExpBordxDateFr2.Visible = False
        txtExpBordxDateTo2.Visible = False
        CboCLNPayor.Enabled = False
        txtExpMemberNoTo2 = ""
        txtExpMemberNo2 = ""
        txtExpBatchNo2 = ""
        CboCLNPayor = ""
        txtExpBordxDate2 = "__/__/____"
        txtExpBordxDateFr2 = "__/__/____"
        txtExpBordxDateTo2 = "__/__/____"
        txtEXPTotalPrincipal2 = 0
        txtEXPTotalSupplementary2 = 0
        
    Case 2
        Label5.Visible = False
        Label6.Visible = False
        Label11.Visible = False
        Label4.Visible = True
        Label9.Visible = True
        Label10.Visible = False
        Label8.Visible = False
        txtExpBordxDate2.Visible = True
        txtExpBatchNo2.Visible = True
        txtExpMemberNo2.Visible = False
        txtExpMemberNoFr2.Visible = False
        txtExpMemberNoTo2.Visible = False
        txtExpBordxDateFr2.Visible = False
        txtExpBordxDateTo2.Visible = False
        CboCLNPayor.Enabled = True
        txtExpMemberNoTo2 = ""
        txtExpMemberNo2 = ""
        txtExpBatchNo2 = ""
        CboCLNPayor = ""
        txtExpBordxDate2 = "__/__/____"
        txtExpBordxDateFr2 = "__/__/____"
        txtExpBordxDateTo2 = "__/__/____"
        txtEXPTotalPrincipal2 = 0
        txtEXPTotalSupplementary2 = 0
    
End Select
End Sub

Private Sub cboEXPCriteriaSel2_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboEXPCriteriaSel2.Text, cboEXPCriteriaSel2.SelStart) + Chr(KeyAscii) + Mid(cboEXPCriteriaSel2.Text, cboEXPCriteriaSel2.SelStart + cboEXPCriteriaSel2.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboEXPCriteriaSel2.ListCount - 1
        
        If NewText = LCase(Left(cboEXPCriteriaSel2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboEXPCriteriaSel2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboEXPCriteriaSel2.Text = ValidValue
               cboEXPCriteriaSel2.SelStart = 0
               cboEXPCriteriaSel2.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub EXPPrincipalListing2()
Dim strsql As String
Dim strsqls As String
Dim sql As String
Dim MBMOther As New ADODB.Recordset
Dim MBM As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim ListMaster As ListItem
Dim isPrinted As Boolean
Dim rsSearch As New ADODB.Recordset
Dim ColumnItem As ListItem
Dim itemToPrint As Integer
Dim CoverQty As Integer
Dim TotalPrincipal As Integer
Dim strsqlCOV As String
Dim MBMCov As New ADODB.Recordset
Dim MBMCPLNCode  As String
Dim tmpMBMNumber As String
Dim MBMOthers As New ADODB.Recordset

TotalSupplementary = 0
TotalPrincipal = 0
MBMCPLNCode = ""
tmpMBMNumber = ""
strsqls = ""

Select Case cboEXPCriteriaSel2.ListIndex
    Case 0 'Card Export Criteria Selection for Single Member
        'lvwCRD.listitems.Clear
        strsql = " Select INSType, PAYCode, MBMBatchNo, MBMBordxDate, MBMGroupCompany, " & _
                 " MBMName, MBMNumber, MBMIcBcPp, MBMPolNumber, MBMEmployeeNo, " & _
                 " PLNType, MBMPolExpDate, MBMStatus, MBMDataStatus, GRPCoCode,MBMDepartment,OPFamilyStatus " & _
                 " From VIEW_CLN_CardExport Where MBMNumber = '" & Trim(txtExpMemberNo2) & "' And MBMDataStatus <> 'D'" 'AND " & _
                 '" PLNType >= '" & Trim(TxtCLNPlanCode1) & "' And PLNType <= '" & Trim(TxtCLNPlanCode2) & "'"
        
        If Len(Trim(TxtCLNPlanCode1)) Then
            strsql = strsql + " AND PLNType >= '" & Trim(TxtCLNPlanCode1) & "'"
        End If
        
        If Len(Trim(TxtCLNPlanCode2)) Then
            strsql = strsql + " AND PLNType <= '" & Trim(TxtCLNPlanCode2) & "'"
        End If
        
        
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Case 1 'Card Export Criteria Selection for Member's No Range
        'lvwCRD.listitems.Clear
        strsql = " SELECT INSType, PAYCode, MBMBatchNo, MBMBordxDate, MBMEmployeeNo, " & _
                 " MBMName, MBMNumber, MBMIcBcPp, MBMPolNumber, MBMGroupCompany, MBMDepartment,OPFamilyStatus," & _
                 " PLNType, MBMPolExpDate, MBMStatus, MBMDataStatus, GRPCoCode From VIEW_CLN_CardExport "
        strsql = strsql & " WHERE MBMNumber >= '" & txtExpMemberNoFr2 & "'"
        strsql = strsql & " AND MBMNumber <= '" & txtExpMemberNoTo2 & "'"
        strsql = strsql & " AND MBMDataStatus <> 'D' "
        strsql = strsql & " AND MBMStatus <> 'C' "
        'strsql = strsql & " AND PLNType >= '" & Trim(TxtCLNPlanCode1) & "' And PLNType <= '" & Trim(TxtCLNPlanCode2) & "'"
        
        If Len(Trim(TxtCLNPlanCode1)) Then
            strsql = strsql + " AND PLNType >= '" & Trim(TxtCLNPlanCode1) & "'"
        End If
        
        If Len(Trim(TxtCLNPlanCode2)) Then
            strsql = strsql + " AND PLNType <= '" & Trim(TxtCLNPlanCode2) & "'"
        End If
        
        strsql = strsql & " ORDER BY MBMDepartment, MBMName "
        
        MBM.Open strsql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    Case 2
        sDate = CStr(Month(txtExpBordxDate2)) & "/" & CStr(Day(txtExpBordxDate2)) & "/" & CStr(Year(txtExpBordxDate2))
        
        strsql = " SELECT INSType, PAYCode, MBMBatchNo, MBMBordxDate, GRPCoCode, " & _
                 " MBMName, MBMNumber, MBMIcBcPp, MBMPolNumber, MBMGroupCompany, MBMEmployeeNo,OPFamilyStatus, " & _
                 " PLNType, MBMPolExpDate, MBMStatus, MBMDataStatus,MBMDepartment from VIEW_CLN_CardExport "
        strsql = strsql & " WHERE MBMBordxDate = '" & sDate & "'"
        strsql = strsql & " AND MBMBatchNo = '" & txtExpBatchNo2 & "'"
        strsql = strsql & " AND PAYCode = '" & Mid(CboCLNPayor, 1, 2) & "'"
        strsql = strsql & " AND MBMDataStatus <> 'D' "
        strsql = strsql & " AND MBMStatus <> 'C' "
        'strsql = strsql & " AND PLNType >= '" & Trim(TxtCLNPlanCode1) & "' And PLNType <= '" & Trim(TxtCLNPlanCode2) & "'"
        
        If Len(Trim(TxtCLNPlanCode1)) Then
            strsql = strsql + " AND PLNType >= '" & Trim(TxtCLNPlanCode1) & "'"
        End If
        
        If Len(Trim(TxtCLNPlanCode2)) Then
            strsql = strsql + " AND PLNType <= '" & Trim(TxtCLNPlanCode2) & "'"
        End If
        
        strsql = strsql & " ORDER BY MBMDepartment, MBMName "
        
        MBM.Open strsql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
   End Select
       
            If MBM.EOF = True Then
                MsgBox "Record not found !", vbInformation + vbOKOnly
                Screen.MousePointer = Default
                Exit Sub
            End If
        Frame4.Enabled = False
        txtSVRPath = crdPath & "CardPrinting\"
        txtPath = "C:\CardPrinting\"
        If cboEXPCriteriaSel2.ListIndex = 0 Then
            txtFilName = txtExpMemberNo2
        ElseIf cboEXPCriteriaSel.ListIndex = 1 Then
            txtFilName = txtExpMemberNoFr2
        Else
            txtFilName = Mid(CboCLNPayor, 1, 2) & Mid(txtExpBordxDate2, 1, 2) & Mid(txtExpBordxDate2, 4, 2) & Mid(txtExpBordxDate2, 7, 4)
        End If
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set txtSVRfile = fso.CreateTextFile((Trim(txtSVRPath) & Trim(txtFilName & ".xls")), True)
        Set txtfile = fso.CreateTextFile((Trim(txtPath) & Trim(txtFilName & ".xls")), True)
        txtfile = (Trim(txtPath) & Trim(txtFilName & ".xls"))
        txtSVRfile = (Trim(txtSVRPath) & Trim(txtFilName & ".xls"))
            
        Set xlApp = CreateObject("Excel.application")
        Set xlBook = xlApp.Workbooks.Open(txtfile)
        Set xlsheet = xlBook.Sheets(1)
         xlsheet.Cells(1, 1).Value = CboCLNPayor
         xlsheet.Cells(2, 1).Value = "DATE :" & Mid(txtExpBordxDate2, 1, 2) & "/" & Mid(txtExpBordxDate2, 4, 2) & "/" & Mid(txtExpBordxDate2, 7, 4)
         xlsheet.Cells(3, 1).Value = "FILE NAME :" & txtFilName
         
         row = 5
         Rcounter = 0
         
         Do Until MBM.EOF
            
                strsqls = ""
                tmpEncoding = ""
                tmpMBMNumber = Split(MBM!MBMNumber, "*")(0)
                strsqls = "Select MBMNumber,MBMSecurityCode,MBMCardType,MBMClientType from MBMSCSecurity where MBMNumber = '" & Trim(tmpMBMNumber) & "'"
                MBMOther.Open strsqls, medixmasterconn, adOpenKeyset, adLockOptimistic
                    
                    If MBMOther.recordCount > 0 Then
                        tmpEncoding = MBMOther!MBMSecurityCode & MBMOther!MBMCardType & MBMOther!MBMClientType & "  hcsb.com.my                   />" & MBMOther!MBMNumber & "-00"
                        '6001mmed  hcsb.com.my                   />dcga8524393-00
                    End If
                MBMOther.Close
                Set MBMOther = Nothing

               ' Display record in List View
               Cnt = 0
               Rcounter = Rcounter + 1
               row = row + 1
               xlApp.DisplayAlerts = False
                 ' Export record to excel format
                    MBMNo = MBM!MBMNumber
                    If Len(MBM!MBMEmployeeNo) Then
                        EmployeeNo = Trim(MBM!MBMEmployeeNo)
                    End If
                    If Len(MBM!GRPCoCode) Then
                        GrpCode = Trim(MBM!GRPCoCode)
                    End If
                    If Len(MBM!MBMGroupCompany) Then
                        GroupCompany = Trim(MBM!MBMGroupCompany)
                    End If
                    ExpDate = Format(MBM!MBMPolExpDate, "dd/mm/yyyy")
                    xlsheet.Cells(row, 1).Value = Rcounter
                    xlsheet.Cells(row, 2).Value = GroupCompany
                    xlsheet.Cells(row, 3).Value = Trim(MBM!MBMName)
                    If MBM!MBMPolNumber <> "" Then
                        xlsheet.Cells(row, 4).Value = Trim(MBM!MBMPolNumber)
                    ElseIf MBM!MBMPolNumber = "" And EmployeeNo <> "" Then
                        xlsheet.Cells(row, 4).Value = EmployeeNo
                    End If
                    xlsheet.Cells(row, 5).Value = Trim(MBM!MBMNumber) & "-" & "00"
                    xlsheet.Cells(row, 6).Value = Format(MBM!MBMPolExpDate, "dd/mm/yyyy")
                    xlsheet.Cells(row, 7).Value = GrpCode
                    xlsheet.Cells(row, 8).Value = IIf(Len(MBM!MBMDepartment), MBM!MBMDepartment, "")
                    xlsheet.Cells(row, 9).Value = tmpEncoding
                    '==========================================================================
                    If Trim(MBM!InsType) = "F" Or Trim(MBM!InsType) = "H" Then
                        strsqlCOV = "Select MBMCNumber, MBMCName, MBMCCoverID, MBMCIcBcPpNo, MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(MBMNo) & "'"
                        MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
                    If MBMCov.recordCount Then
                    Do Until MBMCov.EOF
                        
                        Rcounter = Rcounter + 1
                        row = row + 1
                        MBMCName = Trim(MBMCov!MBMCName)
                        MBMCCoverID = MBMCov!MBMCCoverID
                        MBMCICNo = MBMCov!MBMCICBCPPNo

                        xlsheet.Cells(row, 1).Value = Rcounter
                        xlsheet.Cells(row, 2).Value = GroupCompany
                        xlsheet.Cells(row, 3).Value = MBMCName
                        If MBM!MBMPolNumber <> "" Then
                            xlsheet.Cells(row, 4).Value = Trim(MBM!MBMPolNumber)
                        ElseIf MBM!MBMPolNumber = "" And EmployeeNo <> "" Then
                            xlsheet.Cells(row, 4).Value = EmployeeNo
                        End If
                        xlsheet.Cells(row, 5).Value = Trim(MBM!MBMNumber) & "-" & MBMCCoverID
                        xlsheet.Cells(row, 6).Value = Format(MBM!MBMPolExpDate, "dd/mm/yyyy")
                        xlsheet.Cells(row, 7).Value = GrpCode
                        xlsheet.Cells(row, 8).Value = IIf(Len(MBM!MBMDepartment), MBM!MBMDepartment, "")
                        xlsheet.Cells(row, 9).Value = IIf(Len(MBM!OPFamilyStatus), MBM!OPFamilyStatus, "")
                        xlsheet.Cells(row, 10).Value = tmpEncoding
                        TotalSupplementary = TotalSupplementary + 1
                        txtEXPTotalSupplementary2 = TotalSupplementary
                    MBMCov.MoveNext
                    Loop
                    MBMCov.Close
                    Set MBMCov = Nothing
                    Else
                        Rcounter = Rcounter + 1
                        row = row + 1
                        If MBMCov.recordCount = 0 Then
                            MBMCov.Close
                            Set MBMCov = Nothing
                        End If
                    End If
                    

                    End If
                        
        
        TotalPrincipal = TotalPrincipal + 1
        txtEXPTotalPrincipal2 = TotalPrincipal
        
        MBM.MoveNext
        Loop
        xlBook.Save
        xlBook.SaveAs (Trim(txtSVRPath) & Trim(txtFilName & ".xls"))
        MsgBox "Export successful", vbOKOnly + vbInformation, DialogBoxTitle
        
        Set xlBook = Nothing
        xlApp.Quit
        Set xlApp = Nothing
                
        MBM.Close
        Set MBM = Nothing
        Frame4.Enabled = True

End Sub

Private Sub CboCLNPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCLNPayor.Text, CboCLNPayor.SelStart) + Chr(KeyAscii) + Mid(CboCLNPayor.Text, CboCLNPayor.SelStart + CboCLNPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCLNPayor.ListCount - 1
 
        If NewText = LCase(Left(CboCLNPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCLNPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCLNPayor.Text = ValidValue
                CboCLNPayor.SelStart = 0
                CboCLNPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub txtCardBordxDate_LostFocus()
    If IsDate(txtCardBordxDate) Then
    Else
        If txtCardBordxDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtCardBordxDate.SetFocus
        End If
    End If
End Sub

Private Sub txtCardBordxDateFr_LostFocus()
    If IsDate(txtCardBordxDateFr) Then
    Else
        If txtCardBordxDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtCardBordxDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtCardBordxDateTo_LostFocus()
    If IsDate(txtCardBordxDateTo) Then
    Else
        If txtCardBordxDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtCardBordxDateTo.SetFocus
        End If
    End If
End Sub

Private Sub txtDateSent_LostFocus()
    If IsDate(txtDateSent) Then
    Else
        If txtDateSent <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateSent.SetFocus
        End If
    End If
End Sub


Private Sub txtExpBordxDate2_LostFocus()
    If IsDate(txtExpBordxDate2) Then
    Else
        If txtExpBordxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtExpBordxDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtExpBordxDateFr2_LostFocus()
    If IsDate(txtExpBordxDateFr2) Then
    Else
        If txtExpBordxDateFr2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtExpBordxDateFr2.SetFocus
        End If
    End If
End Sub

Private Sub txtExpBordxDateTo2_LostFocus()
    If IsDate(txtExpBordxDateTo2) Then
    Else
        If txtExpBordxDateTo2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtExpBordxDateTo2.SetFocus
        End If
    End If
End Sub

Private Sub TxtCLNPlanCode1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(TxtCLNPlanCode1.Text, TxtCLNPlanCode1.SelStart) + Chr(KeyAscii) + Mid(TxtCLNPlanCode1.Text, TxtCLNPlanCode1.SelStart + TxtCLNPlanCode1.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To TxtCLNPlanCode1.ListCount - 1
 
        If NewText = LCase(Left(TxtCLNPlanCode1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = TxtCLNPlanCode1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                TxtCLNPlanCode1.Text = ValidValue
                TxtCLNPlanCode1.SelStart = 0
                TxtCLNPlanCode1.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub TxtCLNPlanCode2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(TxtCLNPlanCode2.Text, TxtCLNPlanCode2.SelStart) + Chr(KeyAscii) + Mid(TxtCLNPlanCode2.Text, TxtCLNPlanCode2.SelStart + TxtCLNPlanCode2.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To TxtCLNPlanCode2.ListCount - 1
 
        If NewText = LCase(Left(TxtCLNPlanCode2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = TxtCLNPlanCode2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                TxtCLNPlanCode2.Text = ValidValue
                TxtCLNPlanCode2.SelStart = 0
                TxtCLNPlanCode2.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub MemberListing()
Dim strsql As String
Dim strsqls As String
Dim sql As String
Dim MBMOther As New ADODB.Recordset
Dim MBM As New ADODB.Recordset
Dim ListMaster As ListItem
Dim rsSearch As New ADODB.Recordset
Dim ColumnItem As ListItem
Dim CoverQty As Integer
Dim TotalPrincipal As Integer
Dim strsqlCOV As String
Dim MBMCov As New ADODB.Recordset
Dim MBMCPLNCode  As String
Dim Effdate As String
Dim listrecord As Integer
Dim listloop1 As Integer
Dim TempCount As Integer
Dim TempPLNCode As String
Dim tmpSuppEffDate As String
Dim tmpSuppExpDate As String
Dim Subsidary As String

TotalSupplementary = 0
TotalPrincipal = 0
MBMCPLNCode = ""
'TempCount = 0
        listrecord = lstMemberListing.ListItems.Count
        If listrecord <> 0 Then
            
            For listloop1 = 1 To listrecord
                Call MBMNoCheck(listloop1)
            Next listloop1
        End If
        If Check = True Then
            MsgBox "Membership is on the list!", vbOKOnly + vbCritical
        Else
            If Len(txtExpMemberNo) = "8" Then
                strsql = "Select INSCode,PAYCode,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEFFDate, MBMPayorExpDate,MBMStatus,MBMDataStatus from MBMCrossReference where substring(MBMNumber,1,8) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,10,2) = '01'  "
            ElseIf Len(txtExpMemberNo) = "9" Then
                strsql = "Select INSCode,PAYCode,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEFFDate, MBMPayorExpDate,MBMStatus,MBMDataStatus from MBMCrossReference where substring(MBMNumber,1,9) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,11,2) = '01'  "
            ElseIf Len(txtExpMemberNo) = "11" Then
                strsql = "Select INSCode,PAYCode,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEFFDate, MBMPayorExpDate,MBMStatus,MBMDataStatus from MBMCrossReference where substring(MBMNumber,1,11) = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D' And substring(MBMNumber,13,2) = '01'  "
            Else
                strsql = "Select INSCode,PAYCode,MBMName,MBMNumber,MBMIcBcPp,MBMPolicyNo,PLNCode,MBMPayorEFFDate, MBMPayorExpDate,MBMStatus,MBMDataStatus from MBMCrossReference where MBMNumber = '" & Trim(txtExpMemberNo) & "' And MBMDataStatus <> 'D'  "
            End If
            MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
                    
            If MBM.recordCount = 0 Then
                    MsgBox "Record not found !", vbInformation + vbOKOnly
                    Screen.MousePointer = Default
                    Exit Sub
            End If
            
            'Else
                Do Until MBM.EOF
                    'TempCount = TempCount + 1
                    MBMName = MBM!MBMName
                    MBMNo = MBM!MBMNumber
                    MBMICNo = MBM!MBMIcBcPp
                    MBMPOLNO = MBM!MBMPolicyNo
                    TempPLNCode = MBM!PLNCode
                    Effdate = Format(MBM!MBMPayorEffDate, "dd/mm/yyyy")
                    ExpDate = Format(MBM!MBMPayorEXPDate, "dd/mm/yyyy")
                    
                    strsqls = "Select MBMBranch, MBMGroupCompany, MBMEmployeeNo, MBMDepartment,Subsidiary from MBMOthers  where MBMNumber = '" & Trim(MBM!MBMNumber) & "'"
                    MBMOther.Open strsqls, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                                                        
                    If MBMOther.recordCount Then
                        If Len(MBMOther!MBMGroupCompany) Then
                            GroupCompany = MBMOther!MBMGroupCompany
                        Else
                            GroupCompany = ""
                        End If
                        If Len(MBMOther!MBMEmployeeNo) Then
                            EmployeeNo = MBMOther!MBMEmployeeNo
                        Else
                            EmployeeNo = ""
                        End If
                        If Len(MBMOther!MBMDepartment) Then
                            Department = MBMOther!MBMDepartment
                        Else
                            Department = ""
                        End If
                        If Len(MBMOther!MBMBranch) Then
                            MBMBranch = MBMOther!MBMBranch
                        Else
                            MBMBranch = ""
                        End If
                        If Len(MBMOther!Subsidiary) Then
                            Subsidary = MBMOther!Subsidiary
                        Else
                            Subsidary = ""
                        End If
                    End If
                    MBMOther.Close
                    Set MBMOther = Nothing
                    
                    Set ListMaster = lstMemberListing.ListItems.Add(, , MBMName)
                    
                    If Len(MBMICNo) Then
                        ListMaster.SubItems(1) = MBMICNo
                    End If
                    If Len(MBMPOLNO) Then
                        ListMaster.SubItems(2) = MBMPOLNO
                    End If
                    If Mid(MBMNo, 9, 1) = "*" Then
                        ListMaster.SubItems(3) = Mid(MBMNo, 1, 8) & "-" & "00"
                    ElseIf Mid(MBMNo, 10, 1) = "*" Then
                        ListMaster.SubItems(3) = Mid(MBMNo, 1, 9) & "-" & "00"
                    Else
                        ListMaster.SubItems(3) = MBMNo & "-" & "00"
                    End If
                    If Len(TempPLNCode) Then
                        ListMaster.SubItems(4) = TempPLNCode
                    End If
                    If Len(Effdate) Then
                        ListMaster.SubItems(5) = Effdate
                    End If
                    
                    If Len(ExpDate) Then
                        ListMaster.SubItems(6) = ExpDate
                    End If
                    If Len(GroupCompany) Then
                        ListMaster.SubItems(7) = GroupCompany
                    End If
                    
                    If Len(Department) Then
                        ListMaster.SubItems(8) = Department
                    End If
                    If Len(EmployeeNo) Then
                        ListMaster.SubItems(9) = EmployeeNo
                    End If
                    If Len(MBMBranch) Then
                        ListMaster.SubItems(10) = MBMBranch
                    End If
                    
                    'add subsidary for cmc
                     If Len(Subsidary) Then
                           ListMaster.SubItems(11) = Subsidary
                     End If
            
                    
                    If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
                        Subsidary = ""
                        strsqlCOV = "Select MBMCNumber,MBMCName,MBMCCoverID,MBMCIcBcPpNo,MBMCPLNCode,MBMCEffDate,MBMCExpDate, MBMOthers.Subsidiary as Subsidary from MBMCoveredPersons inner join MBMOthers on MBMOthers.MBMnumber = MBMCoveredPersons.MBMCNumber where MBMCNumber = '" & Trim(MBMNo) & "'"
                        MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
                        If MBMCov.recordCount Then
                    
                            Do Until MBMCov.EOF
                                Subsidary = MBMCov!Subsidary
                                MBMCName = MBMCov!MBMCName
                                MBMCCoverID = MBMCov!MBMCCoverID
                                MBMCICNo = MBMCov!MBMCICBCPPNo
                                If MBMCov!MBMCEffDate <> "" Then
                                    tmpSuppEffDate = MBMCov!MBMCEffDate
                                Else
                                    tmpSuppEffDate = ""
                                End If
                                If MBMCov!MBMCExpDate <> "" Then
                                    tmpSuppExpDate = MBMCov!MBMCExpDate
                                Else
                                    tmpSuppExpDate = ""
                                End If
                                
                                Set ListMaster = lstMemberListing.ListItems.Add(, , MBMCName)
                                
                                If Len(MBMCICNo) Then
                                    ListMaster.SubItems(1) = MBMCICNo
                                End If
                                If Len(MBMPOLNO) Then
                                    ListMaster.SubItems(2) = MBMPOLNO
                                End If
                                    
                                If Mid(MBMNo, 9, 1) = "*" Then
                                    ListMaster.SubItems(3) = Mid(MBMNo, 1, 8) & "-" & MBMCCoverID
                                ElseIf Mid(MBMNo, 10, 1) = "*" Then
                                    ListMaster.SubItems(3) = Mid(MBMNo, 1, 9) & "-" & MBMCCoverID
                                Else
                                    ListMaster.SubItems(3) = MBMNo & "-" & MBMCCoverID
                                End If
                                
                                If Len(TempPLNCode) Then
                                    ListMaster.SubItems(4) = TempPLNCode
                                End If
                                
                                If tmpSuppEffDate = "" Then
                                    ListMaster.SubItems(5) = Effdate
                                Else
                                    ListMaster.SubItems(5) = tmpSuppEffDate
                                End If
                                If tmpSuppExpDate = "" Then
                                    ListMaster.SubItems(6) = ExpDate
                                Else
                                    ListMaster.SubItems(6) = tmpSuppExpDate
                                End If
                                If Len(GroupCompany) Then
                                    ListMaster.SubItems(7) = GroupCompany
                                End If
                                
                                If Len(Department) Then
                                    ListMaster.SubItems(8) = Department
                                End If
                                If Len(EmployeeNo) Then
                                    ListMaster.SubItems(9) = EmployeeNo
                                End If
                                If Len(MBMBranch) Then
                                    ListMaster.SubItems(10) = MBMBranch
                                End If
                                'add subsidary for cmc
                                If Len(Subsidary) Then
                                    ListMaster.SubItems(11) = Subsidary
                                End If
            
                        MBMCov.MoveNext
                        Loop
                        MBMCov.Close
                        Set MBMCov = Nothing
                    Else
                        MBMCov.Close
                        Set MBMCov = Nothing
                    End If
        
                End If
                
                MBM.MoveNext
                Loop
                'MBM.Close
                'Set MBM = Nothing
            End If
        'End If
End Sub

Private Function MBMNoCheck(listloop1 As Integer)
    
    If Len(txtExpMemberNo) = "8" Then
        If Trim(txtExpMemberNo) = Mid(Trim(lstMemberListing.ListItems(listloop1).SubItems(3)), 1, 8) Then
            Check = True
        Else
            Check = False
        End If
    ElseIf Len(txtExpMemberNo) = "9" Then
        If Trim(txtExpMemberNo) = Mid(Trim(lstMemberListing.ListItems(listloop1).SubItems(3)), 1, 9) Then
            Check = True
        Else
            Check = False
        End If
    Else
        If Trim(txtExpMemberNo) = ChkStr(Trim(lstMemberListing.ListItems(listloop1).SubItems(3))) Then
            Check = True
        Else
            Check = False
        End If
    End If
End Function

Private Sub InsertLealet()
Dim strsqlPlutoBor As String
Dim PlutoBor As New ADODB.Recordset
Dim sqlbor As String
Dim BOR As New ADODB.Recordset
Dim sqlinsert As String
Dim PLUTOStatus As Boolean
Dim tmpMBMName As String
Dim tmpMBMAddress1 As String
Dim tmpMBMAddress2 As String
Dim tmpMBMAddress3 As String
Dim tmpMBMCName As String
Dim tmpMBMGroupCompany As String
Dim tmpPrevComp As String
Dim tmpPayorRemarks As String
Dim tmpAllergy As String
Dim tmpPLNRBType As String
Dim tmpPLNDescription As String
Dim tmpSuppLimitStatus As String
Dim tmpMBMCAvailableLimit As Double
Dim tmpMBMAvailableLimit As Double
Dim tmpDepartment As String
Dim tmpBRHCode As String
Dim tmpOccupation As String
Dim tmpMBMCoverID As String

    PLUTOStatus = False
    strsqlPlutoBor = ""
    
    strsqlPlutoBor = "Select MBMNumber from Tmp_Leaflet where MBMNumber = '" & Trim(tmpMemberNo) & "'"
    PlutoBor.Open strsqlPlutoBor, medixmasterconnPluto, adOpenForwardOnly, adLockReadOnly
    
    Do While Not PlutoBor.EOF
        PLUTOStatus = True
        PlutoBor.MoveNext
    Loop
    PlutoBor.Close
    Set PlutoBor = Nothing

    If PLUTOStatus = False Then
        If OPTNorLeaflet.Value = True Then
            sqlbor = "Select * from VIEW_Leaflet where MBMNumber = '" & Trim(tmpMemberNo) & "'"
            BOR.Open sqlbor, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        ElseIf OPTSpeLeaflet.Value = True Then
            sqlbor = "Select * from View_Leaflet_Clinical where MBMNumber = '" & Trim(tmpMemberNo) & "'"
            BOR.Open sqlbor, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        End If
        Do While Not BOR.EOF
            TempMBMName = Replace(Trim(BOR!MBMName), "'", "''")
            tmpMBMAddress1 = Replace(Trim(BOR!MBMAddress1), "'", "''")
            tmpMBMAddress2 = Replace(Trim(BOR!MBMAddress2), "'", "''")
            
            If Len(BOR!MBMAddress3) Then
                tmpMBMAddress3 = Replace(Trim(BOR!MBMAddress3), "'", "''")
            End If
            If Len(BOR!MBMCName) Then
                tmpMBMCName = Replace(Trim(BOR!MBMCName), "'", "''")
            End If
            
            If Len(BOR!MBMGroupCompany) Then
                tmpMBMGroupCompany = Replace(Trim(BOR!MBMGroupCompany), "'", "''")
            End If
            
            If Len(BOR!MBMPrevInsCom) Then
                tmpPrevComp = Replace(Trim(BOR!MBMPrevInsCom), "'", "''")
            Else
                tmpPrevComp = ""
            End If
            
            If Len(BOR!MBMPAYRemarks) Then
                tmpPayorRemarks = Replace(Trim(BOR!MBMPAYRemarks), "'", "''")
            Else
                tmpPayorRemarks = ""
            End If
            
            If Len(BOR!MBMAllergic) Then
                tmpAllergy = Replace(Trim(BOR!MBMAllergic), "'", "''")
            End If
            If Len(BOR!MBMDepartment) Then
                tmpDepartment = Replace(Trim(BOR!MBMDepartment), "'", "''")
            Else
                tmpDepartment = ""
            End If
            
            If Len(BOR!BRNCode) Then
                tmpBRHCode = Replace(Trim(BOR!BRNCode), "'", "''")
            Else
                tmpBRHCode = ""
            End If
            
            If Len(BOR!MBMOccupation) Then
                tmpOccupation = Replace(Trim(BOR!MBMOccupation), "'", "''")
            Else
                tmpOccupation = ""
            End If
            
            tmpMBMAvailableLimit = IIf(IsNumeric(BOR!MBMAvailableLimit), BOR!MBMAvailableLimit, "0")
            
            tmpMBMCAvailableLimit = IIf(IsNumeric(BOR!MBMCAvailableLimit), BOR!MBMCAvailableLimit, "0")
            
            If BOR!MBMCRELCode <> "P" Then
                tmpMBMCoverID = BOR!MBMCCoverID
            Else
                tmpMBMCoverID = "00"
            End If
            sqlinsert = ""
            sqlinsert = "Insert into Tmp_Leaflet(MBMNumber , MBMPolicyNo ,MBMName ,MBMAddress1  ,MBMAddress2 , " & _
                        "MBMAddress3 ,MBMCName ,PAYCompanyName,MBMAgentCode,HLTCode, " & _
                        "MBMEmployeeNo,MBMDepartment,MBMPostCode,STACode,CTYCode,MBMDataStatus, " & _
                        "MBMStatus,BRNCode ,INSCode,MBMOccupation,MBMPrevInsCom,MBMPayRemarks,MBMAllergic, " & _
                        "PLNRBType, PLNDescription, SuppLimitStatus,MBMAvailableLimit,MBMCAvailableLimit,CoverID,MBMIcBcPp  ) " & _
                        "Values('" & BOR!MBMNumber & "','" & BOR!MBMPolicyNo & "', '" & TempMBMName & "',  '" & tmpMBMAddress1 & "',  '" & tmpMBMAddress2 & "', " & _
                        "'" & tmpMBMAddress3 & "', '" & tmpMBMCName & "', '" & tmpMBMGroupCompany & "', '" & BOR!MBMAgentCode & "', '" & BOR!HLTCode & "', " & _
                        "'" & BOR!MBMEmployeeNo & "', '" & tmpDepartment & "', '" & BOR!MBMPostCode & "', '" & BOR!STACode & "', '" & BOR!CTYCode & "',  '" & BOR!MBMDataStatus & "', " & _
                        "'" & BOR!MBMStatus & "', '" & tmpBRHCode & "', '" & BOR!INSCode & "', '" & tmpOccupation & "','" & tmpPrevComp & "', '" & tmpPayorRemarks & "', '" & tmpAllergy & "', " & _
                        "'" & BOR!PLNRBType & "', '" & BOR!PLNDescription & "', '" & BOR!SuppLimitStatus & "'," & tmpMBMAvailableLimit & ", " & tmpMBMCAvailableLimit & ", '" & tmpMBMCoverID & "', '" & BOR!MBMIcBcPp & "') "
                        
            medixmasterconnPluto.Execute sqlinsert
        BOR.MoveNext
        Loop
    BOR.Close
    Set BOR = Nothing
    
    End If

End Sub







