VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseAdjustment 
   Caption         =   "Case Adjustment"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame20 
      Height          =   495
      Index           =   10
      Left            =   120
      TabIndex        =   164
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox cboClaimability 
         Height          =   315
         Left            =   8640
         TabIndex        =   6
         Top             =   120
         Width           =   2895
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         Top             =   160
         Width           =   690
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Prev"
         Height          =   285
         Left            =   3240
         TabIndex        =   2
         Top             =   160
         Width           =   690
      End
      Begin VB.ComboBox cboStatus1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   140
         Width           =   1335
      End
      Begin VB.TextBox txtCaseNo 
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
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         Top             =   160
         Width           =   1515
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4680
         TabIndex        =   4
         Top             =   160
         Width           =   690
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Height          =   300
         Left            =   2640
         TabIndex        =   1
         Top             =   160
         Width           =   570
      End
      Begin VB.Label Label65 
         Caption         =   "Claimability"
         Height          =   255
         Index           =   108
         Left            =   7800
         TabIndex        =   294
         Top             =   165
         Width           =   855
      End
      Begin VB.Label Label65 
         Caption         =   "Case Status"
         Height          =   255
         Index           =   107
         Left            =   5400
         TabIndex        =   293
         Top             =   165
         Width           =   855
      End
      Begin VB.Label Label65 
         Caption         =   "Cas Number"
         Height          =   255
         Index           =   37
         Left            =   120
         TabIndex        =   216
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.Frame Frame20 
      Height          =   720
      Index           =   9
      Left            =   120
      TabIndex        =   162
      Top             =   680
      Width           =   11655
      Begin VB.TextBox txtRenewal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7500
         TabIndex        =   12
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   9
         Top             =   130
         Width           =   1575
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1050
         MaxLength       =   18
         TabIndex        =   7
         Top             =   130
         Width           =   2000
      End
      Begin MSMask.MaskEdBox txtMBMEffDate 
         Height          =   285
         Left            =   8760
         TabIndex        =   14
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPayCode 
         Enabled         =   0   'False
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
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   10
         Top             =   360
         Width           =   555
      End
      Begin VB.TextBox txtMBMName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   8
         Top             =   380
         Width           =   2000
      End
      Begin VB.TextBox txtTakeOver 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   11
         Top             =   380
         Width           =   375
      End
      Begin MSMask.MaskEdBox txtMBMExpiryDate 
         Height          =   285
         Left            =   8760
         TabIndex        =   15
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboPending 
         Height          =   315
         Left            =   6300
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label65 
         Caption         =   "Exp Date"
         Height          =   255
         Index           =   47
         Left            =   8040
         TabIndex        =   226
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Eff Date"
         Height          =   255
         Index           =   46
         Left            =   8040
         TabIndex        =   225
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Renewal"
         Height          =   255
         Index           =   45
         Left            =   6720
         TabIndex        =   224
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Pending"
         Height          =   255
         Index           =   44
         Left            =   5520
         TabIndex        =   223
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Pol Status"
         Height          =   255
         Index           =   43
         Left            =   5520
         TabIndex        =   222
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "T/Over"
         Height          =   255
         Index           =   42
         Left            =   4440
         TabIndex        =   221
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Payor"
         Height          =   255
         Index           =   41
         Left            =   3120
         TabIndex        =   220
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "IC Num"
         Height          =   255
         Index           =   40
         Left            =   3120
         TabIndex        =   219
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Mem Name"
         Height          =   255
         Index           =   39
         Left            =   120
         TabIndex        =   218
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label65 
         Caption         =   "Mem Num"
         Height          =   255
         Index           =   38
         Left            =   120
         TabIndex        =   217
         Top             =   180
         Width           =   975
      End
      Begin VB.Label txtPolStatus 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6300
         TabIndex        =   187
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblExpired 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Expired"
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
         Height          =   285
         Left            =   10005
         TabIndex        =   176
         Top             =   405
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblWaitingPeriod 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Waiting Period"
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
         Height          =   285
         Left            =   10005
         TabIndex        =   175
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6100
      Left            =   120
      TabIndex        =   165
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10769
      _Version        =   393216
      Tabs            =   10
      Tab             =   1
      TabsPerRow      =   10
      TabHeight       =   529
      TabCaption(0)   =   "General Info"
      TabPicture(0)   =   "frmCaseAdjustment.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstMember"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Case Details"
      TabPicture(1)   =   "frmCaseAdjustment.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame20(4)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Case Others"
      TabPicture(2)   =   "frmCaseAdjustment.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "TakeOver Details"
      TabPicture(3)   =   "frmCaseAdjustment.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Benefit"
      TabPicture(4)   =   "frmCaseAdjustment.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtUpgPLNCode"
      Tab(4).Control(1)=   "lstBenefits"
      Tab(4).Control(2)=   "lstNewBenefitDetails"
      Tab(4).Control(3)=   "txtDOC"
      Tab(4).Control(4)=   "Label65(95)"
      Tab(4).Control(5)=   "Label65(94)"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Warranty Exclusion"
      TabPicture(5)   =   "frmCaseAdjustment.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtPrincipalExc"
      Tab(5).Control(1)=   "txtSuppExc"
      Tab(5).Control(2)=   "cboSuppExc"
      Tab(5).Control(3)=   "Label65(111)"
      Tab(5).Control(4)=   "Label65(110)"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Case Mgmt Notes"
      TabPicture(6)   =   "frmCaseAdjustment.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdIntCASRem"
      Tab(6).Control(1)=   "cmdPage"
      Tab(6).Control(2)=   "Frame12"
      Tab(6).Control(3)=   "Frame11"
      Tab(6).Control(4)=   "txtCASRemark"
      Tab(6).Control(5)=   "txtCASRemarkTwo"
      Tab(6).Control(6)=   "txtIntCASRemark"
      Tab(6).Control(7)=   "lblMemo"
      Tab(6).Control(8)=   "lblPage"
      Tab(6).Control(9)=   "Label65(96)"
      Tab(6).Control(10)=   "Label65(97)"
      Tab(6).Control(11)=   "Label65(98)"
      Tab(6).ControlCount=   12
      TabCaption(7)   =   "LG"
      TabPicture(7)   =   "frmCaseAdjustment.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame4"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Audit Review"
      TabPicture(8)   =   "frmCaseAdjustment.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame5"
      Tab(8).Control(1)=   "Frame14"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Worksheet"
      TabPicture(9)   =   "frmCaseAdjustment.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "lstWorksheetListing"
      Tab(9).Control(1)=   "Frame20(12)"
      Tab(9).ControlCount=   2
      Begin VB.Frame Frame20 
         Height          =   700
         Index           =   12
         Left            =   -74880
         TabIndex        =   382
         Top             =   5280
         Width           =   11415
         Begin VB.CommandButton cmdCLMAppUpdate 
            Caption         =   "Update App"
            Enabled         =   0   'False
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
            Left            =   8280
            TabIndex        =   331
            Top             =   240
            Width           =   1080
         End
         Begin VB.ComboBox cboAppStatus 
            Height          =   315
            Left            =   5880
            TabIndex        =   330
            Top             =   240
            Width           =   1815
         End
         Begin MSMask.MaskEdBox txtAppDate 
            Height          =   300
            Left            =   3240
            TabIndex        =   329
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
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
         Begin VB.Label lblClaimNo 
            Height          =   255
            Left            =   840
            TabIndex        =   386
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "Claim No"
            Height          =   255
            Index           =   115
            Left            =   120
            TabIndex        =   385
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Approval Date"
            Height          =   255
            Index           =   113
            Left            =   2160
            TabIndex        =   384
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "Approval Status"
            Height          =   255
            Index           =   112
            Left            =   4680
            TabIndex        =   383
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lstWorksheetListing 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   381
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8493
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ClaimNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Act Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "App Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Pay2HOS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pay2MBM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "BordxNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "BordxDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "BordxAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Payor App Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Payor App Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdIntCASRem 
         Caption         =   "Memo"
         Height          =   375
         Left            =   -64920
         TabIndex        =   368
         Top             =   2880
         Width           =   855
      End
      Begin VB.Frame Frame8 
         Height          =   615
         Left            =   -74880
         TabIndex        =   357
         Top             =   360
         Width           =   11415
         Begin VB.ComboBox cboIntExtTO 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5520
            TabIndex        =   142
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtLinkCASNo 
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
            Left            =   10080
            MaxLength       =   15
            TabIndex        =   144
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtCASTOPlan 
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
            Left            =   8040
            MaxLength       =   4
            TabIndex        =   143
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtCASTakeOverAmt 
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
            Left            =   2880
            MaxLength       =   15
            TabIndex        =   141
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cboCASTakeOveInd 
            Height          =   315
            Left            =   840
            TabIndex        =   140
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "Link Case No"
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
            Index           =   92
            Left            =   9120
            TabIndex        =   373
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "T/O Plan"
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
            Index           =   66
            Left            =   7320
            TabIndex        =   372
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label65 
            Caption         =   "T/O Amt"
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
            Index           =   61
            Left            =   2160
            TabIndex        =   371
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label65 
            Caption         =   "T/O Ind"
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
            Index           =   60
            Left            =   120
            TabIndex        =   370
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label65 
            Caption         =   "Int/Ext TakeOver"
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
            Index           =   109
            Left            =   4200
            TabIndex        =   365
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtUpgPLNCode 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67440
         TabIndex        =   353
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtPrincipalExc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   -74880
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   352
         Top             =   780
         Width           =   11415
      End
      Begin VB.TextBox txtSuppExc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   -74880
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   351
         Top             =   3660
         Width           =   11415
      End
      Begin VB.ComboBox cboSuppExc 
         Height          =   315
         Left            =   -73680
         TabIndex        =   350
         Top             =   3300
         Width           =   5775
      End
      Begin VB.CommandButton cmdPage 
         Caption         =   "Next"
         Height          =   375
         Left            =   -64920
         TabIndex        =   348
         Top             =   5280
         Width           =   855
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame12"
         Height          =   4935
         Left            =   -66720
         TabIndex        =   343
         Top             =   720
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame11"
         Height          =   4935
         Left            =   -72960
         TabIndex        =   342
         Top             =   720
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Height          =   5115
         Left            =   -74880
         TabIndex        =   338
         Top             =   360
         Width           =   11415
         Begin MSComctlLib.ListView lstFollowUp 
            Height          =   2505
            Left            =   120
            TabIndex        =   339
            Top             =   2520
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   4419
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "FollowUp No."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "LG Amt"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "IssuedBy"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Issued Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Treatment Date"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lstLGListing 
            Height          =   2145
            Left            =   120
            TabIndex        =   340
            Top             =   360
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   3784
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LG No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "LG Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "LG Amt"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Cov ID"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Patient"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "DOA"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Hospital"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Issued By"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Remark"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label65 
            Caption         =   "Double Click to reprint LG"
            Height          =   255
            Index           =   99
            Left            =   120
            TabIndex        =   341
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   295
         Top             =   360
         Width           =   11415
         Begin VB.TextBox txtAuditReview 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   4920
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   305
            Top             =   240
            Width           =   6375
         End
         Begin VB.TextBox txtINVNo 
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   296
            Top             =   240
            Width           =   2010
         End
         Begin MSMask.MaskEdBox txtINVDate 
            Height          =   300
            Left            =   1560
            TabIndex        =   297
            Top             =   480
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   529
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
         Begin VB.TextBox txtINVAmt 
            Height          =   285
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   298
            Top             =   735
            Width           =   2010
         End
         Begin VB.TextBox txtDISRec 
            Height          =   285
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   299
            Top             =   990
            Width           =   2010
         End
         Begin MSMask.MaskEdBox txtDLS 
            Height          =   300
            Left            =   1560
            TabIndex        =   300
            Top             =   1245
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   529
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
         Begin MSMask.MaskEdBox txtDLR 
            Height          =   300
            Left            =   1560
            TabIndex        =   301
            Top             =   1515
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   529
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
         Begin VB.ComboBox CboARStatus 
            Height          =   315
            Left            =   1560
            TabIndex        =   302
            Top             =   1785
            Width           =   2010
         End
         Begin MSMask.MaskEdBox TxtDateOpen 
            Height          =   300
            Left            =   1560
            TabIndex        =   303
            Top             =   2040
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   529
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
         Begin MSMask.MaskEdBox TxtDRTC 
            Height          =   300
            Left            =   1560
            TabIndex        =   304
            Top             =   2295
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   529
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
         Begin VB.TextBox TxtARAnswer 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   4920
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   306
            Top             =   1035
            Width           =   6375
         End
         Begin VB.TextBox txtARRemarks 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   4920
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   307
            Top             =   1785
            Width           =   6375
         End
         Begin VB.TextBox TxtPVNo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   308
            Top             =   2865
            Width           =   6375
         End
         Begin VB.TextBox TxtPVDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   309
            Top             =   3120
            Width           =   6375
         End
         Begin VB.TextBox TxtAmtPaid 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   310
            Top             =   3360
            Width           =   6375
         End
         Begin VB.Label Label65 
            Caption         =   "Inv Number"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   325
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label65 
            Caption         =   "Inv date"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   324
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "Inv Amt"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   323
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label65 
            Caption         =   "Disc Received"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   322
            Top             =   1030
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "Date Letter Sent"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   321
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "Date Letter Reply"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   320
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label65 
            Caption         =   "A/R Status"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   319
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Date Open"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   318
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label65 
            Caption         =   "Date Returned to Claim"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   317
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label65 
            Caption         =   "Query"
            Height          =   255
            Index           =   9
            Left            =   3960
            TabIndex        =   316
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Answer"
            Height          =   255
            Index           =   10
            Left            =   3960
            TabIndex        =   315
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Remarks"
            Height          =   255
            Index           =   11
            Left            =   3960
            TabIndex        =   314
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "PV Number"
            Height          =   255
            Index           =   12
            Left            =   3960
            TabIndex        =   313
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "PV Date"
            Height          =   255
            Index           =   13
            Left            =   3960
            TabIndex        =   312
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Amt Paid"
            Height          =   255
            Index           =   14
            Left            =   3960
            TabIndex        =   311
            Top             =   3360
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Height          =   5685
         Left            =   -74880
         TabIndex        =   188
         Top             =   360
         Width           =   11415
         Begin VB.Frame Frame20 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1275
            Index           =   0
            Left            =   120
            TabIndex        =   190
            Top             =   120
            Width           =   8175
            Begin VB.CommandButton cmdInvList 
               Caption         =   "Search"
               Height          =   255
               Left            =   7200
               TabIndex        =   146
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox txtPROCINVDone1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   65
               Top             =   405
               Width           =   5535
            End
            Begin VB.TextBox txtPROCINVDone2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   67
               Top             =   660
               Width           =   5535
            End
            Begin VB.TextBox txtPROCINVDone3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   69
               Top             =   915
               Width           =   5535
            End
            Begin VB.TextBox txtINVCode1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5640
               MaxLength       =   150
               TabIndex        =   66
               Top             =   405
               Width           =   1455
            End
            Begin VB.TextBox txtINVCode2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5640
               MaxLength       =   150
               TabIndex        =   68
               Top             =   660
               Width           =   1455
            End
            Begin VB.TextBox txtINVCode3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5640
               MaxLength       =   150
               TabIndex        =   70
               Top             =   915
               Width           =   1455
            End
            Begin VB.Label Label65 
               Caption         =   "Investigation"
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
               Index           =   105
               Left            =   120
               TabIndex        =   285
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.Frame Frame20 
            Height          =   1275
            Index           =   1
            Left            =   8350
            TabIndex        =   189
            Top             =   120
            Width           =   3015
            Begin VB.TextBox txtMEDSURDone1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   122
               Top             =   405
               Width           =   2895
            End
            Begin VB.TextBox txtMEDSURDone2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   123
               Top             =   660
               Width           =   2895
            End
            Begin VB.TextBox txtMEDSURDone3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   124
               Top             =   915
               Width           =   2895
            End
            Begin VB.Label Label65 
               Caption         =   "Medical Management"
               Height          =   255
               Index           =   83
               Left            =   120
               TabIndex        =   268
               Top             =   120
               Width           =   1695
            End
         End
         Begin VB.Frame Frame20 
            Height          =   2250
            Index           =   2
            Left            =   8350
            TabIndex        =   191
            Top             =   1320
            Width           =   3015
            Begin VB.TextBox txtATTDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   840
               MaxLength       =   100
               TabIndex        =   125
               Top             =   165
               Width           =   2080
            End
            Begin VB.TextBox txtATTDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   840
               MaxLength       =   30
               TabIndex        =   126
               Top             =   420
               Width           =   2080
            End
            Begin VB.TextBox txtATTDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   840
               MaxLength       =   30
               TabIndex        =   127
               Top             =   675
               Width           =   2080
            End
            Begin VB.TextBox txtATTDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   3
               Left            =   840
               MaxLength       =   30
               TabIndex        =   128
               Top             =   915
               Width           =   2080
            End
            Begin VB.TextBox txtATTDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   4
               Left            =   840
               MaxLength       =   30
               TabIndex        =   129
               Top             =   1170
               Width           =   2080
            End
            Begin VB.TextBox txtAnaesDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   840
               MaxLength       =   30
               TabIndex        =   130
               Top             =   1410
               Width           =   2080
            End
            Begin VB.TextBox txtAnaesDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   840
               MaxLength       =   30
               TabIndex        =   131
               Top             =   1665
               Width           =   2080
            End
            Begin VB.TextBox txtAnaesDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   840
               MaxLength       =   30
               TabIndex        =   132
               Top             =   1920
               Width           =   2080
            End
            Begin VB.Label Label65 
               Caption         =   "Anaes 3"
               Height          =   255
               Index           =   91
               Left            =   60
               TabIndex        =   276
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label Label65 
               Caption         =   "Anaes 2"
               Height          =   255
               Index           =   90
               Left            =   60
               TabIndex        =   275
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label65 
               Caption         =   "Anaes 1"
               Height          =   255
               Index           =   89
               Left            =   60
               TabIndex        =   274
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label65 
               Caption         =   "Att Doc 5"
               Height          =   255
               Index           =   88
               Left            =   60
               TabIndex        =   273
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label65 
               Caption         =   "Att Doc 4"
               Height          =   255
               Index           =   87
               Left            =   60
               TabIndex        =   272
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label65 
               Caption         =   "Att Doc 3"
               Height          =   255
               Index           =   86
               Left            =   60
               TabIndex        =   271
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label65 
               Caption         =   "Att Doc 2"
               Height          =   255
               Index           =   85
               Left            =   60
               TabIndex        =   270
               Top             =   450
               Width           =   1215
            End
            Begin VB.Label Label65 
               Caption         =   "Att Doc 1"
               Height          =   255
               Index           =   84
               Left            =   60
               TabIndex        =   269
               Top             =   180
               Width           =   1215
            End
         End
         Begin VB.Frame FrameSMMARate 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2265
            Left            =   3960
            TabIndex        =   247
            Top             =   1320
            Width           =   1095
            Begin VB.TextBox txtSTotalMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   248
               Top             =   1920
               Width           =   975
            End
            Begin VB.TextBox txtSMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   81
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox txtSMMARate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   82
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox txtSMMARate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   83
               Top             =   1080
               Width           =   975
            End
            Begin VB.TextBox txtSMMARate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   84
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox txtSMMARate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   85
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Surg MMA Rate"
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
               Index           =   68
               Left            =   120
               TabIndex        =   254
               Top             =   120
               Width           =   735
            End
            Begin VB.Line Line1 
               X1              =   80
               X2              =   960
               Y1              =   1860
               Y2              =   1860
            End
         End
         Begin VB.Frame FrameSOPCS 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2265
            Left            =   2760
            TabIndex        =   251
            Top             =   1320
            Width           =   1215
            Begin VB.TextBox txtProcCode1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   72
               Top             =   600
               Width           =   1095
            End
            Begin VB.CommandButton cmdOPCS 
               Caption         =   "Search"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   252
               Top             =   320
               Width           =   735
            End
            Begin VB.TextBox txtProcCode2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   74
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox txtProcCode3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   76
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtProcCode4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   78
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox txtProcCode5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   80
               Top             =   1560
               Width           =   1095
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "OPCS"
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
               Index           =   69
               Left            =   360
               TabIndex        =   255
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame FrameAnaesMMA 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2265
            Left            =   5040
            TabIndex        =   249
            Top             =   1320
            Width           =   1095
            Begin VB.TextBox txtTotalAnaesMMARate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   250
               Top             =   1920
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   86
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   87
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   88
               Top             =   1080
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   89
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   90
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Anaes MMA Rate"
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
               Index           =   71
               Left            =   80
               TabIndex        =   257
               Top             =   120
               Width           =   975
            End
            Begin VB.Line Line3 
               X1              =   80
               X2              =   960
               Y1              =   1860
               Y2              =   1860
            End
         End
         Begin VB.Frame Frame20 
            Height          =   2150
            Index           =   3
            Left            =   8350
            TabIndex        =   192
            Top             =   3480
            Width           =   3015
            Begin MSMask.MaskEdBox txtDateRECDoc 
               Height          =   300
               Left            =   4080
               TabIndex        =   147
               Top             =   540
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
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
            Begin VB.TextBox txtREDAmt 
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
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   133
               Top             =   150
               Width           =   1335
            End
            Begin VB.ComboBox cboTopUpLmt 
               Height          =   315
               Left            =   1560
               Sorted          =   -1  'True
               TabIndex        =   134
               Text            =   "N - No"
               Top             =   400
               Width           =   1335
            End
            Begin VB.ComboBox cboCASInd 
               Height          =   315
               Left            =   1560
               TabIndex        =   135
               Top             =   690
               Width           =   1335
            End
            Begin VB.ComboBox cboQuery 
               Height          =   315
               Left            =   1560
               TabIndex        =   136
               Top             =   950
               Width           =   1335
            End
            Begin VB.ComboBox cboPayPending 
               Height          =   315
               Left            =   1560
               TabIndex        =   137
               Text            =   "N - No"
               Top             =   1230
               Width           =   1335
            End
            Begin MSMask.MaskEdBox txtFirstDocDate 
               Height          =   315
               Left            =   1560
               TabIndex        =   138
               Top             =   1500
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
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
            Begin MSMask.MaskEdBox txtSubDocDate 
               Height          =   315
               Left            =   1560
               TabIndex        =   139
               Top             =   1785
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
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
            Begin VB.Label Label65 
               Caption         =   "First Doc Date"
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
               Index           =   59
               Left            =   120
               TabIndex        =   369
               Top             =   1560
               Width           =   1455
            End
            Begin VB.Label Label29 
               Caption         =   "Subseq. Doc Date"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   364
               Top             =   1800
               Width           =   1335
            End
            Begin VB.Label Label65 
               Caption         =   "Pymt P'ding App"
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
               Index           =   67
               Left            =   120
               TabIndex        =   242
               Top             =   1250
               Width           =   1335
            End
            Begin VB.Label Label65 
               Caption         =   "Query Status"
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
               Index           =   65
               Left            =   120
               TabIndex        =   241
               Top             =   980
               Width           =   975
            End
            Begin VB.Label Label65 
               Caption         =   "Special App"
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
               Index           =   64
               Left            =   120
               TabIndex        =   240
               Top             =   700
               Width           =   855
            End
            Begin VB.Label Label65 
               Caption         =   "Topup Lmt"
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
               Index           =   63
               Left            =   120
               TabIndex        =   239
               Top             =   450
               Width           =   855
            End
            Begin VB.Label Label65 
               Caption         =   "Red Amt"
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
               Index           =   62
               Left            =   120
               TabIndex        =   238
               Top             =   200
               Width           =   735
            End
         End
         Begin VB.Frame Frame19 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2265
            Index           =   0
            Left            =   6120
            TabIndex        =   277
            Top             =   1320
            Width           =   1095
            Begin VB.TextBox txtSActFee1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   91
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox txtTotalSActFee 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   278
               Top             =   1920
               Width           =   975
            End
            Begin VB.TextBox txtSActFee2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   92
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox txtSActFee3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   93
               Top             =   1080
               Width           =   975
            End
            Begin VB.TextBox txtSActFee4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   94
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox txtSActFee5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   95
               Top             =   1560
               Width           =   975
            End
            Begin VB.Line Line2 
               X1              =   80
               X2              =   960
               Y1              =   1860
               Y2              =   1860
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Act Surg Fee"
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
               Index           =   100
               Left            =   120
               TabIndex        =   279
               Top             =   120
               Width           =   735
            End
         End
         Begin VB.Frame FrameSProc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2265
            Index           =   0
            Left            =   120
            TabIndex        =   253
            Top             =   1320
            Width           =   2625
            Begin VB.TextBox txtProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   71
               Top             =   600
               Width           =   2535
            End
            Begin VB.TextBox txtProcedure2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   73
               Top             =   840
               Width           =   2535
            End
            Begin VB.TextBox txtProcedure3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   75
               Top             =   1080
               Width           =   2535
            End
            Begin VB.TextBox txtProcedure4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   77
               Top             =   1320
               Width           =   2535
            End
            Begin VB.TextBox txtProcedure5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   79
               Top             =   1560
               Width           =   2535
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Surgical Procedure"
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
               Index           =   70
               Left            =   120
               TabIndex        =   256
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.Frame FrameNSOPCS 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   2740
            TabIndex        =   245
            Top             =   3520
            Width           =   1215
            Begin VB.TextBox txtNSProcCode1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   102
               Top             =   480
               Width           =   1130
            End
            Begin VB.CommandButton cmdOPCS 
               Caption         =   "Search"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   246
               Top             =   150
               Width           =   735
            End
            Begin VB.TextBox txtNSProcCode2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   104
               Top             =   720
               Width           =   1130
            End
            Begin VB.TextBox txtNSProcCode3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   106
               Top             =   960
               Width           =   1130
            End
            Begin VB.TextBox txtNSProcCode4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   108
               Top             =   1200
               Width           =   1130
            End
            Begin VB.TextBox txtNSProcCode5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   110
               Top             =   1440
               Width           =   1130
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "OPCS"
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
               Index           =   103
               Left            =   45
               TabIndex        =   283
               Top             =   120
               Width           =   420
            End
         End
         Begin VB.Frame FrameNSMMARate 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   3960
            TabIndex        =   243
            Top             =   3520
            Width           =   1095
            Begin VB.TextBox txtTotalNSMMARate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   244
               Top             =   1780
               Width           =   975
            End
            Begin VB.TextBox txtNSMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   111
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtNSMMARate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   112
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox txtNSMMARate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   113
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtNSMMARate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   114
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox txtNSMMARate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   115
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Non-Surg MMA Rate"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   104
               Left            =   120
               TabIndex        =   284
               Top             =   120
               Width           =   855
            End
            Begin VB.Line Line5 
               X1              =   45
               X2              =   960
               Y1              =   1740
               Y2              =   1740
            End
         End
         Begin VB.Frame FrameSProc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Index           =   2
            Left            =   120
            TabIndex        =   288
            Top             =   3520
            Width           =   2625
            Begin VB.TextBox txtNSProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   101
               Top             =   480
               Width           =   2535
            End
            Begin VB.TextBox txtNSProcedure2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   103
               Top             =   720
               Width           =   2535
            End
            Begin VB.TextBox txtNSProcedure3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   105
               Top             =   960
               Width           =   2535
            End
            Begin VB.TextBox txtNSProcedure4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   107
               Top             =   1200
               Width           =   2535
            End
            Begin VB.TextBox txtNSProcedure5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   109
               Top             =   1440
               Width           =   2535
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Non- Surgical Procedure"
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
               Index           =   102
               Left            =   120
               TabIndex        =   289
               Top             =   120
               Width           =   1815
            End
         End
         Begin VB.Frame Frame19 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Index           =   1
            Left            =   5060
            TabIndex        =   290
            Top             =   3520
            Width           =   1095
            Begin VB.TextBox txtTotalNSActFee 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   291
               Top             =   1800
               Width           =   975
            End
            Begin VB.TextBox txtNSActFee1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   116
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtNSActFee2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   117
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox txtNSActFee3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   118
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtNSActFee4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   119
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox txtNSActFee5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   120
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Act Non-Surg Fee"
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
               Index           =   106
               Left            =   75
               TabIndex        =   292
               Top             =   120
               Width           =   975
            End
            Begin VB.Line Line6 
               X1              =   75
               X2              =   955
               Y1              =   1740
               Y2              =   1740
            End
         End
         Begin VB.Frame Frame22 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2265
            Left            =   7200
            TabIndex        =   280
            Top             =   1320
            Width           =   1095
            Begin VB.TextBox txtTotalAnaesActFees 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   281
               Top             =   1920
               Width           =   975
            End
            Begin VB.TextBox txtAnaesActFees1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   96
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox txtAnaesActFees2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   97
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox txtAnaesActFees3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   98
               Top             =   1080
               Width           =   975
            End
            Begin VB.TextBox txtAnaesActFees4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   99
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox txtAnaesActFees5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   100
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Act Anaes Fee"
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
               Index           =   101
               Left            =   120
               TabIndex        =   282
               Top             =   120
               Width           =   735
            End
            Begin VB.Line Line4 
               X1              =   80
               X2              =   960
               Y1              =   1860
               Y2              =   1860
            End
         End
         Begin VB.Frame FrameSProc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Index           =   1
            Left            =   6150
            TabIndex        =   286
            Top             =   3520
            Width           =   2155
            Begin VB.TextBox TxtLGRemark 
               Appearance      =   0  'Flat
               Height          =   1605
               Left            =   40
               MaxLength       =   100
               ScrollBars      =   3  'Both
               TabIndex        =   121
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label Label65 
               Caption         =   "LG remarks"
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
               Index           =   72
               Left            =   120
               TabIndex        =   287
               Top             =   240
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame20 
         Height          =   5445
         Index           =   4
         Left            =   120
         TabIndex        =   180
         Top             =   360
         Width           =   11415
         Begin VB.CheckBox chkDiag 
            Height          =   255
            Left            =   11115
            TabIndex        =   195
            Top             =   2205
            Width           =   255
         End
         Begin VB.CheckBox ChkDOD 
            Caption         =   "Check1"
            Height          =   195
            Left            =   11115
            TabIndex        =   30
            Top             =   700
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox ChkDOA 
            Caption         =   "Check1"
            Height          =   255
            Left            =   11115
            TabIndex        =   27
            Top             =   420
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.Frame Frame20 
            Height          =   1365
            Index           =   5
            Left            =   5640
            TabIndex        =   184
            Top             =   2400
            Width           =   5655
            Begin VB.CommandButton CmdDiag 
               Caption         =   ">"
               Height          =   255
               Index           =   2
               Left            =   5280
               TabIndex        =   41
               Top             =   180
               Width           =   225
            End
            Begin VB.TextBox cboFDICDCode1 
               Height          =   315
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   44
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtFinalDiag1 
               Height          =   315
               Left            =   90
               MaxLength       =   150
               TabIndex        =   42
               Top             =   480
               Width           =   3870
            End
            Begin VB.TextBox txtFinalDiag2 
               Height          =   315
               Left            =   90
               MaxLength       =   150
               TabIndex        =   45
               Top             =   735
               Width           =   3870
            End
            Begin VB.TextBox txtFinalDiag3 
               Height          =   285
               Left            =   90
               MaxLength       =   150
               TabIndex        =   48
               Top             =   1005
               Width           =   3870
            End
            Begin MSMask.MaskEdBox txtFDDOS1 
               Height          =   315
               Left            =   3960
               TabIndex        =   43
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
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
            Begin MSMask.MaskEdBox txtFDDOS2 
               Height          =   315
               Left            =   3960
               TabIndex        =   46
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
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
            Begin MSMask.MaskEdBox txtFDDOS3 
               Height          =   315
               Left            =   3960
               TabIndex        =   49
               Top             =   1005
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
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
            Begin VB.TextBox cboFDICDCode2 
               Height          =   315
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   47
               Top             =   735
               Width           =   615
            End
            Begin VB.TextBox cboFDICDCode3 
               Height          =   315
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   50
               Top             =   1005
               Width           =   615
            End
            Begin VB.Label Label65 
               Caption         =   "ICD"
               Height          =   255
               Index           =   57
               Left            =   4920
               TabIndex        =   236
               Top             =   180
               Width           =   615
            End
            Begin VB.Label Label65 
               Caption         =   "Date Onset"
               Height          =   255
               Index           =   56
               Left            =   3960
               TabIndex        =   235
               Top             =   180
               Width           =   855
            End
            Begin VB.Label Label65 
               Caption         =   "Final Diag."
               Height          =   255
               Index           =   55
               Left            =   120
               TabIndex        =   234
               Top             =   180
               Width           =   975
            End
         End
         Begin VB.Frame Frame20 
            Height          =   1365
            Index           =   7
            Left            =   120
            TabIndex        =   182
            Top             =   2400
            Width           =   5415
            Begin VB.CommandButton CmdDiag 
               Caption         =   ">"
               Height          =   255
               Index           =   0
               Left            =   5040
               TabIndex        =   37
               Top             =   180
               Width           =   225
            End
            Begin VB.TextBox cboICDCode 
               Height          =   315
               Index           =   0
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   40
               Top             =   465
               Width           =   615
            End
            Begin VB.TextBox txtPrelimDiag 
               Height          =   285
               Index           =   0
               Left            =   90
               MaxLength       =   150
               TabIndex        =   38
               Top             =   465
               Width           =   3650
            End
            Begin MSMask.MaskEdBox txtPDDOS 
               Height          =   300
               Index           =   0
               Left            =   3720
               TabIndex        =   39
               Top             =   465
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   529
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
            Begin VB.TextBox cboICDCode 
               Height          =   315
               Index           =   1
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   358
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox cboICDCode 
               Height          =   315
               Index           =   2
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   359
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtPrelimDiag 
               Height          =   285
               Index           =   1
               Left            =   90
               MaxLength       =   150
               TabIndex        =   360
               Top             =   720
               Width           =   3650
            End
            Begin VB.TextBox txtPrelimDiag 
               Height          =   285
               Index           =   2
               Left            =   90
               MaxLength       =   150
               TabIndex        =   361
               Top             =   960
               Width           =   3650
            End
            Begin MSMask.MaskEdBox txtPDDOS 
               Height          =   300
               Index           =   1
               Left            =   3720
               TabIndex        =   362
               Top             =   720
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   529
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
            Begin MSMask.MaskEdBox txtPDDOS 
               Height          =   300
               Index           =   2
               Left            =   3720
               TabIndex        =   363
               Top             =   960
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   529
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
            Begin VB.Label Label65 
               Caption         =   "ICD"
               Height          =   255
               Index           =   51
               Left            =   4680
               TabIndex        =   230
               Top             =   180
               Width           =   615
            End
            Begin VB.Label Label65 
               Caption         =   "Date Onset"
               Height          =   255
               Index           =   50
               Left            =   3720
               TabIndex        =   229
               Top             =   180
               Width           =   975
            End
            Begin VB.Label Label65 
               Caption         =   "Prelim. Diag"
               Height          =   255
               Index           =   49
               Left            =   120
               TabIndex        =   228
               Top             =   180
               Width           =   1215
            End
         End
         Begin VB.TextBox txtPatientIC 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   24
            Top             =   130
            Width           =   3255
         End
         Begin VB.ComboBox cboPatientList 
            Enabled         =   0   'False
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
            Left            =   1400
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   120
            Width           =   5000
         End
         Begin MSMask.MaskEdBox txtFirstCall 
            Height          =   300
            Left            =   7800
            TabIndex        =   25
            Top             =   390
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
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
         Begin MSMask.MaskEdBox txtFirstTreated 
            Height          =   300
            Left            =   7800
            TabIndex        =   28
            Top             =   660
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
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
         Begin MSMask.MaskEdBox txtDateAdmitted 
            Height          =   300
            Left            =   9840
            TabIndex        =   26
            Top             =   390
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
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
         Begin MSMask.MaskEdBox txtDateDischarge 
            Height          =   300
            Left            =   9840
            TabIndex        =   29
            Top             =   660
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
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
         Begin VB.ComboBox txtHOSCode 
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
            ItemData        =   "frmCaseAdjustment.frx":0118
            Left            =   1400
            List            =   "frmCaseAdjustment.frx":011A
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   400
            Width           =   5000
         End
         Begin VB.TextBox KIVPatientName 
            Height          =   285
            Left            =   1400
            TabIndex        =   181
            Top             =   120
            Visible         =   0   'False
            Width           =   5000
         End
         Begin VB.TextBox txtReason 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   1400
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   680
            Width           =   5000
         End
         Begin VB.TextBox txtREFBy 
            Height          =   285
            Left            =   1400
            MaxLength       =   100
            TabIndex        =   19
            Top             =   1235
            Width           =   5000
         End
         Begin VB.ComboBox cboStayType 
            Height          =   315
            Left            =   4410
            TabIndex        =   21
            Top             =   1490
            Width           =   2000
         End
         Begin VB.ComboBox CboPayType 
            Height          =   315
            ItemData        =   "frmCaseAdjustment.frx":011C
            Left            =   4410
            List            =   "frmCaseAdjustment.frx":011E
            TabIndex        =   23
            Top             =   1770
            Width           =   2000
         End
         Begin VB.TextBox txtCASLGAmt 
            Alignment       =   1  'Right Justify
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
            Left            =   9840
            TabIndex        =   32
            Top             =   910
            Width           =   1215
         End
         Begin VB.TextBox txtCASReserveAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   7800
            TabIndex        =   31
            Top             =   910
            Width           =   1215
         End
         Begin VB.ComboBox cboCondwillOcc 
            Height          =   315
            Left            =   7800
            TabIndex        =   33
            Top             =   1200
            Width           =   3255
         End
         Begin VB.ComboBox cboFLWTreatment 
            Height          =   315
            Left            =   7800
            TabIndex        =   34
            Top             =   1490
            Width           =   3255
         End
         Begin VB.ComboBox cboNature 
            Height          =   315
            Left            =   1400
            TabIndex        =   20
            Top             =   1490
            Width           =   2000
         End
         Begin VB.ComboBox cboModality 
            Height          =   315
            Left            =   1400
            TabIndex        =   22
            Top             =   1770
            Width           =   2000
         End
         Begin MSMask.MaskEdBox txtFollowUpDate 
            Height          =   300
            Left            =   7800
            TabIndex        =   35
            Top             =   1750
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
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
         Begin VB.ComboBox cboFolUp 
            Height          =   315
            Left            =   9840
            TabIndex        =   36
            Top             =   1750
            Width           =   1215
         End
         Begin VB.Frame Frame20 
            Height          =   1425
            Index           =   6
            Left            =   5640
            TabIndex        =   185
            Top             =   3720
            Width           =   5655
            Begin VB.CommandButton CmdDiag 
               Caption         =   ">"
               Height          =   255
               Index           =   3
               Left            =   1680
               TabIndex        =   61
               Top             =   180
               Width           =   225
            End
            Begin VB.ComboBox cboREPCode1 
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
               Left            =   90
               Style           =   1  'Simple Combo
               TabIndex        =   62
               Top             =   480
               Width           =   5480
            End
            Begin VB.ComboBox cboREPCode2 
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
               Left            =   90
               Style           =   1  'Simple Combo
               TabIndex        =   63
               Top             =   765
               Width           =   5480
            End
            Begin VB.ComboBox cboREPCode3 
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
               Left            =   90
               Style           =   1  'Simple Combo
               TabIndex        =   64
               Top             =   1035
               Width           =   5480
            End
            Begin VB.Label Label65 
               Caption         =   "Repudiation Resons"
               Height          =   255
               Index           =   58
               Left            =   120
               TabIndex        =   237
               Top             =   180
               Width           =   1815
            End
         End
         Begin VB.Frame Frame20 
            Height          =   1425
            Index           =   8
            Left            =   120
            TabIndex        =   183
            Top             =   3720
            Width           =   5415
            Begin VB.CommandButton CmdDiag 
               Caption         =   ">"
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   51
               Top             =   180
               Width           =   225
            End
            Begin VB.TextBox txtUD1 
               Height          =   285
               Left            =   90
               MaxLength       =   150
               TabIndex        =   52
               Top             =   480
               Width           =   3650
            End
            Begin MSMask.MaskEdBox txtUDDOS1 
               Height          =   300
               Left            =   3720
               TabIndex        =   53
               Top             =   465
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   529
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
            Begin VB.TextBox txtUD2 
               Height          =   285
               Left            =   90
               MaxLength       =   150
               TabIndex        =   55
               Top             =   720
               Width           =   3650
            End
            Begin MSMask.MaskEdBox txtUDDOS2 
               Height          =   300
               Left            =   3720
               TabIndex        =   56
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   529
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
            Begin VB.TextBox txtUD3 
               Height          =   285
               Left            =   90
               MaxLength       =   150
               TabIndex        =   58
               Top             =   960
               Width           =   3650
            End
            Begin MSMask.MaskEdBox txtUDDOS3 
               Height          =   300
               Left            =   3720
               TabIndex        =   59
               Top             =   975
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   529
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
            Begin VB.TextBox cboUDICDCode1 
               Height          =   315
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   54
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox cboUDICDCode2 
               Height          =   315
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   57
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox cboUDICDCode3 
               Height          =   315
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   60
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label65 
               Caption         =   "ICD"
               Height          =   255
               Index           =   54
               Left            =   4680
               TabIndex        =   233
               Top             =   180
               Width           =   615
            End
            Begin VB.Label Label65 
               Caption         =   "Date Onset"
               Height          =   255
               Index           =   53
               Left            =   3720
               TabIndex        =   232
               Top             =   180
               Width           =   975
            End
            Begin VB.Label Label65 
               Caption         =   "Underlying Disease"
               Height          =   255
               Index           =   52
               Left            =   120
               TabIndex        =   231
               Top             =   180
               Width           =   1455
            End
         End
         Begin VB.Label Label65 
            Caption         =   "Fol-Up"
            Height          =   255
            Index           =   48
            Left            =   9240
            TabIndex        =   227
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "LG Amt"
            Height          =   255
            Index           =   32
            Left            =   9120
            TabIndex        =   215
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "D.O.D"
            Height          =   255
            Index           =   31
            Left            =   9120
            TabIndex        =   214
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "D.O.A"
            Height          =   255
            Index           =   30
            Left            =   9120
            TabIndex        =   213
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Fol-Up Date"
            Height          =   255
            Index           =   29
            Left            =   6720
            TabIndex        =   212
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Fol-Up Req"
            Height          =   255
            Index           =   28
            Left            =   6720
            TabIndex        =   211
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label65 
            Caption         =   "Recur"
            Height          =   255
            Index           =   27
            Left            =   6720
            TabIndex        =   210
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Reserve Amt"
            Height          =   255
            Index           =   26
            Left            =   6720
            TabIndex        =   209
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "F. Call"
            Height          =   255
            Index           =   25
            Left            =   6720
            TabIndex        =   208
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "F. Call"
            Height          =   255
            Index           =   24
            Left            =   6720
            TabIndex        =   207
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "IC Number"
            Height          =   255
            Index           =   23
            Left            =   6720
            TabIndex        =   206
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Cash Type"
            Height          =   255
            Index           =   22
            Left            =   3480
            TabIndex        =   205
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Stay Type"
            Height          =   255
            Index           =   21
            Left            =   3480
            TabIndex        =   204
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Treatment Mode"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   203
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "Nat of Clm"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   202
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Refer"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   201
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Reason"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   200
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Hospital"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   199
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Patient"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   198
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label53 
            Caption         =   "Copy Diag."
            Height          =   255
            Left            =   10245
            TabIndex        =   194
            Top             =   2205
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1770
         Left            =   -74880
         TabIndex        =   166
         Top             =   360
         Width           =   11415
         Begin VB.TextBox txtMBMPolNo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            MaxLength       =   25
            TabIndex        =   186
            Top             =   120
            Width           =   1335
         End
         Begin VB.TextBox txtAdd1 
            Height          =   285
            Left            =   960
            MaxLength       =   50
            TabIndex        =   153
            Top             =   130
            Width           =   4815
         End
         Begin VB.TextBox txtINSCode 
            BackColor       =   &H80000001&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   14400
            TabIndex        =   167
            Top             =   1200
            Width           =   1695
         End
         Begin MSMask.MaskEdBox txtDOB 
            Height          =   285
            Left            =   7440
            TabIndex        =   158
            Top             =   375
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtAdd2 
            Height          =   285
            Left            =   960
            MaxLength       =   50
            TabIndex        =   178
            Top             =   390
            Width           =   4815
         End
         Begin VB.TextBox txtAdd3 
            Height          =   285
            Left            =   960
            MaxLength       =   50
            TabIndex        =   154
            Top             =   650
            Width           =   4815
         End
         Begin VB.TextBox txtPostCode 
            Height          =   285
            Left            =   960
            MaxLength       =   5
            TabIndex        =   155
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox txtCity 
            Height          =   285
            Left            =   3480
            MaxLength       =   50
            TabIndex        =   156
            Top             =   900
            Width           =   2295
         End
         Begin VB.TextBox txtState 
            Height          =   285
            Left            =   960
            MaxLength       =   50
            TabIndex        =   157
            Top             =   1160
            Width           =   4815
         End
         Begin VB.TextBox txtAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            TabIndex        =   177
            Top             =   630
            Width           =   1335
         End
         Begin VB.TextBox txtAvailableAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            TabIndex        =   159
            Top             =   885
            Width           =   1335
         End
         Begin VB.TextBox txtRecoveryOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            TabIndex        =   160
            Top             =   1140
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            TabIndex        =   161
            Top             =   1395
            Width           =   1335
         End
         Begin VB.Label Label65 
            Caption         =   "Installment"
            Height          =   255
            Index           =   82
            Left            =   6480
            TabIndex        =   267
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Amt O/S"
            Height          =   255
            Index           =   81
            Left            =   6480
            TabIndex        =   266
            Top             =   1150
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Avail. Limit"
            Height          =   255
            Index           =   80
            Left            =   6480
            TabIndex        =   265
            Top             =   920
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Ann. Limit"
            Height          =   255
            Index           =   79
            Left            =   6480
            TabIndex        =   264
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "D.O.B"
            Height          =   255
            Index           =   78
            Left            =   6480
            TabIndex        =   263
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Pol No"
            Height          =   255
            Index           =   77
            Left            =   6480
            TabIndex        =   262
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Town"
            Height          =   255
            Index           =   76
            Left            =   2760
            TabIndex        =   261
            Top             =   915
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "State"
            Height          =   255
            Index           =   75
            Left            =   240
            TabIndex        =   260
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "PCode"
            Height          =   255
            Index           =   74
            Left            =   240
            TabIndex        =   259
            Top             =   920
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Add"
            Height          =   255
            Index           =   73
            Left            =   240
            TabIndex        =   258
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Insured Type"
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
            Left            =   13200
            TabIndex        =   170
            Top             =   1245
            Width           =   1275
         End
         Begin VB.Label lblBelow2000 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Below 2000"
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
            Height          =   285
            Left            =   9120
            TabIndex        =   169
            Top             =   960
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label txtMorW 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Height          =   285
            Left            =   9120
            TabIndex        =   168
            Top             =   360
            Width           =   2175
         End
      End
      Begin MSComctlLib.ListView lstBenefitDetails 
         Height          =   5115
         Left            =   -74880
         TabIndex        =   171
         Top             =   360
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   9022
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Benefits Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Benefit Description"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "No Of Days"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM per Day"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Disability"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMember 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   179
         Top             =   2160
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5953
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cover ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "I/C No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Relation"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOB"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Supp Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Warranty Exc"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame14 
         Caption         =   "Hospital Visit"
         Height          =   975
         Left            =   -74880
         TabIndex        =   326
         Top             =   4080
         Width           =   11415
         Begin VB.TextBox txtHospVisit 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   4920
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   328
            Top             =   240
            Width           =   6375
         End
         Begin VB.TextBox txtOfficer 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   327
            Top             =   600
            Width           =   2730
         End
         Begin MSMask.MaskEdBox txtDV 
            Height          =   315
            Left            =   1200
            TabIndex        =   332
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox txtDOVTime 
            Height          =   255
            Left            =   3000
            TabIndex        =   333
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label65 
            Caption         =   "D.O.Visit"
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   337
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "Off-In-Charge"
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   336
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "Time"
            Height          =   255
            Index           =   35
            Left            =   2520
            TabIndex        =   335
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "Action"
            Height          =   255
            Index           =   36
            Left            =   4080
            TabIndex        =   334
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCASRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4860
         Left            =   -74040
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   344
         Top             =   840
         Width           =   8895
      End
      Begin VB.TextBox txtCASRemarkTwo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4860
         Left            =   -74040
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   349
         Top             =   840
         Width           =   8895
      End
      Begin MSComctlLib.ListView lstBenefits 
         Height          =   4635
         Left            =   -74920
         TabIndex        =   354
         Top             =   840
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   8176
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bnf Desc"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Days"
            Object.Width           =   952
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM PDay"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Dis"
            Object.Width           =   1640
         EndProperty
      End
      Begin MSComctlLib.ListView lstNewBenefitDetails 
         Height          =   4635
         Left            =   -69160
         TabIndex        =   355
         Top             =   840
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   8176
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bnf Desc"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Days"
            Object.Width           =   952
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM PDay"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Dis"
            Object.Width           =   1640
         EndProperty
      End
      Begin MSMask.MaskEdBox txtDOC 
         Height          =   300
         Left            =   -65640
         TabIndex        =   356
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.TextBox txtIntCASRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4860
         Left            =   -74040
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   367
         Top             =   840
         Width           =   8895
      End
      Begin VB.Label lblMemo 
         Caption         =   "Pls Check Memo"
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
         Left            =   -65040
         TabIndex        =   380
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label65 
         Caption         =   "Supplementary"
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
         Index           =   111
         Left            =   -74880
         TabIndex        =   379
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label65 
         Caption         =   "Principal"
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
         Index           =   110
         Left            =   -74760
         TabIndex        =   378
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label65 
         Caption         =   "Eff Date"
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
         Index           =   95
         Left            =   -66360
         TabIndex        =   377
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label65 
         Caption         =   "Plan"
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
         Index           =   94
         Left            =   -67920
         TabIndex        =   376
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblPage 
         Caption         =   "Page 1"
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
         Left            =   -65040
         TabIndex        =   374
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label65 
         Caption         =   "Date"
         Height          =   255
         Index           =   96
         Left            =   -74040
         TabIndex        =   347
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Remarks"
         Height          =   255
         Index           =   97
         Left            =   -72840
         TabIndex        =   346
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label65 
         Caption         =   "Person Incharge"
         Height          =   255
         Index           =   98
         Left            =   -66720
         TabIndex        =   345
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame20 
      Height          =   520
      Index           =   11
      Left            =   120
      TabIndex        =   163
      Top             =   7480
      Width           =   11610
      Begin VB.CommandButton cmdTHDrugs 
         Caption         =   "&T.Home Drugs"
         Height          =   315
         Left            =   5380
         TabIndex        =   366
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdPreUnd 
         Caption         =   "N&on-Disc"
         Height          =   315
         Left            =   4560
         TabIndex        =   197
         Top             =   150
         Width           =   825
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print &Mgmt Note"
         Height          =   315
         Left            =   9360
         TabIndex        =   196
         Top             =   150
         Width           =   1455
      End
      Begin VB.CommandButton cmdFollowUp 
         Caption         =   "&Fol. Up LG"
         Height          =   315
         Left            =   3600
         TabIndex        =   193
         Top             =   150
         Width           =   960
      End
      Begin VB.CommandButton cmdPrior 
         Caption         =   "&Notify"
         Height          =   315
         Left            =   2760
         TabIndex        =   151
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton cmdSaveAdd 
         Caption         =   "Upd &Add"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1785
         TabIndex        =   150
         ToolTipText     =   "Update Address"
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton cmdRepudiate 
         Caption         =   "Rep&udiate"
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   149
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton cmdLG 
         Caption         =   "&LG"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   148
         Top             =   150
         Width           =   800
      End
      Begin VB.TextBox txtHLTCOde 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   174
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtPLNCode 
         Height          =   285
         Left            =   840
         TabIndex        =   173
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtINSCode1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   172
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
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
         Left            =   10800
         TabIndex        =   152
         Top             =   150
         Width           =   720
      End
      Begin VB.CommandButton cmdUpdate1 
         Caption         =   "&Update"
         Enabled         =   0   'False
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
         Left            =   8400
         TabIndex        =   145
         Top             =   150
         Width           =   960
      End
   End
   Begin VB.Label Label65 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Case Adjustment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   93
      Left            =   0
      TabIndex        =   375
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmCaseAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intLastPoz2 As Integer
Dim FirstDiag As Integer
Dim FinalDiag As Integer
Dim UD1 As Integer
Dim SaveOthers As Integer
Dim REPReasons As Integer
Dim Letter As String
Dim LGNumber As String
Dim FOLNumber As String
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim EXCNumber As String
Dim MBMNumber As String
Dim TempCASNo As String
Dim RB As String
Dim patient As String
Dim ReprintLG As Boolean
Dim ReprintBTBG As Boolean
Dim SndLG As Boolean
Dim bTabDone As Boolean
Dim ScanCode As String
Dim HOspitalName As String
Dim PrevDOA As String
Dim PatientName As String
Dim CaseNumber As String
Dim REDAmtInd As Boolean
Dim RedAmt As Double
Dim CaseRegDate As String
Dim TxtPlanCode  As String
Dim MBMCPLNCode As String
Dim sqlstrCAS As String
Dim CASSearch As New ADODB.Recordset
Dim check As Boolean
Dim TotalFinalDiag As String
Dim TotalProc As String
Dim LGNo As String
Dim LGVer As Integer
Dim PreIllDate As String
Dim PreIllCAT As String
Dim PreIllOthers As String
Dim PreIllInfBy As String
Dim UndIllDate As String
Dim UndIllCAT As String
Dim UndIllOthers As String
Dim UndIllInfBy As String
Dim UpdClaimability As Boolean
Public Document, Personnel, Attention As String
Public Underlying As Boolean
Public CASUnderlying As String
Dim MMASurgFees As Double
Dim MMAAnaeFees As Double
Dim MMARateCal As Integer
Dim tmpCLMNumber As String
Dim tmpCLMversion As String
Dim tmpPLNPerVisit As String

Private Sub CboARStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboARStatus, CboARStatus.SelStart) + Chr(KeyAscii) + Mid(cboPatientList, cboPatientList.SelStart + cboPatientList.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboARStatus.ListCount - 1
        
        If NewText = LCase(Left(CboARStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboARStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboARStatus = ValidValue
               CboARStatus.SelStart = 0
               CboARStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCASInd_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCASInd, cboCASInd.SelStart) + Chr(KeyAscii) + Mid(cboCASInd, cboCASInd.SelStart + cboCASInd.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboCASInd.ListCount - 1
        
        If NewText = LCase(Left(cboCASInd.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCASInd.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboCASInd = ValidValue
               cboCASInd.SelStart = 0
               cboCASInd.SelLength = Len(ValidValue)
            End If
        End If
End Sub
Private Sub cboCASTakeOveInd_Click()
    If Mid(cboCASTakeOveInd, 1, 1) = "Y" Then
        txtCASTakeOverAmt.Enabled = True
    Else
        txtCASTakeOverAmt.Enabled = False
    End If

End Sub

Private Sub cboCASTakeOveInd_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCASTakeOveInd, cboCASTakeOveInd.SelStart) + Chr(KeyAscii) + Mid(cboCASTakeOveInd, cboCASTakeOveInd.SelStart + cboCASTakeOveInd.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboCASTakeOveInd.ListCount - 1
        
        If NewText = LCase(Left(cboCASTakeOveInd.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCASTakeOveInd.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboCASTakeOveInd = ValidValue
               cboCASTakeOveInd.SelStart = 0
               cboCASTakeOveInd.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCASTakeOveInd_LostFocus()
    If Mid(cboCASTakeOveInd, 1, 1) = "Y" Then
        txtCASTakeOverAmt.Enabled = True
    Else
        txtCASTakeOverAmt.Enabled = False
    End If
End Sub




Private Sub cboIntExtTO_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboIntExtTO, cboIntExtTO.SelStart) + Chr(KeyAscii) + Mid(cboIntExtTO, cboIntExtTO.SelStart + cboIntExtTO.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboIntExtTO.ListCount - 1
        
        If NewText = LCase(Left(cboIntExtTO.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboIntExtTO.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboIntExtTO = ValidValue
               cboIntExtTO.SelStart = 0
               cboIntExtTO.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboPayPending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayPending, cboPayPending.SelStart) + Chr(KeyAscii) + Mid(cboPayPending, cboPayPending.SelStart + cboPayPending.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboPayPending.ListCount - 1
        
        If NewText = LCase(Left(cboPayPending.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayPending.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboPayPending = ValidValue
               cboPayPending.SelStart = 0
               cboPayPending.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboTopUpLmt_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboTopUpLmt, cboTopUpLmt.SelStart) + Chr(KeyAscii) + Mid(cboTopUpLmt, cboTopUpLmt.SelStart + cboTopUpLmt.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboTopUpLmt.ListCount - 1
        
        If NewText = LCase(Left(cboTopUpLmt.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboTopUpLmt.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboTopUpLmt = ValidValue
               cboTopUpLmt.SelStart = 0
               cboTopUpLmt.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub chkDiag_Click()

    If chkDiag.value = 1 Then
        txtFinalDiag1 = txtPrelimDiag(0)
        txtFinalDiag2 = txtPrelimDiag(1)
        txtFinalDiag3 = txtPrelimDiag(2)
        txtFDDOS1 = txtPDDOS(0)
        txtFDDOS2 = txtPDDOS(1)
        txtFDDOS3 = txtPDDOS(2)
        cboFDICDCode1 = cboICDCode(0)
        cboFDICDCode2 = cboICDCode(1)
        cboFDICDCode3 = cboICDCode(2)
    Else
        txtFinalDiag1 = ""
        txtFinalDiag2 = ""
        txtFinalDiag3 = ""
        txtFDDOS1 = "__/__/____"
        txtFDDOS2 = "__/__/____"
        txtFDDOS3 = "__/__/____"
        cboFDICDCode1 = ""
        cboFDICDCode2 = ""
        cboFDICDCode3 = ""
    End If

End Sub

Private Sub cmdCLMAppUpdate_Click()
Dim strsql As String
Dim CLM As New ADODB.Recordset

    
    If lblClaimNo = "" Then
        MsgBox "Please select claim no", vbOKOnly + vbCritical, DialogBoxTitle
        txtAppDate.SetFocus
        Exit Sub

    ElseIf txtAppDate = "__/__/____" Then
        MsgBox "Please Enter Approval Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtAppDate.SetFocus
        Exit Sub

    ElseIf cboAppStatus = "" Then
        MsgBox "Please Enter Approval Status", vbOKOnly + vbCritical, DialogBoxTitle
        cboAppStatus.SetFocus
        Exit Sub
    
    End If
    medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

    strsql = " Select CasNumber, claimversion, CLMPayorAppStatus, CLMPayorAppDate from Claim Where CasNumber = '" & Trim(tmpCLMNumber) & "' AND claimversion = '" & Trim(tmpCLMversion) & "'"
    CLM.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    With CLM
        !CLMPayorAppStatus = Mid(cboAppStatus, 1, 1)
        !CLMPayorAppDate = txtAppDate
    End With
    CLM.Update
    MsgBox "Record has been updated", vbOKOnly + vbInformation, DialogBoxTitle
    lstWorksheetListing.listitems.Clear
    Call WSListing

    CLM.Close
    Set CLM = Nothing
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
End Sub

Private Sub CmdDiag_Click(a As Integer)

    If a = "3" Then
        If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
            FrmSearchRepudiation.Source = 1
            FrmSearchRepudiation.Show vbModal
        End If
        'Else
        '    frmSearchICD.Options = 0
        '    frmSearchICD.optRep.Visible = False
        '    frmSearchICD.Source = 1
        '    frmSearchICD.show vbModal
        'End If
    Else
        'If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Then
        '    FrmSearchRepudiation.Source = 1
        '    FrmSearchRepudiation.show vbModal
        'Else
            frmSearchICD.Options = 0
            frmSearchICD.optRep.Visible = False
            frmSearchICD.Source = 1
            frmSearchICD.Show vbModal
        'End If
    End If
End Sub

Private Sub cboFolUp_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboFolUp, cboFolUp.SelStart) + Chr(KeyAscii) + Mid(cboFolUp, cboFolUp.SelStart + cboFolUp.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboFolUp.ListCount - 1
        
        If NewText = LCase(Left(cboFolUp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboFolUp.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboFolUp = ValidValue
               cboFolUp.SelStart = 0
               cboFolUp.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub CboPayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayType, CboPayType.SelStart) + Chr(KeyAscii) + Mid(CboPayType, CboPayType.SelStart + CboPayType.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboPayType.ListCount - 1
        
        If NewText = LCase(Left(CboPayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboPayType = ValidValue
               CboPayType.SelStart = 0
               CboPayType.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboPending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPending, cboPending.SelStart) + Chr(KeyAscii) + Mid(cboPending, cboPending.SelStart + cboPending.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboPending.ListCount - 1
        
        If NewText = LCase(Left(cboPending.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPending.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboPending = ValidValue
               cboPending.SelStart = 0
               cboPending.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboQuery_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboQuery, cboQuery.SelStart) + Chr(KeyAscii) + Mid(cboQuery, cboQuery.SelStart + cboQuery.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboQuery.ListCount - 1
        
        If NewText = LCase(Left(cboQuery.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboQuery.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboQuery = ValidValue
               cboQuery.SelStart = 0
               cboQuery.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboStayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboStayType, cboStayType.SelStart) + Chr(KeyAscii) + Mid(cboStayType, cboStayType.SelStart + cboStayType.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboStayType.ListCount - 1
        
        If NewText = LCase(Left(cboStayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboStayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboStayType = ValidValue
               cboStayType.SelStart = 0
               cboStayType.SelLength = Len(ValidValue)
            End If
        End If
End Sub
Private Sub cmdFollowUp_Click()
Dim strsql As String
Dim TempDate As String
Dim RecordSaved
    
    TempDate = Format(Date, "yyyy/mm/dd")
    If Trim(txtCaseNo) = "" Then
        MsgBox "Please Enter Case Number", vbOKOnly + vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
        Exit Sub
    ElseIf lstLGListing.listitems.Count = 0 Then
        MsgBox "Follow up Denial, record LG not found", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    
    If txtFollowUpDate <> "__/__/____" Then
        If IsDate(txtFollowUpDate) = False Then
            MsgBox "Invalid date format!", vbOKOnly + vbCritical, DialogBoxTitle
            txtFollowUpDate.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Please Enter Treatment Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtFollowUpDate.SetFocus
        Exit Sub
    End If
    
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then
        FOLNumber = GetFollowUpCode(ChkStr(CaseNumber))
                  
        strsql = "EXEC Case_SAVEFollowUpLG '" & ChkStr(txtCaseNo) & "', " & _
                "'" & ChkStr(FOLNumber) & "'," & txtCASLGAmt & ", '" & Logon & "', " & _
                "'" & TempDate & "', '" & Format(txtFollowUpDate, "yyyy/mm/dd") & "'"
        medixmasterconnMaint.Execute strsql
        MsgBox "Record Saved...", vbOKOnly + vbExclamation, DialogBoxTitle
        Call FollowUpListing
        Call funcword1
    End If
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    
End Sub


Private Sub cmdIntCASRem_Click()
    txtIntCASRemark.Visible = True
    txtCASRemark.Visible = False
    txtCASRemarkTwo.Visible = False
    lblPage.Caption = "Memo"
End Sub

Private Sub cmdInvList_Click()
    frmSearchInv.Source = 1
    frmSearchInv.Show vbModal
End Sub

Private Sub cmdOPCS_Click(Index As Integer)
    If Index = 1 Then
        frmSearchCPT.Source = 2
    Else
        frmSearchCPT.Source = 3
    End If
        frmSearchCPT.Show vbModal
End Sub

Private Sub cmdPage_Click()
    If cmdPage.Caption = "Next" Then
        txtCASRemark.Visible = False
        txtCASRemarkTwo.Visible = True
        cmdPage.Caption = "Previous"
        lblPage.Caption = "Page 2"
    Else
        txtCASRemark.Visible = True
        txtCASRemarkTwo.Visible = False
        cmdPage.Caption = "Next"
        lblPage.Caption = "Page 1"
    End If
        txtIntCASRemark.Visible = False
End Sub

Private Sub cmdPreUnd_Click()
    If Len(txtCaseNo) Then
        'Load frmCASNonDisclosure
        frmCASNonDisclosure.TempCASNumber = txtCaseNo
        frmCASNonDisclosure.Show
    End If
End Sub

Private Sub cmdPrint_Click()
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If txtCaseNo <> "" Then
        Letter = "M"
        Call funcword
    End If
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub cmdPrior_Click()
Dim Days As String
    
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If IsDate(txtMBMExpiryDate) Then
        Days = CDate(txtMBMExpiryDate) - Date
        If Days <= 45 And Days >= 0 Then
            Letter = "P"
            Call funcword
        ElseIf CDate(txtMBMExpiryDate) < CDate(CaseRegDate) Then
            Letter = "E"
            Call funcword
        End If
    End If
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    cmdShow.Enabled = True
    cmdLG.Enabled = True
    cmdRepudiate.Enabled = True
    cmdSaveAdd.Enabled = True
    cmdPrior.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdUpdate1.Enabled = True
End Sub

Private Sub cmdTHDrugs_Click()
    Letter = "T"
    Call funcword
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

Private Sub FollowUpListing()
Dim rst As New ADODB.Recordset
Dim strsql As String
Dim Follow As ListItem
    
    lstFollowUp.listitems.Clear
    strsql = "EXEC SearchFollowUpLG '" & ChkStr(txtCaseNo) & "'"
    rst.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
    
    Do Until rst.EOF
        With rst
            Set Follow = lstFollowUp.listitems.Add(, , !FollowUpNo)
            Follow.SubItems(1) = !LGAmt
            Follow.SubItems(2) = !FollowUpIssuedBy
            Follow.SubItems(3) = !FollowUpIssuedDate
            Follow.SubItems(4) = IIf(Len(!FollowupDate), !FollowupDate, "")
            .MoveNext
        End With
    Loop
    
End Sub

Private Sub CboClaimability_Change()
    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
        If cboREPCode1 = "" Then
            MsgBox "Please Select Repudiation reason", vbOKOnly + vbCritical, DialogBoxTitle
            cmdRepudiate.Enabled = False
            'Frame18.Enabled = True
            SSTab1.Tab = 1
        Else
            cmdRepudiate.Enabled = True
            'Frame18.Enabled = True
        End If
        cmdLG.Enabled = False
    Else
        cmdRepudiate.Enabled = False
        'Frame18.Enabled = False
        cmdLG.Enabled = True

    End If

End Sub

Private Sub cboClaimability_Click()
           
    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
        If cboREPCode1 = "" Then
            MsgBox "Please Select Repudiation reason", vbOKOnly + vbCritical, DialogBoxTitle
            cmdRepudiate.Enabled = False
            'Frame18.Enabled = True
            SSTab1.Tab = 1
        Else
            cmdRepudiate.Enabled = True
            'Frame18.Enabled = True
        End If
        cmdLG.Enabled = False
    Else
        cmdRepudiate.Enabled = False
        'Frame18.Enabled = False
        cmdLG.Enabled = True

    End If
End Sub

Private Sub CboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboClaimability, cboClaimability.SelStart) + Chr(KeyAscii) + Mid(cboClaimability, cboClaimability.SelStart + cboClaimability.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboClaimability.ListCount - 1
        
        If NewText = LCase(Left(cboClaimability.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboClaimability.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboClaimability = ValidValue
               cboClaimability.SelStart = 0
               cboClaimability.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboCondwillOcc_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCondwillOcc, cboCondwillOcc.SelStart) + Chr(KeyAscii) + Mid(cboCondwillOcc, cboCondwillOcc.SelStart + cboCondwillOcc.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboCondwillOcc.ListCount - 1
        
        If NewText = LCase(Left(cboCondwillOcc.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCondwillOcc.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboCondwillOcc = ValidValue
               cboCondwillOcc.SelStart = 0
               cboCondwillOcc.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCondwillOcc_LostFocus()
    If Len(cboCondwillOcc) > 7 Then
        MsgBox "Condition Will Occ cannot exceed 7 characters!", vbCritical
        cboCondwillOcc.SetFocus
    End If

End Sub

Private Sub cboFLWTreatment_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboFLWTreatment, cboFLWTreatment.SelStart) + Chr(KeyAscii) + Mid(cboFLWTreatment, cboFLWTreatment.SelStart + cboFLWTreatment.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboFLWTreatment.ListCount - 1
        
        If NewText = LCase(Left(cboFLWTreatment.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboFLWTreatment.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboFLWTreatment = ValidValue
               cboFLWTreatment.SelStart = 0
               cboFLWTreatment.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboFLWTreatment_LostFocus()
    If Len(cboFLWTreatment) > 7 Then
        MsgBox "FLW Treatment cannot exceed 7 characters!", vbCritical
        cboFLWTreatment.SetFocus
    End If

End Sub

Private Sub CboModality_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboModality, cboModality.SelStart) + Chr(KeyAscii) + Mid(cboModality, cboModality.SelStart + cboModality.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboModality.ListCount - 1
        
        If NewText = LCase(Left(cboModality.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboModality.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboModality = ValidValue
               cboModality.SelStart = 0
               cboModality.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboModality_LostFocus()
    If Len(cboModality) > 30 Then
        MsgBox "Modality cannot exceed 30 character!", vbCritical
        cboModality.SetFocus
    End If
End Sub

Private Sub cboNature_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboNature, cboNature.SelStart) + Chr(KeyAscii) + Mid(cboNature, cboNature.SelStart + cboNature.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboNature.ListCount - 1
        
        If NewText = LCase(Left(cboNature.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboNature.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboNature = ValidValue
               cboNature.SelStart = 0
               cboNature.SelLength = Len(ValidValue)
            End If
        End If
    End Sub

Private Sub cboNature_LostFocus()
    If Len(cboNature) > 30 Then
        MsgBox "Claim Nature cannot exceed 30 character!", vbCritical
        cboNature.SetFocus
    End If

End Sub

Private Sub cboPatientList_Change()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim strsql As String
Dim strsqlCOV As String
Dim is_integer As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

   If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
    
        If Mid(cboPatientList, 1, 2) = "00" Then
            'strsql = "Select MBMICBcPp from MBMCrossReference where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd.Execute
        
                If MBM.EOF = True Then
                    MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    txtPatientIC = MBM!MBMIcBcPp
                End If
            MBM.Close
            Set MBM = Nothing
        ElseIf Mid(cboPatientList, 1, 2) <> "00" Then
            'strsqlCOV = "Select MBMCICBCPPNo,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            'MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBMCov.CursorLocation = adUseClient
            MBMCov.CursorType = adOpenForwardOnly
            MBMCov.CacheSize = 100
            Set MBMCov = cmd.Execute
            
                If MBMCov.EOF = False Then
                    txtPatientIC = MBMCov!MBMCICBCPPNo
                    cboSuppExc = cboPatientList
                    If Len(MBMCov!MBMCPLNCode) Then
                        MBMCPLNCode = MBMCov!MBMCPLNCode
                    End If
                End If
            MBMCov.Close
            Set MBMCov = Nothing
        End If
    End If
 
End Sub

Private Sub cboPatientList_Click()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim strsql As String
Dim strsqlCOV As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

   If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
    
        If Mid(cboPatientList, 1, 2) = "00" Then
            'strsql = "Select MBMICBcPp from MBMCrossReference where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd.Execute
            
                If MBM.EOF = True Then
                    MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    txtPatientIC = MBM!MBMIcBcPp
                End If
            MBM.Close
            Set MBM = Nothing
        ElseIf Mid(cboPatientList, 1, 2) <> "00" Then
            'strsqlCOV = "Select MBMCICBCPPNo,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            'MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBMCov.CursorLocation = adUseClient
            MBMCov.CursorType = adOpenForwardOnly
            MBMCov.CacheSize = 100
            Set MBMCov = cmd.Execute
            
                If MBMCov.EOF = False Then
                    txtPatientIC = MBMCov!MBMCICBCPPNo
                    cboSuppExc = cboPatientList
                    If Len(MBMCov!MBMCPLNCode) Then
                        MBMCPLNCode = MBMCov!MBMCPLNCode
                    End If
                End If
            MBMCov.Close
            Set MBMCov = Nothing
        End If
    End If
End Sub

Private Sub cboPatientList_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPatientList, cboPatientList.SelStart) + Chr(KeyAscii) + Mid(cboPatientList, cboPatientList.SelStart + cboPatientList.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboPatientList.ListCount - 1
        
        If NewText = LCase(Left(cboPatientList.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPatientList.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboPatientList = ValidValue
               cboPatientList.SelStart = 0
               cboPatientList.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboPatientList_LostFocus()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim strsql As String
Dim strsqlCOV As String
Dim is_integer As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
        
   If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
    
        If Mid(cboPatientList, 1, 2) = "00" Then
            'strsql = "Select MBMICBcPp from MBMCrossReference where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd.Execute
            
                If MBM.EOF = True Then
                    MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    txtPatientIC = MBM!MBMIcBcPp
                End If
            MBM.Close
            Set MBM = Nothing
        ElseIf Mid(cboPatientList, 1, 2) <> "00" Then
            'strsqlCOV = "Select MBMCICBCPPNo,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            'MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd.Execute
            
                If MBMCov.EOF = False Then
                    txtPatientIC = MBMCov!MBMCICBCPPNo
                    cboSuppExc = cboPatientList
                    If Len(MBMCov!MBMCPLNCode) Then
                        MBMCPLNCode = MBMCov!MBMCPLNCode
                    End If
                End If
            MBMCov.Close
            Set MBMCov = Nothing
        End If
    End If
End Sub

Private Sub cboREPCode1_Change()
    If cboREPCode1 <> "" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    End If
End Sub

Private Sub cboREPCode1_Click()
    If cboREPCode1 <> "" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    End If
End Sub

Private Sub cboREPCode1_GotFocus()
    If cboREPCode1 <> "" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    End If
End Sub

Private Sub cboREPCode1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboREPCode1, cboREPCode1.SelStart) + Chr(KeyAscii) + Mid(cboREPCode1, cboREPCode1.SelStart + cboREPCode1.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboREPCode1.ListCount - 1
        
        If NewText = LCase(Left(cboREPCode1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboREPCode1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboREPCode1 = ValidValue
               cboREPCode1.SelStart = 0
               cboREPCode1.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboREPCode1_LostFocus()
    If cboREPCode1 <> "" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    End If
End Sub

Private Sub cboREPCode2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboREPCode2, cboREPCode2.SelStart) + Chr(KeyAscii) + Mid(cboREPCode2, cboREPCode2.SelStart + cboREPCode2.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboREPCode2.ListCount - 1
        
        If NewText = LCase(Left(cboREPCode2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboREPCode2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboREPCode2 = ValidValue
               cboREPCode2.SelStart = 0
               cboREPCode2.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboREPCode3_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboREPCode3, cboREPCode3.SelStart) + Chr(KeyAscii) + Mid(cboREPCode3, cboREPCode3.SelStart + cboREPCode3.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboREPCode3.ListCount - 1
        
        If NewText = LCase(Left(cboREPCode3.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboREPCode3.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboREPCode3 = ValidValue
               cboREPCode3.SelStart = 0
               cboREPCode3.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboSuppExc_Change()
Dim strsql As String
Dim MBMCov As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim conn As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
    

    If txtMBMNumber <> "" And IsNumeric(Mid(cboSuppExc, 1, 1)) = True Then
        'strsql = "Select MBMCnumber , MBMCCoverID , MBMCExclusion from MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Trim(Mid(cboSuppExc, 1, 2)) & "'"
        'MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockPessimistic
        strAppendCase = "2"
        strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Trim(Mid(cboSuppExc, 1, 2)) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        MBMCov.CursorLocation = adUseClient
        MBMCov.CursorType = adOpenForwardOnly
        MBMCov.CacheSize = 100
        Set MBMCov = cmd.Execute
        
        If MBMCov.EOF = False Then
            If Trim(MBMCov!MBMCExclusion) = "" Then
                txtSuppExc = MBMCov!MBMCExclusion
            End If
        End If
        MBMCov.Close
        Set MBMCov = Nothing
    End If
End Sub

Private Sub cboSuppExc_Click()
Dim strsql As String
Dim MBMCov As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

     medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    If txtMBMNumber <> "" And IsNumeric(Mid(cboSuppExc, 1, 1)) Then
        'strsql = "Select MBMCnumber , MBMCCoverID , MBMCExclusion from MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = " & Trim(Mid(cboSuppExc, 1, 2)) & ""
        'MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockPessimistic
        strAppendCase = "2"
        strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = " & Trim(Mid(cboSuppExc, 1, 2)) & ""
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        MBMCov.CursorLocation = adUseClient
        MBMCov.CursorType = adOpenForwardOnly
        MBMCov.CacheSize = 100
        Set MBMCov = cmd.Execute
        
        If MBMCov.EOF = False Then
            If Trim(MBMCov!MBMCExclusion) = "" Then
                txtSuppExc = MBMCov!MBMCExclusion
            End If
        End If
        MBMCov.Close
        Set MBMCov = Nothing
    End If
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub cboSuppExc_GotFocus()
Dim strsql As String
Dim MBMCov As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

     medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
     
    If txtMBMNumber <> "" And IsNumeric(Mid(cboSuppExc, 1, 1)) Then
        'strsql = "Select MBMCnumber , MBMCCoverID , MBMCExclusion from MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = " & Trim(Mid(cboSuppExc, 1, 2)) & ""
        'MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockPessimistic
        strAppendCase = "2"
        strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = " & Trim(Mid(cboSuppExc, 1, 2)) & ""
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        MBMCov.CursorLocation = adUseClient
        MBMCov.CursorType = adOpenForwardOnly
        MBMCov.CacheSize = 100
        Set MBMCov = cmd.Execute
        
        If MBMCov.EOF = False Then
            If Trim(MBMCov!MBMCExclusion) = "" Then
                txtSuppExc = MBMCov!MBMCExclusion
            End If
        End If
        MBMCov.Close
        Set MBMCov = Nothing
    End If
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub cmdExit_Click()
        Unload Me
        Set frmCaseAdjustment = Nothing
End Sub

Private Sub cmdLG_Click()
Dim RecordSaved
Dim CASLG As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
Dim strsqlPLN As String
Dim strsqlBTBG As String
Dim startTrans
Dim ChkVer As String
Dim TempSysDate As String
Dim TempLGAmt As String
Dim strsqlCASLG As String

Dim strsqlMBM As String
Dim Including As String
Dim Terms As String
Dim COPay As Boolean
Dim MRI As Boolean
Dim cmd2 As New ADODB.Command
Dim MBM As New ADODB.Recordset

Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
    
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If Trim(txtMBMNumber) = "" Then
        MsgBox "Please Enter Membership Information", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
    ElseIf Trim(cboPatientList) = "" Then
        MsgBox "Please Enter Patient Information", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
    ElseIf Trim(txtHOSCode) = "" Then
        MsgBox "Please Enter Hospital Information", vbOKOnly + vbCritical, DialogBoxTitle
        txtHOSCode.SetFocus
    ElseIf txtDateAdmitted = "__/__/____" Then
        MsgBox "Please Enter Admission Date", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
    ElseIf cboPatientList = "" Then
        MsgBox "Please Enter Patient's ID", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(txtCASReserveAmount) = "" Then
        MsgBox "Please Enter Reserve Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASReserveAmount.SetFocus
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
        
            Letter = ""
            Letter = "L"
            LGNumber = GetLGCode(ChkStr(CaseNumber))
            TempSysDate = Format(Date, "yyyy/mm/dd")
            TempLGAmt = CDbl(IIf(IsNumeric(txtCASLGAmt) = True, txtCASLGAmt, 0))
            
            strsqlPLN = "EXEC Case_Adjustment '" & Trim(txtMBMNumber) & "',3"
            PLN.Open strsqlPLN, medixmasterconn, adOpenForwardOnly, adLockOptimistic
            If Not PLN.EOF Then
                Letter = "B"
                LGNumber = "B" & Mid(LGNumber, 2, Len(LGNumber))
            End If
            PLN.Close
            Set PLN = Nothing
            
            If txtHLTCOde = "S" Then
                'strsqlMBM = "Select PLNCode, CoPaymentInd, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(TxtPlanCode) & "'"
                strAppendCase = "19"
                StrAppend1 = " AND PLNCode = '" & Trim(TxtPlanCode) & "'"
            Else
                'strsqlMBM = "Select PLNCode, CoPaymentInd, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(TxtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
                strAppendCase = "19"
                StrAppend1 = " AND PLNCode = '" & Trim(TxtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
            End If
            'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
            Set cmd2.ActiveConnection = medixmasterconn
            cmd2.CommandText = "Case_Adjustment"
            cmd2.CommandType = adCmdStoredProc
            cmd2.CommandTimeout = 300
            Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
            cmd2.Parameters.Append Prm1
            Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd2.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd2.Execute
            
            If MBM.EOF = False Then
                If Trim(MBM!PLNPerVisit) = "Y" Then
                    tmpPLNPerVisit = "Y"
                Else
                    tmpPLNPerVisit = ""
                End If
            End If
            MBM.Close
            Set MBM = Nothing
            
            ChkVer = ""
            ChkVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
            strsqlCASLG = "EXEC Case_AdjInsertCASLG '" & ChkStr(CaseNumber) & "', '" & Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-"))) & "', '" & TempSysDate & "', " & TempLGAmt & ", '" & Logon & "', '" & ChkStr(Trim(TxtLGRemark)) & "', '" & Letter & "', '" & tmpPLNPerVisit & "'"
            medixmasterconnMaint.Execute strsqlCASLG
            Call funcword
            cmdLG.Enabled = False
            
        End If
    End If
   
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    SndLG = False
    cmdShow.Enabled = True
    cmdLG.Enabled = True
    cmdRepudiate.Enabled = True
    cmdSaveAdd.Enabled = True
    cmdPrior.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdUpdate1.Enabled = True

End Sub

Private Sub cmdRepudiate_Click()
Dim REP As New ADODB.Recordset
Dim strsql As String
Dim RecordSaved
    
    Document = ""
    Personnel = ""
    Attention = ""
    If Trim(txtMBMNumber) = "" Then
        MsgBox "Please Enter Membership Information", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
    ElseIf Trim(cboPatientList) = "" Then
        MsgBox "Please Enter Patient Information", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
    ElseIf Trim(txtHOSCode) = "" Then
        MsgBox "Please Enter Hospital Information", vbOKOnly + vbCritical, DialogBoxTitle
        txtHOSCode.SetFocus
    ElseIf txtDateAdmitted = "__/__/____" Then
        MsgBox "Please Enter Admission Date", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
    ElseIf cboPatientList = "" Then
        MsgBox "Please Enter Patient's ID", vbOKOnly + vbCritical, DialogBoxTitle
        cboPatientList.SetFocus
    ElseIf Trim(txtCASReserveAmount) = "" Then
        MsgBox "Please Enter Reserve Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASReserveAmount.SetFocus
    ElseIf cboREPCode1 = "" Then
        MsgBox "Please Enter Repudiation Reason", vbOKOnly + vbCritical, DialogBoxTitle
        cboREPCode1.SetFocus
    Else
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "Exec Case_Adjustment '" & Trim(txtCaseNo) & "', 22"
        REP.Open strsql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
        If Not REP.EOF Then
            If REP!CASDocument <> "" And REP!CASPersonnel <> "" And REP!CASAttention <> "" Then
                With REP
                    frmRepudiationSelection.Document = !CASDocument
                    frmRepudiationSelection.Personnel = !CASPersonnel
                    frmRepudiationSelection.Attention = !CASAttention
                End With
            Else
                frmRepudiationSelection.Document = ""
                frmRepudiationSelection.Personnel = ""
                frmRepudiationSelection.Attention = ""
            End If
            
            If REP!CASUnderlying = "Y" Then
                frmRepudiationSelection.CASUnderlying = True
            Else
                frmRepudiationSelection.CASUnderlying = False
            End If
        Else
            frmRepudiationSelection.Document = ""
            frmRepudiationSelection.Personnel = ""
            frmRepudiationSelection.Attention = ""
            frmRepudiationSelection.CASUnderlying = False
        End If
        REP.Close
        Set REP = Nothing
        frmRepudiationSelection.Source = 1
        frmRepudiationSelection.Show vbModal
        
        If Document <> "" And Personnel <> "" And Attention <> "" Then
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                    Screen.MousePointer = vbHourglass
                    txtCASReserveAmount.SetFocus
                    Letter = ""
                    Letter = "R"
                    Call SaveRepudiation
                    Call funcword
                    cmdRepudiate.Enabled = False
                    cmdLG.Enabled = False
                    Screen.MousePointer = vbDefault
                End If
        End If
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
    End If
    
    cmdShow.Enabled = True
    cmdLG.Enabled = True
    cmdRepudiate.Enabled = True
    cmdSaveAdd.Enabled = True
    cmdPrior.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdUpdate1.Enabled = True
End Sub

Private Sub CmdNext_Click()
    Dim CAS As New ADODB.Recordset
    Dim sql As String
    Dim CASStatus As Boolean
    Dim CasNumber As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    Dim strOrderby As String
        
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    'sql = " select CASNumber from CASCrossReference where CASNumber > '" & RemStr(txtCaseNo) & "' " & _
        " Order by CASnumber "
    CAS.MaxRecords = 1
    'CAS.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    strAppendCase = "2"
    strAppend = " AND CASNumber > '" & RemStr(txtCaseNo) & "'"
    strOrderby = " Order by CASnumber "
    strAppend = strAppend + strOrderby
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    CAS.CursorLocation = adUseClient
    CAS.CursorType = adOpenForwardOnly
    CAS.CacheSize = 100
    Set CAS = cmd.Execute
    
    If CAS.EOF = False Then
        CASStatus = True
        CasNumber = CAS!CasNumber
    Else
        CASStatus = False
    End If
    
    CAS.Close
    Set CAS = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    cmdUpdate1.Visible = True

    If CASStatus = True Then
        txtCaseNo.Text = CasNumber
        cmdShow_Click
    Else
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
    End If
    cmdShow.Enabled = True
    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    Else
        cmdRepudiate.Enabled = False
        cmdLG.Enabled = True
    End If
    cmdSaveAdd.Enabled = True
    cmdPrior.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdUpdate1.Enabled = True

End Sub

Private Sub cmdPrev_Click()
    Dim CAS As New ADODB.Recordset
    Dim sql As String
    Dim CASStatus As Boolean
    Dim CasNumber  As String
    Dim StrsqlAppend As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    Dim strOrderby As String
           
    If Trim(txtCaseNo) <> "" Then
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        'sql = " select CASNumber from CASCrossReference where CASnumber < '" & RemStr(txtCaseNo) & "' " & _
            "Order by CASNumber DESC "
        CAS.MaxRecords = 1
        'CAS.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        strAppendCase = "2"
        strAppend = " AND CASnumber < '" & RemStr(txtCaseNo) & "'"
        strOrderby = "Order by CASNumber DESC "
        strAppend = strAppend + strOrderby
        Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        CAS.CursorLocation = adUseClient
        CAS.CursorType = adOpenForwardOnly
        CAS.CacheSize = 100
        Set CAS = cmd.Execute
        
        If CAS.EOF = False Then
            CASStatus = True
            CasNumber = CAS!CasNumber
        Else
            CASStatus = False
        End If
        
        CAS.Close
        Set CAS = Nothing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
        cmdUpdate1.Visible = True

        If CASStatus = True Then
            txtCaseNo.Text = CasNumber
            cmdShow_Click
        Else
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Case Number Format! ", vbCritical + vbOKOnly
        End If
    End If
    cmdShow.Enabled = True
    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    Else
        cmdRepudiate.Enabled = False
        cmdLG.Enabled = True
    End If
    cmdSaveAdd.Enabled = True
    cmdPrior.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdUpdate1.Enabled = True
End Sub

Private Sub cmdSaveAdd_Click()
Dim MBM As New ADODB.Recordset
Dim strsql As String
Dim RecordSaved
Dim strsqlSaveAdd As String
Dim Add1, Add2, Add3
    
    If txtAdd1 = "" Then
        MsgBox "Please Enter Address", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtPostCode = "" Then
        MsgBox "Please Enter Post Code", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtState = "" Then
        MsgBox "Please Enter State", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'strsql = "select MBMAddress1,MBMAddress2,MBMAddress3,CTYCode,MBMPostCode,STACode from MBM where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        'MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        RecordSaved = MsgBox("Do you want to Update Member's address?", vbYesNo + vbExclamation)
        Add1 = ChkStr(Replace(txtAdd1, "'", "''"))
        Add2 = ChkStr(Replace(txtAdd2, "'", "''"))
        Add3 = ChkStr(Replace(txtAdd3, "'", "''"))
        
        If RecordSaved = vbYes Then
            strsqlSaveAdd = " EXEC Case_AdjUpdateAdd '" & Trim(txtMBMNumber) & "', '" & Trim(Add1) & "', " & _
                            "'" & Trim(Add2) & "', '" & Trim(Add3) & "', '" & Trim(ChkStr(txtCity)) & "', " & _
                            "'" & ChkStr(txtPostCode) & "', '" & Trim(ChkStr(txtState)) & "'"
            medixmasterconn.Execute strsqlSaveAdd
            'With MBM
            '    !MBMAddress1 = Trim(ChkStr(txtAdd1))
            '    !MBMAddress2 = Trim(ChkStr(txtAdd2))
            '    !MBMAddress3 = Trim(ChkStr(txtAdd3))
            '    !CTYCode = Trim(ChkStr(txtCity))
            '    !MBMPostCode = ChkStr(txtPostCode)
            '    !STACode = Trim(ChkStr(txtState))
            'End With
            'MBM.Update
            MsgBox "Address Updated Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            'MBM.Close
            'Set MBM = Nothing
        End If
        medixmasterconn.Close
        Set medixmasterconn = Nothing
    End If
    cmdShow.Enabled = True
    cmdLG.Enabled = True
    cmdRepudiate.Enabled = True
    cmdSaveAdd.Enabled = True
    cmdPrior.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdUpdate1.Enabled = True
End Sub

Private Sub cmdSearch_Click()
    frmCASList.Source = 3
    frmCASList.Show vbModal
End Sub

Public Sub cmdShow_Click()
On Error GoTo ErrMsg
Dim getRecord As Integer
    RedAmt = 0
    REDAmtInd = False
    MBMCPLNCode = ""
    LGNo = ""
    LGVer = 0
    TxtLGRemark = ""
    Screen.MousePointer = vbHourglass

    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
    Else
        ScanCode = Trim(txtCaseNo)
        If Mid(ScanCode, 1, 2) = "LO" Then
            txtCaseNo = ""
            Call SearchForm(ScanCode)
        Else
            medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
            medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                TempCASNo = txtCaseNo
                patient = cboPatientList
                Call ClearForm(Me)
                Call ClearEntries
                lstLGListing.listitems.Clear
                lstFollowUp.listitems.Clear
                lstWorksheetListing.listitems.Clear
                'lstSurgProcList.listitems.Clear
                'lstNonSurgProcList.listitems.Clear
                cboPatientList.Clear
                txtCaseNo = TempCASNo
                MMARateCal = 0
                getRecord = showCAS
                MMARateCal = 1
            medixmasterconn.Close
            Set medixmasterconn = Nothing
            medixmasterconnMaint.Close
            Set medixmasterconnMaint = Nothing
            medixmasterconnWS.Close
            Set medixmasterconnWS = Nothing
        End If
    End If
    Screen.MousePointer = Default
    Exit Sub
    
ErrMsg:
    MsgBox Err.Description
    Screen.MousePointer = Default
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
End Sub

Private Function showCAS()
Dim CoverID As String
Dim HOS As New ADODB.Recordset
Dim strsqlHOS As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim RecordSaved
    CaseNumber = ""
    
        sqlstrCAS = "Exec SearchCasAdjustment '" & Trim(txtCaseNo) & "'"
        Set CASSearch = New ADODB.Recordset
        CASSearch.Open sqlstrCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
    If Not CASSearch.EOF Then
        showCAS = 1
        With CASSearch
            txtCaseNo = !CasNumber
            CaseNumber = Trim(!CasNumber)
            txtMBMNumber = Trim(!MBMNumber)
            If Len(!MBMCPLNCode) Then
                MBMCPLNCode = !MBMCPLNCode
            End If
            If !MBMPatientCovID = "00" Then
                cboPatientList = Trim(!MBMPatientCovID) & " - " & Trim(!MBMName)
                PatientName = Trim(!MBMName)
            Else
                cboPatientList = Trim(!MBMPatientCovID) & " - " & Trim(!MBMCName)
                cboSuppExc = Trim(!MBMPatientCovID) & " - " & Trim(!MBMCName)
                PatientName = Trim(!MBMCName)
            End If
            CaseRegDate = CDate(!CaseRegDate)
            If Len(!CASDateAdmitted) Then
                txtDateAdmitted.mask = ""
                txtDateAdmitted = Format(!CASDateAdmitted, "dd/mm/yyyy")
                txtDateAdmitted.mask = "##/##/####"
            End If
            If Len(!CASDischargeDate) Then
                txtDateDischarge.mask = ""
                txtDateDischarge = Format(!CASDischargeDate, "dd/mm/yyyy")
                txtDateDischarge.mask = "##/##/####"
            End If
            If !CASQuery <> "" Then
                cboQuery = !CASQuery
            End If
       'strsqlHOS = "Select * from HOS where HOSCode = '" & Trim(!CASHOSCode) & "'"
       'HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "6"
        strAppend = " AND HOSCode = '" & Trim(!CASHOSCode) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        HOS.CursorLocation = adUseClient
        HOS.CursorType = adOpenForwardOnly
        HOS.CacheSize = 100
        Set HOS = cmd.Execute
        
            If HOS.EOF = False Then
                 If Len(!CASHOSCode) Then
                     txtHOSCode = HOS!HOSName
                 End If
            End If
        HOS.Close
        Set HOS = Nothing
            
            If Len(!CASStayType) Then
                cboStayType = !CASStayType
            End If

            If Len(!CASPAYType) Then
                CboPayType = !CASPAYType
            End If
            
            If Len(!CASAdmissionReason) Then
                txtReason = !CASAdmissionReason
            End If
            If Len(!ClaimNatureCode) Then
                cboNature = !ClaimNatureCode
            End If
            If Len(!ModalityCode) Then
                cboModality = !ModalityCode
            End If
            If Len(!ReferedBy) Then
                txtREFBy = !ReferedBy
            End If
            If Len(!CASReserveAmount) Then
                txtCASReserveAmount = !CASReserveAmount
            End If
            If Len(!CASLGAmt) Then
                txtCASLGAmt = !CASLGAmt
            End If
            If Trim(!CASStatus) = "O" Then
                cboStatus1.ListIndex = 0
            ElseIf Trim(!CASStatus) = "C" Then
                cboStatus1.ListIndex = 1
            ElseIf Trim(!CASStatus) = "D" Then
                cboStatus1.ListIndex = 2
            End If
            
            If Trim(!CASClaimability) = "AA" Then
                cboClaimability.Text = "AA - Accepted"
             ElseIf Trim(!CASClaimability) = "RR" Then
                cboClaimability.Text = "RR - Rejected"
             ElseIf Trim(!CASClaimability) = "CC" Then
                cboClaimability.Text = "CC - Cancelled"
             ElseIf Trim(!CASClaimability) = "SA" Then
                cboClaimability.Text = "SA - Special Approval"
             ElseIf Trim(!CASClaimability) = "PP" Then
                cboClaimability.Text = "PP - Pay & Claim"
             ElseIf Trim(!CASClaimability) = "NN" Then
                cboClaimability.Text = "NN - Not Decided"
             ElseIf Trim(!CASClaimability) = "NA" Then
                cboClaimability.Text = "NA - Not Decided/Accepted"
             ElseIf Trim(!CASClaimability) = "NR" Then
                cboClaimability.Text = "NR - Not Decided/Rejected"
            ElseIf Trim(!CASClaimability) = "PA" Then
                cboClaimability.Text = "PA - Pay&Claim/Accepted"
            ElseIf Trim(!CASClaimability) = "PR" Then
                cboClaimability.Text = "PR - Pay&Claim/Rejected"
            ElseIf Trim(!CASClaimability) = "AR" Then
                cboClaimability.Text = "AR - Accepted/Rejected"
            ElseIf Trim(!CASClaimability) = "RA" Then
                cboClaimability.Text = "RA - Rejected/Accepted"
            ElseIf Trim(!CASClaimability) = "BR" Then
                cboClaimability.Text = "BR - BTBG/Rejected"
            ElseIf Trim(!CASClaimability) = "ER" Then
                cboClaimability.Text = "ER - Expired/Rejected"
                
            End If
             
             If Trim(!CASPending) = "Y" Then
                cboPending.ListIndex = 0
             ElseIf Trim(!CASPending) = "N" Then
                cboPending.ListIndex = 1
             End If
            If Len(!CASInd) Then
                cboCASInd = !CASInd
            End If
            If Len(!LINKCASNo) Then
                TxtLinkCASNo = !LINKCASNo
            End If
            
            If Len(!CASTopupLmt) Then
                cboTopUpLmt = !CASTopupLmt
            End If
            
            txtCASTakeOverAmt = IIf(Len(!CASTakeOverAmt), !CASTakeOverAmt, "")
            
            cboCASTakeOveInd = IIf(Len(!CASTakeOverInd), !CASTakeOverInd, "")
            
            txtCASTOPlan = IIf(Len(!CASTakeOverPlan), !CASTakeOverPlan, "")
        End With
        
            Call showMBM
            MBMNumber = Trim(txtMBMNumber)
            Call CallLimit
            Call showCASRemarks
            Call showOther
            Call showOther2
            Call ShowFirstDiag
            Call ShowFinalDiag
            Call ShowUD
            Call ShowRepudiation
            Call showMBMListing
            Call showMBMCovListing
            Call BNFListing
            Call showAuditReview
            If txtUpgPLNCode <> "" Or txtDOC <> "__/__/____" Then
                Call NewBNFListing
            End If
            
            Call GetAccountSummary
            If txtDOB <> "__/__/____" Then
                txtMorW = Year(Date) - Year(txtDOB)
                txtMorW = txtMorW & " Years Old"
            End If
        
        If Mid(cboStatus1, 1, 1) = "O" Then
            cboStatus1.Enabled = False
            If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
                cmdRepudiate.Enabled = True
                cmdLG.Enabled = False
            Else
                cmdRepudiate.Enabled = False
                cmdLG.Enabled = True
            End If
            cmdSaveAdd.Enabled = True
            cmdUpdate1.Enabled = True
            Frame2.Enabled = True
            'Frame7.Enabled = True
            'Frame15.Enabled = True
            'Frame16.Enabled = True
            'Frame17.Enabled = True
            txtDateDischarge.Enabled = True
            txtCASRemark.Enabled = True
            lblPage.Caption = "Page 1"
        ElseIf Mid(cboStatus1, 1, 1) = "C" Or Mid(cboStatus1, 1, 1) = "D" Then
            cboStatus1.Enabled = False
            cmdLG.Enabled = False
            cmdRepudiate.Enabled = False
            cmdSaveAdd.Enabled = False
            cmdUpdate1.Enabled = False
            Frame6.Enabled = False
            Frame2.Enabled = False
            'Frame7.Enabled = False
            'Frame15.Enabled = False
            'Frame16.Enabled = False
            'Frame17.Enabled = False
            txtDateDischarge.Enabled = False
            txtCASRemark.Enabled = False
        End If
        
        If Mid(cboClaimability, 1, 2) = "AA" And Mid(cboPending, 1, 1) = "Y" Then
            'UpdClaimability = True
            'Call cmdUpdate1_Click
            RecordSaved = MsgBox("Do you want to change Claimability Status?", vbYesNo + vbDefaultButton2 + vbExclamation)
            If RecordSaved = vbYes Then
                Call UpdateClaimability
                MsgBox "Status has been updated sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            End If
        End If
    Else
        cmdSaveAdd.Enabled = False
        cmdUpdate1.Enabled = False
        cmdLG.Enabled = False
        cmdRepudiate.Enabled = False
        MsgBox "No record found!", vbCritical + vbOKOnly
    End If
   
    CASSearch.Close
    Set CASSearch = Nothing
End Function

Private Function showCASRemarks()
    txtCASRemark = ""
    If Not CASSearch.EOF Then
        If Len(CASSearch!CASORemarks) Then
            txtCASRemark = Replace(CASSearch!CASORemarks, "~", Chr(13))
        End If
        If Len(CASSearch!CASORemarksTwo) Then
            txtCASRemark = txtCASRemark + Replace(CASSearch!CASORemarksTwo, "~", Chr(13))
        End If
        
        If Len(CASSearch!CASOTwoRemarks) Then
            txtCASRemarkTwo = Replace(CASSearch!CASOTwoRemarks, "~", Chr(13))
        End If
        If Len(CASSearch!CASOTwoRemarksTwo) Then
            txtCASRemarkTwo = txtCASRemarkTwo + Replace(CASSearch!CASOTwoRemarksTwo, "~", Chr(13))
        End If
        
        If Len(CASSearch!CASORemarksInt) Then
            txtIntCASRemark = Replace(CASSearch!CASORemarksInt, "~", Chr(13))
            lblMemo.Visible = True
        End If
    End If
    
End Function

Private Function showMBM()
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim sqlstr1 As String
Dim strsqlMBMExc As String
Dim MBMCov As New ADODB.Recordset
Dim MBMExc As New ADODB.Recordset

    'sqlstr = " Select MBMNumber, HLTCode, MBMPolicyNo,  MBMName ," & _
    '        " MBMIcBcPp,  MBMMemberType , " & _
    '        " MBMStatus, PAYCode, AGECode , " & _
    '        " INSCode, PLNCode, " & _
    '        " MBMAddress1, MBMAddress2, " & _
    '        " MBMAddress3, MBMPostCode, CTYCode, STACode, " & _
    '        " MBMDOB , BRNCode, MBMTakeOver , MBMRenewFlag,  MBMAvailableLimit, " & _
    '        " MBMPayorEffDate, MBMPayorExpDate, MBMENDtEffDate , MBMNewPlanCode, MBMNewPolEffDate "
    'sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic

    'strsqlMBMExc = " Select MBMNumber, MBMExclusion From MBMOthersTwo Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    'MBMExc.Open strsqlMBMExc, medixmasterconn, adOpenKeyset, adLockOptimistic

    sqlstr = "Exec SearchMBMCas '" & Trim(txtMBMNumber) & "'"
    
    Set MBM = New ADODB.Recordset
    MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    If Not MBM.EOF Then
        With MBM
            If Len(!MBMNumber) Then
                txtMBMNumber = Trim(!MBMNumber)
            End If
            If Len(!MBMName) Then
                txtMBMName = Trim(!MBMName)
            End If
            If Len(!MBMPolicyNo) Then
                txtMBMPolNo = Trim(!MBMPolicyNo)
            End If
            If Len(!MBMIcBcPp) Then
                txtMBMIcBcPP = Trim(!MBMIcBcPp)
            End If
            If Len(!MBMAddress1) Then
                txtAdd1 = Trim(!MBMAddress1)
            End If
            If Len(!MBMAddress2) Then
                txtAdd2 = Trim(!MBMAddress2)
            End If
            If Len(!MBMAddress3) Then
                txtAdd3 = Trim(!MBMAddress3)
            End If
            If Len(!PAYCode) Then
                txtPayCode = Trim(!PAYCode)
            End If
            If Len(Format(!MBMPayorEffDate, "dd/mm/yyyy")) Then
                txtMBMEffDate.mask = ""
                txtMBMEffDate.Text = Trim(Format(MBM!MBMPayorEffDate, "dd/mm/yyyy"))
                txtMBMEffDate.mask = "##/##/####"
            End If
            If Len(!HLTCode) Then
                txtHLTCOde = Trim(!HLTCode)
            End If
            If Len(Format(!MBMPayorEXPDate, "dd/mm/yyyy")) Then
                txtMBMExpiryDate.mask = ""
                txtMBMExpiryDate = Trim(Format(!MBMPayorEXPDate, "dd/mm/yyyy"))
                txtMBMExpiryDate.mask = "##/##/####"
            End If
            If Len(!PLNCode) Then
                txtPLNCode = Trim(!PLNCode)
            End If
            If Len(!INSCode) Then
                txtINSCode1 = Trim(MBM!INSCode)
            End If
            If Len(Format(!MBMDOB, "dd/mm/yyyy")) Then
                txtDOB.mask = ""
                txtDOB = Trim(Format(!MBMDOB, "dd/mm/yyyy"))
                txtDOB.mask = "##/##/####"
            End If
            If Len(!CTYCode) Then
                txtCity = Trim(!CTYCode)
            End If
            If Len(!MBMPostCode) Then
                txtPostCode = Trim(!MBMPostCode)
            End If
                If Len(Trim(!STACode)) Then
                    txtState = Trim(!STACode)
                End If
                If Trim(!MBMTakeover) <> "" Then
                    txtTakeOver = Trim(!MBMTakeover)
                End If
                If Len(!MBMRenewFlag) Then
                    txtRenewal = !MBMRenewFlag
                End If
                TxtPlanCode = Trim(!PLNCode)
                If Len(!MBMNewPlanCode) Then
                    txtUpgPLNCode = Trim(!MBMNewPlanCode)
                End If
                If Len(Format(!MBMNewPolEffDate, "dd/mm/yyyy")) Then
                    txtDOC.mask = ""
                    txtDOC = Trim(Format(!MBMNewPolEffDate, "dd/mm/yyyy"))
                    txtDOC.mask = "##/##/####"
                End If
                
                txtAvailableAnnualLimit = Format(Trim(!MBMAvailableLimit), "#,##0.00")
                txtPolStatus = Trim(!MBMStatus)
                
                If Len(!MBMExclusion) Then
                    txtPrincipalExc = !MBMExclusion
                End If
                
                If CDate(txtMBMExpiryDate) < Date Then
                    lblExpired.Visible = True
                    lblExpired.Caption = "Expired"
                ElseIf CDate(txtMBMExpiryDate) - Date <= 30 Then
                    lblExpired.Visible = True
                    lblExpired.Caption = "< 30 Days to Expiry"
                Else
                    lblExpired.Visible = False
                End If
    
                'sqlstr1 = " select MBMCNumber, MBMCCoverID, MBMCName, " & _
                ' "  MBMCAnnualLimit " & _
                ' " from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    
                'MBMCov.Open sqlstr1, medixmasterconn, adOpenKeyset, adLockOptimistic
    
                cboPatientList.AddItem "00 -" & txtMBMName
                'cboSuppExc.AddItem "0 -" & txtMBMName
            
                If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
                    Do While Not MBM.EOF
                        cboPatientList.AddItem MBM!MBMCCoverID & "-" & MBM!MBMCName
                        cboSuppExc.AddItem MBM!MBMCCoverID & "-" & MBM!MBMCName
                        MBM.MoveNext
                    Loop
                End If
            End With
          End If
    
    MBM.Close
    'MBMCov.Close
    'MBMExc.Close
    Set MBM = Nothing
    'Set MBMCov = Nothing
    'Set MBMExc = Nothing
End Function

Public Sub CallLimit()
Dim sqlstr As String
Dim AnnualLimit As String
Dim PLN As String
Dim INS As String
Dim HLT As String
Dim Effdate As Date
Dim MBM As New ADODB.Recordset
    'sqlstr = "Select PLNCode, INSCode, HLTCode, MBMPayorEffDate from MBM Where MBMNumber = '" & RemStr(txtMBMNumber) & "'"
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'PLN = Trim(MBM!PLNCode)
    'INS = Trim(MBM!INSCode)
    'HLT = Trim(MBM!HLTCode)
    
    Effdate = CDate(txtMBMEffDate)
    
    txtAnnualLimit = ""
    'If Len(Trim(cboPLNCode)) Then
        txtAnnualLimit = Format(GetAnnualLimit(txtINSCode1, txtPLNCode, txtHLTCOde, Effdate), "#,##0.00")
    'End If
            
        'txtAnnualLimit = ""
        'txtAnnualLimit = Format(GetAnnualLimit(INS, PLN, HLT, EffDate), "#,##0.00")

End Sub

Private Function showOther()
Dim sqlstr As String
Dim Other As New ADODB.Recordset
Dim i As Integer
Dim CPTList As ListItem
Dim CPTStatus As Boolean
Dim ProcReq As String
Dim CPTCode As String
Dim ActSurgFees As Double
Dim ActAnaeFees As Double
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
Dim TotalActSurgFees As Double
Dim TotalActAnaeFees As Double
Dim TotalMMASurgFees As Double
Dim TotalMMAAnaeFees As Double
    
    'txtCASRemark = ""
    'txtCASAdmissionReason = ""
    'txtCASRepudiation = ""
    
    'sqlstr = "Select * From CASOthers where CASNumber = '" & RemStr(txtCaseNo) & "'"
    'Other.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TotalActSurgFees = 0
    TotalActAnaeFees = 0
    TotalMMASurgFees = 0
    TotalMMAAnaeFees = 0
    
    If Not CASSearch.EOF Then
        With CASSearch
            If Len(Trim(!CASOCondWillRecur)) Then
                 cboCondwillOcc = !CASOCondWillRecur
            End If
            If Len(Trim(!CASOFollowUpRequired)) Then
                 cboFLWTreatment = !CASOFollowUpRequired
            End If
            If Len(Trim(!CASFolUp)) Then
                 cboFolUp = !CASFolUp
            End If
            
            If Len(!CASFirstCall) Then
                txtFirstCall.mask = ""
                txtFirstCall = Format(!CASFirstCall, "dd/mm/yyyy")
                txtFirstCall.mask = "##/##/####"
            End If
            If Len(!CASFirstTreated) Then
                txtFirstTreated.mask = ""
                txtFirstTreated = Format(!CASFirstTreated, "dd/mm/yyyy")
                txtFirstTreated.mask = "##/##/####"
            End If
            If Len(!CASOPROCRequired1) Then
                txtProcedure1 = !CASOPROCRequired1
            End If
            
            If Len(!CASOPROCRequired2) Then
                txtProcedure2 = !CASOPROCRequired2
            End If

            If Len(!CASOPROCRequired3) Then
                 txtProcedure3 = !CASOPROCRequired3
            End If

            If Len(!CASOPROCRequired4) Then
               txtProcedure4 = !CASOPROCRequired4
            End If

            If Len(!CASOPROCRequired5) Then
                 txtProcedure5 = !CASOPROCRequired5
            End If

            If Len(!CASOProcInvestigationDone1) Then
                 txtPROCINVDone1 = !CASOProcInvestigationDone1
            End If
            
            If Len(!CASOProcInvestigationDone2) Then
                 txtPROCINVDone2 = !CASOProcInvestigationDone2
            End If

            If Len(!CASOProcInvestigationDone3) Then
                 txtPROCINVDone3 = !CASOProcInvestigationDone3
            End If

            If Len(!CASOMedSurMgmt1) Then
                 txtMEDSURDone1 = !CASOMedSurMgmt1
            End If
            
            If Len(!CASOMedSurMgmt2) Then
                 txtMEDSURDone2 = !CASOMedSurMgmt2
            End If

            If Len(!CASOMedSurMgmt3) Then
                 txtMEDSURDone3 = !CASOMedSurMgmt3
            End If
            
            If Len(!CASDRD) Then
                txtDateRECDoc = CDate(!CASDRD)
            End If
            
            If (!CASRedAmt) Then
                txtREDAmt = !CASRedAmt
                REDAmtInd = True
                RedAmt = !CASRedAmt
            End If
            
            If Len(!CASCptCode1) Then
                txtProcCode1 = !CASCptCode1
                'Call SearchMMARate(!CASCptCode1)
                'txtSMMARate1 = MMASurgFees
                'txtAnaesMMARate1 = MMAAnaeFees
            End If
            
            If Len(!CASCptCode2) Then
                txtProcCode2 = !CASCptCode2
                'Call SearchMMARate(!CASCptCode2)
                'txtSMMARate2 = MMASurgFees
                'txtAnaesMMARate2 = MMAAnaeFees
            End If
            
            If Len(!CASCptCode3) Then
                txtProcCode3 = !CASCptCode3
                'Call SearchMMARate(!CASCptCode3)
                'txtSMMARate3 = MMASurgFees
                'txtAnaesMMARate3 = MMAAnaeFees
            End If
            
            If Len(!CASCptCode4) Then
                txtProcCode4 = !CASCptCode4
                'Call SearchMMARate(!CASCptCode4)
                'txtSMMARate4 = MMASurgFees
                'txtAnaesMMARate4 = MMAAnaeFees
            End If
            
            If Len(!CASCptCode5) Then
                txtProcCode5 = !CASCptCode5
                'Call SearchMMARate(!CASCptCode5)
                'txtSMMARate5 = MMASurgFees
                'txtAnaesMMARate5 = MMAAnaeFees
            End If
            
            If Len(!CASInvCode1) Then
                txtINVCode1 = !CASInvCode1
            End If
            
            If Len(!CASInvCode2) Then
                txtINVCode2 = !CASInvCode2
            End If
            
            If Len(!CASInvCode3) Then
                txtINVCode3 = !CASInvCode3
            End If
            
            If Len(!CASPayPending) Then
                cboPayPending = !CASPayPending
            End If
            
            If Len(!CASONonSurgPROCRequired1) Then
                txtNSProcedure1 = !CASONonSurgPROCRequired1
            End If
            
            If Len(!CASONonSurgPROCRequired2) Then
               txtNSProcedure2 = !CASONonSurgPROCRequired2
            End If

            If Len(!CASONonSurgPROCRequired3) Then
                 txtNSProcedure3 = !CASONonSurgPROCRequired3
            End If

            If Len(!CASONonSurgPROCRequired4) Then
               txtNSProcedure4 = !CASONonSurgPROCRequired4
            End If

            If Len(!CASONonSurgPROCRequired5) Then
                 txtNSProcedure5 = !CASONonSurgPROCRequired5
            End If
            
            If Len(!CASNonSurgCptCode1) Then
                txtNSProcCode1 = !CASNonSurgCptCode1
            End If
            
            If Len(!CASNonSurgCptCode2) Then
                txtNSProcCode2 = !CASNonSurgCptCode2
            End If
            
            If Len(!CASNonSurgCptCode3) Then
                txtNSProcCode3 = !CASNonSurgCptCode3
            End If
            
            If Len(!CASNonSurgCptCode4) Then
                txtNSProcCode4 = !CASNonSurgCptCode4
            End If
            
            If Len(!CASNonSurgCptCode5) Then
                txtNSProcCode5 = !CASNonSurgCptCode5
            End If
            
            txtSActFee1 = IIf(Len(!CASProcActSurgFees1), !CASProcActSurgFees1, 0)
            txtSActFee2 = IIf(Len(!CASProcActSurgFees2), !CASProcActSurgFees2, 0)
            txtSActFee3 = IIf(Len(!CASProcActSurgFees3), !CASProcActSurgFees3, 0)
            txtSActFee4 = IIf(Len(!CASProcActSurgFees4), !CASProcActSurgFees4, 0)
            txtSActFee5 = IIf(Len(!CASProcActSurgFees5), !CASProcActSurgFees5, 0)
            
            txtAnaesActFees1 = IIf(Len(!CASProcActAnaeFees1), !CASProcActAnaeFees1, 0)
            txtAnaesActFees2 = IIf(Len(!CASProcActAnaeFees2), !CASProcActAnaeFees2, 0)
            txtAnaesActFees3 = IIf(Len(!CASProcActAnaeFees3), !CASProcActAnaeFees3, 0)
            txtAnaesActFees4 = IIf(Len(!CASProcActAnaeFees4), !CASProcActAnaeFees4, 0)
            txtAnaesActFees5 = IIf(Len(!CASProcActAnaeFees5), !CASProcActAnaeFees5, 0)
            
            txtNSActFee1 = IIf(Len(!CASProcActNonSurgFees1), !CASProcActNonSurgFees1, 0)
            txtNSActFee2 = IIf(Len(!CASProcActNonSurgFees2), !CASProcActNonSurgFees2, 0)
            txtNSActFee3 = IIf(Len(!CASProcActNonSurgFees3), !CASProcActNonSurgFees3, 0)
            txtNSActFee4 = IIf(Len(!CASProcActNonSurgFees4), !CASProcActNonSurgFees4, 0)
            txtNSActFee5 = IIf(Len(!CASProcActNonSurgFees5), !CASProcActNonSurgFees5, 0)
            
        End With
    
    End If
End Function

Private Function showOther2()
    
    If Not CASSearch.EOF Then
        
        With CASSearch
            If Len(!CASAttendDoctor1) Then
                txtATTDoc(0) = Trim(!CASAttendDoctor1)
            End If
            If Len(!CASAttendDoctor2) Then
                txtATTDoc(1) = Trim(!CASAttendDoctor2)
            End If
            If Len(!CASAttendDoctor3) Then
                txtATTDoc(2) = Trim(!CASAttendDoctor3)
            End If
            If Len(!CASAttendDoctor4) Then
                txtATTDoc(3) = Trim(!CASAttendDoctor4)
            End If
            If Len(!CASAttendDoctor5) Then
                txtATTDoc(4) = Trim(!CASAttendDoctor5)
            End If
            If Len(!CASAnaesName1) Then
                txtAnaesDoc(0) = Trim(!CASAnaesName1)
            End If
            If Len(!CASAnaesName2) Then
                txtAnaesDoc(1) = Trim(!CASAnaesName2)
            End If
            If Len(!CASAnaesName3) Then
                txtAnaesDoc(2) = Trim(!CASAnaesName3)
            End If
                                
            If Len(!CASFirstDocDate) Then
                txtFirstDocDate = !CASFirstDocDate
            End If
            If Len(!CASSubSeqDocDate) Then
                txtSubDocDate = !CASSubSeqDocDate
            End If
            cboIntExtTO = IIf(Len(!CASIntExtTO), !CASIntExtTO, "")
                                
        End With
        
    End If
    
End Function

Private Sub cmdUpdate1_Click()
Dim strsqlCAS As String
Dim CAS As New ADODB.Recordset
Dim CASHIS As New ADODB.Recordset
Dim CASHOS As New ADODB.Recordset
Dim DIAHIS As New ADODB.Recordset
Dim DIA As New ADODB.Recordset
Dim AUT As New ADODB.Recordset
Dim sqlstr As String
Dim sqlstr1 As String
Dim diafirstsql As String
Dim diafinalsql As String
Dim diafirst As New ADODB.Recordset
Dim diafinal As New ADODB.Recordset
Dim listrecord As Integer
Dim listloop As Integer
Dim diagnosistr As String
Dim hoshistory As String
Dim RecordSaved
Dim AdjustmentSeqNo As Integer
Dim ListCount As Integer
Dim CaseNumber As String
Dim CASDRD As String
Dim strsqlREP As String
Dim REP As New ADODB.Recordset
Dim StrSqlCRef As String
Dim CASCRef As New ADODB.Recordset
Dim CASMBMNumber As String
Dim CASMBMID As String
Dim HOSCode As String
Dim CASRMK As New ADODB.Recordset
Dim strsqlRMK As String
Dim cmd As New ADODB.Command
Dim cmd1 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim Prm3 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim LINKCASNo, CASInd, DateAdmitted, DateDischarge, LogonDate
Dim CASRemark, CASRemarkTwo, Reason, RefBy
Dim TempTakeOverLmt As Double
Dim strSearchRMK As String
Dim CASRemarksInt As String

    UD1 = 0
    FirstDiag = 0
    FinalDiag = 0
    CASMBMNumber = ""
    CASMBMID = ""
    HOSCode = ""
    TempTakeOverLmt = 0
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No!", vbCritical + vbOKOnly, DialogBoxTitle
        txtCaseNo.SetFocus
    ElseIf cboClaimability = "" Then
        MsgBox "Please Select Claimability!", vbCritical + vbOKOnly, DialogBoxTitle
        cboClaimability.SetFocus
    ElseIf txtMBMNumber = "" Then
        MsgBox "Please Enter Member No!", vbCritical + vbOKOnly, DialogBoxTitle
        txtMBMNumber.SetFocus
    ElseIf txtMBMName = "" Then
        MsgBox "Please Enter Member Name!", vbCritical + vbOKOnly, DialogBoxTitle
        txtMBMName.SetFocus
    ElseIf txtMBMPolNo = "" Then
        MsgBox "Please Enter Policy No!", vbCritical + vbOKOnly, DialogBoxTitle
        txtMBMPolNo.SetFocus
    ElseIf txtDateAdmitted = "__/__/____" Then
        MsgBox "Please Enter Admission Date", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
        txtDateAdmitted.SetFocus
    ElseIf cboTopUpLmt = "" Then
        MsgBox "Please Enter Topup Limit Indicator", vbOKOnly + vbCritical, DialogBoxTitle
        cboTopUpLmt.SetFocus
    ElseIf Mid(cboIntExtTO, 1, 1) = "I" And Trim(txtCASTOPlan) = "" Then
        MsgBox "Please Enter Takeover Plan", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASTOPlan.SetFocus
    'ElseIf txtPROCRequired1 <> "" And txtProcCode1 = "" Then
    '    MsgBox "Please Enter Procedure Code", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtPROCRequired1.SetFocus
    'ElseIf txtPROCRequired2 <> "" And txtProcCode2 = "" Then
    '    MsgBox "Please Enter Procedure Code", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtProcCode2.SetFocus
    'ElseIf txtPROCRequired3 <> "" And txtProcCode3 = "" Then
    '    MsgBox "Please Enter Procedure Code", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtProcCode3.SetFocus
    'ElseIf txtPROCRequired4 <> "" And txtProcCode4 = "" Then
    '    MsgBox "Please Enter Procedure Code", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtProcCode4.SetFocus
    'ElseIf txtPROCRequired5 <> "" And txtProcCode5 = "" Then
    '    MsgBox "Please Enter Procedure Code", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtProcCode5.SetFocus
        
    Else
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                    TempTakeOverLmt = IIf(IsNumeric(txtCASTakeOverAmt) = True, txtCASTakeOverAmt, 0)
                'AdjustmentSeqNo = GetAdjustmentNo(txtMBMNumber, txtMBMPolNo, txtCaseNo, "C") ' Get Adjustment Sequence No
                '---------------------------------------------------------------------------------------------------------------------------------------
                'Duplicate Copy for Case and put into Case History CAS -> CASHIS
    
                'updatecase = "Select * from CAS Where CASNumber ='" & Trim(tempcasNo) & "' And MBMNumber ='" & Trim(MBMNumber) & " '"
                'CAS.Open updatecase, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "7"
                strAppend = " AND CASNumber ='" & Trim(TempCASNo) & "' And MBMNumber ='" & Trim(MBMNumber) & " '"
                Set cmd.ActiveConnection = medixmasterconn
                cmd.CommandText = "Case_Adjustment"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd.Parameters.Append Prm2
                CAS.CursorLocation = adUseClient
                CAS.CursorType = adOpenForwardOnly
                CAS.CacheSize = 100
                Set CAS = cmd.Execute
                
                With CAS
                    CASMBMNumber = !MBMNumber
                    CASMBMID = !MBMPatientCovID
                    HOSCode = !CASHOSCode
                End With
            '----------7-----------------------------------------------------------------------------------------------------------------------------
            
            'Update the Cas Fields
                'With CAS
                    'Dim Test As String
                    '    Test = Mid(cboClaimability, 1, 2)
                '    !CASStatus = Mid(cboStatus1, 1, 1)
                '    !CASClaimability = Mid(cboClaimability, 1, 1)
                    If txtDateAdmitted <> "__/__/____" Then
                        DateAdmitted = Format(txtDateAdmitted, "mm/dd/yyyy")
                    Else
                        DateAdmitted = Empty
                    End If
                    If txtDateDischarge <> "__/__/____" Then
                        DateDischarge = Format(txtDateDischarge, "mm/dd/yyyy")
                    Else
                        DateDischarge = Empty
                    End If
                    If txtDateRECDoc <> "__/__/____" Then
                        CASDRD = Format(txtDateRECDoc, "mm/dd/yyyy")
                    Else
                        CASDRD = Empty
                    End If
                '    !CASRemark = txtCASRemark
                '    !ModalityCode = Mid(cboModality, 1, 1)
                '    !PayTypeCode = Mid(CboClaimType, 1, 2)
                '    !CASStayType = Mid(cboStayType, 1, 1)
                '    !CASPAYType = Mid(CboPayType, 1, 1)
                '    !ClaimNatureCode = Mid(cboNature, 1, 3)
                '    !CASPending = Mid(cboPending, 1, 1)
                '    !AttendDoctor = ChkStr(TempDRName)
                '    !ReferedBy = ChkStr(txtREFBy)
                '    If txtCASReserveAmount <> "" Then
                '        !CASReserveAmount = txtCASReserveAmount
                '    End If
                    If TxtLinkCASNo <> "" Then
                        LINKCASNo = ChkStr(TxtLinkCASNo)
                    Else
                        LINKCASNo = Null
                    End If
                    If Len(cboCASInd) Then
                        CASInd = Mid(cboCASInd, 1, 1)
                    Else
                        CASInd = Null
                    End If
                    LogonDate = Format(Date, "mm/dd/yyyy")
                    CASRemark = ChkStr(Replace(txtCASRemark, "'", "''"))
                    CASRemark = ChkStr(Replace(CASRemark, Chr(13), "~"))
                    CASRemarkTwo = ChkStr(Replace(txtCASRemarkTwo, "'", "''"))
                    CASRemarkTwo = ChkStr(Replace(CASRemarkTwo, Chr(13), "~"))
                    CASRemarksInt = ChkStr(Replace(txtIntCASRemark, "'", "''"))
                    CASRemarksInt = ChkStr(Replace(CASRemarksInt, Chr(13), "~"))
                    
                    
                '    !CASLastUpdateUser = Logon
                '    !CASLastUpdateDate = Date
                '    !CASAdmissionReason = ChkStr(txtReason)
                '    If Len(txtCASLGAmt) Then
                '        !CASLGAmt = txtCASLGAmt
                '    End If
                '    !CASTopupLmt = Mid(cboTopUpLmt, 1, 1)
                'End With
                'CAS.Update
                
                strsqlCAS = "EXEC Case_AdjUpdateCAS '" & Trim(TempCASNo) & "', " & _
                            "'" & Mid(cboStatus1, 1, 1) & "', '" & Mid(cboClaimability, 1, 2) & "', '" & DateAdmitted & "', '" & DateDischarge & "', " & _
                            "'" & Mid(cboModality, 1, 1) & "', " & _
                            "'" & Mid(cboStayType, 1, 1) & "', '" & Mid(CboPayType, 1, 1) & "', " & _
                            "'" & Mid(cboNature, 1, 3) & "', '" & Mid(cboPending, 1, 1) & "', " & _
                            "" & txtCASReserveAmount & ", '" & LINKCASNo & "', " & _
                            "'" & CASInd & "', '" & Logon & "', '" & LogonDate & "', " & txtCASLGAmt & ", '" & Mid(cboTopUpLmt, 1, 1) & "', " & _
                            "'" & CASDRD & "', '" & Mid(cboQuery, 1, 1) & "', " & TempTakeOverLmt & ", '" & Mid(cboCASTakeOveInd, 1, 1) & "', '" & txtCASTOPlan & "'"
                medixmasterconn.Execute strsqlCAS
            
                'strsqlRMK = "Select CASNumber,CASORemarks from CASRemarks Where CASNumber ='" & Trim(tempcasNo) & "'"
                'CASRMK.Open strsqlRMK, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "8"
                strAppend = " AND CASNumber ='" & Trim(TempCASNo) & "'"
                Set cmd1.ActiveConnection = medixmasterconn
                cmd1.CommandText = "Case_Adjustment"
                cmd1.CommandType = adCmdStoredProc
                cmd1.CommandTimeout = 300
                Set Prm1 = cmd1.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd1.Parameters.Append Prm1
                Set Prm2 = cmd1.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd1.Parameters.Append Prm2
                CASRMK.CursorLocation = adUseClient
                CASRMK.CursorType = adOpenForwardOnly
                CASRMK.CacheSize = 100
                Set CASRMK = cmd1.Execute
                
                If CASRMK.EOF = True Then
                    If Len(CASRemark) <= 1900 Then
                        strsqlRMK = "EXEC Case_AdjInsertCASRemark '" & Trim(TempCASNo) & "', '" & CASRemark & "', ''"
                    ElseIf Len(CASRemark) > 1900 Then
                        strsqlRMK = "EXEC Case_AdjInsertCASRemark '" & Trim(TempCASNo) & "', '" & Mid(CASRemark, 1, 1900) & "', '" & Mid(CASRemark, 1901, Len(CASRemark) - 1900) & "'"
                    End If
                Else
                    If Len(CASRemark) <= 1900 Then
                        strsqlRMK = "EXEC Case_AdjUpdateCASRemark '" & Trim(TempCASNo) & "', '" & CASRemark & "', ''"
                    ElseIf Len(CASRemark) > 1900 Then
                        strsqlRMK = "EXEC Case_AdjUpdateCASRemark '" & Trim(TempCASNo) & "', '" & Mid(CASRemark, 1, 1900) & "', '" & Mid(CASRemark, 1901, Len(CASRemark) - 1900) & "'"
                    End If
                End If
                medixmasterconn.Execute strsqlRMK
                CASRMK.Close
                Set CASRMK = Nothing
                
                strSearchRMK = "Select CASNumber from CASRemarksTwo Where CASNumber ='" & Trim(TempCASNo) & "'"
                CASRMK.Open strSearchRMK, medixmasterconn, adOpenKeyset, adLockOptimistic
                If CASRMK.recordCount Then
                    If Len(CASRemarkTwo) <= 1900 Then
                        strsqlRMK = "EXEC Case_AdjUpdateCASRemark2 '" & Trim(TempCASNo) & "', '" & CASRemarkTwo & "', ''"
                    ElseIf Len(CASRemarkTwo) > 1900 Then
                        strsqlRMK = "EXEC Case_AdjUpdateCASRemark2 '" & Trim(TempCASNo) & "', '" & Mid(CASRemarkTwo, 1, 1900) & "', '" & Mid(CASRemarkTwo, 1901, Len(CASRemarkTwo) - 1900) & "'"
                    End If
                Else
                    If Len(CASRemarkTwo) <= 1900 Then
                        strsqlRMK = "EXEC Case_AdjInsertCASRemark2 '" & Trim(TempCASNo) & "', '" & CASRemarkTwo & "', ''"
                    ElseIf Len(CASRemarkTwo) > 1900 Then
                        strsqlRMK = "EXEC Case_AdjInsertCASRemark2 '" & Trim(TempCASNo) & "', '" & Mid(CASRemarkTwo, 1, 1900) & "', '" & Mid(CASRemarkTwo, 1901, Len(CASRemarkTwo) - 1900) & "'"
                    End If
                    
                End If
                    medixmasterconn.Execute strsqlRMK
                CASRMK.Close
                Set CASRMK = Nothing
                
                
                strsqlRMK = ""
                strsqlRMK = "EXEC Case_AdjUpdateCASRemarkInt '" & Trim(TempCASNo) & "', '" & CASRemarksInt & "'"
                medixmasterconn.Execute strsqlRMK
                'StrSqlCRef = "Select * from CASCrossReference Where CASNumber ='" & Trim(tempcasNo) & "' And MBMNumber ='" & Trim(MBMNumber) & " '"
                'CASCRef.Open StrSqlCRef, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                strAppendCase = "3"
                strAppend = " AND CASNumber ='" & Trim(TempCASNo) & "' And MBMNumber ='" & Trim(MBMNumber) & " '"
                Set cmd2.ActiveConnection = medixmasterconnMaint
                cmd2.CommandText = "Case_Adjustment"
                cmd2.CommandType = adCmdStoredProc
                cmd2.CommandTimeout = 300
                Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd2.Parameters.Append Prm1
                Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd2.Parameters.Append Prm2
                CASCRef.CursorLocation = adUseClient
                CASCRef.CursorType = adOpenForwardOnly
                CASCRef.CacheSize = 100
                Set CASCRef = cmd2.Execute
                
                    'With CASCRef
                        '!CASStatus = Mid(cboStatus1, 1, 1)
                        '!CASClaimability = Mid(CboClaimability, 1, 1)
                        If txtDateAdmitted <> "__/__/____" Then
                            DateAdmitted = Format(txtDateAdmitted, "mm/dd/yyyy")
                        Else
                            DateAdmitted = Empty
                        End If
                        If txtDateDischarge <> "__/__/____" Then
                            DateDischarge = Format(txtDateDischarge, "mm/dd/yyyy")
                        Else
                            DateDischarge = Empty
                        End If
                        '!CASPending = Mid(cboPending, 1, 1)
                        If TxtLinkCASNo <> "" Then
                            LINKCASNo = ChkStr(TxtLinkCASNo)
                        Else
                            LINKCASNo = Null
                        End If
                        '!CASTopupLmt = Mid(cboTopUpLmt, 1, 1)
                        '!MBMNumber = CASMBMNumber
                        '!MBMPatientCovID = CASMBMID
                    'End With
                    'CASCRef.Update
                    
                    StrSqlCRef = "'" & Trim(TempCASNo) & "', '" & Mid(cboStatus1, 1, 1) & "', " & _
                                "'" & Mid(cboClaimability, 1, 2) & "', '" & DateAdmitted & "', '" & DateDischarge & "', " & _
                                "'" & Mid(cboPending, 1, 1) & "', '" & LINKCASNo & "', '" & Mid(cboTopUpLmt, 1, 1) & "', " & _
                                "'" & CASMBMNumber & "', '" & CASMBMID & "', '" & Mid(cboPayPending, 1, 1) & "'"
                    
                    If CASCRef.EOF = True Then
                        StrSqlCRef = "EXEC Case_AdjInsertCASRef " + StrSqlCRef
                    Else
                        StrSqlCRef = "EXEC Case_AdjUpdateCASRef " + StrSqlCRef
                    End If
                    
                medixmasterconnMaint.Execute StrSqlCRef
                                                        
                Call SaveFirstDiag
                Call SaveFinalDiag
                Call SaveUD
                Call SaveCASOthers
                Call SaveCASOthers2
                Call saveHos
                Call UpdateBalLmt
                Call UpdateLG
                Call UpdateAuditReview
                If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
                    Call SaveRepudiation
                End If
                
                CAS.Close
                Set CAS = Nothing
                CASCRef.Close
                Set CASCRef = Nothing
                cmdUpdate1.Enabled = False
                MsgBox "Record Updated Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            End If
            Frame2.Enabled = True
            'Frame7.Enabled = True
            'Frame15.Enabled = True
            'Frame16.Enabled = True
            'Frame17.Enabled = True
            cmdUpdate1.Enabled = True
        End If
            medixmasterconn.Close
            Set medixmasterconn = Nothing
            medixmasterconnMaint.Close
            Set medixmasterconnMaint = Nothing
        cmdShow.Enabled = True
        cmdLG.Enabled = True
        cmdRepudiate.Enabled = True
        cmdSaveAdd.Enabled = True
        cmdPrior.Enabled = True
        cmdPrev.Enabled = True
        CmdNext.Enabled = True
        cmdUpdate1.Enabled = True
End Sub

Private Sub SaveFirstDiag()
On Error Resume Next
Dim diafirst As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strsqlDIAFirst As String
Dim DIADOS1, DIADOS2, DIADOS3
Dim PrelimDiag1, PrelimDiag2, PrelimDiag3
   
    'strsql = "Select * from DIAFirst where CasNumber = '" & Trim(txtCaseNo) & "'"
    'diafirst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "9"
    strAppend = " AND CasNumber = '" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    diafirst.CursorLocation = adUseClient
    diafirst.CursorType = adOpenForwardOnly
    diafirst.CacheSize = 100
    Set diafirst = cmd.Execute
        
    'If diafirst.EOF = True Then
        'diafirst.AddNew
        'diafirst!CasNumber = ChkStr(txtCaseNo)
    'End if
    
        'With diafirst
        '    If Trim(Mid(cboICDCode1, 1, 3)) <> "" Then
        '        !ICDCode1 = Mid(cboICDCode1, 1, 3)
        '    End If
                If txtPDDOS(0) <> "__/__/____" Then
                    DIADOS1 = Format(Trim(txtPDDOS(0)), "mm/dd/yyyy")
                    txtPDDOS(0).mask = "##/##/####"
                Else
                    DIADOS1 = Empty
                End If
        '        !DIAPrelimDiag1 = Trim(ChkStr(txtPrelimDiag1))
        '    If Trim(Mid(cboICDCode2, 1, 3)) <> "" Then
        '        !ICDCode2 = Mid(cboICDCode2, 1, 3)
        '    End If
        '======================== Bug ========================
                If txtPDDOS(1) <> "__/__/____" Then
                    DIADOS2 = Format(Trim(txtPDDOS(1)), "mm/dd/yyyy")
                Else
                    DIADOS2 = Empty
                End If
        '======================== Bug =========================
        '        !DIAPrelimDiag2 = Trim(ChkStr(txtPrelimDiag2))
        '    If Trim(Mid(cboICDCode3, 1, 3)) <> "" Then
        '        !ICDCode3 = Mid(cboICDCode3, 1, 3)
        '    End If
                If txtPDDOS(2) <> "__/__/____" Then
                    DIADOS3 = Format(Trim(txtPDDOS(2)), "mm/dd/yyyy")
                Else
                    DIADOS3 = Empty
                End If
        '        !DIAPrelimDiag3 = Trim(ChkStr(txtPrelimDiag3))
        'diafirst.Update
        'End With
        PrelimDiag1 = ChkStr(Replace(txtPrelimDiag(0), "'", "''"))
        PrelimDiag2 = ChkStr(Replace(txtPrelimDiag(1), "'", "''"))
        PrelimDiag3 = ChkStr(Replace(txtPrelimDiag(2), "'", "''"))
        
        strsqlDIAFirst = "'" & Trim(txtCaseNo) & "', '" & Mid(cboICDCode(0), 1, 3) & "', '" & Mid(cboICDCode(1), 1, 3) & "', " & _
                        "'" & Mid(cboICDCode(2), 1, 3) & "', '" & Trim(PrelimDiag1) & "', '" & Trim(PrelimDiag2) & "', " & _
                        "'" & Trim(PrelimDiag3) & "', '" & DIADOS1 & "', '" & DIADOS2 & "', '" & DIADOS3 & "'"
        
    If diafirst.EOF = True Then
        strsqlDIAFirst = "EXEC Case_AdjInsertDIAFirst " + strsqlDIAFirst
    Else
        strsqlDIAFirst = "EXEC Case_AdjUpdateDIAFirst " + strsqlDIAFirst
    End If
    medixmasterconn.Execute strsqlDIAFirst
    diafirst.Close
    Set diafirst = Nothing
    
End Sub

Private Sub SaveRepudiation()
Dim REP As New ADODB.Recordset
Dim CAS As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cnn As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strsqlREP As String
Dim strsqlCAS As String
Dim LogonDate
    
    ' INSERT INTO CASRepudiation TABLE
    strAppendCase = "10"
    strAppend = " AND CASNumber = '" & Trim(TempCASNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    REP.CursorLocation = adUseClient
    REP.CursorType = adOpenForwardOnly
    REP.CacheSize = 100
    Set REP = cmd.Execute
        
    'strsqlREP = "select * from CASRepudiation where CASNumber = '" & Trim(tempcasNo) & "'"
    'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'strsqlCAS = "Select *  From CAS where CASNumber = '" & Trim(txtCaseNo) & "'"
    'CAS.Open strsqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    strAppendCase = "7"
    strAppend = " AND CASNumber = '" & Trim(txtCaseNo) & "'"
    Set cnn.ActiveConnection = medixmasterconn
    cnn.CommandText = "Case_Adjustment"
    cnn.CommandType = adCmdStoredProc
    cnn.CommandTimeout = 300
    Set Prm1 = cnn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cnn.Parameters.Append Prm1
    Set Prm2 = cnn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cnn.Parameters.Append Prm2
    CAS.CursorLocation = adUseClient
    CAS.CursorType = adOpenForwardOnly
    CAS.CacheSize = 100
    Set CAS = cnn.Execute
    
    LogonDate = Format(Date, "mm/dd/yyyy")
    
        'If REP.EOF = True Then
            'REP.AddNew
            'REP!CasNumber = ChkStr(txtCaseNo)
        'End If
           
            strsqlREP = "'" & Trim(TempCASNo) & "', '" & IIf(Len(cboREPCode1), Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - IIf(InStr(1, cboREPCode1, "-") = 0, 0, 1))), "") & "', " & _
                        "'" & IIf(Len(cboREPCode2), Trim(Mid(cboREPCode2, 1, InStr(1, cboREPCode2, "-") - IIf(InStr(1, cboREPCode2, "-") = 0, 0, 1))), "") & "', '" & IIf(Len(cboREPCode3), Trim(Mid(cboREPCode3, 1, InStr(1, cboREPCode3, "-") - IIf(InStr(1, cboREPCode3, "-") = 0, 0, 1))), "") & "', " & _
                        "'" & Logon & "', '" & LogonDate & "', '" & IIf(Len(Document), IIf(InStr(1, Document, "-") = 0, Trim(Document), Trim(Mid(Document, 1, InStr(1, Document, "-") - IIf(InStr(1, Document, "-") = 0, 0, 1)))), "") & "', " & _
                        "'" & IIf(Len(Personnel), IIf(InStr(1, Personnel, "-") = 0, Trim(Personnel), Trim(Mid(Personnel, 1, InStr(1, Personnel, "-") - IIf(InStr(1, Personnel, "-") = 0, 0, 1)))), "") & "', " & _
                        "'" & IIf(Len(Attention), IIf(InStr(1, Attention, "-") = 0, Trim(Attention), Trim(Mid(Attention, 1, InStr(1, Attention, "-") - IIf(InStr(1, Attention, "-") = 0, 0, 1)))), "") & "', " & _
                        "'" & CASUnderlying & "'"
            
            strsqlCAS = "EXEC Case_AdjUpdateCAS1 '" & Trim(txtCaseNo) & "', '" & Mid(cboClaimability, 1, 2) & "'"
            medixmasterconn.Execute strsqlCAS
            'With REP
            '    !CASREPCode1 = Trim(Mid(cboREPCode1, 1, 6))
            '    !CASREPCode2 = Trim(Mid(cboREPCode2, 1, 6))
            '    !CASREPCode3 = Trim(Mid(cboREPCode3, 1, 6))
            '    !CASREPIssuedBy = Logon
            '    !CASREPDate = Date
            'End With
            'REP.Update
            '    With CAS
            '        !CASClaimability = Mid(CboClaimability, 1, 1)
            '    End With
            '    CAS.Update
            
        If REP.EOF = True Then
            strsqlREP = "EXEC Case_AdjInsertCASRep " + strsqlREP
        Else
            strsqlREP = "EXEC Case_AdjUpdateCASRep " + strsqlREP
        End If
        medixmasterconn.Execute strsqlREP
    REP.Close
    CAS.Close
    Set REP = Nothing
    Set CAS = Nothing
End Sub

Private Sub SaveFinalDiag()
On Error Resume Next
Dim diafinal As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strsqlDIAFinal As String
Dim LogonDate
Dim DIADOS1, DIADOS2, DIADOS3 As String
Dim FinalDiag1, FinalDiag2, FinalDiag3
    
TotalFinalDiag = 0

' INSERT INTO DIAFinal TABLE
    'strsql = "select *  from DIAFinal where casNumber ='" & txtCaseNo & "'"
    'diafinal.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "11"
    strAppend = " AND casNumber ='" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    diafinal.CursorLocation = adUseClient
    diafinal.CursorType = adOpenForwardOnly
    diafinal.CacheSize = 100
    Set diafinal = cmd.Execute
    
    LogonDate = Format(Date, "mm/dd/yyyy")
    FinalDiag1 = ChkStr(Replace(txtFinalDiag1, "'", "''"))
    FinalDiag2 = ChkStr(Replace(txtFinalDiag2, "'", "''"))
    FinalDiag3 = ChkStr(Replace(txtFinalDiag3, "'", "''"))
    
    'If diafinal.EOF = True Then
        'diafinal.AddNew
        'diafinal!CasNumber = ChkStr(txtCaseNo)
    'End if
    
     'With diafinal
     '    !DIALastAdjustmentUser = Logon
     '    !DIALastAdjustmentDate = Date
         
         If cboFDICDCode1 <> "" Then
     '       !ICDCode1 = Mid(ChkStr(cboFDICDCode1), 1, 3)
            TotalFinalDiag = CDbl(TotalFinalDiag) + 1
         End If
         If txtFDDOS1 <> "__/__/____" Then
             DIADOS1 = Format(txtFDDOS1, "mm/dd/yyyy")
         Else
             DIADOS1 = Empty
         End If
     '    !DIAFinalDiag1 = ChkStr(txtFinalDiag1)
         If cboFDICDCode2 <> "" Then
     '       !ICDCode2 = Mid(ChkStr(cboFDICDCode2), 1, 3)
            TotalFinalDiag = CDbl(TotalFinalDiag) + 1
         End If
         If txtFDDOS2 <> "__/__/____" Then
             DIADOS2 = Format(txtFDDOS2, "mm/dd/yyyy")
         Else
             DIADOS2 = Empty
         End If
     '    !DIAFinalDiag2 = ChkStr(txtFinalDiag2)
         If cboFDICDCode3 <> "" Then
     '       !ICDCode3 = Mid(ChkStr(cboFDICDCode3), 1, 3)
            TotalFinalDiag = CDbl(TotalFinalDiag) + 1
         End If
         If txtFDDOS3 <> "__/__/____" Then
             DIADOS3 = Format(txtFDDOS3, "mm/dd/yyyy")
         Else
             DIADOS3 = Empty
         End If
      '   !DIAFinalDiag3 = ChkStr(txtFinalDiag3)
     'End With
     'diafinal.Update
    strsqlDIAFinal = "'" & ChkStr(txtCaseNo) & "', '" & Mid(ChkStr(cboFDICDCode1), 1, 3) & "', " & _
                        "'" & Mid(ChkStr(cboFDICDCode2), 1, 3) & "', '" & Mid(ChkStr(cboFDICDCode3), 1, 3) & "', '" & FinalDiag1 & "', " & _
                        "'" & FinalDiag2 & "', '" & FinalDiag3 & "', '" & DIADOS1 & "', '" & DIADOS2 & "', '" & DIADOS3 & "', " & _
                        "'" & Logon & "', '" & LogonDate & "'"
    
   
    If diafinal.EOF = True Then
        strsqlDIAFinal = "EXEC Case_AdjInsertDIAFinal " + strsqlDIAFinal
    Else
        strsqlDIAFinal = "EXEC Case_AdjUpdateDIAFinal " + strsqlDIAFinal
    End If
    medixmasterconn.Execute strsqlDIAFinal
    diafinal.Close
    Set diafinal = Nothing

End Sub

Private Sub SaveUD()
On Error Resume Next
Dim CASUD As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strsqlCASUD As String
Dim CASUDDOS1, CASUDDOS2, CASUDDOS3
Dim UD1, UD2, UD3
    
    'strsql = "select * from CAS_UD where CASNumber = '" & Trim(tempcasNo) & "'"
    'CASUD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "12"
    strAppend = " AND CASNumber = '" & Trim(TempCASNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    CASUD.CursorLocation = adUseClient
    CASUD.CursorType = adOpenForwardOnly
    CASUD.CacheSize = 100
    Set CASUD = cmd.Execute
        
    'If CASUD.EOF = True Then
        'CASUD.AddNew
        'CASUD!CasNumber = ChkStr(txtCaseNo)
    'End If
        'strsqlCASUD = "EXEC Case_AdjUpdateUD '" & Trim(tempcasNo) & "', '" & Mid(cboUDICDCode1, 1, 3) & "', '" & Mid(cboUDICDCode2, 1, 3) & "', " & _
                    "'" & Mid(cboUDICDCode3, 1, 3) & "', '" & ChkStr(txtUD1) & "', '" & ChkStr(txtUD2) & "', '" & ChkStr(txtUD3) & "', " & _
                    "'" & txtUDDOS1 & "', '" & txtUDDOS2 & "', '" & txtUDDOS3 & "'"
    'With CASUD
    '    If Trim(cboUDICDCode1) <> "" Then
    '        !CASUDCode1 = Mid(cboUDICDCode1, 1, 3)
    '    End If
        If txtUDDOS1 <> "__/__/____" Then
            CASUDDOS1 = Format(txtUDDOS1, "mm/dd/yyyy")
        Else
            CASUDDOS1 = Empty
        End If
    '        !CASUDReasons1 = ChkStr(txtUD1)
    '    If Trim(cboUDICDCode2) <> "" Then
    '        !CASUDCode2 = Mid(cboUDICDCode2, 1, 3)
    '    End If
        If txtUDDOS2 <> "__/__/____" Then
            CASUDDOS2 = Format(txtUDDOS2, "mm/dd/yyyy")
        Else
            CASUDDOS2 = Empty
        End If
    '    !CASUDReasons2 = txtUD2
    '    If Trim(cboUDICDCode3) <> "" Then
    '        !CASUDCode3 = Mid(cboUDICDCode3, 1, 3)
    '    End If
        If txtUDDOS3 <> "__/__/____" Then
            CASUDDOS3 = Format(txtUDDOS3, "mm/dd/yyyy")
        Else
            CASUDDOS3 = Empty
        End If
    '        !CASUDReasons3 = ChkStr(txtUD3)
    'CASUD.Update
    'End With
    UD1 = ChkStr(Replace(txtUD1, "'", "''"))
    UD2 = ChkStr(Replace(txtUD2, "'", "''"))
    UD3 = ChkStr(Replace(txtUD3, "'", "''"))
    
    strsqlCASUD = "'" & Trim(TempCASNo) & "', '" & Mid(cboUDICDCode1, 1, 3) & "', '" & Mid(cboUDICDCode2, 1, 3) & "', " & _
                    "'" & Mid(cboUDICDCode3, 1, 3) & "', '" & UD1 & "', '" & UD2 & "', '" & UD3 & "', " & _
                    "'" & CASUDDOS1 & "', '" & CASUDDOS2 & "', '" & CASUDDOS3 & "'"
    
    If CASUD.EOF = True Then
        strsqlCASUD = "EXEC Case_AdjInsertUD " + strsqlCASUD
    Else
        strsqlCASUD = "EXEC Case_AdjUpdateUD " + strsqlCASUD
    End If
    medixmasterconn.Execute strsqlCASUD
    CASUD.Close
    Set CASUD = Nothing
End Sub

Private Sub SaveCASOthers()
Dim sqlstrCASOthers As String
Dim Other As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim CASRedAmt, FirstCall, FirstTreated As String
Dim strsqlCASOthers As String
Dim MEDSURDone1, MEDSURDone2, MEDSURDone3
Dim PROCINVDone1, PROCINVDone2, PROCINVDone3
Dim PROCRequired1, PROCRequired2, PROCRequired3, PROCRequired4, PROCRequired5
Dim CASCptCode1, CASCptCode2, CASCptCode3, CASCptCode4, CASCptCode5 As String
Dim CASONonSurgPROCRequired1, CASONonSurgPROCRequired2, CASONonSurgPROCRequired3, CASONonSurgPROCRequired4, CASONonSurgPROCRequired5 As String
Dim CASNonSurgCptCode1, CASNonSurgCptCode2, CASNonSurgCptCode3, CASNonSurgCptCode4, CASNonSurgCptCode5 As String
    
    TotalProc = 0
    
    'sqlstr = "Select * From CASOthers where CASNumber = '" & RemStr(txtCaseNo) & "'"
    'Other.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "13"
    strAppend = " AND CASNumber = '" & RemStr(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    Other.CursorLocation = adUseClient
    Other.CursorType = adOpenForwardOnly
    Other.CacheSize = 100
    Set Other = cmd.Execute


        If txtFirstCall <> "__/__/____" Then
            FirstCall = Format(txtFirstCall, "mm/dd/yyyy")
        Else
            FirstCall = Empty
        End If
        
        If txtFirstTreated <> "__/__/____" Then
             FirstTreated = Format(txtFirstTreated, "mm/dd/yyyy")
        Else
             FirstTreated = Empty
        End If
        
        If txtProcedure1 <> "" Then
            TotalProc = CDbl(TotalProc) + 1
        End If

        If txtProcedure2 <> "" Then
            TotalProc = CDbl(TotalProc) + 1
        End If

        If txtProcedure3 <> "" Then
            TotalProc = CDbl(TotalProc) + 1
        End If
        
        If txtProcedure4 <> "" Then
            TotalProc = CDbl(TotalProc) + 1
        End If
            
        If txtProcedure5 <> "" Then
            TotalProc = CDbl(TotalProc) + 1
        End If
            
        If IsNumeric(txtREDAmt) = True Then
            CASRedAmt = txtREDAmt
        Else
            CASRedAmt = 0
        End If
        
        MEDSURDone1 = ChkStr(Replace(txtMEDSURDone1, "'", "''"))
        MEDSURDone2 = ChkStr(Replace(txtMEDSURDone2, "'", "''"))
        MEDSURDone3 = ChkStr(Replace(txtMEDSURDone3, "'", "''"))
        PROCINVDone1 = ChkStr(Replace(txtPROCINVDone1, "'", "''"))
        PROCINVDone2 = ChkStr(Replace(txtPROCINVDone2, "'", "''"))
        PROCINVDone3 = ChkStr(Replace(txtPROCINVDone3, "'", "''"))
        
        PROCRequired1 = ChkStr(Replace(txtProcedure1, "'", "''"))
        PROCRequired2 = ChkStr(Replace(txtProcedure2, "'", "''"))
        PROCRequired3 = ChkStr(Replace(txtProcedure3, "'", "''"))
        PROCRequired4 = ChkStr(Replace(txtProcedure4, "'", "''"))
        PROCRequired5 = ChkStr(Replace(txtProcedure5, "'", "''"))
        CASCptCode1 = IIf(Len(txtProcCode1), txtProcCode1, "")
        CASCptCode2 = IIf(Len(txtProcCode2), txtProcCode2, "")
        CASCptCode3 = IIf(Len(txtProcCode3), txtProcCode3, "")
        CASCptCode4 = IIf(Len(txtProcCode4), txtProcCode4, "")
        CASCptCode5 = IIf(Len(txtProcCode5), txtProcCode5, "")
        
        CASONonSurgPROCRequired1 = ChkStr(Replace(txtNSProcedure1, "'", "''"))
        CASONonSurgPROCRequired2 = ChkStr(Replace(txtNSProcedure2, "'", "''"))
        CASONonSurgPROCRequired3 = ChkStr(Replace(txtNSProcedure3, "'", "''"))
        CASONonSurgPROCRequired4 = ChkStr(Replace(txtNSProcedure4, "'", "''"))
        CASONonSurgPROCRequired5 = ChkStr(Replace(txtNSProcedure5, "'", "''"))
        
        CASNonSurgCptCode1 = IIf(Len(txtNSProcCode1), txtNSProcCode1, "")
        CASNonSurgCptCode2 = IIf(Len(txtNSProcCode2), txtNSProcCode2, "")
        CASNonSurgCptCode3 = IIf(Len(txtNSProcCode3), txtNSProcCode3, "")
        CASNonSurgCptCode4 = IIf(Len(txtNSProcCode4), txtNSProcCode4, "")
        CASNonSurgCptCode5 = IIf(Len(txtNSProcCode5), txtNSProcCode5, "")
        
        
        strsqlCASOthers = "'" & ChkStr(txtCaseNo) & "', " & _
                        "'" & PROCINVDone1 & "', '" & PROCINVDone2 & "', '" & PROCINVDone3 & "', " & _
                        "'" & MEDSURDone1 & "', '" & MEDSURDone2 & "', '" & MEDSURDone3 & "', " & _
                        "'" & Mid(cboCondwillOcc, 1, 1) & "', '" & Mid(cboFLWTreatment, 1, 1) & "', '" & FirstCall & "', " & _
                        "'" & FirstTreated & "', " & CASRedAmt & ", " & _
                        "'" & ChkStr(txtINVCode1) & "','" & ChkStr(txtINVCode2) & "', '" & ChkStr(txtINVCode3) & "', '" & Mid(cboFolUp, 1, 1) & "', '" & Mid(cboPayPending, 1, 1) & "', " & _
                        "'" & PROCRequired1 & "', '" & PROCRequired2 & "', '" & PROCRequired3 & "', '" & PROCRequired4 & "', '" & PROCRequired5 & "', " & _
                        "'" & CASCptCode1 & "', '" & CASCptCode2 & "', '" & CASCptCode3 & "', '" & CASCptCode4 & "', '" & CASCptCode5 & "', " & _
                        "'" & CASONonSurgPROCRequired1 & "', '" & CASONonSurgPROCRequired2 & "', '" & CASONonSurgPROCRequired3 & "', '" & CASONonSurgPROCRequired4 & "', '" & CASONonSurgPROCRequired5 & "', " & _
                        "'" & CASNonSurgCptCode1 & "', '" & CASNonSurgCptCode2 & "', '" & CASNonSurgCptCode3 & "', '" & CASNonSurgCptCode4 & "', '" & CASNonSurgCptCode5 & "'"

    
    If Other.EOF = True Then
        strsqlCASOthers = "EXEC Case_AdjInsertCASOthers " + strsqlCASOthers
    Else
        strsqlCASOthers = "EXEC Case_AdjUpdateCASOthers " + strsqlCASOthers
    End If
    
    medixmasterconn.Execute strsqlCASOthers
    
    Other.Close
    Set Other = Nothing
End Sub

Private Sub SaveCASOthers2()
Dim CASOthersTwo As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim TotalDoctor As String
Dim TotalAnaes As String
Dim TempDRName As String
Dim TempDRName2 As String
Dim TempDRName3 As String
Dim TempDRName4 As String
Dim TempDRName5 As String
Dim TempAnaesName As String
Dim TempAnaesName2 As String
Dim TempAnaesName3 As String
Dim strsqlCASOthers2 As String
Dim Reason, RefBy As String
Dim TmpFirstDocDate As String
Dim TmpSubSeqDocDate As String
Dim TmpIntExtTO As String

TotalDoctor = 0
TotalAnaes = 0

TempDRName = Replace(txtATTDoc(0), "'", "''")
TempDRName2 = Replace(txtATTDoc(1), "'", "''")
TempDRName3 = Replace(txtATTDoc(2), "'", "''")
TempDRName4 = Replace(txtATTDoc(3), "'", "''")
TempDRName5 = Replace(txtATTDoc(4), "'", "''")
TempAnaesName = Replace(txtAnaesDoc(0), "'", "''")
TempAnaesName2 = Replace(txtAnaesDoc(1), "'", "''")
TempAnaesName3 = Replace(txtAnaesDoc(2), "'", "''")
Reason = ChkStr(Replace(txtReason, "'", "''"))
RefBy = ChkStr(Replace(txtREFBy, "'", "''"))

    ' INSERT INTO CASOthersTwo TABLE
    'CASOthersTwo.Open "Select CASNumber, * from CASOthersTwo where CASNumber = '" & Trim(txtCaseNo) & "'", medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "14"
    strAppend = " AND CASNumber = '" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    CASOthersTwo.CursorLocation = adUseClient
    CASOthersTwo.CursorType = adOpenForwardOnly
    CASOthersTwo.CacheSize = 100
    Set CASOthersTwo = cmd.Execute
        
    'If CASOthersTwo.EOF = True Then
        'CASOthersTwo.AddNew
        'CASOthersTwo!CasNumber = ChkStr(txtCaseNo)
    'End If
        
        'With CASOthersTwo
            If TempDRName <> "" Then
        '        !CASAttendDoctor1 = ChkStr(TempDRName)
                TotalDoctor = CDbl(TotalDoctor) + 1
        '    Else
        '        !CASAttendDoctor1 = Null
            End If
            If TempDRName2 <> "" Then
        '        !CASAttendDoctor2 = ChkStr(TempDRName2)
                TotalDoctor = CDbl(TotalDoctor) + 1
        '    Else
        '        !CASAttendDoctor2 = Null
            End If
            If TempDRName3 <> "" Then
        '        !CASAttendDoctor3 = ChkStr(TempDRName3)
                TotalDoctor = CDbl(TotalDoctor) + 1
        '    Else
        '        !CASAttendDoctor3 = Null
            End If
            If TempDRName4 <> "" Then
        '        !CASAttendDoctor4 = ChkStr(TempDRName4)
                TotalDoctor = CDbl(TotalDoctor) + 1
        '    Else
        '        !CASAttendDoctor4 = Null
            End If
            If TempDRName5 <> "" Then
        '        !CASAttendDoctor5 = ChkStr(TempDRName5)
                TotalDoctor = CDbl(TotalDoctor) + 1
        '    Else
        '        !CASAttendDoctor5 = Null
            End If
            If TempAnaesName <> "" Then
        '        !CASAnaesName1 = ChkStr(TempAnaesName)
                TotalAnaes = CDbl(TotalAnaes) + 1
        '    Else
        '        !CASAnaesName1 = Null
            End If
            If TempAnaesName2 <> "" Then
        '        !CASAnaesName2 = ChkStr(TempAnaesName2)
                TotalAnaes = CDbl(TotalAnaes) + 1
        '    Else
        '        !CASAnaesName2 = Null
            End If
            If TempAnaesName3 <> "" Then
        '        !CASAnaesName3 = ChkStr(TempAnaesName3)
                TotalAnaes = CDbl(TotalAnaes) + 1
        '    Else
        '        !CASAnaesName3 = Null
            End If
                
            If txtFirstDocDate <> "__/__/____" Then
                TmpFirstDocDate = Format(txtFirstDocDate, "mm/dd/yyyy")
            Else
                TmpFirstDocDate = Empty
            End If
        
            If txtSubDocDate <> "__/__/____" Then
                TmpSubSeqDocDate = Format(txtSubDocDate, "mm/dd/yyyy")
            Else
                TmpSubSeqDocDate = Empty
            End If
            TmpIntExtTO = IIf(Len(cboIntExtTO), Mid(cboIntExtTO, 1, 1), "")
                
        '    !CASTotalDoctors = TotalDoctor
        '    !CASTotalAnaes = TotalAnaes
        '    !CASTotalFinalDiag = TotalFinalDiag
        '    !CASTotalProc = TotalProc
            
        'End With
        'CASOthersTwo.Update
        strsqlCASOthers2 = "'" & ChkStr(txtCaseNo) & "', '" & ChkStr(TempDRName) & "', " & _
                        "'" & ChkStr(TempDRName2) & "', '" & ChkStr(TempDRName3) & "', '" & ChkStr(TempDRName4) & "', " & _
                        "'" & ChkStr(TempDRName5) & "', '" & ChkStr(TempAnaesName) & "', '" & ChkStr(TempAnaesName2) & "', " & _
                        "'" & ChkStr(TempAnaesName3) & "', '" & TotalDoctor & "', '" & TotalAnaes & "', '" & TotalFinalDiag & "', " & _
                        "'" & TotalProc & "', '" & Reason & "', '" & RefBy & "', " & _
                        "'" & TmpFirstDocDate & "', '" & TmpSubSeqDocDate & "', '" & TmpIntExtTO & "'"
                        
    If CASOthersTwo.EOF = True Then
        strsqlCASOthers2 = "EXEC Case_AdjInsertCASOthers2 " + strsqlCASOthers2
    Else
        strsqlCASOthers2 = "EXEC Case_AdjUpdateCASOthers2 " + strsqlCASOthers2
    End If
    medixmasterconn.Execute strsqlCASOthers2
    CASOthersTwo.Close
    Set CASOthersTwo = Nothing
End Sub

Private Sub ComboListing()
Dim HOS As New ADODB.Recordset
Dim SRV As New ADODB.Recordset
Dim Nature As New ADODB.Recordset
Dim Modality As New ADODB.Recordset
Dim PAYMode As New ADODB.Recordset
Dim PAYTYPE As New ADODB.Recordset
Dim STAYTYPE As New ADODB.Recordset
Dim CASRep As New ADODB.Recordset
Dim SPEAPP As New ADODB.Recordset
Dim cmd As New ADODB.Command

    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchSpecApp"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    SPEAPP.CursorLocation = adUseClient
    SPEAPP.CursorType = adOpenForwardOnly
    SPEAPP.CacheSize = 100
    Set SPEAPP = cmd.Execute
 
    Do Until SPEAPP.EOF
        With SPEAPP
            cboCASInd.AddItem !SpecAppCode & " - " & !SpecAppDesc
            SPEAPP.MoveNext
        End With
    Loop
    SPEAPP.Close
    Set SPEAPP = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchNatureCode"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Nature.CursorLocation = adUseClient
    Nature.CursorType = adOpenForwardOnly
    Nature.CacheSize = 100
    Set Nature = cmd.Execute
     
    Do Until Nature.EOF
        With Nature
            cboNature.AddItem !ClaimNatureCode & " - " & !ClaimNatureName
            Nature.MoveNext
        End With
    Loop
    Nature.Close
    Set Nature = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPAYType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAYTYPE.CursorLocation = adUseClient
    PAYTYPE.CursorType = adOpenForwardOnly
    PAYTYPE.CacheSize = 100
    Set PAYTYPE = cmd.Execute
    
    Do Until PAYTYPE.EOF
        With PAYTYPE
            CboPayType.AddItem !PayTypeCode & " - " & Trim(!PayTypeName)
            PAYTYPE.MoveNext
        End With
    Loop
    PAYTYPE.Close
    Set PAYTYPE = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchStayType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    STAYTYPE.CursorLocation = adUseClient
    STAYTYPE.CursorType = adOpenForwardOnly
    STAYTYPE.CacheSize = 100
    Set STAYTYPE = cmd.Execute
    
    Do Until STAYTYPE.EOF
        With STAYTYPE
            cboStayType.AddItem !PayTypeCode & " - " & Trim(!PayTypeName)
            STAYTYPE.MoveNext
        End With
    Loop

    STAYTYPE.Close
    Set STAYTYPE = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchTreatMode"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Modality.CursorLocation = adUseClient
    Modality.CursorType = adOpenForwardOnly
    Modality.CacheSize = 100
    Set Modality = cmd.Execute

    Do Until Modality.EOF
        With Modality
            cboModality.AddItem !ModalityCode & " - " & Trim(!ModalityName)
            Modality.MoveNext
        End With
    Loop
    Modality.Close
    Set Modality = Nothing

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHOS"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cmd.Execute

    Do Until HOS.EOF
        With HOS
            txtHOSCode.AddItem Trim(!HOSName)
            HOS.MoveNext
        End With
    Loop
    HOS.Close
    Set HOS = Nothing
    
    cboStatus1.AddItem "O" & " - " & "Open"
    cboStatus1.AddItem "C" & " - " & "Close"
    cboStatus1.AddItem "D" & " - " & "Delete"

    cboClaimability.AddItem "AA" & " - " & "Accepted"
    cboClaimability.AddItem "RR" & " - " & "Repudiated"
    cboClaimability.AddItem "CC" & " - " & "Cancelled"
    cboClaimability.AddItem "SA" & " - " & "Special Approval"
    cboClaimability.AddItem "PP" & " - " & "Pay & Claim"
    cboClaimability.AddItem "NN" & " - " & "Not Decided"
    cboClaimability.AddItem "NA" & " - " & "Not Decided/Accepted"
    cboClaimability.AddItem "NR" & " - " & "Not Decided/Repudiated"
    cboClaimability.AddItem "PA" & " - " & "Pay&Claim/Accepted"
    cboClaimability.AddItem "PR" & " - " & "Pay&Claim/Rejected"
    cboClaimability.AddItem "AR" & " - " & "Accepted/Rejected"
    cboClaimability.AddItem "RA" & " - " & "Rejected/Accepted"
    cboClaimability.AddItem "BR" & " - " & "BTBG/Rejected"
    cboClaimability.AddItem "ER" & " - " & "Expired/Rejected"
    cboPending.AddItem "Y - Yes"
    cboPending.AddItem "N - No"
    
    cboCondwillOcc.AddItem "Y - YES"
    cboCondwillOcc.AddItem "N - NO"
    
    cboFLWTreatment.AddItem "Y - YES"
    cboFLWTreatment.AddItem "N - NO"

    cboTopUpLmt.AddItem "N - No"
    cboTopUpLmt.AddItem "Y - Yes"

    cboQuery.AddItem "Y - YES"
    cboQuery.AddItem "N - No"
    
    cboFolUp.AddItem "Y - YES"
    cboFolUp.AddItem "N - NO"
    
    'With cboDiagRepCode
    '    .AddItem "Prelim. Diag"
    '    .AddItem "Final Diag"
    '    .AddItem "Underlying Disease"
    '    If Mid(cboDiagRepCode, 1, 1) = "R" Then
    '        .AddItem "Repudiation Reason"
    '    End If
    'End With
    
    'With cboInvestProcCode
    '    .AddItem "Investigation"
    '    .AddItem "Procedures"
    'End With
    
    With cboPayPending
        .AddItem "Y - Yes"
        .AddItem "N - No"
    End With

    cboCASTakeOveInd.AddItem "Y - Yes"
    cboCASTakeOveInd.AddItem "N - No"

    CboARStatus.AddItem "C - Close"
    CboARStatus.AddItem "O - Open"
    
    cboIntExtTO.AddItem "I - Internal"
    cboIntExtTO.AddItem "E - External"
    
    cboAppStatus.AddItem "N - No"
    cboAppStatus.AddItem "Y - Yes"
medixmasterconn.Close
Set medixmasterconn = Nothing

End Sub

Private Sub CheckClaim(ErrorStatus As String)
On Error Resume Next
    Dim CLM As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String

    'strclm = "SELECT CASNumber, MBMNumber, CLMStatus From CLM WHERE (CASNumber = '" & txtCaseNo & "') AND (MBMNumber = '" & txtMBMNumber & "') AND (CLMStatus = 'O')"
    'CLM.Open strclm, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "16"
    strAppend = " AND (CASNumber = '" & txtCaseNo & "') AND (MBMNumber = '" & txtMBMNumber & "') AND (CLMStatus = 'O')"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        CLM.CursorLocation = adUseClient
        CLM.CursorType = adOpenForwardOnly
        CLM.CacheSize = 100
        Set CLM = cmd.Execute
        
    If CLM.EOF = False Then
        If ErrorStatus = "D" Then
            MsgBox "Claim Worksheet Already Registered Please Use Close Case", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf ErrorStatus = "C" Then
            MsgBox "Close All Claim Before Close The Case", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    End If
    CLM.Close
End Sub

Private Sub Form_Load()
Dim intFullHeight As Integer
Dim intDisplayHeight As Integer

    SndLG = False
    Call ComboListing
    FirstDiag = 0
    FinalDiag = 0
    'cmdRepudiate.Enabled = False

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
    
End Sub

Private Sub ShowFinalDiag()
    Dim diafinal As New ADODB.Recordset
    Dim strsql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String

    'strsql = "select * from DIAFinal where CASNumber = '" & txtCaseNo & "'"
    'diafinal.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "11"
    strAppend = " AND CASNumber = '" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        diafinal.CursorLocation = adUseClient
        diafinal.CursorType = adOpenForwardOnly
        diafinal.CacheSize = 100
        Set diafinal = cmd.Execute
        
        If diafinal.EOF = False Then
                With diafinal
                    If Len(!ICDCode1) Then
                        cboFDICDCode1 = !ICDCode1
                    End If
                    If Len(!DIADOS1) Then
                        txtFDDOS1.mask = ""
                        txtFDDOS1 = Format(!DIADOS1, "dd/mm/yyyy")
                        txtFDDOS1.mask = "##/##/####"
                    End If
                    If Len(!DIAFinalDiag1) Then
                        txtFinalDiag1 = !DIAFinalDiag1
                    End If
                    If Len(!ICDCode2) Then
                        cboFDICDCode2 = !ICDCode2
                    End If
                    If Len(!DIADOS2) Then
                        txtFDDOS2.mask = ""
                        txtFDDOS2 = Format(!DIADOS2, "dd/mm/yyyy")
                        txtFDDOS2.mask = "##/##/####"
                    End If
                    If Len(!DIAFinalDiag2) Then
                        txtFinalDiag2 = !DIAFinalDiag2
                    End If
                    If Len(!ICDCode3) Then
                        cboFDICDCode3 = !ICDCode3
                    End If
                    If Len(!DIADOS3) Then
                        txtFDDOS3.mask = ""
                        txtFDDOS3 = Format(!DIADOS3, "dd/mm/yyyy")
                        txtFDDOS3.mask = "##/##/####"
                    End If
                    If Len(!DIAFinalDiag3) Then
                        txtFinalDiag3 = !DIAFinalDiag3
                    End If
        
                End With
            End If
    diafinal.Close
    Set diafinal = Nothing
End Sub

Private Sub ShowFirstDiag()
    Dim diafirst As New ADODB.Recordset
    Dim strsql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
    'strsql = "Select * from DIAFirst where CasNumber = '" & Trim(txtCaseNo) & "'"
    'diafirst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "9"
    strAppend = " AND CasNumber = '" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        diafirst.CursorLocation = adUseClient
        diafirst.CursorType = adOpenForwardOnly
        diafirst.CacheSize = 100
        Set diafirst = cmd.Execute
        
        If diafirst.EOF = False Then
            Do Until diafirst.EOF
            With diafirst
                If Len(!ICDCode1) Then
                    cboICDCode(0) = !ICDCode1
                End If
                If Len(!DIADOS1) Then
                    txtPDDOS(0).mask = ""
                    txtPDDOS(0) = Format(!DIADOS1, "dd/mm/yyyy")
                    txtPDDOS(0).mask = "##/##/####"
                End If
                If Len(!DIAPrelimDiag1) Then
                    txtPrelimDiag(0) = !DIAPrelimDiag1
                End If
                If Len(!ICDCode2) Then
                    cboICDCode(1) = !ICDCode2
                End If
                If Len(!DIADOS2) Then
                    txtPDDOS(1).mask = ""
                    txtPDDOS(1) = Format(!DIADOS2, "dd/mm/yyyy")
                    txtPDDOS(1).mask = "##/##/####"
                End If
                If Len(!DIAPrelimDiag2) Then
                    txtPrelimDiag(1) = !DIAPrelimDiag2
                End If
                If Len(!ICDCode3) Then
                    cboICDCode(2) = !ICDCode3
                End If
                If Len(!DIADOS3) Then
                    txtPDDOS(2).mask = ""
                    txtPDDOS(2) = Format(!DIADOS3, "dd/mm/yyyy")
                    txtPDDOS(2).mask = "##/##/####"
                End If
                If Len(!DIAPrelimDiag3) Then
                    txtPrelimDiag(2) = !DIAPrelimDiag3
                End If
            End With
            diafirst.MoveNext
            Loop
            End If
    diafirst.Close
    Set diafirst = Nothing
End Sub

Private Sub ShowUD()
    Dim CASUD As New ADODB.Recordset
    Dim strsql As String
    
        If Not CASSearch.EOF Then
            With CASSearch
                If Len(!CASUDCode1) Then
                    cboUDICDCode1 = !CASUDCode1
                End If
                If Len(!CASUDDOS1) Then
                    txtUDDOS1.mask = ""
                    txtUDDOS1 = Format(!CASUDDOS1, "dd/mm/yyyy")
                    txtUDDOS1.mask = "##/##/####"
                End If
                If Len(!CASUDReasons1) Then
                    txtUD1 = !CASUDReasons1
                End If
                If Len(!CASUDCode2) Then
                    cboUDICDCode2 = !CASUDCode2
                End If
                If Len(!CASUDDOS2) Then
                    txtUDDOS2.mask = ""
                    txtUDDOS2 = Format(!CASUDDOS2, "dd/mm/yyyy")
                    txtUDDOS2.mask = "##/##/####"
                End If
                If Len(!CASUDReasons2) Then
                    txtUD2 = !CASUDReasons2
                End If
                If Len(!CASUDCode3) Then
                    cboUDICDCode3 = !CASUDCode3
                End If
                If Len(!CASUDDOS3) Then
                    txtUDDOS3.mask = ""
                    txtUDDOS3 = Format(!CASUDDOS3, "dd/mm/yyyy")
                    txtUDDOS3.mask = "##/##/####"
                End If
                If Len(!CASUDReasons3) Then
                    txtUD3 = !CASUDReasons3
                End If
               
            End With
        End If
End Sub










Private Sub lstFollowUp_Click()
    txtFollowUpDate = lstFollowUp.SelectedItem.SubItems(4)
End Sub

Private Sub lstFollowUp_DblClick()
    If lstFollowUp.listitems.Count <> 0 Then
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        FOLNumber = lstFollowUp.SelectedItem.Text
        Call funcword1
        
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
    End If
End Sub

Private Sub lstLGListing_Click()

    If lstLGListing.listitems.Count Then
        
        LGNo = ""
        LGVer = 0
        
        LGNo = Trim(Mid(lstLGListing.SelectedItem.SubItems(9), 1, IIf(InStr(1, lstLGListing.SelectedItem.SubItems(9), "-") = 0, 10, InStr(1, lstLGListing.SelectedItem.SubItems(9), "-") - 1)))
        LGVer = Trim(Mid(lstLGListing.SelectedItem.SubItems(9), InStr(1, lstLGListing.SelectedItem.SubItems(9), "-") + 1, Len(lstLGListing.SelectedItem.SubItems(9)) - InStr(1, lstLGListing.SelectedItem.SubItems(9), "-")))
             
        TxtLGRemark = Trim(lstLGListing.SelectedItem.SubItems(8))
        
    End If
End Sub

Private Sub lstLGListing_DblClick()
Dim cmd As New ADODB.Command
Dim LG As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strsqlLG As String
Dim ChkVer As String

    ChkVer = ""
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    strAppendCase = "7"
    strAppend = " AND CASENo ='" & Trim(LGNo) & "' AND LGNo = '" & Trim(LGVer) & "'"
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    LG.CursorLocation = adUseClient
    LG.CursorType = adOpenForwardOnly
    LG.CacheSize = 100
    Set LG = cmd.Execute
    
    If lstLGListing.listitems.Count Then
        If LG!LGStatus = "B" Then
            Letter = "B"
        ElseIf LG!LGStatus = "L" Then
            Letter = "L"
        'ElseIf LG!LGStatus = "V" Then
        '    Letter = "V"
        End If
        If LG!PLNPerVisit = "Y" Then
            tmpPLNPerVisit = LG!PLNPerVisit
        End If
        ReprintLG = True
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call funcword2
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        ReprintLG = False
    End If
    LG.Close
    Set LG = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
End Sub

Private Sub lstWorksheetListing_Click()
Dim lstWS As ListItem


    lblClaimNo = ""
    txtAppDate = "__/__/____"
    cboAppStatus = ""
    
    If lstWorksheetListing.listitems.Count Then
            Set lstWS = lstWorksheetListing.SelectedItem
            lblClaimNo = IIf(Len(Trim(lstWorksheetListing.SelectedItem)), lstWorksheetListing.SelectedItem, "")
            cboAppStatus = IIf(Len(Trim(lstWS.SubItems(8))), lstWS.SubItems(8), "")
            txtAppDate = IIf(Len(lstWS.SubItems(9)), Format(lstWS.SubItems(9), "dd/mm/yyyy"), "__/__/____")
            tmpCLMNumber = Trim(Mid(lstWorksheetListing.SelectedItem, 1, IIf(InStr(1, lstWorksheetListing.SelectedItem, "-") = 0, 10, InStr(1, lstWorksheetListing.SelectedItem, "-") - 1)))
            tmpCLMversion = Trim(Mid(lstWorksheetListing.SelectedItem, InStr(1, lstWorksheetListing.SelectedItem, "-") + 1, Len(lstWorksheetListing.SelectedItem) - InStr(1, lstWorksheetListing.SelectedItem, "-")))

    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 7 Then
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Call LGListing
        Call FollowUpListing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
    ElseIf SSTab1.Tab = 9 Then
        medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Call WSListing
        medixmasterconnWS.Close
        Set medixmasterconnWS = Nothing
    End If
End Sub

Private Sub SSTab1_DblClick()
    If SSTab1.Tab = 7 Then
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Call LGListing
        Call FollowUpListing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
    End If
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtAnaesActFees1_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    D = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalAnaesActFees = f

End Sub

Private Sub txtAnaesActFees2_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    D = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalAnaesActFees = f

End Sub

Private Sub txtAnaesActFees3_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    D = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalAnaesActFees = f

End Sub

Private Sub txtAnaesActFees4_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    D = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalAnaesActFees = f

End Sub

Private Sub txtAnaesActFees5_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    D = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalAnaesActFees = f

End Sub

Private Sub txtAnaesMMARate1_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    D = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalAnaesMMARate = f

End Sub

Private Sub txtAnaesMMARate2_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    D = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalAnaesMMARate = f

End Sub

Private Sub txtAnaesMMARate3_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    D = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalAnaesMMARate = f

End Sub

Private Sub txtAnaesMMARate4_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    D = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalAnaesMMARate = f

End Sub

Private Sub txtAnaesMMARate5_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    D = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalAnaesMMARate = f

End Sub

Private Sub TxtCaseNo_GotFocus()
    cmdShow.Default = True
    cmdUpdate1.Enabled = False
End Sub

Private Sub txtCaseNo_LostFocus()
txtCaseNo = ChkStr(txtCaseNo)
cmdShow.Default = False
End Sub

Private Sub txtCASRemark_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub

Private Sub txtCASRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 126 Then
        MsgBox "You are not allow to use this symbol as a Keyword", vbOKOnly + vbCritical, DialogBoxTitle
        KeyAscii = 0
    End If
End Sub

Private Sub txtCasREmark_LostFocus()
    KeyPreview = True
    cmdShow.Default = True
End Sub

Private Sub txtCASRemarkTwo_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub

Private Sub txtCASReserveAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtCASLGAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub txtDateAdmitted_GotFocus()
PrevDOA = txtDateAdmitted
End Sub

Private Sub txtDateAdmitted_LostFocus()
    If IsDate(txtDateAdmitted) Then
        If (txtDateDischarge <> "__/__/____") Then
            If (CDate(txtDateDischarge) < CDate(txtDateAdmitted)) Then
                MsgBox "Admission Date cannot Greater than Discharge Date! ", vbCritical
                txtDateAdmitted = PrevDOA
                txtDateAdmitted.SetFocus
            End If
        End If
        If txtDateAdmitted <> "__/__/____" Then
            If CDate(txtDateAdmitted) > Date And ChkDOA.value = 1 Then
                MsgBox "Admission Date cannot Greater than system date! ", vbCritical
                txtDateAdmitted.SetFocus
            End If
        End If

    Else
        If txtDateAdmitted <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateAdmitted.SetFocus
        End If
    End If
End Sub

Private Sub txtDateDischarge_LostFocus()
    If IsDate(txtDateDischarge) Then
        If txtDateAdmitted <> "__/__/____" Then
            If (CDate(txtDateDischarge) < CDate(txtDateAdmitted)) Then
                MsgBox "Discharged Date cannot less then Admission Date ! ", vbCritical
                txtDateDischarge.SetFocus
            End If
        End If
        If txtDateDischarge <> "__/__/____" Then
            If CDate(txtDateDischarge) > Date And ChkDOD.value = 1 Then
                MsgBox "Discharge Date cannot greater than system date! ", vbCritical
                txtDateDischarge.SetFocus
            End If
        End If
    Else
        If txtDateDischarge <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateDischarge.SetFocus
        End If
    End If
End Sub

Private Sub TxtDateOpen_LostFocus()
    If IsDate(TxtDateOpen) Then
    Else
    If TxtDateOpen <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
         TxtDateOpen.SetFocus
    End If
    End If
End Sub

Private Sub txtDateRECDoc_LostFocus()
    If IsDate(txtDateRECDoc) Then
    Else
    If txtDateRECDoc <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtDateRECDoc.SetFocus
    End If
    End If
End Sub



Private Sub txtDLR_LostFocus()
    If IsDate(txtDLR) Then
    Else
        If txtDLR <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
             txtDLR.SetFocus
        End If
    End If
End Sub

Private Sub txtDLS_LostFocus()
    If IsDate(txtDLS) Then
    Else
        If txtDLS <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDLS.SetFocus
        End If
    End If
End Sub

Private Sub TxtDRTC_LostFocus()
    If IsDate(TxtDRTC) Then
    Else
    If TxtDRTC <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
         TxtDRTC.SetFocus
    End If
    End If

End Sub

Private Sub txtDV_LostFocus()
    If IsDate(txtDV) Then
    Else
    If txtDV <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtDV.SetFocus
    End If
    End If
End Sub

Private Sub txtFDDOS1_LostFocus()
    If IsDate(txtFDDOS1) Then
    Else
    If txtFDDOS1 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtFDDOS1.SetFocus
    End If
    End If
End Sub

Private Sub txtFDDOS2_LostFocus()
    If IsDate(txtFDDOS2) Then
    Else
    If txtFDDOS2 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtFDDOS2.SetFocus
    End If
    End If
End Sub

Private Sub txtFDDOS3_LostFocus()
    If IsDate(txtFDDOS3) Then
    Else
    If txtFDDOS3 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtFDDOS3.SetFocus
    End If
    End If
End Sub

Private Sub txtFirstCall_LostFocus()
    If IsDate(txtFirstCall) Then
    Else
    If txtFirstCall <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtFirstCall.SetFocus
    End If
    End If
End Sub

Private Sub txtFirstDocDate_LostFocus()
    If IsDate(txtFirstDocDate) Then
        If txtSubDocDate <> "__/__/____" Then
            If (CDate(txtSubDocDate) < CDate(txtFirstDocDate)) Then
                MsgBox "First Doc. Date cannot greater than subsequent Doc. Date! ", vbCritical
                txtFirstDocDate.SetFocus
            End If
        End If
    Else
        If txtFirstDocDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtFirstDocDate.SetFocus
        End If
    End If
End Sub

Private Sub txtFirstTreated_LostFocus()
    If IsDate(txtFirstTreated) Then
    Else
    If txtFirstTreated <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtFirstTreated.SetFocus
    End If
    End If
End Sub

Private Sub txtHOSCode_LostFocus()
    If Len(txtHOSCode) > 60 Then
        MsgBox "Patient cannot exceed 60 character!", vbCritical
        txtHOSCode.SetFocus
    End If
End Sub

Private Sub txtIntCASRemark_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub

Private Sub txtInvDate_LostFocus()
    If IsDate(txtINVDate) Then
    Else
    If txtINVDate <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtINVDate.SetFocus
    End If
    End If

End Sub

Private Sub txtNSActFee1_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    D = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalNSActFee = f
End Sub

Private Sub txtNSActFee2_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    D = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalNSActFee = f
End Sub

Private Sub txtNSActFee3_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    D = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalNSActFee = f
End Sub

Private Sub txtNSActFee4_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    D = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalNSActFee = f
End Sub

Private Sub txtNSActFee5_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    D = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalNSActFee = f
End Sub

Private Sub txtNSMMARate1_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    D = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalNSMMARate = f
End Sub

Private Sub txtNSMMARate2_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    D = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalNSMMARate = f

End Sub

Private Sub txtNSMMARate3_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    D = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalNSMMARate = f

End Sub

Private Sub txtNSMMARate4_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    D = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalNSMMARate = f

End Sub

Private Sub txtNSMMARate5_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    D = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalNSMMARate = f

End Sub

Private Sub txtNSProcCode1_Change()
    If Len(Trim(txtNSProcCode1)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode1)
        Else
            Call SearchMMARateCal(txtNSProcCode1)
        End If
        txtNSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate1 = 0
    End If
End Sub

Private Sub txtNSProcCode2_Change()
    If Len(Trim(txtNSProcCode2)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode2)
        Else
            Call SearchMMARateCal(txtNSProcCode2)
        End If
        txtNSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate2 = 0
    End If

End Sub

Private Sub txtNSProcCode3_Change()
    If Len(Trim(txtNSProcCode3)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode3)
        Else
            Call SearchMMARateCal(txtNSProcCode3)
        End If
        txtNSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate3 = 0
    End If
End Sub

Private Sub txtNSProcCode4_Change()
    If Len(Trim(txtNSProcCode4)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode4)
        Else
            Call SearchMMARateCal(txtNSProcCode4)
        End If
        txtNSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate4 = 0
    End If
End Sub

Private Sub txtNSProcCode5_Change()
    If Len(Trim(txtNSProcCode5)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode5)
        Else
            Call SearchMMARateCal(txtNSProcCode5)
        End If
        txtNSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate5 = 0
    End If

End Sub

Private Sub txtPDDOS_LostFocus(Index As Integer)
    
    If IsDate(txtPDDOS(0)) Then
    Else
        If txtPDDOS(0) <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPDDOS(0).SetFocus
        End If
    End If
    
    If IsDate(txtPDDOS(1)) Then
    Else
        If txtPDDOS(1) <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPDDOS(1).SetFocus
        End If
    End If

    If IsDate(txtPDDOS(2)) Then
    Else
        If txtPDDOS(2) <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPDDOS(2).SetFocus
        End If
    End If

End Sub

Private Sub txtProcCode1_Change()
    If Len(Trim(txtProcCode1)) Then
    If MMARateCal = 0 Then
        Call SearchMMARate(txtProcCode1)
    Else
        Call SearchMMARateCal(txtProcCode1)
    End If
        txtSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate1 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate1 = 0
         txtAnaesMMARate1 = 0
    End If
End Sub

Private Sub txtProcCode2_Change()
    If Len(Trim(txtProcCode2)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtProcCode2)
        Else
            Call SearchMMARateCal(txtProcCode2)
        End If
        txtSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate2 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate2 = 0
         txtAnaesMMARate2 = 0
    End If

End Sub

Private Sub txtProcCode3_Change()
    If Len(Trim(txtProcCode3)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtProcCode3)
        Else
            Call SearchMMARateCal(txtProcCode3)
        End If
        txtSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate3 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate3 = 0
         txtAnaesMMARate3 = 0
    End If
End Sub

Private Sub txtProcCode4_Change()
    If Len(Trim(txtProcCode4)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtProcCode4)
        Else
            Call SearchMMARateCal(txtProcCode4)
        End If
        txtSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate4 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate4 = 0
         txtAnaesMMARate4 = 0
    End If
End Sub

Private Sub txtProcCode5_Change()
    If Len(Trim(txtProcCode5)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtProcCode5)
        Else
            Call SearchMMARateCal(txtProcCode5)
        End If
        txtSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate5 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate5 = 0
         txtAnaesMMARate5 = 0
    End If
End Sub



Private Sub txtReason_Click()
    KeyPreview = False
End Sub

Private Sub txtReason_LostFocus()
    KeyPreview = True
End Sub

Private Sub txtSActFee1_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    D = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalSActFee = f

End Sub

Private Sub txtSActFee2_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    D = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalSActFee = f

End Sub

Private Sub txtSActFee3_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    D = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalSActFee = f

End Sub

Private Sub txtSActFee4_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    D = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalSActFee = f

End Sub

Private Sub txtSActFee5_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    D = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtTotalSActFee = f

End Sub

Private Sub txtSMMARate1_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    D = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtSTotalMMARate1 = f

End Sub

Private Sub txtSMMARate2_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    D = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtSTotalMMARate1 = f

End Sub

Private Sub txtSMMARate3_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    D = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtSTotalMMARate1 = f

End Sub

Private Sub txtSMMARate4_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    D = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtSTotalMMARate1 = f

End Sub

Private Sub txtSMMARate5_Change()
Dim a, B, C, D, e, f As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    D = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(D) + CDbl(e)
    txtSTotalMMARate1 = f

End Sub

Private Sub txtState_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SSTab1.Tab = 1
        cboPatientList.SetFocus
    End If
End Sub

Private Sub txtState_LostFocus()
    SSTab1.Tab = 1
End Sub

Private Sub txtSubDocDate_LostFocus()
    If IsDate(txtSubDocDate) Then
        If txtFirstDocDate = "__/__/____" Then
            MsgBox "Please Enter First Doc. Date! ", vbCritical
            txtSubDocDate.Text = "__/__/____"
            txtFirstDocDate.SetFocus
        Else
            If (CDate(txtSubDocDate) < CDate(txtFirstDocDate)) Then
                MsgBox "First Doc. Date cannot greater than subsequent Doc. Date! ", vbCritical
                txtFirstDocDate.SetFocus
            End If
        End If
    Else
        If txtSubDocDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtSubDocDate.SetFocus
        End If
    End If

End Sub

Private Sub txtUDDOS1_LostFocus()
    If IsDate(txtUDDOS1) Then
    Else
    If txtUDDOS1 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtUDDOS1.SetFocus
    End If
    End If
End Sub

Private Sub txtUDDOS2_LostFocus()
    If IsDate(txtUDDOS2) Then
    Else
    If txtUDDOS2 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtUDDOS2.SetFocus
    End If
    End If
End Sub

Private Sub txtUDDOS3_LostFocus()
    If IsDate(txtUDDOS3) Then
    Else
    If txtUDDOS3 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtUDDOS3.SetFocus
    End If
    End If
End Sub

Private Sub saveHos()
On Error Resume Next
    Dim strsql As String
    Dim strsqlCAS As String
    Dim CAS As New ADODB.Recordset
    Dim strsqlHOS As String
    Dim HOS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim cnn As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
    ' INSERT INTO DIAFinal TABLE
'    strsql = "Select * from CAS_HOS where CaseCode = '" & Trim(tempcasNo) & "'"
 '   HOS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
 '           If HOS.recordCount = 0 Then
  '          HOS.AddNew
   '             HOS!CaseCode = ChkStr(tempcasNo)
    '        Else
     '       End If
'            With HOS
 '               !HosCode = Mid(ChkStr(txtHOSCode), 1, 6)
  '          End With
   '         HOS.Update
  '  HOS.Close
   ' Set HOS = Nothing
    
    'strsqlCAS = "Select * from CAS where CASNumber = '" & Trim(tempcasNo) & "'"
    'CAS.Open strsqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "7"
    strAppend = " AND CASNumber = '" & Trim(TempCASNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    CAS.CursorLocation = adUseClient
    CAS.CursorType = adOpenForwardOnly
    CAS.CacheSize = 100
    Set CAS = cmd.Execute
        
    HOspitalName = Replace(txtHOSCode, "'", "''")
    
    'strsqlHOS = "select HosCode, HosName from HOS where HosName = '" & Trim(HOspitalName) & "'"
    'HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "23"
    strAppend = " AND HosName = '" & Trim(HOspitalName) & "'"
    Set cnn.ActiveConnection = medixmasterconn
    cnn.CommandText = "Case_Adjustment"
    cnn.CommandType = adCmdStoredProc
    cnn.CommandTimeout = 300
    Set Prm1 = cnn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cnn.Parameters.Append Prm1
    Set Prm2 = cnn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cnn.Parameters.Append Prm2
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cnn.Execute
    
    If CAS.EOF = False Then
        If HOS.EOF = False Then
            strsqlHOS = "EXEC Case_AdjUpdateHOS '" & Trim(TempCASNo) & "', '" & ChkStr(HOS!HOSCode) & "'"
            medixmasterconn.Execute strsqlHOS
            'CAS!CASHOSCode = ChkStr(HOS!HOSCode)
            'CAS.Update
        End If
    End If
    
    CAS.Close
    Set CAS = Nothing
    HOS.Close
    Set HOS = Nothing
End Sub

Private Sub BNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim BNF As ListItem
    Dim sql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
    If Len(Trim(txtHLTCOde)) <> 0 _
     And Len(Trim(txtINSCode1)) <> 0 And Len(Trim(txtPLNCode)) <> 0 Then
        If Len(MBMCPLNCode) And Mid(cboPatientList, 1, 2) <> "00" Then
            strAppendCase = "17"
            strAppend = " AND PLNCOde ='" & Trim(MBMCPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(txtHLTCOde) & "'" & _
              " AND INSCODe ='" & Trim(txtINSCode1) & "'"
            'sql = " select BNFCode, BNFNoOFDays , " & _
              " BNFclaimableAmount, BNFdescription , BNFRMPerDis " & _
              " from View_BNF where PLNCOde ='" & Trim(MBMCPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(txtHLTCode) & "'" & _
              " AND INSCODe ='" & Trim(txtINSCode1) & "'"
        Else
            strAppendCase = "17"
            strAppend = " AND PLNCOde ='" & Trim(txtPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(txtHLTCOde) & "'" & _
              " AND INSCODe ='" & Trim(txtINSCode1) & "'"
            'sql = " select BNFCode, BNFNoOFDays , " & _
              " BNFclaimableAmount, BNFdescription , BNFRMPerDis " & _
              " from View_BNF where PLNCOde ='" & Trim(txtPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(txtHLTCode) & "'" & _
              " AND INSCODe ='" & Trim(txtINSCode1) & "'"
        End If
            'rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Adjustment"
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
            Set BNF = lstBenefits.listitems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFdescription) <> "" Then
                    BNF.SubItems(1) = rst2!BNFdescription
                End If
                BNF.SubItems(2) = IIf(IsNull(rst2!BNFNoOfDays), 0, rst2!BNFNoOfDays)
                 If Trim(rst2!BNFClaimAbleAmount) <> "" Then
                    BNF.SubItems(3) = Format(rst2!BNFClaimAbleAmount, "###,###,###.00")
                End If
                If Trim(rst2!BNFRMperDis) <> "" Then
                    BNF.SubItems(4) = Format(rst2!BNFRMperDis, "###,###,###.00")
                End If
                If rst2!BNFCode = "101" Then
                    RB = BNF.SubItems(3)
                End If
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub NewBNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim BNF As ListItem
    Dim sql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String

    If lstNewBenefitDetails.listitems.Count = 0 And Len(Trim(txtHLTCOde)) <> 0 _
     And Len(Trim(txtINSCode1)) <> 0 And Len(Trim(txtUpgPLNCode)) <> 0 Then
        'sql = " select BNFCode, BNFNoOFDays , " & _
          " BNFclaimableAmount, BNFdescription , BNFRMPerDis " & _
          " from View_BNF where PLNCOde ='" & Trim(txtUpgPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(txtHLTCOde) & "'" & _
          " AND INSCODe ='" & Trim(txtINSCode1) & "'"
        'rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "17"
        strAppend = " AND PLNCOde ='" & Trim(txtUpgPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(txtHLTCOde) & "'" & _
          " AND INSCODe ='" & Trim(txtINSCode1) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
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
            Set BNF = lstNewBenefitDetails.listitems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFdescription) <> "" Then
                    BNF.SubItems(1) = rst2!BNFdescription
                End If
                BNF.SubItems(2) = IIf(IsNull(rst2!BNFNoOfDays), 0, rst2!BNFNoOfDays)
                 If Trim(rst2!BNFClaimAbleAmount) <> "" Then
                    BNF.SubItems(3) = Format(rst2!BNFClaimAbleAmount, "###,###,###.00")
                End If
                If Trim(rst2!BNFRMperDis) <> "" Then
                    BNF.SubItems(4) = Format(rst2!BNFRMperDis, "###,###,###.00")
                End If
                If rst2!BNFCode = "101" Then
                    RB = BNF.SubItems(3)
                End If
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub LGListing()
    Dim LG As New ADODB.Recordset
    Dim LGList As ListItem
    Dim sql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
    If Len(Trim(txtMBMNumber)) <> 0 Then
        lstLGListing.listitems.Clear
        
        'sql = " Select CASeNo, LGNo, LGDate, " & _
            " LGAmt, LGIssuedBy, LGRemark  From CASLG " & _
            " where CASENo ='" & Trim(txtCaseNo) & "'"
    
        'LG.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        strAppendCase = "7"
        strAppend = " AND CASENo ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        LG.CursorLocation = adUseClient
        LG.CursorType = adOpenForwardOnly
        LG.CacheSize = 100
        Set LG = cmd.Execute
        
        If Not LG.EOF Then
            Do Until LG.EOF
                If LG!LGStatus = "B" Then
                    Set LGList = lstLGListing.listitems.Add(, , Trim("B" & LG!CaseNo & "-" & LG!LGNo))
                Else
                    Set LGList = lstLGListing.listitems.Add(, , Trim("L" & LG!CaseNo & "-" & LG!LGNo))
                End If
                With LG
                    If Len(!LGDate) Then
                        LGList.SubItems(1) = Format(!LGDate, "dd/mm/yyyy")
                    End If
                    If Len(!LGAmt) Then
                        LGList.SubItems(2) = Format(!LGAmt, "#,##0.00")
                    End If
                    If Len(cboPatientList) Then
                        LGList.SubItems(3) = Trim(Mid(ChkStr(cboPatientList), 1, InStr(1, ChkStr(cboPatientList), "-") - 1))
                    End If
                    If Len(PatientName) Then
                        LGList.SubItems(4) = Trim(PatientName)
                    End If
                    If Len(txtDateAdmitted) Then
                       LGList.SubItems(5) = Format(txtDateAdmitted, "dd/mm/yyyy")
                    End If
                    If Len(txtHOSCode) Then
                        LGList.SubItems(6) = Trim(txtHOSCode)
                    End If
                    If Len(!LGIssuedBy) Then
                        LGList.SubItems(7) = Trim(!LGIssuedBy)
                    End If
                    If Len(!LGRemark) Then
                        LGList.SubItems(8) = Trim(!LGRemark)
                    End If
                    If Len(!CaseNo) Then
                        LGList.SubItems(9) = Trim(!CaseNo & "-" & !LGNo)
                    End If
                    
                    .MoveNext
                End With
                Loop
                LG.Close
                Set LG = Nothing
            End If
        End If
End Sub

Private Sub ShowRepudiation()
Dim REP As New ADODB.Recordset
Dim REPCode As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strsqlREP As String
    ' INSERT INTO CASRepudiation TABLE
     
    'strsqlREP = "select * from CASRepudiation where CASNumber = '" & txtCaseNo & "'"
    'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "10"
    strAppend = " AND CASNumber = '" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        REP.CursorLocation = adUseClient
        REP.CursorType = adOpenForwardOnly
        REP.CacheSize = 100
        Set REP = cmd.Execute
        
    If Not REP.EOF Then
        With REP
            If Len(Trim(!CASREPCode1)) Then
                strsqlREP = "EXEC Case_Inquiry '" & Trim(!CASREPCode1) & "',7"
                REPCode.Open strsqlREP, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                With REPCode
                    cboREPCode1 = Trim(!REPCode) & " - " & Trim(!RepDescription)
                End With
                REPCode.Close
                Set REPCode = Nothing
            End If
            If Len(Trim(!CASREPCode2)) Then
                strsqlREP = "EXEC Case_Inquiry '" & Trim(!CASREPCode2) & "',7"
                REPCode.Open strsqlREP, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                With REPCode
                    cboREPCode2 = Trim(!REPCode) & " - " & Trim(!RepDescription)
                End With
                REPCode.Close
                Set REPCode = Nothing
            End If
            If Len(Trim(!CASREPCode3)) Then
                strsqlREP = "EXEC Case_Inquiry '" & Trim(!CASREPCode3) & "',7"
                REPCode.Open strsqlREP, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                With REPCode
                    cboREPCode3 = Trim(!REPCode) & " - " & Trim(!RepDescription)
                End With
                REPCode.Close
                Set REPCode = Nothing
            End If
            If Len(Trim(!CASDocument)) Then
                Document = !CASDocument
            End If
            If Len(Trim(!CASPersonnel)) Then
                Personnel = !CASPersonnel
            End If
            If Len(Trim(!CASAttention)) Then
                Attention = !CASAttention
            End If
        End With
    End If
    REP.Close
    Set REP = Nothing
End Sub

Public Sub ClearEntries(Optional KIV As Boolean)
   On Error Resume Next
   Dim ctl As Control

     For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
        ElseIf TypeOf ctl Is ComboBox Then
            If ctl.Enabled = True Then
                ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is MaskEdBox Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
        End If
    Next ctl
        cboStatus1 = "O - Open"
        lstMember.listitems.Clear
        lstBenefits.listitems.Clear
        lstNewBenefitDetails.listitems.Clear
        cboSuppExc.Clear
        txtCASLGAmt = 2500
        txtCASReserveAmount = 1600
        cboNature = "ILL - ILLNess"
        cboModality = "S - Surgical"
        cmdPage.Caption = "Next"
        txtCASRemark.Visible = True
        lblMemo.Visible = False
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
            ctl.value = False
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    
    End If
Next ctl
End Function

Private Function bookmark(Optional strAppend As String) As Integer
Dim strsqlPay As String
Dim PAY As New ADODB.Recordset
Dim strsqlREP As String
Dim REP As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim strsqlHOS As String
Dim strsqlLG As String
Dim LG As New ADODB.Recordset
Dim REPCat As New ADODB.Recordset
Dim strsqlREPCat As String
Dim LGNo As String
Dim LGVersion As String
Dim ContactPerson As String
Dim MBMExpiryDate As Long
Dim MBM As New ADODB.Recordset
Dim strsqlMBM As String
Dim Including As String
Dim Terms As String
Dim COPay As Boolean
Dim MRI As Boolean
Dim cmd As New ADODB.Command
Dim cmd1 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim cmd5 As New ADODB.Command
Dim cmd6 As New ADODB.Command
Dim cmd7 As New ADODB.Command
Dim cmd8 As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String

    If Letter = "T" Then
        bookmark = False
        If Len(txtCaseNo) Then
            objdoc.Bookmarks("CaseNo").Range.Text = txtCaseNo
        End If

        If Logon <> "" Then
            objdoc.Bookmarks("ppby").Range.Text = Logon
        End If
        
        If Trim(cboPatientList) <> "" Then
            objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
        'Else
        '    objdoc.Bookmarks("patient").Range.Text = txtPatientList
        End If

        If txtMBMNumber <> "" Then
            objdoc.Bookmarks("mbmno").Range.Text = txtMBMNumber
        End If
        
        If txtMBMPolNo <> "" Then
            objdoc.Bookmarks("PolicyNo").Range.Text = txtMBMPolNo
        End If
        
        If txtDateAdmitted <> "" Then
            objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
        End If
                
        If txtDateDischarge <> "" Then
            objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
        End If
                
        bookmark = True
    
    Else
    
    ContactPerson = ""
    LGNo = ""
    LGVersion = ""
    Including = ""
    Terms = ""
    
    'If txtHLTCOde = "S" Then
        'strsqlMBM = "Select PLNCode, CoPaymentInd, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(TxtPlanCode) & "'"
    'Else
        'strsqlMBM = "Select PLNCode, CoPaymentInd, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(TxtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
    'End If
    'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'strsqlPAY = "Select * from PAY where PAYCode = '" & Trim(txtPayCode) & "'"
    'PAY.Open strsqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "18"
    StrAppend1 = " AND PAYCode = '" & Trim(txtPayCode) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
        
    HOspitalName = Replace(txtHOSCode, "'", "''")
    
    'strsqlHOS = "select * from HOS where HOSName = '" & Trim(HOspitalName) & "'"
    'HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "6"
    StrAppend1 = " AND HOSName = '" & Trim(HOspitalName) & "'"
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "Case_Adjustment"
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd1.Parameters.Append Prm2
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cmd1.Execute
    
    If Letter = "L" Then
        bookmark = False
        
        If Trim(cboPatientList) <> "" Then
          objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
        Else
          objdoc.Bookmarks("patient").Range.Text = " Please key in patient name here!"
        End If
        
        If ReprintLG = False Then
            If LGNumber <> "" Then
                objdoc.Bookmarks("LGno").Range.Text = LGNumber
            End If
            
            LGVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
            If LGVer <> "1" Then
                objdoc.Bookmarks("Note").Range.Text = "NOTE: THIS GUARANTEE LETTER SUPERCEDES ALL PREVIOUS GUARANTEE LETTERS"
            End If
            
            If txtCASLGAmt <> "" Then
                objdoc.Bookmarks("amount").Range.Text = Format(txtCASLGAmt, "#,##0.00")
            End If
            
        Else
            If Len(lstLGListing.SelectedItem.Text) Then
                    objdoc.Bookmarks("LGno").Range.Text = Trim(lstLGListing.SelectedItem.Text)
            End If
            
            LGVer = Trim(Mid(lstLGListing.SelectedItem.SubItems(9), InStr(1, lstLGListing.SelectedItem.SubItems(9), "-") + 1, Len(lstLGListing.SelectedItem.SubItems(9)) - InStr(1, lstLGListing.SelectedItem.SubItems(9), "-")))
            If LGVer <> "1" Then
                objdoc.Bookmarks("Note").Range.Text = "NOTE: THIS GUARANTEE LETTER SUPERCEDES ALL PREVIOUS GUARANTEE LETTERS"
            End If
            
            objdoc.Bookmarks("amount").Range.Text = Format(lstLGListing.SelectedItem.SubItems(2), "#,##0.00")
                
        End If
        
        objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
        
        If txtMBMNumber <> "" Then
            objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
        End If
        If txtHOSCode <> "" Then
            objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
        End If
        If txtDateAdmitted <> "" Then
            objdoc.Bookmarks("dateofadmission").Range.Text = txtDateAdmitted
        End If
        If txtMBMPolNo <> "" Then
            objdoc.Bookmarks("policyno").Range.Text = txtMBMPolNo
        End If
        
        If RB <> "" Then
            objdoc.Bookmarks("RB").Range.Text = RB
        End If
        
        If Len(HOS!HOSContact) Then
            objdoc.Bookmarks("Attention").Range.Text = HOS!HOSContact
        End If
    
        If Len(HOS!HOSFaxNo) Then
            objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
        End If
        
        If lstLGListing.listitems.Count <> 0 Then
            objdoc.Bookmarks("Remark").Range.Text = Trim(lstLGListing.SelectedItem.SubItems(8))
        Else
            If TxtLGRemark <> "" Then
                objdoc.Bookmarks("Remark").Range.Text = Trim(TxtLGRemark)
            End If
        End If
        
        If txtHLTCOde = "S" Then
            'strsqlMBM = "Select PLNCode, CoPaymentInd, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(TxtPlanCode) & "'"
            strAppendCase = "19"
            StrAppend1 = " AND PLNCode = '" & Trim(TxtPlanCode) & "'"
        Else
            'strsqlMBM = "Select PLNCode, CoPaymentInd, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(TxtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
            strAppendCase = "19"
            StrAppend1 = " AND PLNCode = '" & Trim(TxtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
        End If
        'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
        Set cmd2.ActiveConnection = medixmasterconn
        cmd2.CommandText = "Case_Adjustment"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
        cmd2.Parameters.Append Prm1
        Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd2.Parameters.Append Prm2
        MBM.CursorLocation = adUseClient
        MBM.CursorType = adOpenForwardOnly
        MBM.CacheSize = 100
        Set MBM = cmd2.Execute
        
        If MBM.EOF = False Then
        
            If MBM!PLNMeal = "Y" Then
                Including = Including + "Meal"
            End If
        
            If MBM!PLNNursing = "Y" Then
                Including = Including + " " & ",Nursing Charges"
            End If
        
            If MBM!PLNTax = "Y" Then
                Including = Including + " " & ",Tax"
            End If
        
            If MBM!CoPaymentInd = "Y" Then
                COPay = True
            Else
                COPay = False
            End If
        
            If MBM!PLNMRI = "Y" Then
                MRI = True
            Else
                MRI = False
            End If
        Else
            COPay = False
            MRI = False
        End If
        
        If Mid(txtPLNCode, 1, 1) = "4" Then
            Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                     vbNewLine & "       physician." & _
                     vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                     vbNewLine & "4)    The following items are not covered: -" & _
                     vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                     vbNewLine & "           (b) Non-Medical items." & _
                     vbNewLine & "           (c) Medical Report Charges." & _
                     vbNewLine & "           (d) Registration Fees, Admission Fees, Billing Fees and Admission Kit." & _
                     vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Lenses, Pacemaker, Braces, Corset " & _
                     vbNewLine & "           (f) Lodger" & _
                     vbNewLine & "           (g) Supplies & Services not related to illness injury." & _
                     vbNewLine & "5)    Co-Payment - In the event the Room and Board limit is exceeded, the Insurance will pay only 80% of " & _
                     vbNewLine & "       other eligible costs."
        
        ElseIf COPay = True And MRI = True Then

            Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                     vbNewLine & "       physician." & _
                     vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                     vbNewLine & "4)    The following items are not covered: -" & _
                     vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                     vbNewLine & "           (b) Non-Medical items." & _
                     vbNewLine & "           (c) Medical Report Charges." & _
                     vbNewLine & "           (d) Registration Fees, Admission Fees, Billing Fees, Dispensing Fees and Admission Kit." & _
                     vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Lenses, Pacemaker, Braces, Corset, Arm sling and Cervical" & _
                     vbNewLine & "                Collar." & _
                     vbNewLine & "           (f) Lodger" & _
                     vbNewLine & "           (g) Supplies & Services not related to illness injury." & _
                     vbNewLine & "5)    Co-Payment - In the event the Room and Board limit is exceeded, the Insurance will pay only 80% of " & _
                     vbNewLine & "       other eligible costs." & _
                     vbNewLine & "6)    Insurance will pay only 80% of MRI, CT Scan and Heart Scan charges."
                            
        ElseIf COPay = False And MRI = True Then
            
            Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                     vbNewLine & "       physician." & _
                     vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                     vbNewLine & "4)    The following items are not covered: -" & _
                     vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                     vbNewLine & "           (b) Non-Medical items." & _
                     vbNewLine & "           (c) Medical Report Charges." & _
                     vbNewLine & "           (d) Registration Fees, Admission Fees, Billing Fees, Dispensing Fees and Admission Kit." & _
                     vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Lenses, Pacemaker, Braces, Corset, Arm sling and Cervical" & _
                     vbNewLine & "                Collar." & _
                     vbNewLine & "           (f) Lodger" & _
                     vbNewLine & "           (g) Supplies & Services not related to illness injury." & _
                     vbNewLine & "5)    Insurance will pay only 80% of MRI, CT Scan and Heart Scan charges."
                                 
        ElseIf COPay = True And MRI = False Then
                 
            Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                     vbNewLine & "       physician." & _
                     vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                     vbNewLine & "4)    The following items are not covered: -" & _
                     vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                     vbNewLine & "           (b) Non-Medical items." & _
                     vbNewLine & "           (c) Medical Report Charges." & _
                     vbNewLine & "           (d) Registration Fees, Admission Fees, Billing Fees, Dispensing Fees and Admission Kit." & _
                     vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Lenses, Pacemaker, Braces, Corset, Arm sling and Cervical" & _
                     vbNewLine & "                Collar." & _
                     vbNewLine & "           (f) Lodger" & _
                     vbNewLine & "           (g) Supplies & Services not related to illness injury." & _
                     vbNewLine & "5)    Co-Payment - In the event the Room and Board limit is exceeded, the Insurance will pay only 80% of " & _
                     vbNewLine & "       other eligible costs."
            
        ElseIf COPay = False And MRI = False Then
            
            Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                     vbNewLine & "       physician." & _
                     vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                     vbNewLine & "4)    The following items are not covered: -" & _
                     vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                     vbNewLine & "           (b) Non-Medical items." & _
                     vbNewLine & "           (c) Medical Report Charges." & _
                     vbNewLine & "           (d) Registration Fees, Admission Fees, Billing Fees, Dispensing Fees and Admission Kit." & _
                     vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Lenses, Pacemaker, Braces, Corset, Arm sling and Cervical" & _
                     vbNewLine & "                Collar." & _
                     vbNewLine & "           (f) Lodger" & _
                     vbNewLine & "           (g) Supplies & Services not related to illness injury."
                    
        End If
        
        objdoc.Bookmarks("including").Range.Text = Including
        objdoc.Bookmarks("string").Range.Text = Terms
        bookmark = True
        
    ElseIf Letter = "B" Then
        bookmark = False
        
        If Trim(cboPatientList) <> "" Then
            objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
        Else
            objdoc.Bookmarks("patient").Range.Text = " Please key in patient name here!"
        End If
        
        If ReprintBTBG = False Then
            If LGNumber <> "" Then
                objdoc.Bookmarks("BTBNo").Range.Text = LGNumber
            End If
        Else
            If Len(lstLGListing.SelectedItem.Text) Then
                objdoc.Bookmarks("BTBNo").Range.Text = Trim(lstLGListing.SelectedItem.Text)
            End If
        End If
        
        objdoc.Bookmarks("BTBID").Range.Text = Trim(Logon)
        
        If txtMBMNumber <> "" Then
            objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
        End If
        If txtHOSCode <> "" Then
            objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
        End If
        If txtDateAdmitted <> "" Then
            objdoc.Bookmarks("dateofadmission").Range.Text = txtDateAdmitted
        End If
        If txtMBMPolNo <> "" Then
            objdoc.Bookmarks("policyno").Range.Text = txtMBMPolNo
        End If
        
        If RB <> "" Then
            objdoc.Bookmarks("RB").Range.Text = RB
        End If
        
        If Len(HOS!HOSContact) Then
            objdoc.Bookmarks("Attention").Range.Text = HOS!HOSContact
        End If
    
        If Len(HOS!HOSFaxNo) Then
            objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
        End If
                
        If txtHLTCOde = "S" Then
            'strsqlMBM = "Select PLNCode, CoPaymentInd, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(TxtPlanCode) & "'"
            strAppendCase = "19"
            StrAppend1 = " AND PLNCode = '" & Trim(TxtPlanCode) & "'"
        Else
            'strsqlMBM = "Select PLNCode, CoPaymentInd, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(TxtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
            strAppendCase = "19"
            StrAppend1 = " AND PLNCode = '" & Trim(TxtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
        End If
        'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
        Set cmd2.ActiveConnection = medixmasterconn
        cmd2.CommandText = "Case_Adjustment"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
        cmd2.Parameters.Append Prm1
        Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd2.Parameters.Append Prm2
        MBM.CursorLocation = adUseClient
        MBM.CursorType = adOpenForwardOnly
        MBM.CacheSize = 100
        Set MBM = cmd2.Execute
        
        If MBM.EOF = False Then
        
            If MBM!PLNMeal = "Y" Then
                Including = Including + "Meal"
            End If
        
            If MBM!PLNNursing = "Y" Then
                Including = Including + " " & ",Nursing Charges"
            End If
        
            If MBM!PLNTax = "Y" Then
                Including = Including + " " & ",Tax"
            End If
        End If

        Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                     vbNewLine & "        physician." & _
                      vbNewLine & "3)    Items of personal nature such as telephone calls, newspaper, laundry, television etc."
        
        objdoc.Bookmarks("including").Range.Text = Including
        objdoc.Bookmarks("string").Range.Text = Terms
        bookmark = True
        
    ElseIf Letter = "R" Then
        bookmark = False
        
        objdoc.Bookmarks("PrincipalName").Range.Text = txtMBMName
        objdoc.Bookmarks("Add1").Range.Text = txtAdd1
        objdoc.Bookmarks("Add2").Range.Text = txtAdd2 & txtAdd3
        objdoc.Bookmarks("PostCode").Range.Text = txtPostCode & "" & txtCity
        objdoc.Bookmarks("State").Range.Text = txtState
        
        If txtCaseNo <> "" Then
            objdoc.Bookmarks("CASENo").Range.Text = txtCaseNo
        End If

        If Trim(cboPatientList) <> "" Then
          objdoc.Bookmarks("patientname").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
        Else
          objdoc.Bookmarks("patientname").Range.Text = " Please key in patient name here!"
        End If
        
        If txtMBMNumber <> "" Then
            objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
        End If
        
        If txtFinalDiag1 <> "" Then
            Dim strsql As String
            strsql = strsql + txtFinalDiag1
            If txtFinalDiag2 <> "" Then
                strsql = strsql + ", " + txtFinalDiag2
                If txtFinalDiag3 <> "" Then
                    strsql = strsql + ", " + txtFinalDiag3
                End If
            End If
            objdoc.Bookmarks("Reasons").Range.Text = strsql
        End If
        
        If Underlying = True Then
            If txtUD1 <> "" Then
                Dim strsqlUD As String
                strsqlUD = " With underlying illness as "
                strsqlUD = strsqlUD + txtUD1
                If txtUD2 <> "" Then
                    strsqlUD = strsqlUD + ", " + txtUD2
                    If txtUD3 <> "" Then
                        strsqlUD = strsqlUD + ", " + txtUD3
                    End If
                End If
                objdoc.Bookmarks("Underlying").Range.Text = strsqlUD + "."
            End If
        End If
        
        If txtHOSCode <> "" Then
            objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
        End If
        If txtHOSCode <> "" Then
            objdoc.Bookmarks("hospital2").Range.Text = Trim(txtHOSCode)
        End If
        If txtMBMExpiryDate <> "" Then
            objdoc.Bookmarks("ExpDate").Range.Text = txtMBMExpiryDate
        End If
        If txtDateAdmitted <> "" Then
            objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
        End If
        If txtMBMPolNo <> "" Then
            objdoc.Bookmarks("policynumber").Range.Text = txtMBMPolNo
        End If
        If txtMBMName <> "" Then
            objdoc.Bookmarks("InsuredName").Range.Text = txtMBMName
        End If
        If Document <> "" Then
            objdoc.Bookmarks("REPDocument").Range.Text = Trim(Mid(Document, InStr(1, Document, "-") + 1, Len(Document)))
        End If
        If Personnel <> "" Then
            objdoc.Bookmarks("REPPersonnel").Range.Text = Trim(Mid(Personnel, InStr(1, Personnel, "-") + 1, Len(Personnel)))
        End If
        If Attention <> "" Then
            objdoc.Bookmarks("REPAttention").Range.Text = Trim(Mid(Attention, InStr(1, Attention, "-") + 1, Len(Attention)))
        End If
        
        If Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - 1)) <> "RC6" Then
            If Len(Trim(cboREPCode1)) Then
                'strsqlREP = "Select * from REPCode where RepCode = '" & Trim(Mid(cboREPCode1, 1, 7)) & "'"
                'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "20"
                StrAppend1 = " AND RepCode = '" & Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - 1)) & "'"
                Set cmd3.ActiveConnection = medixmasterconn
                cmd3.CommandText = "Case_Adjustment"
                cmd3.CommandType = adCmdStoredProc
                cmd3.CommandTimeout = 300
                Set Prm1 = cmd3.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd3.Parameters.Append Prm1
                Set Prm2 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd3.Parameters.Append Prm2
                REP.CursorLocation = adUseClient
                REP.CursorType = adOpenForwardOnly
                REP.CacheSize = 100
                Set REP = cmd3.Execute
        
                'strsqlREPCat = "Select * from REPCodeCategory where REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                'REPCat.Open strsqlREPCat, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "21"
                StrAppend1 = " AND REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                Set cmd4.ActiveConnection = medixmasterconn
                cmd4.CommandText = "Case_Adjustment"
                cmd4.CommandType = adCmdStoredProc
                cmd4.CommandTimeout = 300
                Set Prm1 = cmd4.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd4.Parameters.Append Prm1
                Set Prm2 = cmd4.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd4.Parameters.Append Prm2
                REPCat.CursorLocation = adUseClient
                REPCat.CursorType = adOpenForwardOnly
                REPCat.CacheSize = 100
                Set REPCat = cmd4.Execute
                
                If Len(REPCat!RepCatName) Then
                    objdoc.Bookmarks("REPCatNamePol1").Range.Text = Trim(REPCat!RepCatName)
                    objdoc.Bookmarks("REPCatName1").Range.Text = Trim(REPCat!RepCatName)
                End If
                
                If Len(REPCat!REPCatDescription) Then
                    objdoc.Bookmarks("REPCat1Desc").Range.Text = Trim(REPCat!REPCatDescription)
                End If
                
                objdoc.Bookmarks("REPReason1").Range.Text = Trim(REP!RepDescription)
        
                REP.Close
                REPCat.Close
                Set REP = Nothing
                Set REPCat = Nothing
            End If
            
            If Len(Trim(cboREPCode2)) Then
                'strsqlREP = "Select * from REPCode where RepCode = '" & Trim(Mid(cboREPCode2, 1, 7)) & "'"
                'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "20"
                StrAppend1 = " AND RepCode = '" & Trim(Mid(cboREPCode2, 1, InStr(1, cboREPCode2, "-") - 1)) & "'"
                Set cmd5.ActiveConnection = medixmasterconn
                cmd5.CommandText = "Case_Adjustment"
                cmd5.CommandType = adCmdStoredProc
                cmd5.CommandTimeout = 300
                Set Prm1 = cmd5.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd5.Parameters.Append Prm1
                Set Prm2 = cmd5.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd5.Parameters.Append Prm2
                REP.CursorLocation = adUseClient
                REP.CursorType = adOpenForwardOnly
                REP.CacheSize = 100
                Set REP = cmd5.Execute
                
                'strsqlREPCat = "Select * from REPCodeCategory where REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                'REPCat.Open strsqlREPCat, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "21"
                StrAppend1 = " AND REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                Set cmd6.ActiveConnection = medixmasterconn
                cmd6.CommandText = "Case_Adjustment"
                cmd6.CommandType = adCmdStoredProc
                cmd6.CommandTimeout = 300
                Set Prm1 = cmd6.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd6.Parameters.Append Prm1
                Set Prm2 = cmd6.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd6.Parameters.Append Prm2
                REPCat.CursorLocation = adUseClient
                REPCat.CursorType = adOpenForwardOnly
                REPCat.CacheSize = 100
                Set REPCat = cmd6.Execute
                
                If Not REPCat.EOF Then
                    objdoc.Bookmarks("REPCatNamePol2").Range.Text = ", " & Trim(REPCat!RepCatName)
                    objdoc.Bookmarks("REPCatName2").Range.Text = Trim(REPCat!RepCatName)
                End If
                
                If Len(REPCat!REPCatDescription) Then
                    objdoc.Bookmarks("REPCat2Desc").Range.Text = Trim(REPCat!REPCatDescription)
                End If
    
                objdoc.Bookmarks("REPReason2").Range.Text = Trim(REP!RepDescription)
                
                REP.Close
                REPCat.Close
                Set REP = Nothing
                Set REPCat = Nothing
            End If
            
            If Len(Trim(cboREPCode3)) Then
                'strsqlREP = "Select * from REPCode where RepCode = '" & Trim(Mid(cboREPCode3, 1, 7)) & "'"
                'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "20"
                StrAppend1 = " AND RepCode = '" & Trim(Mid(cboREPCode3, 1, InStr(1, cboREPCode3, "-") - 1)) & "'"
                Set cmd7.ActiveConnection = medixmasterconn
                cmd7.CommandText = "Case_Adjustment"
                cmd7.CommandType = adCmdStoredProc
                cmd7.CommandTimeout = 300
                Set Prm1 = cmd7.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd7.Parameters.Append Prm1
                Set Prm2 = cmd7.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd7.Parameters.Append Prm2
                REP.CursorLocation = adUseClient
                REP.CursorType = adOpenForwardOnly
                REP.CacheSize = 100
                Set REP = cmd7.Execute
                
                'strsqlREPCat = "Select * from REPCodeCategory where REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                'REPCat.Open strsqlREPCat, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "21"
                StrAppend1 = " AND REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                Set cmd8.ActiveConnection = medixmasterconn
                cmd8.CommandText = "Case_Adjustment"
                cmd8.CommandType = adCmdStoredProc
                cmd8.CommandTimeout = 300
                Set Prm1 = cmd8.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd8.Parameters.Append Prm1
                Set Prm2 = cmd8.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd8.Parameters.Append Prm2
                REPCat.CursorLocation = adUseClient
                REPCat.CursorType = adOpenForwardOnly
                REPCat.CacheSize = 100
                Set REPCat = cmd8.Execute
                
                If Not REP.EOF Then
                    objdoc.Bookmarks("REPCatNamePol3").Range.Text = ", " & Trim(REPCat!RepCatName)
                    objdoc.Bookmarks("REPCatName3").Range.Text = Trim(REPCat!RepCatName)
                End If
                
                If Len(REPCat!REPCatDescription) Then
                    objdoc.Bookmarks("REPCat3Desc").Range.Text = REPCat!REPCatDescription
                End If
    
                objdoc.Bookmarks("REPReason3").Range.Text = Trim(REP!RepDescription)
                    
                REP.Close
                REPCat.Close
                Set REP = Nothing
                Set REPCat = Nothing
            End If
            
        End If
        
        If txtPayCode <> "" Then
            objdoc.Bookmarks("Payor").Range.Text = PAY!PAYCompanyName
        End If
        
        If txtPayCode <> "" Then
            objdoc.Bookmarks("Payor2").Range.Text = PAY!PAYCompanyName
        End If
        
        If Len(PAY!PAYAdmissionContact) Then
            objdoc.Bookmarks("Pic").Range.Text = Trim(PAY!PAYAdmissionContact)
        End If
        
        'strsqlLG = "Select CASENo from CASLG where CASENo = '" & Trim(txtCaseNo) & "'"
        'LG.Open strsqlLG, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        'If LG.recordCount > 0 Then
        '    objdoc.Bookmarks("LDDesc").Range.Text = "MediExpress has facilitated your admission and settled your hospitalization " & _
        '                                            "charges amounting to RM2,500.00. We enclose a copy of the bill for your records. " & _
        '                                            "We would be pleased if you could let us have your payment to MediExpress(M) " & _
        '                                            "Sdn. Bhd. within 14 days herewith."
        'End If
        'LG.Close
        'Set LG = Nothing
        bookmark = True
    ElseIf Letter = "P" Then
        bookmark = False
        If Not PAY.EOF Then
            objdoc.Bookmarks("Payor").Range.Text = PAY!PAYCompanyName
            objdoc.Bookmarks("ADD1").Range.Text = PAY!PAYAddress1
            objdoc.Bookmarks("ADD2").Range.Text = PAY!PAYAddress2
            objdoc.Bookmarks("PostCode").Range.Text = PAY!PAYAddress3
            ContactPerson = PAY!PAYMAINContact
        End If
        objdoc.Bookmarks("REGDate").Range.Text = Date
        objdoc.Bookmarks("Month").Range.Text = Format(txtMBMExpiryDate, "MMMM YYYY")
            
        If txtCaseNo <> "" Then
            objdoc.Bookmarks("CASENo").Range.Text = txtCaseNo
        End If
    
        If Trim(cboPatientList) <> "" Then
          objdoc.Bookmarks("patientname").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
        Else
          objdoc.Bookmarks("patientname").Range.Text = " Please key in patient name here!"
        End If
        If txtMBMPolNo <> "" Then
            objdoc.Bookmarks("policynumber").Range.Text = txtMBMPolNo
        End If
        If txtMBMExpiryDate <> "__/__/____" Then
            objdoc.Bookmarks("EXPDate").Range.Text = txtMBMExpiryDate
        End If
        If txtMBMNumber <> "" Then
            objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
        End If
        If txtDateAdmitted <> "__/__/____" Then
            objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
        End If
        If txtDateDischarge <> "__/__/____" Then
            objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
        End If
        If txtHOSCode <> "" Then
            objdoc.Bookmarks("hospital").Range.Text = txtHOSCode
        End If
        If txtReason <> "" Then
            objdoc.Bookmarks("Diag").Range.Text = txtReason
        End If
        
        bookmark = True
    
    ElseIf Letter = "E" Then
        bookmark = False
        If Not PAY.EOF Then
            objdoc.Bookmarks("Payor").Range.Text = PAY!PAYCompanyName
            objdoc.Bookmarks("ADD1").Range.Text = PAY!PAYAddress1
            objdoc.Bookmarks("ADD2").Range.Text = PAY!PAYAddress2
            objdoc.Bookmarks("PostCode").Range.Text = PAY!PAYAddress3
            ContactPerson = PAY!PAYMAINContact
        End If
        objdoc.Bookmarks("REGDate").Range.Text = Date
        objdoc.Bookmarks("Month").Range.Text = Format(txtMBMExpiryDate, "MMMM YYYY")
            
        If txtCaseNo <> "" Then
            objdoc.Bookmarks("CASENo").Range.Text = txtCaseNo
        End If
        If Trim(cboPatientList) <> "" Then
          objdoc.Bookmarks("patientname").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
        Else
          objdoc.Bookmarks("patientname").Range.Text = " Please key in patient name here!"
        End If
        If txtMBMPolNo <> "" Then
            objdoc.Bookmarks("policynumber").Range.Text = txtMBMPolNo
        End If
        If txtMBMExpiryDate <> "__/__/____" Then
            objdoc.Bookmarks("EXPDate").Range.Text = txtMBMExpiryDate
        End If
        If txtMBMNumber <> "" Then
            objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
        End If
        If txtDateAdmitted <> "__/__/____" Then
            objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
        End If
        If txtDateDischarge <> "__/__/____" Then
            objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
        End If
        If txtHOSCode <> "" Then
            objdoc.Bookmarks("hospital").Range.Text = txtHOSCode
        End If
        If txtReason <> "" Then
            objdoc.Bookmarks("Diag").Range.Text = txtReason
        End If
                
        bookmark = True
        
    ElseIf Letter = "M" Then
        If Trim(txtCaseNo) <> "" Then
            objdoc.Bookmarks("CASE").Range.Text = Trim(txtCaseNo)
        End If
            
        If Trim(txtCASRemark) <> "" Then
          objdoc.Bookmarks("Note").Range.Text = Trim(txtCASRemark) & Trim(txtCASRemarkTwo)
        End If
    End If

    HOS.Close
    Set HOS = Nothing
    End If
End Function

Private Function bookmark1(Optional strAppend As String) As Integer
Dim Address As New ADODB.Recordset
Dim CurrentDate, MyDate, MyDay As String
Dim i As Integer
Dim strsql As String
        
        bookmark1 = False
    
        If Trim(txtHOSCode) <> "" Then
            objdoc.Bookmarks("CLNName").Range.Text = Trim(ChkStr(txtHOSCode))
        Else
            objdoc.Bookmarks("CLNName").Range.Text = " Please Select a Hospital"
        End If
            
        If FOLNumber <> "" Then
            objdoc.Bookmarks("LGNo").Range.Text = FOLNumber
        End If
        
        If txtMBMNumber <> "" Then
            objdoc.Bookmarks("MBMNo").Range.Text = txtMBMNumber
        End If
        
        If cboPatientList <> "" Then
            objdoc.Bookmarks("Patient").Range.Text = StrConv(Mid(cboPatientList, 6, Len(cboPatientList)), vbProperCase)
        End If
        
        If cboPatientList <> "" Then
            For i = 1 To lstMember.listitems.Count
                If StrComp(Mid(cboPatientList, 1, 2), lstMember.listitems.Item(i)) = 0 Then
                    objdoc.Bookmarks("ICNo").Range.Text = lstMember.listitems(i).SubItems(2)
                End If
            Next
        End If
        
        If txtATTDoc(0) <> "" Then
            objdoc.Bookmarks("Doctor").Range.Text = StrConv(txtATTDoc(0), vbProperCase)
        Else
            objdoc.Bookmarks("Doctor").Range.Text = "Not Specified"
        End If
        
        If IsDate(txtFollowUpDate) = True Or txtFollowUpDate <> "__/__/____" Then
            MyDay = Format(txtFollowUpDate, "dddd")
            MyDate = Format(txtFollowUpDate, "dd mmmm yyyy")
            CurrentDate = Format(Date, "dd mmmm yyyy")
            
            If Len(MyDate) <> "" Then
                objdoc.Bookmarks("Date1").Range.Text = MyDay & ", " & MyDate
            End If
                
            If Len(CurrentDate) <> "" Then
                objdoc.Bookmarks("Date").Range.Text = CurrentDate
            End If
        End If
        If txtHOSCode <> "" Then
            strsql = "EXEC SearchHOS"
            Address.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
            
            Do Until Address.EOF
                If StrComp(ChkStr(txtHOSCode), Address!HOSName) = 0 Then
                    If Trim(Address!HOSAdd1) <> "" Then
                        objdoc.Bookmarks("Add1").Range.Text = StrConv(Address!HOSAdd1, vbProperCase)
                    End If
                    If Trim(Address!HOSAdd2) <> "" Then
                        objdoc.Bookmarks("Add2").Range.Text = StrConv(Address!HOSAdd2, vbProperCase)
                    End If
                    If Trim(Address!HOSAdd3) <> "" Then
                        objdoc.Bookmarks("Add3").Range.Text = StrConv(Address!HOSAdd3, vbProperCase)
                    End If
                    If Not Address!HOSContact = "" Then
                        objdoc.Bookmarks("ContactPerson").Range.Text = Address!HOSContact
                    Else
                        objdoc.Bookmarks("ContactPerson").Range.Text = " Please enter contact person's name "
                    End If
                    If Not Address!HOSFaxNo = "" Then
                        objdoc.Bookmarks("FaxNo").Range.Text = Address!HOSFaxNo
                    End If
                End If
                Address.MoveNext
            Loop
            Address.Close
            Set Address = Nothing
        End If
    bookmark1 = True
End Function

Private Sub funcword()
On Error GoTo openerror

    'If Letter = "L" Then
        'If txtHLTCOde = "S" Then
            'strsqlMBM = "Select PLNCode, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(TxtPlanCode) & "'"
        'Else
            'strsqlMBM = "Select PLNCode, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(TxtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
        'End If
        'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    'End If
    
    ' make sure only one word application launched
    'LocCardPath = crdPath
    If Letter = "L" And tmpPLNPerVisit <> "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\New LG-2.DOT\")
                                                
        End If
        
        If bookmark("New LG-2.DOT") = True Then
                
            objdoc.SaveAs (crdPath & "Template\LG\" & Trim(LGNumber) & ".doc")
            Me.Show
            If MsgBox("Do you want to print this document?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
                 objdoc.SaveAs (crdPath & "Template\LG\" & Trim(LGNumber) & ".doc")
                 objdoc.PrintOut
                 objdoc.Close
            End If
        End If
        Exit Sub
    ElseIf Letter = "L" And tmpPLNPerVisit = "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\LGVisit.DOT\")
                                                
        End If
        
        If bookmark("LGVisit.DOT") = True Then
                
            objdoc.SaveAs (crdPath & "Template\LG\" & Trim(LGNumber) & ".doc")
            Me.Show
            If MsgBox("Do you want to print this document?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
                 objdoc.SaveAs (crdPath & "Template\LG\" & Trim(LGNumber) & ".doc")
                 objdoc.PrintOut
                 objdoc.Close
            End If
        End If
        Exit Sub
        
     ElseIf Letter = "B" And tmpPLNPerVisit <> "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\BTBG.DOT\")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
              vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              objdoc.Save
              objdoc.Close
            ElseIf vbCancel Then
              Exit Sub
            End If
        End If
        
        If bookmark("BTBG.DOT") = True Then
                Me.Show
        End If
        Exit Sub
        
     ElseIf Letter = "B" And tmpPLNPerVisit = "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\BTBGVisit.DOT\")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
              vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              objdoc.Save
              objdoc.Close
            ElseIf vbCancel Then
              Exit Sub
            End If
        End If
        
        If bookmark("BTBGVisit.DOT") = True Then
                Me.Show
        End If
        Exit Sub
        
        'ElseIf bookmark("LG2805012.DOT") = True Then
            'objdoc.SaveAs (crdPath & "Template\LG\" & Trim(LGNumber) & ".doc")
            'Me.show
            'If MsgBox("Do you want to print this document?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
                 'objdoc.SaveAs (crdPath & "Template\LG\" & Trim(LGNumber) & ".doc")
                 'objdoc.PrintOut
                 'objdoc.Close
                 'Exit Sub
            'Else
                
                'Exit Sub
            'End If
        'ElseIf bookmark("LG2805013.DOT") = True Then
            'objdoc.SaveAs (crdPath & "Template\LG\" & Trim(LGNumber) & ".doc")
            'Me.show
            'If MsgBox("Do you want to print this document?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
                 'objdoc.SaveAs (crdPath & "Template\LG\" & Trim(LGNumber) & ".doc")
                 'objdoc.PrintOut
                 'objdoc.Close
                 'Exit Sub
            'Else
                
                'Exit Sub
            'End If
        'End If
    
    ElseIf Letter = "R" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                'If Trim(Mid(cboREPCode1, 1, 5)) = "RC6" Then
                '    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\IMB.DOT\")
                '    Set objdoc = Nothing
                '    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetterCOR2.DOT\")
                'Else
                    'Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\IMB.DOT\")
                    Set objdoc = Nothing
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetter.DOT\")
                'End If
        ElseIf Not objdoc Is Nothing Then
              If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
                vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              ElseIf vbCancel Then
                Exit Sub
              End If
                'Set objdoc = Nothing
                'Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\IMB.DOT\")
                Set objdoc = Nothing
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "REPLetter.DOT")
        End If
        bookmark ("REPLetter.DOT")
        Me.Show

    ElseIf Letter = "P" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\PriorLetter.DOT\")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
              vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              objdoc.Save
              objdoc.Close
            ElseIf vbCancel Then
              Exit Sub
            End If
        End If
        
        If bookmark("PriorLetterNew.dot") = True Then
            Me.Show
        End If
        Exit Sub
        
    ElseIf Letter = "E" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\PriorExpLetter.dot\")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
              vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              objdoc.Save
              objdoc.Close
            ElseIf vbCancel Then
              Exit Sub
            End If
        End If
        
        If bookmark("PriorExpLetter.dot") = True Then
            Me.Show
        End If
        Exit Sub
    
    ElseIf Letter = "M" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "Case\CaseRemark.DOT\")
        End If
        
        If bookmark("CaseRemark.dot") = True Then
            Me.Show
        End If
        Exit Sub
    ElseIf Letter = "T" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\Take Home Drugs.doc\")
        End If
        
        
        If bookmark("Take Home Drugs.doc") = True Then
            Me.Show
        End If
        Exit Sub
        
    End If
    Exit Sub
    
openerror:
    Me.Show
    If Err.Number = "5356" Then
        MsgBox Err.Description & "Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If
    'End If
End Sub

Private Sub funcword1()
On Error GoTo openerror

        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\Guarantee Letter.dot\")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
                vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
            ElseIf vbCancel Then
                Exit Sub
            End If
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\Guarantee Letter.dot\")
        End If
        If bookmark1("Guarantee Letter.dot") = True Then
            Me.Show
        End If
        Exit Sub
    
openerror:
    Me.Show
    If Err.Number = "5356" Then
        MsgBox Err.Description & "Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Sub funcword2()
'Dim MBM As New ADODB.Recordset
'Dim strsqlMBM As String
On Error GoTo openerror
    
    'If Letter = "L" Then
        'strsqlMBM = "Select PLNCode, LGInd from PLN where PLNCode = '" & Trim(TxtPlanCode) & "'"
        'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    'End If
    
    ' make sure only one word application launched
    'LocCardPath = crdPath
    If Letter = "L" And tmpPLNPerVisit <> "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\New LG-2.DOT\")
                
         End If
     
        If bookmark("New LG-2.DOT") = True Then
           Me.Show
            If MsgBox("Do you want to print this document?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
                 objdoc.PrintOut
                 objdoc.Close
                 Exit Sub
            Else
                Exit Sub
            End If
        End If
    ElseIf Letter = "L" And tmpPLNPerVisit = "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\LGVisit.DOT\")
                
         End If
     
        If bookmark("LGVisit.DOT") = True Then
           Me.Show
            If MsgBox("Do you want to print this document?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
                 objdoc.PrintOut
                 objdoc.Close
                 Exit Sub
            Else
                Exit Sub
            End If
        End If
     
     'ElseIf bookmark("LG2805012.DOT") = True Then
         'Me.show
         'If MsgBox("Do you want to print this document?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
              'objdoc.PrintOut
              'objdoc.Close
              'Exit Sub
         'Else
             'Exit Sub
         'End If
     'ElseIf bookmark("LG2805013.DOT") = True Then
         'Me.show
         'If MsgBox("Do you want to print this document?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
              'objdoc.PrintOut
              'objdoc.Close
              'Exit Sub
         'Else
             'Exit Sub
         'End If
     'End If
    ElseIf Letter = "B" And tmpPLNPerVisit <> "Y" Then
        
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\BTBG.DOT\")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
                vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
                objdoc.Save
                objdoc.Close
            ElseIf vbCancel Then
                Exit Sub
            End If
        End If
            
        If bookmark("BTBG.DOT") = True Then
            Me.Show
        End If
        Exit Sub
        
    ElseIf Letter = "B" And tmpPLNPerVisit = "Y" Then
        
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\BTBGVisit.DOT\")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
                vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
                objdoc.Save
                objdoc.Close
            ElseIf vbCancel Then
                Exit Sub
            End If
        End If
            
        If bookmark("BTBGVisit.DOT") = True Then
            Me.Show
        End If
        Exit Sub
        
    ElseIf Letter = "R" Then
    Set objword = Nothing
    If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            If Trim(Mid(cboREPCode1, 1, 5)) = "RC6" Then
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\IMB.DOT\")
                Set objdoc = Nothing
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetterCOR1.DOT\")
            Else
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\IMB.DOT\")
                Set objdoc = Nothing
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetter1.DOT\")
            End If
    ElseIf Not objdoc Is Nothing Then
          If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
            vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
            objdoc.Save
            objdoc.Close
          ElseIf vbCancel Then
            Exit Sub
          End If
            Set objdoc = Nothing
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\IMB.DOT\")
            Set objdoc = Nothing
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "REPLetter.DOT")
    End If
    bookmark ("REPLetter.DOT")
        Me.Show
        
        Else
            Exit Sub
        End If
    
    Exit Sub
openerror:
    Me.Show
    If Err.Number = "5356" Then
        MsgBox Err.Description & ". Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If
    
End Sub

Private Sub showMBMListing()
    Dim MBM As New ADODB.Recordset
    Dim sqlstr As String
    Dim MBMListing As ListItem
    'sqlstr = " Select MBMNumber, HLTCode, MBMPolicyNo,  MBMName ," & _
    '        " MBMIcBcPp,  MBMMemberType , " & _
    '        " MBMStatus, PAYCode, AGECode , " & _
    '        " INSCode, PLNCode, " & _
    '        " MBMAddress1, MBMAddress2, " & _
    '        " MBMAddress3, MBMPostCode, CTYCode, STACode,  " & _
    '        " MBMDOB , BRNCode, MBMTakeOver ,  MBMAvailableLimit, " & _
    '        " MBMPayorEffDate, MBMPayorExpDate, MBMENDtEffDate  "
    'sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
        sqlstr = "Exec SearchMBMCas '" & Trim(txtMBMNumber) & "'"
        
        Set MBM = New ADODB.Recordset
        MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic

        lstMember.listitems.Clear
        'Do While Not MBM.EOF
        If Not MBM.EOF Then
            With MBM
                Set MBMListing = lstMember.listitems.Add(, , "00")
                If Len(Trim(!MBMName)) <> 0 Then
                    MBMListing.SubItems(1) = ChkStr(!MBMName)
                End If
                If Len(Trim(!MBMIcBcPp)) Then
                    MBMListing.SubItems(2) = !MBMIcBcPp
                End If
                If Len(Trim(!MBMMemberType)) Then
                    MBMListing.SubItems(3) = !MBMMemberType
                End If
                If Len(Trim(!MBMDOB)) Then
                    MBMListing.SubItems(4) = !MBMDOB
                End If
            End With
        End If
    MBM.Close
    Set MBM = Nothing
End Sub

' * Start

Private Sub showMBMCovListing()
    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String
                
    'sqlstr = "select MBMCNumber, MBMCCoverID, MBMCName," & _
    '    " MBMCICBCPPNo , MBMCDOB, MBMCSex, MBMCAdultChild, MBMCAgeband, " & _
    '    " MBMCRELCode,   " & _
    '    " MBMCStatus , MBMCAnnualLimit " & _
    '    " from  MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    'MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    sqlstr = "Exec SearchMBMCas '" & Trim(txtMBMNumber) & "'"
    
    Set MBMCov = New ADODB.Recordset
    MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic

    If Len(MBMCov!MBMCCoverID) Then
        Dim MBMListing As ListItem
        Do While Not MBMCov.EOF
            Set MBMListing = lstMember.listitems.Add(, , MBMCov!MBMCCoverID)
            With MBMCov
                If Len(Trim(!MBMCName)) <> 0 Then
                    MBMListing.SubItems(1) = ChkStr(!MBMCName)
                End If
                If Len(Trim(!MBMCICBCPPNo)) <> 0 Then
                    MBMListing.SubItems(2) = !MBMCICBCPPNo
                End If
                If Len(Trim(!MBMCRELCode)) Then
                    MBMListing.SubItems(3) = !MBMCRELCode
                End If
                
                If Len(Trim(!MBMCDOB)) Then
                    MBMListing.SubItems(4) = !MBMCDOB
                End If

            End With
            MBMCov.MoveNext
        Loop
    End If
    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub cboSuppExc_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboSuppExc, cboSuppExc.SelStart) + Chr(KeyAscii) + Mid(cboSuppExc, cboSuppExc.SelStart + cboSuppExc.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboSuppExc.ListCount - 1
        
        If NewText = LCase(Left(cboSuppExc.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboSuppExc.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboSuppExc = ValidValue
               cboSuppExc.SelStart = 0
               cboSuppExc.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub GetAccountSummary()
    Dim Acc As New ADODB.Recordset
    Dim strsql As String
    Dim Paid2MBM, PaidByMem, PayableMem As Double
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String

    'strsql = " select casnumber, totalApprove, totalPayable2Hos, " & _
            " totalPaid2Hos, totalPayable2Mbm, totalPaidByMbm , CLMFinalApprove, " & _
            " totalPaid2MBM , totalExcess , totalReimburse , CLMFinalApprove from view_account_summary where CASNumber = '" & Trim(txtCaseNo) & "'"
    'Acc.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    strAppendCase = "1"
    strAppend = " AND CASNumber = '" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        Acc.CursorLocation = adUseClient
        Acc.CursorType = adOpenForwardOnly
        Acc.CacheSize = 100
        Set Acc = cmd.Execute
        
    If Not Acc.EOF Then
        With Acc
            If IsNumeric(!totalPayable2Mbm) Then
                PayableMem = Format(!totalPayable2Mbm, "#,#.00;(#,#.00)")
            Else
                PayableMem = 0
            End If
            
            If IsNumeric(PayableMem) = False Then
                PayableMem = 0
                
            End If
            If Trim(!totalPaid2MBM) <> "" Then
                Paid2MBM = Format(!totalPaid2MBM, "#,#.00")
            End If
            If IsNumeric(Paid2MBM) = False Then
                Paid2MBM = 0
            End If
            If Trim(!totalPaidByMbm) <> "" Then
                PaidByMem = Format(!totalPaidByMbm, "#,#.00")
            End If
            If IsNumeric(PaidByMem) = False Then
                PaidByMem = 0
            End If
            txtRecoveryOS = Format(CDbl(PayableMem) - CDbl(Paid2MBM) + CDbl(PaidByMem), "#,#.00;(#,#.00)")
        End With
    End If
    Acc.Close
    Set Acc = Nothing

End Sub

Private Sub UpdateBalLmt()
Dim strsqlBalLmt As String
Dim MBM As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim BalLmt As Double
Dim MBMLimit

    BalLmt = 0
    
    If IsNumeric(txtREDAmt) = True Then
        'strsqlMBM = "select MBMAvailableLimit from MBM where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "15"
        strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        MBM.CursorLocation = adUseClient
        MBM.CursorType = adOpenForwardOnly
        MBM.CacheSize = 100
        Set MBM = cmd.Execute
        
            If MBM.EOF = False Then
                If IsNumeric(MBM!MBMAvailableLimit) = False Then
                    MBMLimit = 0
                End If
                
                If REDAmtInd = True Then
                    BalLmt = CDbl(MBM!MBMAvailableLimit) + CDbl(RedAmt)
                Else
                    BalLmt = CDbl(MBM!MBMAvailableLimit)
                End If
                
                If CDbl(BalLmt) > CDbl(txtREDAmt) Then
                    MBMLimit = CDbl(BalLmt) - CDbl(txtREDAmt)
                Else
                    MBMLimit = 0
                End If
                'MBM.Update
            End If
            strsqlBalLmt = "EXEC Case_AdjUpdateBalLmt '" & Trim(txtMBMNumber) & "', " & MBMLimit & ""
            medixmasterconn.Execute strsqlBalLmt
        MBM.Close
        Set MBM = Nothing
    End If
End Sub

Private Sub UpdateLG()
Dim strsqlRMK As String
Dim LG As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim LGRemark
        
    'strsql = "Select CASENo, LGNo, LGRemark from CASLG where CASENo = '" & Trim(LGNo) & "' AND LGNo = '" & Trim(LGVer) & "'"
    'LG.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    strAppendCase = "4"
    strAppend = " AND CASENo = '" & Trim(LGNo) & "' AND LGNo = '" & Trim(LGVer) & "'"
    Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        LG.CursorLocation = adUseClient
        LG.CursorType = adOpenForwardOnly
        LG.CacheSize = 100
        Set LG = cmd.Execute
        
        If LG.EOF = False Then
            
            If TxtLGRemark <> "" Then
                LGRemark = Replace(TxtLGRemark, "'", "''")
                strsqlRMK = "EXEC Case_AdjUpdateLG '" & Trim(LGNo) & "', '" & Trim(LGVer) & "', '" & Trim(LGRemark) & "'"
                'LG!LGRemark = Trim(ChkStr(TxtLGRemark))
            Else
                LGRemark = Empty
                strsqlRMK = "EXEC Case_AdjUpdateLG '" & Trim(LGNo) & "', '" & Trim(LGVer) & "', '" & LGRemark & "'"
            End If
            medixmasterconnMaint.Execute strsqlRMK
        End If
    LG.Close
    Set LG = Nothing
    
End Sub

Private Sub ExportFeeDetails()
On Error Resume Next
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.Worksheet

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

    Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Case\Fee Details.xlt")
    Set objWorksheet = objWorkbook.Sheets(1)
    'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    objWorksheet.Cells(5, "B") = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
    objWorksheet.Cells(6, "B") = Trim(txtMBMNumber)
    
    Screen.MousePointer = vbDefault
    'objWorksheet.Columns.AutoFit
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Exit Sub
ExitFunction:
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Screen.MousePointer = vbDefault
    objExcel.Quit
    Exit Sub
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    GoTo ExitFunction
End Sub

Private Sub UpdateClaimability()
Dim strsql As String
Dim ACPeriod As New ADODB.Recordset
    
    cboClaimability.Text = "NN - Not Decided"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    strsql = " update CAS set " & _
             " CASClaimability = '" & Mid(cboClaimability, 1, 2) & "' " & _
             " Where CASNumber = '" & Trim(txtCaseNo) & "'"
    
    medixmasterconn.Execute strsql
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing

End Sub

Private Sub UpdateAuditReview()
Dim RecordSaved
Dim cmd As New ADODB.Command
Dim rst2 As New ADODB.Recordset
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strsqlCASAuditReview, strsqlCASRef As String
Dim InvDate, CASLetterSent, CASHospReply, CASDOV, CASDOVTime As String
Dim CASDTC, CASAROpenDate, CASInvAmt, CASDisRec As String
Dim AuditReview, answer, Remarks, HospVisit
Dim TempDate As String
    
    cmdUpdate1.Enabled = False
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No!", vbCritical + vbOKOnly, DialogBoxTitle
        txtCaseNo.SetFocus
    Else
        'RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        'If RecordSaved = vbYes Then
            
            strAppend = " AND CASNumber = '" & txtCaseNo & "'"
            strAppendCase = "19"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_AuditReview"
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
               If txtINVDate <> "__/__/____" Then
                   InvDate = Format(txtINVDate, "mm/dd/yyyy")
               Else
                   InvDate = Empty
               End If
               
               If Len(txtINVAmt) Then
                   CASInvAmt = txtINVAmt
               Else
                   CASInvAmt = 0
               End If
               
               If txtDLS <> "__/__/____" Then
                   CASLetterSent = Format(txtDLS, "mm/dd/yyyy")
               Else
                   CASLetterSent = Empty
               End If
               
               If txtDLR <> "__/__/____" Then
                   CASHospReply = Format(txtDLR, "mm/dd/yyyy")
               Else
                   CASHospReply = Empty
               End If
             
               If txtDV <> "__/__/____" Then
                    CASDOV = Format(txtDV, "mm/dd/yyyy")
               Else
                    CASDOV = Null
               End If
                
               If txtDOVTime <> "__:__" Then
                   CASDOVTime = Format(txtDOVTime, "hh:mm")
               End If
             
               If Len(txtDISRec) Then
                   CASDisRec = ChkStr(txtDISRec)
               Else
                   CASDisRec = 0
               End If
               
               If TxtDRTC <> "__/__/____" Then
                    CASDTC = Format(TxtDRTC, "mm/dd/yyyy")
               Else
                    CASDTC = Empty
               End If
               
               If TxtDateOpen <> "__/__/____" Then
                    CASAROpenDate = Format(TxtDateOpen, "mm/dd/yyyy")
               Else
                   CASAROpenDate = Empty
               End If
               
            TempDate = Format(Date, "yyyy/mm/dd")
            answer = ChkStr(Replace(TxtARAnswer, "'", "''"))
            Remarks = ChkStr(Replace(txtARRemarks, "'", "''"))
            HospVisit = ChkStr(Replace(txtHospVisit, "'", "''"))
            AuditReview = ChkStr(Replace(txtAuditReview, "'", "''"))
            
            strsqlCASAuditReview = "'" & ChkStr(Trim(TempCASNo)) & "', '" & ChkStr(txtINVNo) & "', " & _
                            "'" & InvDate & "', " & CASInvAmt & ", '" & CASLetterSent & "', '" & CASHospReply & "', " & _
                            "'" & CASDOV & "', '" & CASDOVTime & "', '" & ChkStr(txtOfficer) & "', '" & HospVisit & "', " & _
                            "" & CASDisRec & ", '" & CASDTC & "', '" & CASAROpenDate & "', " & _
                            "'" & answer & "', '" & Remarks & "', '" & AuditReview & "', '" & Mid(CboARStatus, 1, 1) & "', " & _
                            "'" & Logon & "', '" & TempDate & "'"
            
            If rst2.EOF Then
                strsqlCASAuditReview = "EXEC Case_AuditInsert " + strsqlCASAuditReview
            Else
                strsqlCASAuditReview = "EXEC Case_AuditUpdate " + strsqlCASAuditReview
            End If
            rst2.Close
            Set rst2 = Nothing
            
            medixmasterconn.Execute strsqlCASAuditReview
            
            
            strsqlCASRef = "EXEC Case_AuditUpdate '" & Trim(TempCASNo) & "', '" & Trim(MBMNumber) & "', '" & Mid(CboARStatus, 1, 1) & "'"
            medixmasterconnMaint.Execute strsqlCASRef
            
            'MsgBox "Record Updated Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
        End If
    'End If
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    cmdUpdate1.Enabled = True

End Sub

Private Function showAuditReview()
Dim sqlstr As String
Dim Other As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    'sqlstr = "Select * From CASOthers where CASNumber = '" & RemStr(txtCaseNo) & "'"
    'Other.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "10"
    strAppend = " AND CASNumber = '" & RemStr(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    Other.CursorLocation = adUseClient
    Other.CursorType = adOpenForwardOnly
    Other.CacheSize = 100
    Set Other = cmd.Execute
    
    If Not Other.EOF Then
        Do Until Other.EOF
        With Other
            If Len(!CASDOV) Then
                 txtDV = Format(!CASDOV, "dd/mm/yyyy")
            End If
            
            If Len(!CASDOVTime) Then
                 txtDOVTime = Format(!CASDOVTime, "hh:mm")
            End If
            
            If Len(!CASVRemarks) Then
                txtHospVisit = !CASVRemarks
            End If
            
            If Len(!CASARAnswer) Then
                TxtARAnswer = !CASARAnswer
            End If
            
            If Len(!CASDTC) Then
                 TxtDRTC = Format(!CASDTC, "dd/mm/yyyy")
            End If
            
            If Len(!CASAROpenDate) Then
                 TxtDateOpen = Format(!CASAROpenDate, "dd/mm/yyyy")
            End If
            
            If Len(!CASARRemarks) Then
                 txtARRemarks = !CASARRemarks
            End If
            
            If Len(!CASReviewAudit) Then
                txtAuditReview = !CASReviewAudit
            End If
            
            If Len(!CASInvNo) Then
                txtINVNo = !CASInvNo
            End If
            
            If Len(!InvDate) Then
                txtINVDate.mask = ""
                txtINVDate = Format(!InvDate, "dd/mm/yyyy")
                txtINVDate.mask = "##/##/####"
            End If
            
            If Len(!CASInvAmt) Then
                txtINVAmt = !CASInvAmt
            End If
                        
            If Len(!Officer) Then
                txtOfficer = !Officer
            End If
                        
            If Len(!CASLetterSent) Then
                txtDLS.mask = ""
                txtDLS = Format(!CASLetterSent, "dd/mm/yyyy")
                txtDLS.mask = "##/##/####"
            End If

            If Len(!CASHospReply) Then
                txtDLR.mask = ""
                txtDLR = Format(!CASHospReply, "dd/mm/yyyy")
                txtDLR.mask = "##/##/####"
            End If

            If Len(!CASDisRec) Then
                txtDISRec = !CASDisRec
            End If
            
            If Len(!CASARStatus) Then
                CboARStatus = !CASARStatus
            End If
        End With
        Other.MoveNext
        Loop
    End If
    Other.Close
    Set Other = Nothing
End Function

Private Function SearchMMARate(OPCSCode As String)
Dim strsqlMMA As String
Dim MAA As New ADODB.Recordset

    strsqlMMA = "Select SurgeonsFees, AnaesthetistsFees  from VIEW_MMA_Schedule where  OPCSCode = '" & Trim(OPCSCode) & "'"
    MAA.Open strsqlMMA, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        If MAA.recordCount Then
            MMASurgFees = IIf(Len(MAA!SurgeonsFees), MAA!SurgeonsFees, 0)
            MMAAnaeFees = IIf(Len(MAA!AnaesthetistsFees), MAA!AnaesthetistsFees, 0)
        Else
            MMASurgFees = 0
            MMAAnaeFees = 0
        End If
    MAA.Close
    Set MAA = Nothing
End Function

Private Sub SearchMMARateCal(CPTCode As String)
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
    
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    strsqlMMA = "Select SurgeonsFees, AnaesthetistsFees  from VIEW_MMA_Schedule where  OPCSCode = '" & Trim(CPTCode) & "'"
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        If MMA.recordCount Then
            MMASurgFees = IIf(Len(MMA!SurgeonsFees), MMA!SurgeonsFees, 0)
            MMAAnaeFees = IIf(Len(MMA!AnaesthetistsFees), MMA!AnaesthetistsFees, 0)
        Else
            MMASurgFees = 0
            MMAAnaeFees = 0
        End If
    MMA.Close
    Set MMA = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing

End Sub

Private Sub WSListing()
Dim ws As New ADODB.Recordset
Dim WSList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    If lstWorksheetListing.listitems.Count = 0 Then
            'sql = " Select CASNumber, ClaimVersion, CLMTotalApproved, CLMTotalActual,CLMPending, " & _
            "CLMPayableByMedix2H,CLMPayableByMedix2M, BordxRefNo, ClaimCloseDate, BordxDate " & _
            " From VIEW_WS_Listing " & _
            " where CASNumber ='" & Trim(txtCaseNo) & "'"
                
        'ws.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        strAppendCase = "2"
        strAppend = " AND CASNumber ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        ws.CursorLocation = adUseClient
        ws.CursorType = adOpenForwardOnly
        ws.CacheSize = 100
        Set ws = cmd.Execute
        
        If Not ws.EOF Then
        Do Until ws.EOF
            With ws
                Set WSList = lstWorksheetListing.listitems.Add(, , !CasNumber & "-" & !claimversion)
                    If Len(!CLMTotalActual) Then
                        WSList.SubItems(1) = Format(!CLMTotalActual, "#,##0.00")
                    End If
                    If Len(!CLMTotalApproved) Then
                        WSList.SubItems(2) = Format(!CLMTotalApproved, "#,##0.00")
                    End If
                    If Len(!CLMPayableByMedix2H) Then
                        WSList.SubItems(3) = Format(!CLMPayableByMedix2H, "#,##0.00")
                    End If
                    If Len(!CLMPayableByMedix2M) Then
                        WSList.SubItems(4) = Format(!CLMPayableByMedix2M, "#,##0.00")
                    End If
                    If Len(!BordxRefNo) Then
                        WSList.SubItems(5) = !BordxRefNo
                    End If
                    If Len(!BordxDate) Then
                        WSList.SubItems(6) = !BordxDate
                    End If
                    If Len(!CLMBordxAmt) Then
                        WSList.SubItems(7) = Format(!CLMBordxAmt, "#,##0.00")
                    End If
                    
                    If Len(!CLMPayorAppStatus) Then
                        WSList.SubItems(8) = !CLMPayorAppStatus
                    End If
                    If Len(!CLMPayorAppDate) Then
                        WSList.SubItems(9) = !CLMPayorAppDate
                    End If
                    
                .MoveNext
                End With
        Loop
       
        ws.Close
        Set ws = Nothing
    End If
End If
End Sub




