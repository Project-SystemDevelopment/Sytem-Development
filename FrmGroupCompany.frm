VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGroupCompany 
   Caption         =   "Group Company Maintainance"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   7920
      Width           =   11715
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
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
         Left            =   2640
         TabIndex        =   10
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   960
         TabIndex        =   8
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Left            =   1800
         TabIndex        =   7
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
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
         Left            =   9720
         TabIndex        =   6
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
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
         Height          =   315
         Left            =   10560
         TabIndex        =   5
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
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
         Left            =   3480
         TabIndex        =   4
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
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
         Left            =   8880
         TabIndex        =   3
         Top             =   180
         Width           =   840
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
         Height          =   315
         Left            =   4320
         TabIndex        =   2
         Top             =   180
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7275
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   12832
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Group Comapny"
      TabPicture(0)   =   "FrmGroupCompany.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label46"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstCompanyGrpList"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Subsidiry"
      TabPicture(1)   =   "FrmGroupCompany.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(1)=   "LstviewSubsidary"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Department"
      TabPicture(2)   =   "FrmGroupCompany.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstviewDepartment"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Cost Centre"
      TabPicture(3)   =   "FrmGroupCompany.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstViewCostcentre"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   94
         Top             =   600
         Width           =   11295
         Begin VB.ComboBox cmbDeptState 
            Height          =   315
            Left            =   6960
            Sorted          =   -1  'True
            TabIndex        =   151
            Top             =   1005
            Width           =   3975
         End
         Begin VB.TextBox txtDeptEmail 
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
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   120
            Top             =   1620
            Width           =   3975
         End
         Begin VB.TextBox txtDeptPIC 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   118
            Top             =   1560
            Width           =   3975
         End
         Begin VB.TextBox txtDeptPhone 
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
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   116
            Top             =   1320
            Width           =   3975
         End
         Begin VB.TextBox txtDeptID 
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
            Left            =   6960
            MaxLength       =   20
            TabIndex        =   114
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox txtdeptPostcode 
            Height          =   285
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   112
            Top             =   750
            Width           =   3975
         End
         Begin VB.TextBox txtDeptCity 
            Height          =   285
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   110
            Top             =   495
            Width           =   3975
         End
         Begin VB.ComboBox cmbDepGrpcompany 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   107
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox txtDeptAdd3 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   100
            Top             =   1260
            Width           =   3975
         End
         Begin VB.TextBox txtDeptadd2 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   99
            Top             =   1005
            Width           =   3975
         End
         Begin VB.TextBox txtDeptAdd1 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   98
            Top             =   750
            Width           =   3975
         End
         Begin VB.TextBox txtDeptName 
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
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   97
            Top             =   495
            Width           =   3975
         End
         Begin VB.ComboBox cmbDeptClaim 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   96
            Top             =   2160
            Width           =   3975
         End
         Begin VB.ComboBox cmbDeptMCO 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   95
            Top             =   1800
            Width           =   3975
         End
         Begin VB.Label Label27 
            Caption         =   "Email ID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   121
            Top             =   1620
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "PIC"
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
            Index           =   22
            Left            =   120
            TabIndex        =   119
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label24 
            Caption         =   "Phone"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   117
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label23 
            Caption         =   "Department ID"
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
            Left            =   5760
            TabIndex        =   115
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Post Code"
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
            Index           =   14
            Left            =   5760
            TabIndex        =   113
            Top             =   750
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "City"
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
            Index           =   13
            Left            =   5760
            TabIndex        =   111
            Top             =   495
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Address3"
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
            Index           =   12
            Left            =   120
            TabIndex        =   109
            Top             =   1260
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Address2"
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
            Index           =   11
            Left            =   120
            TabIndex        =   108
            Top             =   1005
            Width           =   1695
         End
         Begin VB.Label Label22 
            Caption         =   "Department Name  *"
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
            TabIndex        =   106
            Top             =   495
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Group Company *"
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
            TabIndex        =   105
            Top             =   255
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Address1   *"
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
            Index           =   9
            Left            =   120
            TabIndex        =   104
            Top             =   750
            Width           =   1695
         End
         Begin VB.Label Label21 
            Caption         =   "State"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   103
            Top             =   1005
            Width           =   1695
         End
         Begin VB.Label Label20 
            Caption         =   "Claim Billing Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   102
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "MCO Billing Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   101
            Top             =   1800
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView lstCompanyGrpList 
         Height          =   2745
         Left            =   120
         TabIndex        =   72
         Top             =   3960
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4842
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
         NumItems        =   21
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Group Company code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Group Company Name"
            Object.Width           =   14111
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Payor Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "BusinessChannel"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Address1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Address2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Address3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "TelNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "FaxNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "EffDt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Conversion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ContactPerson"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "CEOName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Claim Billig Mode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Mco Billing Mode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Grp Com Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Claims PIC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "MCO PIC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "TIMS Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "R&B Tax"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Pre & Post"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame11 
         Height          =   3375
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   11175
         Begin VB.TextBox txtClaimsPic 
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
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   76
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox txtMCOPIC 
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
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   73
            Top             =   3000
            Width           =   3855
         End
         Begin VB.ComboBox cmbMcoBilling 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   71
            Top             =   3000
            Width           =   3975
         End
         Begin VB.ComboBox cmbclaimBill 
            Height          =   315
            ItemData        =   "FrmGroupCompany.frx":0070
            Left            =   1920
            List            =   "FrmGroupCompany.frx":0072
            Sorted          =   -1  'True
            TabIndex        =   69
            Top             =   2640
            Width           =   3975
         End
         Begin VB.ComboBox CboGrpComStatus 
            Height          =   315
            ItemData        =   "FrmGroupCompany.frx":0074
            Left            =   7440
            List            =   "FrmGroupCompany.frx":0076
            Sorted          =   -1  'True
            TabIndex        =   48
            Top             =   2160
            Width           =   1335
         End
         Begin VB.ComboBox cboCoTax 
            Height          =   315
            Left            =   7440
            Sorted          =   -1  'True
            TabIndex        =   47
            Top             =   1815
            Width           =   1335
         End
         Begin VB.ComboBox cboCoPrePost 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   45
            Top             =   1845
            Width           =   1150
         End
         Begin VB.ComboBox cboCoConversion 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   44
            Top             =   1560
            Width           =   1150
         End
         Begin VB.TextBox txtCoHomepage 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   43
            Top             =   1305
            Width           =   3500
         End
         Begin VB.TextBox txtCoEmail 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   42
            Top             =   930
            Width           =   3500
         End
         Begin VB.TextBox txtCoFax 
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
            Left            =   9600
            MaxLength       =   15
            TabIndex        =   41
            Top             =   555
            Width           =   1330
         End
         Begin VB.TextBox txtCoPhone 
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
            Left            =   7440
            MaxLength       =   15
            TabIndex        =   40
            Top             =   555
            Width           =   1275
         End
         Begin VB.TextBox txtCoCEO 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   39
            Top             =   2280
            Width           =   3975
         End
         Begin VB.TextBox txtCoAdmContact 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   38
            Top             =   2025
            Width           =   3975
         End
         Begin VB.TextBox txtCoContact 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   37
            Top             =   1770
            Width           =   3975
         End
         Begin VB.TextBox txtCoAdd3 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   36
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox txtCoAdd2 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   35
            Top             =   1200
            Width           =   3975
         End
         Begin VB.TextBox txtGrpCoCode 
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
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   34
            Top             =   255
            Width           =   930
         End
         Begin VB.TextBox txtCompanyName 
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
            Left            =   5160
            MaxLength       =   250
            TabIndex        =   33
            Top             =   255
            Width           =   5775
         End
         Begin VB.TextBox txtCoAdd1 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   32
            Top             =   960
            Width           =   3975
         End
         Begin VB.ComboBox CboCoGrpPayor 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   31
            Top             =   600
            Width           =   3975
         End
         Begin VB.ComboBox cmbChannel 
            Height          =   315
            ItemData        =   "FrmGroupCompany.frx":0078
            Left            =   9720
            List            =   "FrmGroupCompany.frx":007A
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   2640
            Width           =   1155
         End
         Begin VB.ComboBox cmbTims 
            Height          =   315
            ItemData        =   "FrmGroupCompany.frx":007C
            Left            =   9795
            List            =   "FrmGroupCompany.frx":0086
            Sorted          =   -1  'True
            TabIndex        =   29
            Top             =   2160
            Width           =   1150
         End
         Begin MSMask.MaskEdBox txtCoEffDate 
            Height          =   285
            Left            =   7440
            TabIndex        =   46
            Top             =   1560
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label16 
            Caption         =   "Claims PIC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6120
            TabIndex        =   75
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "MCO PIC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6120
            TabIndex        =   74
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Mco Billing Mode *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   70
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Claim Billing Mode *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   68
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label72 
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   66
            Top             =   1770
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Payor Code           *"
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
            Index           =   9
            Left            =   120
            TabIndex        =   65
            Top             =   555
            Width           =   1695
         End
         Begin VB.Label Label73 
            Caption         =   "Homepage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6120
            TabIndex        =   64
            Top             =   1320
            Width           =   870
         End
         Begin VB.Label Label74 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6120
            TabIndex        =   63
            Top             =   945
            Width           =   495
         End
         Begin VB.Label Label75 
            Caption         =   "Tel No"
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
            Left            =   6120
            TabIndex        =   62
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label76 
            Caption         =   "Fax No"
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
            Left            =   8880
            TabIndex        =   61
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Group Company Name *"
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
            Left            =   3120
            TabIndex        =   60
            Top             =   285
            Width           =   2055
         End
         Begin VB.Label Label2 
            Caption         =   "Company Address    *"
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
            Index           =   10
            Left            =   120
            TabIndex        =   59
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Grp. Company Code *"
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
            Index           =   15
            Left            =   120
            TabIndex        =   58
            Top             =   255
            Width           =   1815
         End
         Begin VB.Label Label77 
            Caption         =   "CEO Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   57
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Conversion"
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
            Left            =   8880
            TabIndex        =   56
            Top             =   1590
            Width           =   930
         End
         Begin VB.Label Label2 
            Caption         =   "Eff. Date  *"
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
            Left            =   6120
            TabIndex        =   55
            Top             =   1575
            Width           =   1050
         End
         Begin VB.Label Label78 
            Caption         =   "Adm. Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   54
            Top             =   2025
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "R&&B Tax"
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
            Index           =   0
            Left            =   6120
            TabIndex        =   53
            Top             =   1860
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "Pre&&Post"
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
            Index           =   20
            Left            =   8880
            TabIndex        =   52
            Top             =   1860
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "Grp Com Status *"
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
            Index           =   21
            Left            =   6120
            TabIndex        =   51
            Top             =   2220
            Width           =   1290
         End
         Begin VB.Label Label2 
            Caption         =   "Business Type *"
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
            Index           =   23
            Left            =   8520
            TabIndex        =   50
            Top             =   2640
            Width           =   1170
         End
         Begin VB.Label Label2 
            Caption         =   "TIMS Client"
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
            Index           =   24
            Left            =   8880
            TabIndex        =   49
            Top             =   2175
            Width           =   810
         End
      End
      Begin MSComctlLib.ListView lstViewCostcentre 
         Height          =   3435
         Left            =   -74760
         TabIndex        =   22
         Top             =   3600
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6059
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
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Costcenter ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Costcenter Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Address 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Address 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Address 3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "City"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Post Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "State"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Phone"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "PIC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Email ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Claim Billing Mode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "MCO Billing Mode"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame3 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   21
         Top             =   480
         Width           =   11055
         Begin VB.ComboBox cmbCostState 
            Height          =   315
            Left            =   6960
            Sorted          =   -1  'True
            TabIndex        =   150
            Top             =   1125
            Width           =   3975
         End
         Begin VB.TextBox txtCostcenEmail 
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
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   148
            Top             =   1740
            Width           =   3975
         End
         Begin VB.TextBox txtCostcenPIC 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   146
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox txtCostcenPhone 
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
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   144
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox txtCostcenID 
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
            Left            =   6960
            MaxLength       =   20
            TabIndex        =   142
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox txtCostcenPostcode 
            Height          =   285
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   140
            Top             =   870
            Width           =   3975
         End
         Begin VB.TextBox txtCostcenCity 
            Height          =   285
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   138
            Top             =   615
            Width           =   3975
         End
         Begin VB.ComboBox cmbCostGrpCompany 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   135
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox txtCostcenAdd3 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   128
            Top             =   1380
            Width           =   3975
         End
         Begin VB.TextBox txtCostcenAdd2 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   127
            Top             =   1125
            Width           =   3975
         End
         Begin VB.TextBox txtCostcenAdd1 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   126
            Top             =   870
            Width           =   3975
         End
         Begin VB.TextBox txtCostcentName 
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
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   125
            Top             =   615
            Width           =   3975
         End
         Begin VB.ComboBox cmbClaimcostcenter 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   124
            Top             =   2280
            Width           =   3975
         End
         Begin VB.ComboBox cmbMCOCostcen 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   123
            Top             =   1920
            Width           =   3975
         End
         Begin VB.Label Label30 
            Caption         =   "Email ID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   149
            Top             =   1740
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "PIC"
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
            Index           =   25
            Left            =   120
            TabIndex        =   147
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label29 
            Caption         =   "Phone"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   145
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "Costcentre ID"
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
            Left            =   5760
            TabIndex        =   143
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Post Code"
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
            Index           =   19
            Left            =   5760
            TabIndex        =   141
            Top             =   870
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "City"
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
            Index           =   18
            Left            =   5760
            TabIndex        =   139
            Top             =   615
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Address3"
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
            Index           =   17
            Left            =   120
            TabIndex        =   137
            Top             =   1380
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Address2"
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
            Index           =   16
            Left            =   120
            TabIndex        =   136
            Top             =   1125
            Width           =   1695
         End
         Begin VB.Label Label25 
            Caption         =   "Costcentre Name  *"
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
            TabIndex        =   134
            Top             =   615
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Group Company *"
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
            TabIndex        =   133
            Top             =   375
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Address1   *"
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
            Index           =   15
            Left            =   120
            TabIndex        =   132
            Top             =   870
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "State"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   131
            Top             =   1125
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Claim Billing Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   130
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "MCO Billing Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   129
            Top             =   1920
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView lstviewDepartment 
         Height          =   3435
         Left            =   -74880
         TabIndex        =   20
         Top             =   3480
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6059
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
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Dept ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Department Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Address 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Address 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Address 3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "City"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Post Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "State"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Phone"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "PIC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Email ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Claim Billing Mode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "MCO Billing Mode"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   11055
         Begin VB.ComboBox cmbSubstate 
            Height          =   315
            Left            =   6960
            Sorted          =   -1  'True
            TabIndex        =   152
            Top             =   1200
            Width           =   3975
         End
         Begin VB.TextBox txtsubEmail 
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
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   90
            Top             =   1860
            Width           =   3975
         End
         Begin VB.TextBox txtSubPIC 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   88
            Top             =   1560
            Width           =   3975
         End
         Begin VB.TextBox txtsubPhone 
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
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   86
            Top             =   1560
            Width           =   3975
         End
         Begin VB.TextBox txtSubID 
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
            Left            =   6960
            MaxLength       =   20
            TabIndex        =   84
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox txtSubPostcode 
            Height          =   285
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   82
            Top             =   870
            Width           =   3975
         End
         Begin VB.TextBox txtsubCity 
            Height          =   285
            Left            =   6960
            MaxLength       =   50
            TabIndex        =   80
            Top             =   615
            Width           =   3975
         End
         Begin VB.ComboBox cmbSubGrpcompany 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   77
            Top             =   240
            Width           =   3975
         End
         Begin VB.ComboBox cmbSubMCO 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   1800
            Width           =   3975
         End
         Begin VB.ComboBox cmbsubClaims 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   2160
            Width           =   3975
         End
         Begin VB.TextBox txtSubName 
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
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   15
            Top             =   500
            Width           =   3975
         End
         Begin VB.TextBox txtSubadd1 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   14
            Top             =   755
            Width           =   3975
         End
         Begin VB.TextBox txtSubadd2 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1010
            Width           =   3975
         End
         Begin VB.TextBox txtSubAdd3 
            Height          =   285
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   12
            Top             =   1265
            Width           =   3975
         End
         Begin VB.Label Label19 
            Caption         =   "Email ID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   91
            Top             =   1860
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "PIC"
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
            Index           =   8
            Left            =   120
            TabIndex        =   89
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label18 
            Caption         =   "Phone"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   87
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "Subsidary ID"
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
            Left            =   5760
            TabIndex        =   85
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Post Code"
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
            Index           =   6
            Left            =   5760
            TabIndex        =   83
            Top             =   870
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "City"
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
            Index           =   5
            Left            =   5760
            TabIndex        =   81
            Top             =   615
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Address3"
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
            Index           =   4
            Left            =   120
            TabIndex        =   79
            Top             =   1265
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Address2"
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
            Index           =   3
            Left            =   120
            TabIndex        =   78
            Top             =   1010
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "MCO Billing Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   25
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Claim Billing Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   24
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label28 
            Caption         =   "State"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5760
            TabIndex        =   19
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Address1   *"
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
            Index           =   7
            Left            =   120
            TabIndex        =   18
            Top             =   755
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Group Company *"
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
            TabIndex        =   17
            Top             =   255
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "Subsidary Name  *"
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
            TabIndex        =   16
            Top             =   500
            Width           =   1815
         End
      End
      Begin MSComctlLib.ListView LstviewSubsidary 
         Height          =   3345
         Left            =   -74760
         TabIndex        =   93
         Top             =   3600
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5900
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
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Subsidary ID"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "GroupCompany"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Subsidary Name"
            Object.Width           =   14111
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Address1"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Address2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Address3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "City"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Post Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "State"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Phone Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "PIC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Email ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Claim Billing Mode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "MCO Billing Mode"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "SUBSIDARY LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -74760
         TabIndex        =   92
         Top             =   3240
         Width           =   1875
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "COMPANY GROUP LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   67
         Top             =   3720
         Width           =   1875
      End
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable        * - Mandatory Fields"
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
      TabIndex        =   122
      Top             =   7560
      Width           =   4095
   End
   Begin VB.Label Label14 
      Caption         =   "CEO Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   23
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "FrmGroupCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Public GRPCoCriteria As String
Private Sub CmdAdd_Click()
EditFlag = False
Call ClearForm(Me)
Call DisableCmdButton(Me)
Call EntriesEnable(True)
cmdSave.Enabled = True
cmdAdd.Enabled = True
cmdEdit.Enabled = True
    Select Case SSTab1.Tab
        Case 0
            txtGrpCoCode.Enabled = False
            txtCompanyName.SetFocus
        Case 1
            txtSubID.Enabled = False
        Case 2
            txtDeptID.Enabled = False
        Case 3
            txtCostcenID.Enabled = False
    End Select
End Sub
Private Sub EntriesEnable(status As Boolean)
    txtGrpCoCode.Enabled = status
    txtCompanyName.Enabled = status
    CboCoGrpPayor.Enabled = status
    txtCoAdd1.Enabled = status
    txtCoAdd2.Enabled = status
    txtCoAdd3.Enabled = status
    txtCoContact.Enabled = status
    txtCoAdmContact.Enabled = status
    txtCoCEO.Enabled = status
    txtCoPhone.Enabled = status
    txtCoFax.Enabled = status
    txtCoEmail.Enabled = status
    txtCoEffDate.Enabled = status
    cboCoTax.Enabled = status
    cboCoConversion.Enabled = status
    cboCoPrePost.Enabled = status
    CboGrpComStatus.Enabled = status
End Sub
'************************End EntriesEnable*********************************

Private Sub cmdcancel_Click()
 Unload Me
End Sub

Private Sub CmdClear_Click()
  Call ClearForm(Me)
    lstCompanyGrpList.ListItems.Clear
    'txtTotalRecord = 0
    cmdEdit.Enabled = False
End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass
    If SSTab1.Tab = 0 Then
        Call DeleteCompanyGrp
    End If
cmdAdd.Enabled = True
Screen.MousePointer = Default
End Sub
Private Sub DeleteCompanyGrp()
Dim DeleteCoGrp As New ADODB.Recordset
Dim rsChkCo As New ADODB.Recordset
Dim StrSql, sqlCompany As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

EditFlag = False
If txtGrpCoCode = "" Then
   MsgBox "Please Enter Company Group Code", vbOKOnly + vbCritical, DialogBoxTitle
   txtGrpCoCode.SetFocus
ElseIf lstCompanyGrpList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCompany = "Select CoGRPCode from CompanyGRP where CoGRPCode = '" & RemStr(txtGrpCoCode) & "'"
   
   rsChkCo.MaxRecords = 1
   rsChkCo.Open sqlCompany, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   If rsChkCo.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Company Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsChkCo.recordCount = 0 Then
           StrSql = "Delete From CompanyGRP Where CoGRPCode = '" & ChkStr(lstCompanyGrpList.SelectedItem.Text) & "'"
           DeleteCoGrp.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call CompanyGRPListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtGrpCoCode.SetFocus
         End If
     End If
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
Exit Sub

End Sub

Private Sub CmdEdit_Click()
 Select Case SSTab1.Tab
        Case 0 'Group company
            Call EnableGrpCompany
        Case 1  'subsidary
             EditFlag = True
        Case 2  'Department
             EditFlag = True
        Case 3    'Cost Center
             EditFlag = True
    End Select
End Sub

Public Sub EnableGrpCompany()
    txtCoAdd1.Enabled = True
    txtCoAdd2.Enabled = True
    txtCoAdd3.Enabled = True
    txtCoContact.Enabled = True
    txtCoAdmContact.Enabled = True
    txtCoCEO.Enabled = True
    cmbclaimBill.Enabled = True
    cmbMcoBilling.Enabled = True
    txtCoPhone.Enabled = True
    txtCoFax.Enabled = True
    txtCoEmail.Enabled = True
    txtCoHomepage.Enabled = True
    txtCoEffDate.Enabled = True
    cboCoConversion.Enabled = True
    cboCoTax.Enabled = True
    cboCoPrePost.Enabled = True
    CboGrpComStatus.Enabled = True
    cmbTims.Enabled = True
    cmbChannel.Enabled = True
    EditFlag = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdPrint_Click()
EditFlag = False
    Call ClearForm(Me)
    Call EntriesEnable(True)
    Call PRVChargesComboListing
    Select Case SSTab1.Tab
        Case 0
             If lstCompanyGrpList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstCompanyGrpList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
    End Select
End Sub
Private Sub PRVChargesComboListing()
Dim SRV As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim HLT As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchSRV"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    SRV.CursorLocation = adUseClient
    SRV.CursorType = adOpenForwardOnly
    SRV.CacheSize = 100
    
    Set cmd2.ActiveConnection = medixmasterconn
    cmd2.CommandText = "SearchHLTType"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 15
    HLT.CursorLocation = adUseClient
    HLT.CursorType = adOpenForwardOnly
    HLT.CacheSize = 100
    
    Set cmd3.ActiveConnection = medixmasterconn
    cmd3.CommandText = "SearchINS"
    cmd3.CommandType = adCmdStoredProc
    cmd3.CommandTimeout = 15
    INS.CursorLocation = adUseClient
    INS.CursorType = adOpenForwardOnly
    INS.CacheSize = 100
    
    Set SRV = cmd.Execute
    Set HLT = cmd2.Execute
    Set INS = cmd3.Execute

    SRV.Close
    Set SRV = Nothing
    INS.Close
    Set INS = Nothing
    HLT.Close
    Set HLT = Nothing
End Sub

Private Sub CmdRefresh_Click()
Call ClearForm(Me)
Call EntriesEnable(True)
Select Case SSTab1.Tab
    Case 0
        Call CompanyGRPComboListing
        Call BindBillingMode
End Select
End Sub
Public Sub BindBillingMode()
Dim sqlstr As String
Dim BillingMode As New ADODB.Recordset
sqlstr = ""

sqlstr = "SELECT Prefix,Description FROM Tbl_BillingModes"
BillingMode.Open sqlstr, medixmasterconn, adOpenDynamic, adLockOptimistic

cmbMcoBilling.Clear
cmbclaimBill.Clear

Do Until BillingMode.EOF
    cmbclaimBill.AddItem BillingMode!Prefix & " - " & BillingMode!Description
    cmbMcoBilling.AddItem BillingMode!Prefix & " - " & BillingMode!Description
    BillingMode.MoveNext
Loop
BillingMode.Close
Set BillingMode = Nothing
End Sub
Private Sub CompanyGRPComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
        
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    ' Set up a command object for the stored procedure.
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
        
    ' Create a recordset by executing the command.
    Set PAY = cmd.Execute
        
    CboCoGrpPayor.Clear

    Do Until PAY.EOF
        CboCoGrpPayor.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        PAY.MoveNext
    Loop
            
    PAY.Close
    Set PAY = Nothing
    
    cboCoConversion.Clear
    cboCoConversion.AddItem "N" & " - " & "No"
    cboCoConversion.AddItem "Y" & " - " & "Yes"
    
    
    
    cboCoTax.Clear
    cboCoTax.AddItem "N" & " - " & "No"
    cboCoTax.AddItem "Y" & " - " & "Yes"
    
    cboCoPrePost.Clear
    cboCoPrePost.AddItem "N" & " - " & "No"
    cboCoPrePost.AddItem "Y" & " - " & "Yes"
    
    CboGrpComStatus.Clear
    CboGrpComStatus.AddItem "B" & " - " & "Both"
    CboGrpComStatus.AddItem "I" & " - " & "In-Patient"
    CboGrpComStatus.AddItem "O" & " - " & "Out-Patient"
    
    
    
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End Sub

Private Sub cmdSave_Click()
  Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SaveCompanyGrp
        Case 1 'Subsidary
            Call SaveSubsidary
        Case 2 'Department
        
        Case 3 'costcenter
        
    End Select
    cmdAdd.Enabled = True
    cmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub
Private Sub SaveSubsidary()
Dim RecordSaved
Dim SubsidaryRS As New ADODB.Recordset
Dim StrSql As String
Dim strsql2 As String
Dim subsidaryCode As String

If cmbSubGrpcompany = "" Then
    MsgBox "Please Select Group Company:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbSubGrpcompany.SetFocus
ElseIf Len(Trim(txtSubName)) <= 0 Then
    MsgBox "Please Enter Subsidary Name:", vbOKOnly + vbCritical, DialogBoxTitle
    txtSubName.SetFocus
ElseIf Len(Trim(txtSubadd1)) <= 0 Then
    MsgBox "Please Enter Address 1:", vbOKOnly + vbCritical, DialogBoxTitle
    txtSubadd1.SetFocus
ElseIf Len(Trim(txtSubPostcode)) <= 0 Then
    MsgBox "Please Enter Post code:", vbOKOnly + vbCritical, DialogBoxTitle
    txtSubPostcode.SetFocus
ElseIf cmbSubstate = "" Then
    MsgBox "Please Select State:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbSubstate.SetFocus
ElseIf cmbSubMCO = "" Then
    MsgBox "Please Select MCO Billing Mode:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbSubMCO.SetFocus
ElseIf cmbsubClaims = "" Then
    MsgBox "Please Select Claim Billing Mode:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbsubClaims.SetFocus
Else
        StrSql = "Select * from Tbl_Subsidary where SubsidaryName ='" + Trim(txtSubName) + "' and GroupCompID ='" + cmbSubGrpcompany + "'"
        SubsidaryRS.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If (SubsidaryRS.recordCount > 0) And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            SubsidaryRS.Close
            Set SubsidaryRS = Nothing
            Exit Sub
        Else
            SubsidaryRS.Close
            Set SubsidaryRS = Nothing
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                 If EditFlag = False Then
                      SubsidaryRS.Open "Tbl_Subsidary", medixmasterconn, adOpenKeyset, adLockOptimistic
                      SubsidaryRS.AddNew
                     Call GenerateSequance("S", subsidaryCode)
                 Else
                      strsql2 = "Select * from Tbl_Subsidary where SubID ='" + Trim(txtSubID) + "' and GroupCompID ='" + Mid(cmbSubGrpcompany, 1, 4) + "'"
                      SubsidaryRS.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                 End If
                        With SubsidaryRS
                            !GroupCompID = Mid(ChkStr(cmbSubGrpcompany), 1, 4)
                            !SubID = IIf(Len(txtSubID), txtSubID, subsidaryCode)
                            !SubsidaryName = IIf(Len(txtSubName), txtSubName, "")
                            !SubAdd1 = IIf(Len(txtSubadd1), ChkStr(txtSubadd1), Null)
                            !SubAdd2 = IIf(Len(txtSubadd2), ChkStr(txtSubadd2), Null)
                            !SubAdd3 = IIf(Len(txtSubAdd3), ChkStr(txtSubAdd3), Null)
                            !City = IIf(Len(txtsubCity), ChkStr(txtsubCity), Null)
                            !PostCode = IIf(Len(txtSubPostcode), ChkStr(txtSubPostcode), Null)
                            !state = IIf(Len(cmbSubstate), ChkStr(cmbSubstate), Null)
                            !PhoneNumber = IIf(Len(txtsubPhone), ChkStr(txtsubPhone), Null)
                            !PIC = IIf(Len(txtSubPIC), ChkStr(txtSubPIC), Null)
                            !EmailID = IIf(Len(txtsubEmail), ChkStr(txtsubEmail), Null)
                            !ClaimBillingMode = IIf(Len(cmbSubMCO), Mid(cmbSubMCO, 1, 1), Null)
                            !PremiumeBillingMode = IIf(Len(cmbsubClaims), Mid(cmbsubClaims, 1, 1), Null)
                            !LastUpdateUser = Logon
                            !LastUpdateDate = Format(Date, "YYYY-MM-DD")
                        End With
                        SubsidaryRS.Update
                        SubsidaryRS.Close
                        Set SubsidaryRS = Nothing
            End If
        End If
  End If
End Sub
Private Sub SaveDepartment()
Dim RecordSaved
Dim SubsidaryRS As New ADODB.Recordset
Dim StrSql As String
Dim strsql2 As String
Dim subsidaryCode As String

If cmbDepGrpcompany = "" Then
    MsgBox "Please Select Group Company:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbDepGrpcompany.SetFocus
ElseIf Len(Trim(txtDeptName)) <= 0 Then
    MsgBox "Please Enter Subsidary Name:", vbOKOnly + vbCritical, DialogBoxTitle
    txtDeptName.SetFocus
ElseIf Len(Trim(txtDeptAdd1)) <= 0 Then
    MsgBox "Please Enter Address 1:", vbOKOnly + vbCritical, DialogBoxTitle
    txtDeptAdd1.SetFocus
ElseIf Len(Trim(txtdeptPostcode)) <= 0 Then
    MsgBox "Please Enter Post code:", vbOKOnly + vbCritical, DialogBoxTitle
    txtdeptPostcode.SetFocus
ElseIf cmbDeptState = "" Then
    MsgBox "Please Select State:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbDeptState.SetFocus
ElseIf cmbDeptMCO = "" Then
    MsgBox "Please Select MCO Billing Mode:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbDeptMCO.SetFocus
ElseIf cmbDeptClaim = "" Then
    MsgBox "Please Select Claim Billing Mode:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbDeptClaim.SetFocus
Else
        StrSql = "Select * from Tbl_Department where DeptName ='" + Trim(txtDeptName) + "' and GroupCompID ='" + Mid(ChkStr(cmbDepGrpcompany), 1, 4) + "'"
        SubsidaryRS.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If (SubsidaryRS.recordCount > 0) And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            SubsidaryRS.Close
            Set SubsidaryRS = Nothing
            Exit Sub
        Else
            SubsidaryRS.Close
            Set SubsidaryRS = Nothing
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                 If EditFlag = False Then
                      SubsidaryRS.Open "Tbl_Department", medixmasterconn, adOpenKeyset, adLockOptimistic
                      SubsidaryRS.AddNew
                     Call GenerateSequance("D", subsidaryCode)
                 Else
                      strsql2 = "Select * from Tbl_Department where DeptID ='" + Trim(txtSubID) + "' and GroupCompID ='" + Mid(ChkStr(cmbDepGrpcompany), 1, 4) + "'"
                      SubsidaryRS.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                 End If
                        With SubsidaryRS
                            !GroupCompID = Mid(ChkStr(cmbDepGrpcompany), 1, 4)
                            !DeptID = IIf(Len(txtSubID), txtSubID, subsidaryCode)
                            !DeptName = IIf(Len(txtSubName), txtSubName, "")
                            !DeptAdd1 = IIf(Len(txtSubadd1), ChkStr(txtSubadd1), Null)
                            !DeptAdd2 = IIf(Len(txtSubadd2), ChkStr(txtSubadd2), Null)
                            !DeptAdd3 = IIf(Len(txtSubAdd3), ChkStr(txtSubAdd3), Null)
                            !DeptCity = IIf(Len(txtsubCity), ChkStr(txtsubCity), Null)
                            !PostCode = IIf(Len(txtSubPostcode), ChkStr(txtSubPostcode), Null)
                            !state = IIf(Len(cmbSubstate), ChkStr(cmbSubstate), Null)
                            !PhoneNumber = IIf(Len(txtsubPhone), ChkStr(txtsubPhone), Null)
                            !PIC = IIf(Len(txtSubPIC), ChkStr(txtSubPIC), Null)
                            !EmailID = IIf(Len(txtsubEmail), ChkStr(txtsubEmail), Null)
                            !ClaimBillingMode = IIf(Len(cmbSubMCO), Mid(cmbSubMCO, 1, 1), Null)
                            !PremiumeBillingMode = IIf(Len(cmbsubClaims), Mid(cmbsubClaims, 1, 1), Null)
                            !LastUpdateUser = Logon
                            !LastUpdateDate = Format(Date, "YYYY-MM-DD")
                        End With
                        SubsidaryRS.Update
                        SubsidaryRS.Close
                        Set SubsidaryRS = Nothing
            End If
        End If
  End If
End Sub

Private Sub SaveCostcenter()
Dim RecordSaved
Dim SubsidaryRS As New ADODB.Recordset
Dim StrSql As String
Dim strsql2 As String
Dim subsidaryCode As String

If cmbCostGrpCompany = "" Then
    MsgBox "Please Select Group Company:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbCostGrpCompany.SetFocus
ElseIf Len(Trim(txtCostcentName)) <= 0 Then
    MsgBox "Please Enter Cost center Name:", vbOKOnly + vbCritical, DialogBoxTitle
    txtCostcentName.SetFocus
ElseIf Len(Trim(txtCostcenAdd1)) <= 0 Then
    MsgBox "Please Enter Address 1:", vbOKOnly + vbCritical, DialogBoxTitle
    txtCostcenAdd1.SetFocus
ElseIf Len(Trim(txtCostcenPostcode)) <= 0 Then
    MsgBox "Please Enter Post code:", vbOKOnly + vbCritical, DialogBoxTitle
    txtCostcenPostcode.SetFocus
ElseIf cmbCostState = "" Then
    MsgBox "Please Select State:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbCostState.SetFocus
ElseIf cmbMCOCostcen = "" Then
    MsgBox "Please Select MCO Billing Mode:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbMCOCostcen.SetFocus
ElseIf cmbClaimcostcenter = "" Then
    MsgBox "Please Select Claim Billing Mode:", vbOKOnly + vbCritical, DialogBoxTitle
    cmbClaimcostcenter.SetFocus
Else
        StrSql = "Select * from Tbl_CostCenter where CostCenterName ='" + Trim(txtDeptName) + "' and GroupCompID ='" + Mid(ChkStr(cmbCostGrpCompany), 1, 4) + "'"
        SubsidaryRS.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If (SubsidaryRS.recordCount > 0) And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            SubsidaryRS.Close
            Set SubsidaryRS = Nothing
            Exit Sub
        Else
            SubsidaryRS.Close
            Set SubsidaryRS = Nothing
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                 If EditFlag = False Then
                      SubsidaryRS.Open "Tbl_CostCenter", medixmasterconn, adOpenKeyset, adLockOptimistic
                      SubsidaryRS.AddNew
                     Call GenerateSequance("C", subsidaryCode)
                 Else
                      strsql2 = "Select * from Tbl_CostCenter where CostCenterID ='" + Trim(txtCostcenID) + "' and GroupCompID ='" + Mid(ChkStr(cmbCostGrpCompany), 1, 4) + "'"
                      SubsidaryRS.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                 End If
                        With SubsidaryRS
                            !GroupCompID = Mid(ChkStr(cmbDepGrpcompany), 1, 4)
                            !CostCenterID = IIf(Len(txtCostcenID), txtCostcenID, subsidaryCode)
                            !CostCenterName = IIf(Len(txtSubName), txtSubName, "")
                            !CostAdd1 = IIf(Len(txtSubadd1), ChkStr(txtSubadd1), Null)
                            !CostAdd2 = IIf(Len(txtSubadd2), ChkStr(txtSubadd2), Null)
                            !CostAdd3 = IIf(Len(txtSubAdd3), ChkStr(txtSubAdd3), Null)
                            !City = IIf(Len(txtsubCity), ChkStr(txtsubCity), Null)
                            !PostCode = IIf(Len(txtSubPostcode), ChkStr(txtSubPostcode), Null)
                            !state = IIf(Len(cmbSubstate), ChkStr(cmbSubstate), Null)
                            !PhoneNumber = IIf(Len(txtsubPhone), ChkStr(txtsubPhone), Null)
                            !PIC = IIf(Len(txtSubPIC), ChkStr(txtSubPIC), Null)
                            !EmailID = IIf(Len(txtsubEmail), ChkStr(txtsubEmail), Null)
                            !ClaimBillingMode = IIf(Len(cmbSubMCO), Mid(cmbSubMCO, 1, 1), Null)
                            !PremiumeBillingMode = IIf(Len(cmbsubClaims), Mid(cmbsubClaims, 1, 1), Null)
                            !LastUpdateUser = Logon
                            !LastUpdateDate = Format(Date, "YYYY-MM-DD")
                        End With
                        SubsidaryRS.Update
                        SubsidaryRS.Close
                        Set SubsidaryRS = Nothing
            End If
        End If
  End If
End Sub

Public Sub GenerateSequance(SeqType As String, subsidaryCode)
Dim COMSeq As New ADODB.Recordset
Dim GRPCoCode As String
Dim StrSql As String
Dim Prefix As String
Dim Sequance As String

StrSql = "Select Prefix,Sequence,SeqCode from Tbl_SubsidarySeq where SeqCode = '" & Trim(SeqType) & "'"
COMSeq.Open StrSql, medixmasterconn, adOpenDynamic, adLockOptimistic

If COMSeq.EOF = True Then
    With COMSeq
        .AddNew
        !Sequence = "0001"
        Sequance = COMSeq!Sequence
        Prefix = COMSeq!Prefix
        .Update
        
    End With
Else
    
    If COMSeq!Sequence = "9999" Then
        If Mid(Prefix, 1, 1) = "Z" Then
            COMSeq!Prefix = Chr(Asc(Mid(COMSeq!Prefix, 1, 1)) + 1) & "A"
        ElseIf Mid(Prefix, 2, 1) = "Z" Then
            COMSeq!Prefix = Mid(COMSeq!Prefix, 1, 1) & Chr(Asc(Mid(COMSeq!Prefix, 2, 1)) + 1)
        End If
        COMSeq!Sequance = "0001"
    Else
        COMSeq!Sequence = Format(COMSeq!Sequence + 1, "0000")
    End If
        Sequance = COMSeq!Sequence
        Prefix = COMSeq!Prefix
    COMSeq.Update
End If
COMSeq.Close
Set COMSeq = Nothing

subsidaryCode = Prefix & Sequance

End Sub

'**********************************Start Group Company**************************************
Private Sub SaveCompanyGrp()
Dim RecordSaved
Dim CompanyGrp As New ADODB.Recordset
Dim SaveCompanyGrp As New ADODB.Recordset
Dim CISCompanyGrp As New ADODB.Recordset
Dim SaveCISCompanyGrp As New ADODB.Recordset
Dim StrSql As String
Dim strsql2 As String
Dim COMSeq As String
    If txtCompanyName = "" Then
        MsgBox "Please Enter Company Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtCompanyName.SetFocus
    ElseIf txtCoAdd1 = "" Then
        MsgBox "Please Enter Group Company Address", vbOKOnly + vbCritical, DialogBoxTitle
        txtCoAdd1.SetFocus
    ElseIf txtCoEffDate = "__/__/____" Then
        MsgBox "Please Enter Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtCoEffDate.SetFocus
    ElseIf CboGrpComStatus = "" Then
        MsgBox "Please Enter Status", vbOKOnly + vbCritical, DialogBoxTitle
        CboGrpComStatus.SetFocus
    ElseIf cmbclaimBill = "" Then
        MsgBox "Please Select Claim Billing Mode", vbOKOnly + vbCritical, DialogBoxTitle
        cmbclaimBill.SetFocus
    ElseIf cmbMcoBilling = "" Then
        MsgBox "Please Select MCO Billing Mode", vbOKOnly + vbCritical, DialogBoxTitle
        cmbMcoBilling.SetFocus
    ElseIf cmbChannel = "" Then
        MsgBox "Please Select Business Type", vbOKOnly + vbCritical, DialogBoxTitle
        cmbChannel.SetFocus
    Else
        StrSql = "Select * From CompanyGRP Where CoGRPName = '" & Trim(txtCompanyName) & "'"
        CompanyGrp.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        strsql2 = "Select * From GRPCompany Where GRPCoName = '" & Trim(txtCompanyName) & "'"
        CISCompanyGrp.Open strsql2, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If (CompanyGrp.recordCount > 0 Or CISCompanyGrp.recordCount > 0) And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            CompanyGrp.Close
            CISCompanyGrp.Close
            txtCompanyName.SetFocus
            txtCompanyName.SelStart = 0
            txtCompanyName.SelLength = Len(txtCompanyName)
            Set CompanyGrp = Nothing
            Set CISCompanyGrp = Nothing
            'Exit Sub
        Else
            CompanyGrp.Close
            Set CompanyGrp = Nothing
            CISCompanyGrp.Close
            Set CISCompanyGrp = Nothing
            
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    If Mid(CboGrpComStatus, 1, 1) = "B" Then
                        CompanyGrp.Open "CompanyGRP", medixmasterconn, adOpenKeyset, adLockOptimistic
                        CompanyGrp.AddNew
                        CISCompanyGrp.Open "GRPCompany", MediConnCISDB, adOpenKeyset, adLockOptimistic
                        CISCompanyGrp.AddNew
                    ElseIf Mid(CboGrpComStatus, 1, 1) = "I" Then
                        CompanyGrp.Open "CompanyGRP", medixmasterconn, adOpenKeyset, adLockOptimistic
                        CompanyGrp.AddNew
                    ElseIf Mid(CboGrpComStatus, 1, 1) = "O" Then
                        CISCompanyGrp.Open "GRPCompany", MediConnCISDB, adOpenKeyset, adLockOptimistic
                        CISCompanyGrp.AddNew
                    End If
                    COMSeq = GenerateCoGrpSeq(Mid(txtCompanyName, 1, 1))
                Else
                    If Mid(CboGrpComStatus, 1, 1) = "O" Then
                        strsql2 = "Select * From GRPCompany Where GRPCoCode = '" & Trim(txtGrpCoCode) & "'"
                        CISCompanyGrp.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                        COMSeq = txtGrpCoCode
                    ElseIf Mid(CboGrpComStatus, 1, 1) = "I" Then
                        StrSql = "Select * From CompanyGRP Where CoGRPCode = '" & Trim(txtGrpCoCode) & "'"
                        CompanyGrp.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        COMSeq = txtGrpCoCode
                    Else
                        StrSql = "Select * From CompanyGRP Where CoGRPCode = '" & Trim(txtGrpCoCode) & "'"
                        CompanyGrp.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        COMSeq = txtGrpCoCode
                        
                        strsql2 = "Select * From GRPCompany Where GRPCoCode = '" & Trim(txtGrpCoCode) & "'"
                        CISCompanyGrp.Open strsql2, MediConnCISDB, adOpenKeyset, adLockOptimistic
                        COMSeq = txtGrpCoCode
                    End If
                End If
                
                    If Mid(CboGrpComStatus, 1, 1) = "B" Then
                    
                        With CompanyGrp
                            !CoGRPCode = ChkStr(COMSeq)
                            !CoGRPName = ChkStr(txtCompanyName)
                            !CoGRPPaycode = Mid(CboCoGrpPayor, 1, 2)
                            !CoGRPAdd1 = IIf(Len(txtCoAdd1), ChkStr(txtCoAdd1), Null)
                            !CoGRPAdd2 = IIf(Len(txtCoAdd2), ChkStr(txtCoAdd2), Null)
                            !CoGRPAdd3 = IIf(Len(txtCoAdd3), ChkStr(txtCoAdd3), Null)
                            !CoGRPAdmContact = IIf(Len(txtCoAdmContact), ChkStr(txtCoAdmContact), Null)
                            !CoGRPContact = IIf(Len(txtCoContact), ChkStr(txtCoContact), Null)
                            !CoGRPCEO = IIf(Len(txtCoCEO), ChkStr(txtCoCEO), Null)
                            !CoGRPPhone = IIf(Len(txtCoPhone), ChkStr(txtCoPhone), Null)
                            !CoGRPFax = IIf(Len(txtCoFax), ChkStr(txtCoFax), Null)
                            !CoGRPEmail = IIf(Len(txtCoEmail), ChkStr(txtCoEmail), Null)
                            !CoGRPEffDate = IIf(IsDate(txtCoEffDate), txtCoEffDate, Null)
                            !CoGRPTax = Mid(cboCoTax, 1, 1)
                            !CoGRPConversion = Mid(cboCoConversion, 1, 1)
                            !CoGRPPrePost = Mid(cboCoPrePost, 1, 1)
                            !CoGRPStatus = Mid(CboGrpComStatus, 1, 1)
                            !BusinessChanel = IIf(Len(cmbChannel), Mid(cmbChannel, 1, 1), Null)
                            If Mid(CboCoGrpPayor, 1, 2) = "TK" Or Mid(CboCoGrpPayor, 1, 2) = "TS" Then
                                !CoGrpTISM = "Y"
                            Else
                                !CoGrpTISM = "N"
                            End If
                            !ClaimsBillingMode = Mid(cmbclaimBill, 1, 1)
                            !McoBillingMode = Mid(cmbMcoBilling, 1, 1)
                            !ClaimsPIC = IIf(Len(txtClaimsPic), txtClaimsPic, "")
                            !MCOPIC = IIf(Len(txtMCOPIC), txtMCOPIC, "")
                            !CoGRPLastUpdateUser = Logon
                            !CoGRPLastUpdateDate = Date
                        End With
                        CompanyGrp.Update
                        CompanyGrp.Close
                        Set CompanyGrp = Nothing
                        With CISCompanyGrp
                            !GRPCoCode = ChkStr(COMSeq)
                            !GRPCoName = ChkStr(txtCompanyName)
                            !GRPCoPaycode = Mid(CboCoGrpPayor, 1, 2)
                            !GRPCoAdd1 = IIf(Len(txtCoAdd1), ChkStr(txtCoAdd1), Null)
                            !GRPCoAdd2 = IIf(Len(txtCoAdd2), ChkStr(txtCoAdd2), Null)
                            !GRPCoAdd3 = IIf(Len(txtCoAdd3), ChkStr(txtCoAdd3), Null)
                            !GRPCoAdmContact = IIf(Len(txtCoAdmContact), ChkStr(txtCoAdmContact), Null)
                            !GRPCoContact = IIf(Len(txtCoContact), ChkStr(txtCoContact), Null)
                            !GRPCoCEO = IIf(Len(txtCoCEO), ChkStr(txtCoCEO), Null)
                            !GRPCoPhone = IIf(Len(txtCoPhone), ChkStr(txtCoPhone), Null)
                            !GRPCoFax = IIf(Len(txtCoFax), ChkStr(txtCoFax), Null)
                            !GRPCoEmail = IIf(Len(txtCoEmail), ChkStr(txtCoEmail), Null)
                            !GRPCoEffDate = IIf(IsDate(txtCoEffDate), txtCoEffDate, Null)
                            !GRPCoTax = Mid(cboCoTax, 1, 1)
                            !GRPCoConversion = Mid(cboCoConversion, 1, 1)
                            !GRPCoPrePost = Mid(cboCoPrePost, 1, 1)
                            !GRPCoStatus = Mid(CboGrpComStatus, 1, 1)
                            !BusinessChanel = IIf(Len(cmbChannel), Mid(cmbChannel, 1, 1), Null)
                            !GRPTIMS = IIf(Len(cmbTims), Mid(cmbTims, 1, 1), Null)
                            !ClaimsBillingMode = Mid(cmbclaimBill, 1, 1)
                            !McoBillingMode = Mid(cmbMcoBilling, 1, 1)
                            !ClaimsPIC = IIf(Len(txtClaimsPic), txtClaimsPic, "")
                            !MCOPIC = IIf(Len(txtMCOPIC), txtMCOPIC, "")
                            !GRPCoLastUpdateUser = Logon
                            !GRPCoLastUpdateDate = Date
                        End With
                        CISCompanyGrp.Update
                        CISCompanyGrp.Close
                        Set CISCompanyGrp = Nothing
                    ElseIf Mid(CboGrpComStatus, 1, 1) = "I" Then
                        With CompanyGrp
                            !CoGRPCode = ChkStr(COMSeq)
                            !CoGRPName = ChkStr(txtCompanyName)
                            !CoGRPPaycode = Mid(CboCoGrpPayor, 1, 2)
                            !CoGRPAdd1 = IIf(Len(txtCoAdd1), ChkStr(txtCoAdd1), Null)
                            !CoGRPAdd2 = IIf(Len(txtCoAdd2), ChkStr(txtCoAdd2), Null)
                            !CoGRPAdd3 = IIf(Len(txtCoAdd3), ChkStr(txtCoAdd3), Null)
                            !CoGRPAdmContact = IIf(Len(txtCoAdmContact), ChkStr(txtCoAdmContact), Null)
                            !CoGRPContact = IIf(Len(txtCoContact), ChkStr(txtCoContact), Null)
                            !CoGRPCEO = IIf(Len(txtCoCEO), ChkStr(txtCoCEO), Null)
                            !CoGRPPhone = IIf(Len(txtCoPhone), ChkStr(txtCoPhone), Null)
                            !CoGRPFax = IIf(Len(txtCoFax), ChkStr(txtCoFax), Null)
                            !CoGRPEmail = IIf(Len(txtCoEmail), ChkStr(txtCoEmail), Null)
                            !CoGRPEffDate = IIf(IsDate(txtCoEffDate), txtCoEffDate, Null)
                            !CoGRPTax = Mid(cboCoTax, 1, 1)
                            !CoGRPConversion = Mid(cboCoConversion, 1, 1)
                            !CoGRPPrePost = Mid(cboCoPrePost, 1, 1)
                            !CoGRPStatus = Mid(CboGrpComStatus, 1, 1)
                            !BusinessChanel = IIf(Len(cmbChannel), Mid(cmbChannel, 1, 1), Null)
                            !CoGrpTISM = IIf(Len(cmbTims), Mid(cmbTims, 1, 1), Null)
                            !ClaimsBillingMode = Mid(cmbclaimBill, 1, 1)
                            !McoBillingMode = Mid(cmbMcoBilling, 1, 1)
                            !ClaimsPIC = IIf(Len(txtClaimsPic), txtClaimsPic, "")
                            !MCOPIC = IIf(Len(txtMCOPIC), txtMCOPIC, "")
                            !CoGRPLastUpdateUser = Logon
                            !CoGRPLastUpdateDate = Date
                        End With
                        CompanyGrp.Update
                        CompanyGrp.Close
                        Set CompanyGrp = Nothing
                    ElseIf Mid(CboGrpComStatus, 1, 1) = "O" Then
                        With CISCompanyGrp
                            !GRPCoCode = ChkStr(COMSeq)
                            !GRPCoName = ChkStr(txtCompanyName)
                            !GRPCoPaycode = Mid(CboCoGrpPayor, 1, 2)
                            !GRPCoAdd1 = IIf(Len(txtCoAdd1), ChkStr(txtCoAdd1), Null)
                            !GRPCoAdd2 = IIf(Len(txtCoAdd2), ChkStr(txtCoAdd2), Null)
                            !GRPCoAdd3 = IIf(Len(txtCoAdd3), ChkStr(txtCoAdd3), Null)
                            !GRPCoAdmContact = IIf(Len(txtCoAdmContact), ChkStr(txtCoAdmContact), Null)
                            !GRPCoContact = IIf(Len(txtCoContact), ChkStr(txtCoContact), Null)
                            !GRPCoCEO = IIf(Len(txtCoCEO), ChkStr(txtCoCEO), Null)
                            !GRPCoPhone = IIf(Len(txtCoPhone), ChkStr(txtCoPhone), Null)
                            !GRPCoFax = IIf(Len(txtCoFax), ChkStr(txtCoFax), Null)
                            !GRPCoEmail = IIf(Len(txtCoEmail), ChkStr(txtCoEmail), Null)
                            !GRPCoEffDate = IIf(IsDate(txtCoEffDate), txtCoEffDate, Null)
                            !GRPCoTax = Mid(cboCoTax, 1, 1)
                            !GRPCoConversion = Mid(cboCoConversion, 1, 1)
                            !GRPCoPrePost = Mid(cboCoPrePost, 1, 1)
                            !GRPCoStatus = Mid(CboGrpComStatus, 1, 1)
                            !BusinessChanel = IIf(Len(cmbChannel), Mid(cmbChannel, 1, 1), Null)
                            !GRPTIMS = IIf(Len(cmbTims), Mid(cmbTims, 1, 1), Null)
                            !ClaimsBillingMode = Mid(cmbclaimBill, 1, 1)
                            !McoBillingMode = Mid(cmbMcoBilling, 1, 1)
                            !ClaimsPIC = IIf(Len(txtClaimsPic), txtClaimsPic, "")
                            !MCOPIC = IIf(Len(txtMCOPIC), txtMCOPIC, "")
                            !GRPCoLastUpdateUser = Logon
                            !GRPCoLastUpdateDate = Date
                        End With
                        CISCompanyGrp.Update
                        CISCompanyGrp.Close
                        Set CISCompanyGrp = Nothing
                        
                    End If
                        Call ClearForm(Me)
                        Call CompanyGRPListing
                        If EditFlag = False Then
                            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                        Else
                            MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                        End If
            End If
            EditFlag = False
        End If
    End If
    Call EntriesEnable(True)
End Sub
Private Sub Form_Load()
    cmdSave.Enabled = True
    GRPCoCriteria = ""
    Call BindChannel
    Select Case SSTab1.Tab
    Case 0
        Call CompanyGRPComboListing
        Call BindBillingMode
    End Select
End Sub
Private Sub BindChannel()
Dim StrSql As String
Dim CompanyGrp As New ADODB.Recordset
    cmbChannel.Clear
    StrSql = "Select PreFix,Description from HISMaintenance..Tbl_BusinessChaneel "
    CompanyGrp.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until CompanyGrp.EOF
       cmbChannel.AddItem (CompanyGrp!Prefix + " - " + CompanyGrp!Description)
    CompanyGrp.MoveNext
    Loop
    CompanyGrp.Close
    Set CompanyGrp = Nothing
End Sub

Private Sub CmdSearch_Click()
cmdSearch.Enabled = False
   Screen.MousePointer = vbHourglass
   Select Case SSTab1.Tab
        Case 0 'Group compnay
            Call SearchCompanyGrp
            Call ClearForm(Me)
            txtGrpCoCode.Enabled = True
            txtCompanyName.Enabled = True
            txtGrpCoCode.SetFocus
        Case 1  'Subsidary
            Call SearchSubsidary
        Case 2  'Department
            Call SearchDepartment
        Case 3  'costcenter
            Call SearchCostcenter
    End Select
    Screen.MousePointer = Default
    cmdAdd.Enabled = True
    cmdSearch.Enabled = True
End Sub

Private Sub SearchSubsidary()
Dim strAppend As String
    Screen.MousePointer = vbHourglass
    strAppend = ""
    If Len(Trim(txtSubID)) Then
        strAppend = strAppend + " And SubID ='" & ChkStr(txtSubID) & "'"
    End If
    If Len(Trim(txtSubName)) Then
        strAppend = strAppend + " And SubsidaryName like '%" & ChkStr(txtSubName) & "%'"
    End If
    If Len(Trim(cmbSubGrpcompany)) Then
        strAppend = strAppend + " And GroupCompID = '" & Mid(cmbSubGrpcompany, 1, 4) & "'"
    End If
    GRPCoCriteria = strAppend
    Call SubsidaryListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub SearchCostcenter()
Dim strAppend As String
    Screen.MousePointer = vbHourglass
    strAppend = ""
    If Len(Trim(txtSubID)) Then
        strAppend = strAppend + " And CostCenterID ='" & ChkStr(txtCostcenID) & "'"
    End If
    If Len(Trim(txtSubName)) Then
        strAppend = strAppend + " And CostCenterName like '%" & ChkStr(txtCostcentName) & "%'"
    End If
    If Len(Trim(cmbSubGrpcompany)) Then
        strAppend = strAppend + " And GroupCompID = '" & Mid(cmbCostGrpCompany, 1, 4) & "'"
    End If
    GRPCoCriteria = strAppend
    Call CostcenterListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub SearchDepartment()
Dim strAppend As String
    Screen.MousePointer = vbHourglass
    strAppend = ""
    If Len(Trim(txtSubID)) Then
        strAppend = strAppend + " And DeptID ='" & ChkStr(txtDeptID) & "'"
    End If
    If Len(Trim(txtSubName)) Then
        strAppend = strAppend + " And DeptName like '%" & ChkStr(txtDeptName) & "%'"
    End If
    If Len(Trim(cmbSubGrpcompany)) Then
        strAppend = strAppend + " And GroupCompID = '" & Mid(cmbDepGrpcompany, 1, 4) & "'"
    End If
    GRPCoCriteria = strAppend
    Call DepartmentListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub DepartmentListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim CompanyGrpMaster As ListItem
Dim sql As String
Dim TotalGRPCOMCount As Integer
    
    sql = "Select * from Tbl_Department where 0 = 0 "
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    lstviewDepartment.ListItems.Clear
    
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set CompanyGrpMaster = lstviewDepartment.ListItems.Add(, , rst2!DeptID)
                CompanyGrpMaster.SubItems(1) = rst2!DeptName
                
                If Len(rst2!GroupCompID) Then
                    CompanyGrpMaster.SubItems(2) = rst2!GroupCompID
                End If
                
                If Len(rst2!DeptAdd1) Then
                    CompanyGrpMaster.SubItems(3) = rst2!DeptAdd1
                End If
                               
                If Len(rst2!DeptAdd2) Then
                    CompanyGrpMaster.SubItems(4) = rst2!DeptAdd2
                End If
                
                If Len(rst2!DeptAdd3) Then
                    CompanyGrpMaster.SubItems(5) = rst2!DeptAdd3
                End If
                
                If Len(rst2!DeptCity) Then
                    CompanyGrpMaster.SubItems(6) = rst2!DeptCity
                End If
                
                If Len(rst2!PostCode) Then
                    CompanyGrpMaster.SubItems(7) = rst2!PostCode
                End If
                
                If Len(rst2!state) Then
                    CompanyGrpMaster.SubItems(8) = rst2!state
                End If
                
                If Len(rst2!PhoneNumber) Then
                    CompanyGrpMaster.SubItems(9) = rst2!PhoneNumber
                End If
                
                If Len(rst2!PIC) Then
                    CompanyGrpMaster.SubItems(10) = rst2!PIC
                End If
                
                If Len(rst2!EmailID) Then
                    CompanyGrpMaster.SubItems(11) = rst2!EmailID
                End If
                
                If Len(rst2!ClaimBillingMode) Then
                    CompanyGrpMaster.SubItems(12) = rst2!ClaimBillingMode
                End If
                
                If Len(rst2!PremiumeBillingMode) Then
                    CompanyGrpMaster.SubItems(13) = rst2!PremiumeBillingMode
                End If
                
                rst2.MoveNext
            Loop
    rst2.Close
    End If
End Sub

Private Sub CostcenterListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim CompanyGrpMaster As ListItem
Dim sql As String
Dim TotalGRPCOMCount As Integer
    
    sql = "Select  * from Tbl_CostCenter where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
              
    TotalGRPCOMCount = 0
    lstViewCostcentre.ListItems.Clear
    
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set CompanyGrpMaster = lstViewCostcentre.ListItems.Add(, , rst2!CostCenterID)
                CompanyGrpMaster.SubItems(1) = rst2!CostCenterName
                
                If Len(rst2!GroupCompID) Then
                    CompanyGrpMaster.SubItems(2) = rst2!GroupCompID
                End If
                
                If Len(rst2!CostAdd1) Then
                    CompanyGrpMaster.SubItems(3) = rst2!CostAdd1
                End If
                               
                If Len(rst2!CostAdd2) Then
                    CompanyGrpMaster.SubItems(4) = rst2!CostAdd2
                End If
                
                If Len(rst2!CostAdd3) Then
                    CompanyGrpMaster.SubItems(5) = rst2!CostAdd3
                End If
                
                If Len(rst2!City) Then
                    CompanyGrpMaster.SubItems(6) = rst2!City
                End If
                
                If Len(rst2!PostCode) Then
                    CompanyGrpMaster.SubItems(7) = rst2!PostCode
                End If
                
                If Len(rst2!state) Then
                    CompanyGrpMaster.SubItems(8) = rst2!state
                End If
                
                If Len(rst2!PhoneNumber) Then
                    CompanyGrpMaster.SubItems(9) = rst2!PhoneNumber
                End If
                
                If Len(rst2!PIC) Then
                    CompanyGrpMaster.SubItems(10) = rst2!PIC
                End If
                
                If Len(rst2!EmailID) Then
                    CompanyGrpMaster.SubItems(11) = rst2!EmailID
                End If
                
                If Len(rst2!ClaimBillingMode) Then
                    CompanyGrpMaster.SubItems(12) = rst2!ClaimBillingMode
                End If
               
                If Len(rst2!McoBillingMode) Then
                    CompanyGrpMaster.SubItems(13) = rst2!McoBillingMode
                End If
                
                rst2.MoveNext
        TotalGRPCOMCount = TotalGRPCOMCount + 1
        Loop
    rst2.Close
    End If
End Sub

Private Sub SubsidaryListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim CompanyGrpMaster As ListItem
Dim sql As String
Dim TotalGRPCOMCount As Integer

    sql = "Select * from Tbl_Subsidary where 0 = 0 "
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
              
    TotalGRPCOMCount = 0
    LstviewSubsidary.ListItems.Clear
    
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set CompanyGrpMaster = LstviewSubsidary.ListItems.Add(, , rst2!SubID)
                
                CompanyGrpMaster.SubItems(1) = rst2!GroupCompID
                
                If Len(rst2!SubsidaryName) Then
                    CompanyGrpMaster.SubItems(2) = rst2!SubsidaryName
                End If
                If Len(rst2!SubAdd1) Then
                    CompanyGrpMaster.SubItems(3) = rst2!SubAdd1
                End If
                If Len(rst2!SubAdd2) Then
                    CompanyGrpMaster.SubItems(4) = rst2!SubAdd2
                End If
                If Len(rst2!SubAdd3) Then
                    CompanyGrpMaster.SubItems(5) = rst2!SubAdd3
                End If
                If Len(rst2!City) Then
                    CompanyGrpMaster.SubItems(6) = rst2!City
                End If
                If Len(rst2!PostCode) Then
                    CompanyGrpMaster.SubItems(7) = rst2!PostCode
                End If
                If Len(rst2!state) Then
                    CompanyGrpMaster.SubItems(8) = rst2!state
                End If
                If Len(rst2!PhoneNumber) Then
                    CompanyGrpMaster.SubItems(9) = rst2!PhoneNumber
                End If
                If Len(rst2!PIC) Then
                    CompanyGrpMaster.SubItems(10) = rst2!PIC
                End If
                If Len(rst2!EmailID) Then
                    CompanyGrpMaster.SubItems(11) = rst2!EmailID
                End If
                If Len(rst2!ClaimBillingMode) Then
                    CompanyGrpMaster.SubItems(12) = rst2!ClaimBillingMode
                End If
                If Len(rst2!PremiumeBillingMode) Then
                    CompanyGrpMaster.SubItems(13) = rst2!PremiumeBillingMode
                End If
                rst2.MoveNext
        Loop
    rst2.Close
    End If
End Sub



Private Sub CompanyGRPListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim CompanyGrpMaster As ListItem
Dim sql As String
Dim TotalGRPCOMCount As Integer
    
    sql = "Select DISTINCT * from VIEW_GRPCompany where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
              
    TotalGRPCOMCount = 0
    lstCompanyGrpList.ListItems.Clear
    
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set CompanyGrpMaster = lstCompanyGrpList.ListItems.Add(, , rst2!GrpCode)
                CompanyGrpMaster.SubItems(1) = rst2!GrpName
                
                If Len(rst2!PAYCode) Then
                    CompanyGrpMaster.SubItems(2) = rst2!PAYCode
                End If
                
                If Len(rst2!Business) Then
                    CompanyGrpMaster.SubItems(3) = rst2!Business + " - " + rst2!Decr
                End If
                               
                If Len(rst2!Address1) Then
                    CompanyGrpMaster.SubItems(4) = rst2!Address1
                End If
                
                If Len(rst2!Address2) Then
                    CompanyGrpMaster.SubItems(5) = rst2!Address2
                End If
                
                If Len(rst2!Address3) Then
                    CompanyGrpMaster.SubItems(6) = rst2!Address3
                End If
                
                If Len(rst2!Phone) Then
                    CompanyGrpMaster.SubItems(7) = rst2!Phone
                End If
                
                If Len(rst2!fax) Then
                    CompanyGrpMaster.SubItems(8) = rst2!fax
                End If
                
                If Len(rst2!EffDt) Then
                    CompanyGrpMaster.SubItems(9) = rst2!EffDt
                End If
                
                If Len(rst2!Conversion) Then
                    CompanyGrpMaster.SubItems(10) = rst2!Conversion
                End If
                
                If Len(rst2!ContactPerson) Then
                    CompanyGrpMaster.SubItems(11) = rst2!ContactPerson
                End If
                
                If Len(rst2!CEO) Then
                    CompanyGrpMaster.SubItems(12) = rst2!CEO
                End If
                
                If Len(rst2!ClaimsBillingMode) Then
                    CompanyGrpMaster.SubItems(13) = rst2!ClaimsBillingMode
                End If
                
                If Len(rst2!McoBillingMode) Then
                    CompanyGrpMaster.SubItems(14) = rst2!McoBillingMode
                End If
                
                If Len(rst2!GRPStatus) Then
                    CompanyGrpMaster.SubItems(15) = rst2!GRPStatus
                End If
                
                If Len(rst2!ClaimsPIC) Then
                    CompanyGrpMaster.SubItems(16) = rst2!ClaimsPIC
                End If
                
                If Len(rst2!MCOPIC) Then
                    CompanyGrpMaster.SubItems(17) = rst2!MCOPIC
                End If
                
                If Len(rst2!TIMSClinet) Then
                    CompanyGrpMaster.SubItems(18) = rst2!TIMSClinet
                Else
                    CompanyGrpMaster.SubItems(18) = "N"
                End If
                
                If Len(rst2!RB) Then
                    CompanyGrpMaster.SubItems(19) = rst2!RB
                End If
                
                If Len(rst2!Prepost) Then
                    CompanyGrpMaster.SubItems(20) = rst2!Prepost
                End If
                rst2.MoveNext
        TotalGRPCOMCount = TotalGRPCOMCount + 1
        Loop
    rst2.Close
    End If
End Sub
Private Sub SearchCompanyGrp()
Dim strAppend As String
    Screen.MousePointer = vbHourglass
    strAppend = ""
    If Len(Trim(txtGrpCoCode)) Then
        strAppend = strAppend + " And Grpcode like '%" & ChkStr(txtGrpCoCode) & "%'"
    End If
    If Len(Trim(txtCompanyName)) Then
        strAppend = strAppend + " And GrpName like '%" & ChkStr(txtCompanyName) & "%'"
    End If
    GRPCoCriteria = strAppend
    Call CompanyGRPListing(strAppend)
    Screen.MousePointer = Default
End Sub
Private Sub lstCompanyGrpList_Click()
Dim rst3 As New ADODB.Recordset
Dim StrSql As String
        
    If lstCompanyGrpList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        StrSql = "Select * from VIEW_GRPCompany Where Grpcode = '" & lstCompanyGrpList.SelectedItem.Text & "'"
        rst3.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        txtGrpCoCode = rst3!GrpCode
        txtCompanyName = rst3!GrpName
        
        If Len(rst3!PAYCode) Then
            CboCoGrpPayor = rst3!PAYCode
        End If
        
        If Len(rst3!Address1) Then
            txtCoAdd1 = rst3!Address1
        End If
        If Len(rst3!Address2) Then
            txtCoAdd2 = rst3!Address2
        End If
        If Len(rst3!Address3) Then
            txtCoAdd3 = rst3!Address3
        End If
        
        If Len(rst3!Admcontact) Then
            txtCoAdmContact = rst3!Admcontact
        End If
        If Len(rst3!ContactPerson) Then
            txtCoContact = rst3!ContactPerson
        End If
        If Len(rst3!CEO) Then
            txtCoCEO = rst3!CEO
        End If
        If Len(rst3!Phone) Then
            txtCoPhone = rst3!Phone
        End If
        If Len(rst3!fax) Then
            txtCoFax = rst3!fax
        End If
        If Len(rst3!Email) Then
            txtCoEmail = rst3!Email
        End If
        If Len(rst3!HomePage) Then
            txtCoHomepage = rst3!HomePage
        End If
        If Len(rst3!EffDt) Then
            txtCoEffDate = rst3!EffDt
        End If
        If Len(rst3!RB) Then
            cboCoTax = rst3!RB
        End If
        If Len(rst3!Conversion) Then
            cboCoConversion = rst3!Conversion
        End If
        If Len(rst3!Prepost) Then
            cboCoPrePost = rst3!Prepost
        End If
        If Len(rst3!GRPStatus) Then
            CboGrpComStatus = rst3!GRPStatus
        End If
        If Len(rst3!Business) Then
            cmbChannel = rst3!Business + " - " + rst3!Decr
        End If
        
        If Len(rst3!TIMSClinet) Then
            cmbTims = rst3!TIMSClinet
        End If
        If Len(rst3!ClaimsBillingMode) Then
            cmbclaimBill = rst3!ClaimsBillingMode
        End If
        If Len(rst3!McoBillingMode) Then
            cmbMcoBilling = rst3!McoBillingMode
        End If
        
        If Len(rst3!ClaimsPIC) Then
            txtClaimsPic = rst3!ClaimsPIC
        End If
        If Len(rst3!MCOPIC) Then
            txtMCOPIC = rst3!MCOPIC
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
    End If
End Sub

Private Sub lstViewCostcentre_ItemClick(ByVal Item As MSComctlLib.ListItem)
  txtCostcenID = Item
  txtCostcentName = Item.SubItems(1)
  cmbCostGrpCompany = Item.SubItems(2)
  txtCostcenAdd1 = Item.SubItems(3)
  txtCostcenAdd2 = Item.SubItems(4)
  txtCostcenAdd3 = Item.SubItems(5)
  txtCostcenCity = Item.SubItems(6)
  txtCostcenPostcode = Item.SubItems(7)
  cmbCostState = Item.SubItems(8)
  txtCostcenPhone = Item.SubItems(9)
  txtCostcenPIC = Item.SubItems(10)
  txtCostcenEmail = Item.SubItems(11)
  cmbMCOCostcen = Item.SubItems(12)
  cmbClaimcostcenter = Item.SubItems(13)
  cmdEdit.Enabled = True
End Sub

Private Sub lstviewDepartment_ItemClick(ByVal Item As MSComctlLib.ListItem)
  txtDeptID = Item
  txtDeptName = Item.SubItems(1)
  cmbDepGrpcompany = Item.SubItems(2)
  txtDeptAdd1 = Item.SubItems(3)
  txtDeptadd2 = Item.SubItems(4)
  txtDeptAdd3 = Item.SubItems(5)
  txtDeptCity = Item.SubItems(6)
  txtdeptPostcode = Item.SubItems(7)
  cmbDeptState = Item.SubItems(8)
  txtDeptPhone = Item.SubItems(9)
  txtDeptPIC = Item.SubItems(10)
  txtDeptEmail = Item.SubItems(11)
  cmbDeptMCO = Item.SubItems(12)
  cmbDeptClaim = Item.SubItems(13)
  cmdEdit.Enabled = True
End Sub

Private Sub LstviewSubsidary_ItemClick(ByVal Item As MSComctlLib.ListItem)
  txtSubID = Item
  cmbSubGrpcompany = Item.SubItems(1)
  txtSubName = Item.SubItems(2)
  txtSubadd1 = Item.SubItems(3)
  txtSubadd2 = Item.SubItems(4)
  txtSubAdd3 = Item.SubItems(5)
  txtsubCity = Item.SubItems(6)
  txtSubPostcode = Item.SubItems(7)
  cmbSubstate = Item.SubItems(8)
  txtsubPhone = Item.SubItems(9)
  txtSubPIC = Item.SubItems(10)
  txtsubEmail = Item.SubItems(11)
  cmbSubMCO = Item.SubItems(12)
  cmbsubClaims = Item.SubItems(13)
  cmdEdit.Enabled = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

 Select Case SSTab1.Tab
        Case 0 'Group company
            Call BindBillingSubsidary(cmbclaimBill, cmbMcoBilling)
        Case 1 'Subsidary
            Call BindBillingSubsidary(cmbSubMCO, cmbsubClaims)
            Call BindGroupcompany(cmbSubGrpcompany)
            Call BindState(cmbSubstate)
        Case 2 'Department
            Call BindBillingSubsidary(cmbDeptMCO, cmbDeptClaim)
            Call BindGroupcompany(cmbDepGrpcompany)
            Call BindState(cmbDeptState)
        Case 3 'Cost centre
            Call BindBillingSubsidary(cmbMCOCostcen, cmbClaimcostcenter)
            Call BindGroupcompany(cmbCostGrpCompany)
            Call BindState(cmbCostState)
    End Select
End Sub

Public Sub BindBillingSubsidary(cmb1 As ComboBox, cmb2 As ComboBox)

    Dim sqlstr As String
    Dim BillingMode As New ADODB.Recordset
    sqlstr = ""
    
    sqlstr = "SELECT Prefix,Description FROM Tbl_BillingModes"
    BillingMode.Open sqlstr, medixmasterconn, adOpenDynamic, adLockOptimistic
    
    cmb1.Clear
    cmb2.Clear
    
    Do Until BillingMode.EOF
        cmb1.AddItem BillingMode!Prefix & " - " & BillingMode!Description
        cmb2.AddItem BillingMode!Prefix & " - " & BillingMode!Description
        BillingMode.MoveNext
    Loop
    BillingMode.Close
    Set BillingMode = Nothing

End Sub

Public Sub BindGroupcompany(cmb1 As ComboBox)
    Dim sqlstr As String
    Dim BillingMode As New ADODB.Recordset
    sqlstr = ""
    sqlstr = "SELECT Distinct Grpcode,GrpName FROM VIEW_GRPCompany order by Grpcode"
    BillingMode.Open sqlstr, medixmasterconn, adOpenDynamic, adLockOptimistic
    cmb1.Clear
    Do Until BillingMode.EOF
        cmb1.AddItem BillingMode!GrpCode & " - " & BillingMode!GrpName
        BillingMode.MoveNext
    Loop
    BillingMode.Close
    Set BillingMode = Nothing
End Sub

Public Sub BindState(cmb1 As ComboBox)
    Dim sqlstr As String
    Dim BillingMode As New ADODB.Recordset
    sqlstr = ""
    sqlstr = "Select [State] as Stat from tblstate order by [State]"
    BillingMode.Open sqlstr, medixmasterconn, adOpenDynamic, adLockOptimistic
    cmb1.Clear
    Do Until BillingMode.EOF
        cmb1.AddItem BillingMode!Stat
        BillingMode.MoveNext
    Loop
    BillingMode.Close
    Set BillingMode = Nothing
End Sub
