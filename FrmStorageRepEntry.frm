VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStorageRepEntry 
   Caption         =   "Storage - Repudiated Case"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   37
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
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
         Left            =   5640
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSMask.MaskEdBox MaskstrDORqst 
         Height          =   285
         Left            =   8640
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox MaskstrDODelivery 
         Height          =   285
         Left            =   8640
         TabIndex        =   7
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox MaskstrDOReturn 
         Height          =   285
         Left            =   8640
         TabIndex        =   8
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.TextBox TxtstrCasNo2 
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
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtstrCasNo 
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
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox CboFileStatus1 
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
         ItemData        =   "FrmStorageRepEntry.frx":0000
         Left            =   3960
         List            =   "FrmStorageRepEntry.frx":0002
         TabIndex        =   4
         Text            =   "A - All"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox CboCarton1 
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
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox CboUser1 
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
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   765
         Width           =   4215
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
         Left            =   3120
         TabIndex        =   45
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Request By"
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
         TabIndex        =   44
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "File Status"
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
         Left            =   3000
         TabIndex        =   43
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Case No"
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
         TabIndex        =   42
         Top             =   255
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "Carton No"
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
         TabIndex        =   41
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Date Of Request"
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
         Left            =   7080
         TabIndex        =   40
         Top             =   255
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Date Of Delivery"
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
         Left            =   7080
         TabIndex        =   39
         Top             =   525
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Date Of Return"
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
         Left            =   7080
         TabIndex        =   38
         Top             =   795
         Width           =   1455
      End
   End
   Begin VB.Frame FrameuPDATE 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   5760
      Width           =   11655
      Begin VB.TextBox TxtCase1 
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
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   11
         Top             =   255
         Width           =   1575
      End
      Begin VB.TextBox TxtRemarks 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8880
         MaxLength       =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CboCarton2 
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
         ItemData        =   "FrmStorageRepEntry.frx":0004
         Left            =   1200
         List            =   "FrmStorageRepEntry.frx":000B
         TabIndex        =   13
         Top             =   525
         Width           =   1575
      End
      Begin VB.ComboBox CboFileStatus2 
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
         ItemData        =   "FrmStorageRepEntry.frx":0012
         Left            =   3960
         List            =   "FrmStorageRepEntry.frx":0014
         TabIndex        =   12
         Text            =   "I - In"
         Top             =   255
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtDORqst 
         Height          =   285
         Left            =   6600
         TabIndex        =   16
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox txtDOD 
         Height          =   285
         Left            =   6600
         TabIndex        =   17
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox txtDOReturn 
         Height          =   285
         Left            =   6600
         TabIndex        =   18
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.ComboBox CboCartonStatus 
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
         Left            =   3960
         TabIndex        =   14
         Top             =   525
         Width           =   1095
      End
      Begin VB.ComboBox CboUser2 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   780
         Width           =   3855
      End
      Begin VB.ComboBox CboTmpStatus 
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
         Left            =   3960
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Carton Status"
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
         Left            =   2880
         TabIndex        =   51
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Case No"
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
         TabIndex        =   49
         Top             =   255
         Width           =   990
      End
      Begin VB.Label Label8 
         Caption         =   "Remarks"
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
         Left            =   8040
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Date Of Return"
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
         Left            =   5280
         TabIndex        =   35
         Top             =   795
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Date Of Delivery"
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
         Left            =   5280
         TabIndex        =   34
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Date Of Request"
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
         Left            =   5280
         TabIndex        =   33
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "Carton No"
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
         TabIndex        =   30
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "File Status"
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
         Left            =   2880
         TabIndex        =   29
         Top             =   280
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "Request By"
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
         TabIndex        =   28
         Top             =   825
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView Lststr 
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Case No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Carton No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Carton Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "File Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Date Register"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Register By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Request By"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Date Request"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Date Delivery"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Date Return"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Duration"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Last Update By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Last Update Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   46
      Top             =   6960
      Width           =   11715
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1080
         TabIndex        =   56
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdListing 
         Caption         =   "&Listing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8520
         TabIndex        =   24
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "&Print File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9160
         TabIndex        =   25
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton CmdAdd 
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
         Height          =   350
         Left            =   4980
         TabIndex        =   31
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5660
         TabIndex        =   20
         Top             =   240
         Width           =   705
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4920
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6360
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
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
         Height          =   350
         Left            =   9960
         TabIndex        =   26
         Top             =   240
         Width           =   735
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
         Height          =   350
         Left            =   10680
         TabIndex        =   27
         Top             =   240
         Width           =   735
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
         Height          =   350
         Left            =   7800
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "St&op"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7080
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblRecChk 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Left            =   4320
         TabIndex        =   54
         Top             =   280
         Width           =   400
      End
      Begin VB.Label Label12 
         Caption         =   "No. of Rec checked"
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
         Left            =   2880
         TabIndex        =   53
         Top             =   280
         Width           =   1455
      End
      Begin VB.Label Label14 
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
         Left            =   1680
         TabIndex        =   48
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "of"
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
         Left            =   840
         TabIndex        =   47
         Top             =   285
         Width           =   255
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   " Storage - Repudiate Case"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "FrmStorageRepEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim bExit As Boolean
Dim EditFlag As Boolean
Dim listloop1 As Integer
Dim Check As Boolean
Dim CheckCnt As Integer
Dim CaseNoArray() As String

Private Sub CboCarton1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCarton1.Text, CboCarton1.SelStart) + Chr(KeyAscii) + Mid(CboCarton1.Text, CboCarton1.SelStart + CboCarton1.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCarton1.ListCount - 1
 
        If NewText = LCase(Left(CboCarton1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCarton1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCarton1.Text = ValidValue
                CboCarton1.SelStart = 0
                CboCarton1.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCarton2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCarton2.Text, CboCarton2.SelStart) + Chr(KeyAscii) + Mid(CboCarton2.Text, CboCarton2.SelStart + CboCarton2.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCarton2.ListCount - 1
 
        If NewText = LCase(Left(CboCarton2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCarton2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCarton2.Text = ValidValue
                CboCarton2.SelStart = 0
                CboCarton2.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCarton2_LostFocus()
Dim QQ As Integer

    QQ = CboCarton2.ListIndex
    CboTmpStatus.ListIndex = QQ
    If Trim(CboTmpStatus) = "M" Then
        CboCartonStatus.ListIndex = 0
    ElseIf Trim(CboTmpStatus) = "S" Then
        CboCartonStatus.ListIndex = 1
    ElseIf Trim(CboTmpStatus) = "R" Then
        CboCartonStatus.ListIndex = 2
    End If
End Sub

Private Sub CboFileStatus1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboFileStatus1.Text, CboFileStatus1.SelStart) + Chr(KeyAscii) + Mid(CboFileStatus1.Text, CboFileStatus1.SelStart + CboFileStatus1.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboFileStatus1.ListCount - 1
 
        If NewText = LCase(Left(CboFileStatus1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboFileStatus1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboFileStatus1.Text = ValidValue
                CboFileStatus1.SelStart = 0
                CboFileStatus1.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub CboFileStatus2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboFileStatus2.Text, CboFileStatus2.SelStart) + Chr(KeyAscii) + Mid(CboFileStatus2.Text, CboFileStatus2.SelStart + CboFileStatus2.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboFileStatus2.ListCount - 1
 
        If NewText = LCase(Left(CboFileStatus2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboFileStatus2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboFileStatus2.Text = ValidValue
                CboFileStatus2.SelStart = 0
                CboFileStatus2.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboUser1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboUser1.Text, CboUser1.SelStart) + Chr(KeyAscii) + Mid(CboUser1.Text, CboUser1.SelStart + CboUser1.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboUser1.ListCount - 1
 
        If NewText = LCase(Left(CboUser1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboUser1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboUser1.Text = ValidValue
                CboUser1.SelStart = 0
                CboUser1.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboUser2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboUser2.Text, CboUser2.SelStart) + Chr(KeyAscii) + Mid(CboUser2.Text, CboUser2.SelStart + CboUser2.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboUser2.ListCount - 1
 
        If NewText = LCase(Left(CboUser1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboUser2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboUser2.Text = ValidValue
                CboUser2.SelStart = 0
                CboUser2.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub ChkScan_Click()
    If ChkScan.Value = 0 Then
        Label15.Visible = True
        TxtstrCasNo2.Visible = True
        ChkScan.Left = 5760
        ChkScan.Top = 220
    Else
        Call EntriesEnable(True)
        Label15.Visible = False
        TxtstrCasNo2.Visible = False
        Lststr.ListItems.Clear
        ChkScan.Left = 3000
        ChkScan.Top = 220
    End If
    TxtstrCasNo.SetFocus
End Sub

Private Sub cmdAdd_Click()
    EditFlag = False
    If CheckCnt > 0 Then
        TxtCase1 = ""
        CboCarton2.Text = ""
        CboCartonStatus.Text = ""
        CboFileStatus2.Text = ""
        CboUser2.Text = ""
        txtDORqst = "__/__/____"
        txtDOD = "__/__/____"
        txtDOReturn = "__/__/____"
        TxtRemarks = ""
        FrameuPDATE.Enabled = True
        CboFileStatus2.SetFocus
        CmdSave.Enabled = True
    Else
        MsgBox "Please checked at least 1 Claim ! ", vbOKOnly + vbCritical
    End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    Check = False
    EditFlag = False
    Call EntriesEnable(True)
    CmdDelete.Enabled = False
    Lststr.ListItems.Clear
    txtTotalRecord = 0
    CheckCnt = 0
    LblRecChk = ""
    ChkScan.Value = 1
    Label15.Visible = False
    TxtstrCasNo2.Visible = False
    CboFileStatus1.Text = "A - All"
    ChkScan.Left = 3000
    ChkScan.Top = 220
    ReDim CaseNoArray(0)
End Sub

Private Sub CmdClear_LostFocus()
    TxtstrCasNo.SetFocus
End Sub

Private Sub CmdDelete_Click()
Dim rst2 As New ADODB.Recordset
Dim sql2 As String
Dim RecordSaved
    
    Screen.MousePointer = vbHourglass
    
    RecordSaved = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then
        sql2 = ""
        sql2 = "Delete From StorageRep Where StrCaseNo = '" & TxtstrCasNo & "'"
        rst2.Open sql2, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
        CmdSave.Enabled = True
        Call ClearForm(Me)
        Call DisableCmdButton(Me)
        Call EntriesEnable(True)
        Call StorageListing
        TxtstrCasNo.SetFocus
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEdit_Click()
    FrameuPDATE.Enabled = False
    If Trim(TxtCase1) <> "" Then
        EditFlag = True
        CmdDelete.Enabled = False
        CmdSave.Enabled = True
        CmdSave.Caption = "&Update"
        FrameuPDATE.Enabled = True
        CboFileStatus2.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function StorageSearchRecord()
Dim strAppend As String

    strAppend = ""
    
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtstrCasNo)) Then
            strAppend = strAppend + " AND CASNumber = '" & RemStr(Trim(TxtstrCasNo)) & "'"
        End If
    Else
        If Len(Trim(TxtstrCasNo)) Then
            strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtstrCasNo)) & "'"
        End If
        If Len(Trim(TxtstrCasNo2)) Then
            strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtstrCasNo2)) & "'"
        End If
    End If

    If Len(Trim(CboCarton1)) Then
        strAppend = strAppend + " AND STRCartonNo = '" & RemStr(Trim(CboCarton1)) & "'"
    End If
    
    If Trim(CboFileStatus1) <> "" And Mid(CboFileStatus1, 1, 1) <> "A" Then
        strAppend = strAppend + " AND STRFileStatus = '" & RemStr(Trim(Mid(CboFileStatus1, 1, 1))) & "'"
    End If
    
    If Len(Trim(CboUser1)) Then
        strAppend = strAppend + " AND STRRequestBy = '" & RemStr(Trim(CboUser1)) & "'"
    End If
    
    If IsDate(MaskstrDORqst) = True Then
        strAppend = strAppend + " And DORequest = '" & Format(MaskstrDORqst, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDODelivery) = True Then
        strAppend = strAppend + " And DODelivery = '" & Format(MaskstrDODelivery, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(MaskstrDOReturn) = True Then
        strAppend = strAppend + " And DOReturn = '" & Format(MaskstrDOReturn, "mm/dd/yyyy") & "'"
    End If
    
    Call StorageListing(strAppend)
    
End Function

Private Sub StorageListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As Integer
Dim Record As ListItem
Dim sql As String
Dim Current As String
    
    rstCount = 0
    Current = 0
    If ChkScan.Value = 0 Then
        Lststr.ListItems.Clear
    End If

    sql = " Select * from View_Storage_CLmRep where 0 = 0 "
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
        
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdAdd.Enabled = True
    Else
    
    Do
        Do Until rst2.EOF
        rstCount = rst2.recordCount
        lblTotal = rstCount
            With rst2
                Set Record = Lststr.ListItems.Add(, , rst2!CasNumber)
                
                Record.SubItems(1) = IIf(Len(!STRCartonNo), !STRCartonNo, "")
                Record.SubItems(2) = IIf(Len(!CRTCartonStatus), !CRTCartonStatus, "")
                Record.SubItems(3) = IIf(Len(!STRFileStatus), !STRFileStatus, "")
                
                If Len(!CaseRegDate) Then
                    Record.SubItems(4) = Format(!CaseRegDate, "dd/mm/yyyy")
                End If
                Record.SubItems(5) = IIf(Len(!CaseOpenBy), !CaseOpenBy, "")
                
                Record.SubItems(6) = IIf(Len(!STRRequestBy), !STRRequestBy, "")
                If Len(!DORequest) Then
                    Record.SubItems(7) = Format(!DORequest, "dd/mm/yyyy")
                End If
                If Len(!DODelivery) Then
                    Record.SubItems(8) = Format(!DODelivery, "dd/mm/yyyy")
                End If
                If Len(!DOReturn) Then
                    Record.SubItems(9) = Format(!DOReturn, "dd/mm/yyyy")
                End If
                If Len(!DOReturn) And Len(!DODelivery) Then
                    Record.SubItems(10) = DateValue(!DOReturn) - DateValue(!DODelivery)
                End If
                Record.SubItems(11) = IIf(Len(!STRRemarks), !STRRemarks, "")
                Record.SubItems(12) = IIf(Len(!LastUpdateBy), !LastUpdateBy, "")
                
                If Len(!LastUpdateDate) Then
                    Record.SubItems(13) = Format(!LastUpdateDate, "dd/mm/yyyy")
                End If
               
                Current = Current + 1
                lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
            Loop
        Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdExit.Enabled = True
End Sub

Private Sub CmdListing_Click()
    
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    If Trim(CboCarton1) = "" Then
        MsgBox "Please select Carton No. ", vbOKOnly + vbCritical, DialogBoxTitle
        CboCarton1.SetFocus
    Else
        Dim objExcel As Excel.Application
        Dim objWorkbook As Excel.Workbook
        Dim objWorksheet As Excel.WORKSHEET
        Dim AA, BB As Integer
        Dim yy, Cnt As Integer
        Dim strsql As String
        Dim Sto As New ADODB.Recordset
        Dim ECol1 As String
        Dim ECol2 As String
        Dim ECol3 As String
           
        Screen.MousePointer = vbHourglass
    
            If objExcel Is Nothing Then
                Set objExcel = CreateObject("Excel.application")
                    objExcel.Interactive = False
                If objExcel.Visible = False Then
                    objExcel.Visible = True
                End If
                objExcel.WindowState = xlMaximized
                Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Storage\StorageTracking.xlt")
                Set objWorksheet = objWorkbook.Sheets(1)
                objExcel.DisplayAlerts = False
                Set objWorksheet = objWorkbook.Worksheets.Item(1)
            End If
            
            If Err.Number > 0 Then
                MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
                GoTo ExitFunction
            End If
            
            On Error GoTo HANDLE_ERROR
            strsql = "select STRCaseNo, STRFileStatus from StorageRep WHERE STRCartonNO  = '" & CboCarton1 & "'"
            Sto.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            'Set objExcel = New excel.Application
            
            If Sto.recordCount <> 0 Then
                RecordSelected = Sto.recordCount
                    objWorksheet.Cells(2, "F") = Format(Date, "MM/DD/YYYY")
                    objWorksheet.Cells(2, "H") = CboCarton1
                    BB = 0
                    yy = 0
                    Cnt = 0
                    ECol1 = Chr(65)
                    ECol2 = Chr(66)
                    ECol3 = Chr(67)
                    
                    Do Until Sto.EOF
                        BB = BB + 1
                        If BB > 50 Then
                            yy = yy + 4
                            BB = 1
                            ECol1 = Chr(65 + yy)
                            ECol2 = Chr(66 + yy)
                            ECol3 = Chr(67 + yy)
                            objWorksheet.Cells(4, ECol1) = "No"
                            objWorksheet.Cells(4, ECol2) = "Case No"
                            objWorksheet.Cells(4, ECol3) = "Checking"
                        End If
                        
                        Cnt = Cnt + 1
                        objWorksheet.Cells(4 + CDbl(BB), ECol1) = Cnt
                        objWorksheet.Cells(4 + CDbl(BB), ECol2) = Sto!StrCaseNo
                        objWorksheet.Cells(4 + CDbl(BB), ECol3) = Sto!STRFileStatus
                        
                        Sto.MoveNext
                    
                Loop
                
            End If
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
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
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
End Sub

Private Sub CmdPrintFile_Click()
    'Call ExportToExcelWorksheet(Lststr)
    Screen.MousePointer = vbHourglass
        
    If Lststr.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(Lststr)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
Dim AA, CASVer As Integer
Dim RecordSaved
Dim CASNo As String
'    FrameuPDATE.Enabled = True
    Screen.MousePointer = vbHourglass
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        If Trim(CboCarton2) = "" Then
            MsgBox "Please Enter Carton Number", vbOKOnly + vbCritical, DialogBoxTitle
            CboCarton2.SetFocus
        ElseIf Trim(Mid(CboFileStatus2, 1, 1)) = "" Then
            MsgBox "Please Enter File Status", vbOKOnly + vbCritical, DialogBoxTitle
            CboFileStatus1.SetFocus
        ElseIf CboUser2 <> "" And txtDORqst = "__/__/____" Then
            MsgBox "Please Enter Date Of Request", vbOKOnly + vbCritical, DialogBoxTitle
            txtDORqst.SetFocus
        ElseIf CboUser2 <> "" And Format(txtDOD, "YYYYMMDD") < Format(txtDORqst, "YYYYMMDD") Then
            MsgBox "Invalid Date ! Date of Delivery must greater than Date of Request ", vbOKOnly + vbCritical, DialogBoxTitle
            txtDOD = "__/__/____"
            txtDOD.SetFocus
        ElseIf CboUser2 <> "" And Format(txtDOReturn, "YYYYMMDD") < Format(txtDOD, "YYYYMMDD") Then
            MsgBox "Invalid Date ! Date of Return must greater than Date of Delivery ", vbOKOnly + vbCritical, DialogBoxTitle
            txtDOReturn = "__/__/____"
            txtDOReturn.SetFocus
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = True Then
                   Erase CaseNoArray
                   ReDim Preserve CaseNoArray(1)
                   CaseNoArray(1) = Lststr.SelectedItem.Text
                End If
                For AA = 0 To UBound(CaseNoArray)
                    If Trim(CaseNoArray(AA)) <> "" Then
                        
                        Call SaveStore(CaseNoArray(AA))
                    End If
                Next AA
                Erase CaseNoArray
                Arraycnt = 0
                ReDim CaseNoArray(0)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                Call ClearStoreInfo
                CheckCnt = 0
                LblRecChk = ""
                If ChkScan.Value = 1 Then
                    TxtstrCasNo = Lststr.SelectedItem.Text
                End If
                Lststr.ListItems.Clear
                Call StorageSearchRecord
            End If
        End If
    'medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
    CmdAdd.Enabled = True
    CmdSave.Enabled = False
    CmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            If ChkScan.Value = 1 Then
                If TxtstrCasNo = "" Then
                    MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
                    TxtstrCasNo.SetFocus
                Else
                    listrecord = Lststr.ListItems.Count
                    If listrecord <> 0 Then
                        
                        For listloop1 = 1 To listrecord
                            Call CaseNoCheck
                        Next listloop1
                    End If
                    If Check = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    ElseIf Check = False Then
                        Call StorageSearchRecord
                    End If
                    Check = False
                    Screen.MousePointer = Default
                    TxtstrCasNo = ""
                    TxtstrCasNo.SetFocus
                End If
                TxtstrCasNo.SetFocus
            Else
                Call StorageSearchRecord
            End If
    'medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
intExit = 1
End Sub

Private Sub SaveStore(CASNo As String)
Dim RecordSaved
Dim SaveSTORAGE As New ADODB.Recordset
Dim strsql As String
    

    strsql = "Select * From StorageRep Where STRCaseNo = '" & Trim(CASNo) & "'"
    
    SaveSTORAGE.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If SaveSTORAGE.recordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
        SaveSTORAGE.Close
        CboCarton2.SetFocus
        CboCarton2.SelStart = 0
        CboCarton2.SelLength = Len(CboCarton2)
        Set SaveSTORAGE = Nothing
        Exit Sub
    Else
        If EditFlag = False Then
            SaveSTORAGE.AddNew
        End If
        With SaveSTORAGE
            !StrCaseNo = CASNo
            !STRCartonNo = ChkStr(CboCarton2)
            !STRFileStatus = Trim(Mid(CboFileStatus2, 1, 1))
            !STRRequestBy = ChkStr(CboUser2)
            If CboUser2 <> "" Then
                !DORequest = txtDORqst
                !DODelivery = IIf(txtDOD <> "__/__/____", txtDOD, Null)
                !DOReturn = IIf(txtDOReturn <> "__/__/____", txtDOReturn, Null)
            Else
                !DODelivery = Null
                !DOReturn = Null
                !DORequest = Null
            End If
            !STRRemarks = ChkStr(TxtRemarks)
            !LastUpdateBy = Logon
            !LastUpdateDate = Date
        End With
        SaveSTORAGE.Update
        
        SaveSTORAGE.Close
        Set SaveSTORAGE = Nothing
        
    End If
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtstrCasNo)) = ChkStr(Lststr.ListItems(listloop1).Text) Then
        Check = True
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Check = False
    CheckCnt = 0
    Label15.Visible = False
    TxtstrCasNo2.Visible = False
    ChkScan.Left = 3000
    ChkScan.Top = 220
    EditFlag = False
    CmdDelete.Enabled = False
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    Call ComboListing
    
    'medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing

End Sub

Private Sub Lststr_Click()

    FrameuPDATE.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = False
    CmdSave.Caption = "Sa&ve"
    
'    EditFlag = False
    TxtCase1 = ""
    CboCarton2.Text = ""
    CboCartonStatus.Text = ""
    CboFileStatus2.Text = ""
    CboUser2.Text = ""
    txtDORqst = "__/__/____"
    txtDOD = "__/__/____"
    txtDOReturn = "__/__/____"
    TxtRemarks = ""
    
    If Lststr.ListItems.Count <> 0 Then
        
        If Trim(Lststr.SelectedItem.SubItems(1)) <> "" Then
            CmdEdit.Enabled = True
            Call ClearStoreInfo
            TxtCase1 = Lststr.SelectedItem.Text
            CboCarton2.Text = Lststr.SelectedItem.SubItems(1)
            If Lststr.SelectedItem.SubItems(2) = "M" Then
                CboCartonStatus.ListIndex = 0
            ElseIf Lststr.SelectedItem.SubItems(2) = "S" Then
                CboCartonStatus.ListIndex = 1
            ElseIf Lststr.SelectedItem.SubItems(2) = "R" Then
                CboCartonStatus.ListIndex = 2
            End If
            If Lststr.SelectedItem.SubItems(3) = "I" Then
                CboFileStatus2.ListIndex = 0
            ElseIf Lststr.SelectedItem.SubItems(3) = "O" Then
                CboFileStatus2.ListIndex = 1
            End If
            CboUser2.Text = Lststr.SelectedItem.SubItems(6)
            txtDORqst = IIf(Len(Lststr.SelectedItem.SubItems(7)), Lststr.SelectedItem.SubItems(7), "__/__/____")
            txtDOD = IIf(Len(Lststr.SelectedItem.SubItems(8)), Lststr.SelectedItem.SubItems(8), "__/__/____")
            txtDOReturn = IIf(Len(Lststr.SelectedItem.SubItems(9)), Lststr.SelectedItem.SubItems(9), "__/__/____")
            TxtRemarks = IIf(Len(Lststr.SelectedItem.SubItems(11)), Lststr.SelectedItem.SubItems(11), "")
            
'        Else
'            MsgBox "Claim has not been stored!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    End If
End Sub

Private Sub Lststr_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Lststr.SortKey = ColumnHeader.Index - 1
    Lststr.Sorted = True
    If Lststr.SortOrder = lvwAscending Then
        Lststr.SortOrder = lvwDescending
    Else
        Lststr.SortOrder = lvwAscending
    End If
End Sub

Private Sub MaskstrDODelivery_LostFocus()
    If IsDate(MaskstrDODelivery) Then
        If Format(MaskstrDODelivery, "YYYYMMDD") < Format(MaskstrDORqst, "YYYYMMDD") Then
            MsgBox "Delivery Date must Greater than Request Date ! ", vbCritical
            MaskstrDODelivery.SetFocus
        End If
    Else
        If MaskstrDODelivery <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskstrDODelivery = "__/__/____"
            MaskstrDODelivery.SetFocus
        End If
    End If
End Sub

Private Sub MaskstrDOReturn_LostFocus()
    If IsDate(MaskstrDOReturn) Then
        If Format(MaskstrDOReturn, "YYYYMMDD") < Format(MaskstrDODelivery, "YYYYMMDD") Then
            MsgBox "Return Date must Greater than Delviery Date ! ", vbCritical
            MaskstrDOReturn.SetFocus
        End If
    Else
        If MaskstrDOReturn <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskstrDOReturn = "__/__/____"
            MaskstrDOReturn.SetFocus
        End If
    End If
End Sub

Private Sub MaskstrDORqst_LostFocus()
    If IsDate(MaskstrDORqst) = False Then
        If MaskstrDORqst <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            MaskstrDORqst = "__/__/____"
            MaskstrDORqst.SetFocus
        End If
    End If
End Sub


Private Sub txtDOD_LostFocus()
    If IsDate(txtDOD) Then
        If Format(txtDOD, "YYYYMMDD") < Format(txtDORqst, "YYYYMMDD") Then
            MsgBox "Delivery Date must Greater than Request Date ! ", vbCritical
            txtDOD.SetFocus
        End If
    Else
        If txtDOD <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOD = "__/__/____"
            txtDOD.SetFocus
        End If
    End If
End Sub

Private Sub txtDOReturn_LostFocus()
    If IsDate(txtDOReturn) Then
        If Format(txtDOReturn, "YYYYMMDD") < Format(txtDOD, "YYYYMMDD") Then
            MsgBox "Return Date must Greater than Delivery Date ! ", vbCritical
            txtDOReturn.SetFocus
        End If
    Else
        If txtDOReturn <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOReturn = "__/__/____"
            txtDOReturn.SetFocus
        End If
    End If
End Sub

Private Sub txtDORqst_LostFocus()
    If IsDate(txtDORqst) = False Then
        If txtDORqst <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDORqst = "__/__/____"
            txtDORqst.SetFocus
        End If
    End If
End Sub

Private Sub TxtstrCasNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtstrCasNo_LostFocus()
    TxtstrCasNo = ChkStr(TxtstrCasNo)
    CmdSearch.Default = False
End Sub

Private Sub EntriesEnable(status As Boolean)
    '----- STORAGE
    TxtstrCasNo.Enabled = status
    CboCarton1.Enabled = status
    CboFileStatus1.Enabled = status
    CboUser1.Enabled = status
    MaskstrDORqst.Enabled = status
    MaskstrDODelivery.Enabled = status
    MaskstrDOReturn.Enabled = status
    
End Sub

Private Sub ComboListing()
Dim USRRec As New ADODB.Recordset
Dim CRTRec As New ADODB.Recordset
Dim CMD As New ADODB.Command
    
    CboFileStatus1.Clear
    CboFileStatus1.AddItem "A - All"
    CboFileStatus1.AddItem "I - In"
    CboFileStatus1.AddItem "O - Out"
    
    CboFileStatus2.Clear
    CboFileStatus2.AddItem "I - In"
    CboFileStatus2.AddItem "O - Out"
    
    CboCartonStatus.Clear
    CboCartonStatus.AddItem "M - Medix"
    CboCartonStatus.AddItem "S - Store"
    CboCartonStatus.AddItem "R - Return"
    
    Set CMD.ActiveConnection = medixmasterconnWS
    CMD.CommandText = "SearchCarton"
    CMD.CommandType = adCmdStoredProc
    CMD.CommandTimeout = 15
    CRTRec.CursorLocation = adUseClient
    CRTRec.CursorType = adOpenForwardOnly
    CRTRec.CacheSize = 100
    Set CRTRec = CMD.Execute
    CboCarton1.Clear
    CboCarton2.Clear
    CboTmpStatus.Clear
    Do Until CRTRec.EOF
        With CRTRec
            CboCarton1.AddItem !CRTCartonNo
            CboCarton2.AddItem !CRTCartonNo
            CboTmpStatus.AddItem !CRTCartonStatus
            CRTRec.MoveNext
        End With
    Loop
    
    CRTRec.Close
    Set CRTRec = Nothing
    
    Set CMD.ActiveConnection = medixmasterconnMaint
    CMD.CommandText = "SearchUser"
    CMD.CommandType = adCmdStoredProc
    CMD.CommandTimeout = 15
    USRRec.CursorLocation = adUseClient
    USRRec.CursorType = adOpenForwardOnly
    USRRec.CacheSize = 100
    Set USRRec = CMD.Execute
    CboUser1.Clear
    CboUser2.Clear
    Do Until USRRec.EOF
        With USRRec
            CboUser1.AddItem IIf(Len(!USRName), !USRName, "")
            CboUser2.AddItem IIf(Len(!USRName), !USRName, "")
            USRRec.MoveNext
        End With
    Loop
    USRRec.Close
    Set USRRec = Nothing
End Sub

Private Sub TxtRemarks_LostFocus()
    TxtRemarks = ChkStr(TxtRemarks)
End Sub

Private Sub Lststr_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CmdSave.Enabled = True
    FrameuPDATE.Enabled = True
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Call ProcessItemChecked(Item)
  '  medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
End Sub

Private Function ProcessItemChecked(lstRec As ListItem, Optional Partial As Boolean)
Dim BB, newbound  As Integer
Dim tempArray() As String
Dim listloop As Integer
Dim ListCount As Integer
Dim TotalAmt As Variant
Dim strsql As String
Dim StoRec As New ADODB.Recordset
Dim CaseNo As String
Dim tmpSelCasNo As String

Static Arraycnt As Integer
    
    CaseNo = lstRec.Text
    strsql = "Select * From StorageRep where STRCASeNo = '" & Trim(CaseNo) & "'"
    StoRec.Open strsql, medixmasterconnWS, adOpenKeyset, adLockPessimistic

    If lstRec.Checked = True Then
        If StoRec.recordCount = 0 Then
            ReDim Preserve CaseNoArray(Arraycnt)
            CaseNoArray(Arraycnt) = lstRec.Text
            Arraycnt = Arraycnt + 1
            CheckCnt = CheckCnt + 1
            LblRecChk = CheckCnt
        Else
            MsgBox "Case has been stored!", vbCritical
            lstRec.Checked = False
        End If
    Else
        Arraycnt = Arraycnt - 1
        CheckCnt = CheckCnt - 1
        LblRecChk = IIf(CheckCnt < 0, 0, CheckCnt)

        If UBound(CaseNoArray) <> 0 Then
            For BB = 0 To UBound(CaseNoArray)
                tmpSelCasNo = lstRec.Text
                If tmpSelCasNo <> CaseNoArray(BB) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = CaseNoArray(BB)
                    newbound = newbound + 1
                End If
            Next BB
            
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim CaseNoArray(UBound(tempArray))
            For BB = 0 To UBound(tempArray)
                CaseNoArray(BB) = tempArray(BB)
            Next BB
            Erase tempArray
            ReDim tempArray(0)
        End If
    End If
    StoRec.Close
    Set StoRec = Nothing
End Function

Private Sub ClearStoreInfo()
    CboCarton1.Text = ""
    CboFileStatus1.Text = ""
    CboUser1.Text = ""
    MaskstrDORqst = "__/__/____"
    MaskstrDODelivery = "__/__/____"
    MaskstrDOReturn = "__/__/____"
    TxtCase1 = ""
    CboCarton2.Text = ""
    CboCartonStatus.Text = ""
    CboFileStatus2.Text = ""
    CboUser2.Text = ""
    txtDORqst = "__/__/____"
    txtDOD = "__/__/____"
    txtDOReturn = "__/__/____"
    TxtRemarks = ""

End Sub


