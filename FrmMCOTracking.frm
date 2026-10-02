VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMCOTracking 
   Caption         =   "MCO Fees Tracking"
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
   Begin MSComctlLib.ListView lstMCOList 
      Height          =   720
      Left            =   0
      TabIndex        =   117
      Top             =   6570
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1270
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
      NumItems        =   26
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   124
      Top             =   7200
      Width           =   11775
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7200
         TabIndex        =   167
         Top             =   200
         Width           =   720
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
         Height          =   435
         Left            =   7920
         TabIndex        =   130
         Top             =   200
         Width           =   600
      End
      Begin VB.CommandButton cmdSearch 
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
         Height          =   435
         Left            =   6480
         TabIndex        =   129
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   435
         Left            =   11040
         TabIndex        =   128
         Top             =   200
         Width           =   600
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "S&eletion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8520
         TabIndex        =   127
         Top             =   200
         Width           =   960
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9480
         TabIndex        =   126
         Top             =   200
         Width           =   600
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   435
         Left            =   10080
         TabIndex        =   125
         Top             =   200
         Width           =   975
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   135
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reports Type"
      Height          =   2175
      Left            =   9240
      TabIndex        =   118
      Top             =   480
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "MCO Fee Inv."
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "MCO Fee Not Inv."
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "MCO Fee Refund"
         Height          =   255
         Left            =   120
         TabIndex        =   121
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "MCO Fee Not Refund"
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton Option5 
         Caption         =   "MCO Fee Net"
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Selection Criteria"
      Height          =   6045
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   9255
      Begin VB.TextBox TxtPolYear2 
         Height          =   300
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   73
         Top             =   4080
         Width           =   1120
      End
      Begin VB.TextBox TxtPCode2 
         Height          =   300
         Left            =   5880
         MaxLength       =   5
         TabIndex        =   72
         Top             =   3840
         Width           =   1120
      End
      Begin VB.TextBox TxtEndtBatch2 
         Height          =   300
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   71
         Top             =   3585
         Width           =   1120
      End
      Begin VB.TextBox TxtBatch2 
         Height          =   300
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   69
         Top             =   3060
         Width           =   1120
      End
      Begin VB.TextBox TxtAgeband2 
         Height          =   300
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   63
         Top             =   1440
         Width           =   1120
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   300
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   62
         Top             =   1170
         Width           =   1120
      End
      Begin VB.TextBox TxtPolYear1 
         Height          =   300
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   61
         Top             =   4095
         Width           =   1120
      End
      Begin VB.TextBox TxtPCode1 
         Height          =   300
         Left            =   4320
         MaxLength       =   5
         TabIndex        =   60
         Top             =   3840
         Width           =   1120
      End
      Begin VB.TextBox TxtEndtBatch1 
         Height          =   300
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   59
         Top             =   3585
         Width           =   1120
      End
      Begin VB.TextBox TxtBatch1 
         Height          =   300
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   57
         Top             =   3060
         Width           =   1120
      End
      Begin VB.TextBox TxtAgeband1 
         Height          =   300
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   51
         Top             =   1440
         Width           =   1120
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   300
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   50
         Top             =   1170
         Width           =   1120
      End
      Begin VB.TextBox txtPolNo 
         Height          =   300
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   49
         Top             =   4095
         Width           =   1450
      End
      Begin VB.ComboBox cboPrincipalType 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   900
         Width           =   2700
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   4320
         TabIndex        =   47
         Top             =   630
         Width           =   2700
      End
      Begin VB.ComboBox cboPayorCode 
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
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   345
         Width           =   2700
      End
      Begin VB.TextBox TxtTakeCode1 
         Height          =   300
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   45
         Top             =   3765
         Width           =   1450
      End
      Begin VB.ComboBox CboState 
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
         TabIndex        =   44
         Top             =   3480
         Width           =   1450
      End
      Begin VB.ComboBox CboOCP 
         Height          =   315
         Left            =   1200
         TabIndex        =   43
         Top             =   3195
         Width           =   1450
      End
      Begin VB.ComboBox CboRenewal 
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
         TabIndex        =   42
         Top             =   2940
         Width           =   1450
      End
      Begin VB.ComboBox CboStatus 
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
         TabIndex        =   41
         Text            =   "A- Active"
         Top             =   2655
         Width           =   1450
      End
      Begin VB.ComboBox cboDataStatus 
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
         TabIndex        =   40
         Top             =   2370
         Width           =   1450
      End
      Begin VB.ComboBox CboSex 
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
         TabIndex        =   39
         Top             =   2040
         Width           =   1450
      End
      Begin VB.ComboBox CboRace 
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
         TabIndex        =   38
         Top             =   1770
         Width           =   1450
      End
      Begin VB.ComboBox CboMarital 
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
         TabIndex        =   37
         Top             =   1485
         Width           =   1450
      End
      Begin VB.ComboBox CboAdultChild 
         Height          =   315
         Left            =   1200
         TabIndex        =   36
         Top             =   1200
         Width           =   1450
      End
      Begin VB.ComboBox CboMbmType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   930
         Width           =   1450
      End
      Begin VB.ComboBox cboINSCode 
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
         TabIndex        =   34
         Top             =   645
         Width           =   1450
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   33
         Text            =   "S"
         Top             =   360
         Width           =   1450
      End
      Begin VB.CheckBox ChkRenew 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   32
         Top             =   2985
         Width           =   375
      End
      Begin VB.CheckBox chkExpDate 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   31
         Top             =   1965
         Width           =   375
      End
      Begin VB.CheckBox chkDOR 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   30
         Top             =   2505
         Width           =   375
      End
      Begin VB.CheckBox chkBordxDate 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   29
         Top             =   2805
         Width           =   375
      End
      Begin VB.CheckBox chkBatch 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   28
         Top             =   3060
         Width           =   375
      End
      Begin VB.CheckBox chkPCode 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   27
         Top             =   3855
         Width           =   375
      End
      Begin VB.CheckBox chkPolYear 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   26
         Top             =   4125
         Width           =   375
      End
      Begin VB.CheckBox ChkEffDate 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   25
         Top             =   1695
         Width           =   375
      End
      Begin VB.CheckBox ChkOCP 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   24
         Top             =   3225
         Width           =   495
      End
      Begin VB.CheckBox ChkState 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   23
         Top             =   3495
         Width           =   495
      End
      Begin VB.CheckBox ChkDataStatus 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   22
         Top             =   2400
         Width           =   375
      End
      Begin VB.CheckBox ChkStatus 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   21
         Top             =   2685
         Width           =   375
      End
      Begin VB.CheckBox ChkPlanCode 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   20
         Top             =   1155
         Width           =   375
      End
      Begin VB.CheckBox ChkAgeband 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   19
         Top             =   1425
         Width           =   375
      End
      Begin VB.CheckBox ChkDateJoined 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   18
         Top             =   2235
         Width           =   375
      End
      Begin VB.CheckBox ChkEndtBatch 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   17
         Top             =   3570
         Width           =   375
      End
      Begin VB.CheckBox ChkTakeCode 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   16
         Top             =   3770
         Width           =   375
      End
      Begin VB.CheckBox ChkPolNo 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   15
         Top             =   4125
         Width           =   375
      End
      Begin VB.CheckBox ChkHltType 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox ChkInsType 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   13
         Top             =   650
         Width           =   375
      End
      Begin VB.CheckBox ChkMemType 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   12
         Top             =   940
         Width           =   375
      End
      Begin VB.CheckBox ChkAdultChild 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   11
         Top             =   1220
         Width           =   375
      End
      Begin VB.CheckBox ChkMarital 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   10
         Top             =   1500
         Width           =   375
      End
      Begin VB.CheckBox ChkRace 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   9
         Top             =   1790
         Width           =   375
      End
      Begin VB.CheckBox ChkSex 
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
         Height          =   280
         Left            =   2715
         TabIndex        =   8
         Top             =   2070
         Width           =   375
      End
      Begin VB.CheckBox chkEndtBordx 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   7
         Top             =   3330
         Width           =   375
      End
      Begin VB.CheckBox ChkPayor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   280
         Left            =   7050
         TabIndex        =   6
         Top             =   350
         Width           =   375
      End
      Begin VB.CheckBox ChkGrpCom 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   5
         Top             =   620
         Width           =   375
      End
      Begin VB.CheckBox ChkSuppType 
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
         Height          =   280
         Left            =   7050
         TabIndex        =   4
         Top             =   890
         Width           =   375
      End
      Begin MSMask.MaskEdBox TxtEffDate1 
         Height          =   300
         Left            =   4320
         TabIndex        =   52
         Top             =   1710
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtExpDate1 
         Height          =   300
         Left            =   4320
         TabIndex        =   53
         Top             =   1980
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoined1 
         Height          =   300
         Left            =   4320
         TabIndex        =   54
         Top             =   2250
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR1 
         Height          =   300
         Left            =   4320
         TabIndex        =   55
         Top             =   2505
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordx1 
         Height          =   300
         Left            =   4320
         TabIndex        =   56
         Top             =   2805
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtEndtBordx1 
         Height          =   300
         Left            =   4320
         TabIndex        =   58
         Top             =   3330
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtEffDate2 
         Height          =   300
         Left            =   5880
         TabIndex        =   64
         Top             =   1710
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtExpDate2 
         Height          =   300
         Left            =   5880
         TabIndex        =   65
         Top             =   1980
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoined2 
         Height          =   300
         Left            =   5880
         TabIndex        =   66
         Top             =   2250
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR2 
         Height          =   300
         Left            =   5880
         TabIndex        =   67
         Top             =   2505
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordx2 
         Height          =   300
         Left            =   5880
         TabIndex        =   68
         Top             =   2805
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtEndtBordx2 
         Height          =   300
         Left            =   5880
         TabIndex        =   70
         Top             =   3330
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame4 
         Height          =   1470
         Left            =   120
         TabIndex        =   136
         Top             =   4440
         Width           =   9015
         Begin VB.CheckBox Check4 
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
            Height          =   280
            Left            =   7080
            TabIndex        =   166
            Top             =   1080
            Width           =   375
         End
         Begin VB.CheckBox Check3 
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
            Height          =   280
            Left            =   7080
            TabIndex        =   165
            Top             =   780
            Width           =   375
         End
         Begin VB.CheckBox Check2 
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
            Height          =   280
            Left            =   7080
            TabIndex        =   164
            Top             =   480
            Width           =   375
         End
         Begin VB.CheckBox Check1 
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
            Height          =   280
            Left            =   7080
            TabIndex        =   163
            Top             =   200
            Width           =   375
         End
         Begin VB.ComboBox Combo1 
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
            TabIndex        =   153
            Top             =   180
            Width           =   1450
         End
         Begin VB.TextBox Text4 
            Height          =   300
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   146
            Top             =   780
            Width           =   1120
         End
         Begin VB.TextBox Text3 
            Height          =   300
            Left            =   5880
            MaxLength       =   2
            TabIndex        =   145
            Top             =   780
            Width           =   1120
         End
         Begin VB.TextBox Text2 
            Height          =   300
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   138
            Top             =   180
            Width           =   1120
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Left            =   5880
            MaxLength       =   2
            TabIndex        =   137
            Top             =   180
            Width           =   1120
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   300
            Left            =   4320
            TabIndex        =   147
            Top             =   1070
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   300
            Left            =   5880
            TabIndex        =   148
            Top             =   1070
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox Combo2 
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
            TabIndex        =   155
            Top             =   480
            Width           =   1450
         End
         Begin VB.ComboBox Combo3 
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
            TabIndex        =   157
            Top             =   780
            Width           =   1450
         End
         Begin VB.ComboBox Combo4 
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
            TabIndex        =   159
            Top             =   1070
            Width           =   1450
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   300
            Left            =   4320
            TabIndex        =   139
            Top             =   480
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   300
            Left            =   5880
            TabIndex        =   140
            Top             =   480
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label60 
            Caption         =   "CN Print Ind."
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
            TabIndex        =   160
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label59 
            Caption         =   "Inv. Print Ind."
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
            TabIndex        =   158
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label58 
            Caption         =   "CN Ind."
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
            TabIndex        =   156
            Top             =   550
            Width           =   975
         End
         Begin VB.Label Label57 
            Caption         =   "Inv. Ind."
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
            TabIndex        =   154
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label56 
            Caption         =   "CN Date"
            Height          =   255
            Left            =   3315
            TabIndex        =   152
            ToolTipText     =   "Date of Registered"
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label55 
            Caption         =   "To"
            Height          =   255
            Left            =   5520
            TabIndex        =   151
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label54 
            Caption         =   "To"
            Height          =   255
            Left            =   5520
            TabIndex        =   150
            Top             =   900
            Width           =   375
         End
         Begin VB.Label Label53 
            Caption         =   "CN No"
            Height          =   255
            Left            =   3315
            TabIndex        =   149
            Top             =   900
            Width           =   975
         End
         Begin VB.Label Label28 
            Caption         =   "Inv. Date"
            Height          =   255
            Left            =   3315
            TabIndex        =   144
            ToolTipText     =   "Date of Registered"
            Top             =   550
            Width           =   975
         End
         Begin VB.Label Label27 
            Caption         =   "To"
            Height          =   255
            Left            =   5520
            TabIndex        =   143
            Top             =   550
            Width           =   495
         End
         Begin VB.Label Label26 
            Caption         =   "To"
            Height          =   255
            Left            =   5520
            TabIndex        =   142
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Inv. No"
            Height          =   255
            Left            =   3315
            TabIndex        =   141
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label40 
         Caption         =   "Policy Status"
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
         TabIndex        =   116
         Top             =   2715
         Width           =   1020
      End
      Begin VB.Label Label73 
         Caption         =   "Payor "
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
         Left            =   3315
         TabIndex        =   115
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Sex"
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
         TabIndex        =   114
         Top             =   2115
         Width           =   690
      End
      Begin VB.Label Label24 
         Caption         =   "Marital"
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
         TabIndex        =   113
         Top             =   1545
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Renewal Ind"
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
         TabIndex        =   112
         Top             =   3000
         Width           =   930
      End
      Begin VB.Label Label21 
         Caption         =   "Policy Year"
         Height          =   255
         Left            =   3315
         TabIndex        =   111
         Top             =   4155
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   110
         Top             =   4155
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "Age Band"
         Height          =   255
         Left            =   3315
         TabIndex        =   109
         Top             =   1455
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   108
         Top             =   1470
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Post Code"
         Height          =   255
         Left            =   3315
         TabIndex        =   107
         Top             =   3885
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   106
         Top             =   3900
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Batch"
         Height          =   255
         Left            =   3315
         TabIndex        =   105
         Top             =   3105
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   104
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Bordx"
         Height          =   255
         Left            =   3315
         TabIndex        =   103
         Top             =   2850
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   102
         Top             =   2850
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   101
         Top             =   2565
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   100
         Top             =   2010
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   99
         Top             =   1755
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "DOR"
         Height          =   255
         Left            =   3315
         TabIndex        =   98
         ToolTipText     =   "Date of Discharged"
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   3315
         TabIndex        =   97
         ToolTipText     =   "Date of Admission"
         Top             =   2010
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   3315
         TabIndex        =   96
         ToolTipText     =   "Date of Registered"
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   3315
         TabIndex        =   95
         Top             =   945
         Width           =   975
      End
      Begin VB.Label Label74 
         Caption         =   "Ins Type"
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
         TabIndex        =   94
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   240
         TabIndex        =   93
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Hlt Type"
         Height          =   255
         Left            =   240
         TabIndex        =   92
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label72 
         Caption         =   "Data Status"
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
         TabIndex        =   91
         Top             =   2415
         Width           =   1020
      End
      Begin VB.Label Label30 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   90
         Top             =   3405
         Width           =   375
      End
      Begin VB.Label Label31 
         Caption         =   "Endt. Bordx"
         Height          =   255
         Left            =   3315
         TabIndex        =   89
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   88
         Top             =   3645
         Width           =   375
      End
      Begin VB.Label Label33 
         Caption         =   "Endt. Batch"
         Height          =   255
         Left            =   3315
         TabIndex        =   87
         Top             =   3615
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Adult/Child"
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label34 
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
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   3555
         Width           =   810
      End
      Begin VB.Label Label35 
         Caption         =   "Race"
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
         TabIndex        =   84
         Top             =   1830
         Width           =   690
      End
      Begin VB.Label Label36 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   83
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label37 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   3315
         TabIndex        =   82
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "Occupation"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   3270
         Width           =   975
      End
      Begin VB.Label Label42 
         Caption         =   "Group Co."
         Height          =   255
         Left            =   3315
         TabIndex        =   80
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label43 
         Caption         =   "Date Joined"
         Height          =   255
         Left            =   3315
         TabIndex        =   79
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label45 
         Caption         =   "To"
         Height          =   255
         Left            =   5520
         TabIndex        =   78
         Top             =   2295
         Width           =   375
      End
      Begin VB.Label Label47 
         Caption         =   "Take Code"
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label50 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   4155
         Width           =   855
      End
      Begin VB.Label Label51 
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
         Left            =   2760
         TabIndex        =   75
         Top             =   240
         Width           =   375
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
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   7080
         TabIndex        =   74
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11805
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MCO Fees Tracking"
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
         TabIndex        =   2
         Top             =   90
         Width           =   3285
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report Selection"
      Height          =   3735
      Left            =   9240
      TabIndex        =   0
      Top             =   2760
      Width           =   2535
      Begin VB.OptionButton Option11 
         Caption         =   "Listed By Credit Note"
         Height          =   255
         Left            =   240
         TabIndex        =   162
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Listed By End. Bordx."
         Height          =   255
         Left            =   240
         TabIndex        =   161
         Top             =   2040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Summarised By Ins. Type"
         Height          =   255
         Left            =   240
         TabIndex        =   134
         Top             =   1620
         Width           =   2175
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Listed By Inv. No"
         Height          =   255
         Left            =   240
         TabIndex        =   133
         Top             =   1260
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Listed By Bordx."
         Height          =   255
         Left            =   240
         TabIndex        =   132
         Top             =   870
         Width           =   1695
      End
      Begin VB.OptionButton Option6 
         Caption         =   "By Membership"
         Height          =   255
         Left            =   240
         TabIndex        =   131
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmMCOTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call ListViewHeader1
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

Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub Option1_Click()
    Option6.value = True
    Call ListViewHeader1
    Option2.Visible = True
    Option7.Visible = True
    Option8.Visible = True
    Option9.Visible = True
    Option9.Top = 1620
    Option11.Visible = False
End Sub

Private Sub Option2_Click()
    Option6.value = True
    Call ListViewHeader1
    Option7.Visible = True
    Option8.Visible = False
    Option9.Visible = True
    Option9.Top = 1260
    Option10.Visible = False
    Option11.Visible = False
End Sub

Private Sub Option3_Click()
    Option6.value = True
    Call ListViewHeader1
    Option7.Visible = False
    Option9.Top = 1620
    Option9.Visible = True
    Option10.Visible = True
    Option10.Top = 870
    Option8.Visible = False
    Option11.Visible = True
    Option11.Top = 1260
End Sub

Private Sub Option4_Click()
    Option6.value = True
    Call ListViewHeader1
    Option7.Visible = True
    Option8.Visible = False
    Option9.Top = 1260
    Option9.Visible = True
    Option10.Visible = False
    Option11.Visible = False
End Sub

Private Sub Option5_Click()
    Option6.value = True
    Call ListViewHeader1
    Option7.Visible = False
    Option8.Visible = False
    Option9.Visible = True
    Option9.Top = 870
    Option10.Visible = False
    Option11.Visible = False
End Sub

Private Sub Option6_Click()
    If Option1.value = True Then
        Call ListViewHeader1
    ElseIf Option2.value = True Then
        Call ListViewHeader1
    ElseIf Option3.value = True Then
        Call ListViewHeader5
    ElseIf Option4.value = True Then
        Call ListViewHeader5
    ElseIf Option5.value = True Then
        Call ListViewHeader5
    End If
End Sub

Private Sub Option7_Click()
    If Option1.value = True Then
        Call ListViewHeader2
    ElseIf Option2.value = True Then
        Call ListViewHeader2
    ElseIf Option4.value = True Then
        Call ListViewHeader6
    End If
End Sub

Private Sub Option8_Click()
    If Option1.value = True Then
        Call ListViewHeader3
    End If
End Sub

Private Sub Option9_Click()
    If Option1.value = True Then
        Call ListViewHeader4
    ElseIf Option2.value = True Then
        Call ListViewHeader4
    ElseIf Option3.value = True Then
        Call ListViewHeader8
    ElseIf Option4.value = True Then
        Call ListViewHeader8
    ElseIf Option5.value = True Then
        Call ListViewHeader9
    End If
End Sub

Private Sub Option10_Click()
    If Option3.value = True Then
        Call ListViewHeader6
    End If
End Sub

Private Sub Option11_Click()
    If Option3.value = True Then
        Call ListViewHeader7
    End If
End Sub

' ================= ListView Header ============================
Private Sub ListViewHeader1()
    lstMCOList.ColumnHeaders(1).Text = "Bordx Date"
    lstMCOList.ColumnHeaders(2).Text = "Bordx Batch"
    lstMCOList.ColumnHeaders(2).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(3).Text = "Member No"
    lstMCOList.ColumnHeaders(3).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(4).Text = "Member Name"
    lstMCOList.ColumnHeaders(4).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(5).Text = "Inv. No"
    lstMCOList.ColumnHeaders(5).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(6).Text = "Inv. Date"
    lstMCOList.ColumnHeaders(6).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(7).Text = "HLT Type"
    lstMCOList.ColumnHeaders(7).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(8).Text = "Adult/Child"
    lstMCOList.ColumnHeaders(8).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(9).Text = "Eff. Date"
    lstMCOList.ColumnHeaders(9).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(10).Text = "Exp Date"
    lstMCOList.ColumnHeaders(10).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(11).Text = "MCO Basic"
    lstMCOList.ColumnHeaders(11).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(12).Text = "MCO Gross"
    lstMCOList.ColumnHeaders(12).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader2()
    lstMCOList.ColumnHeaders(1).Text = "Payor"
    lstMCOList.ColumnHeaders(2).Text = "Bordx Date"
    lstMCOList.ColumnHeaders(2).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(3).Text = "Batch No"
    lstMCOList.ColumnHeaders(3).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(4).Text = "Invoice No"
    lstMCOList.ColumnHeaders(4).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(5).Text = "Invoice Date"
    lstMCOList.ColumnHeaders(5).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(6).Text = "MCO Gross"
    lstMCOList.ColumnHeaders(6).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader3()
    lstMCOList.ColumnHeaders(1).Text = "Payor"
    lstMCOList.ColumnHeaders(2).Text = "Invoice No"
    lstMCOList.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(3).Text = "Invoice Date"
    lstMCOList.ColumnHeaders(3).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(4).Text = "Bordx Date"
    lstMCOList.ColumnHeaders(4).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(5).Text = "Batch No"
    lstMCOList.ColumnHeaders(5).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(6).Text = "MCO Gross"
    lstMCOList.ColumnHeaders(6).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader4()
    lstMCOList.ColumnHeaders(1).Text = "Payor"
    lstMCOList.ColumnHeaders(2).Text = "Bordx Date"
    lstMCOList.ColumnHeaders(2).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(3).Text = "Batch No"
    lstMCOList.ColumnHeaders(3).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(4).Text = "Invoice No"
    lstMCOList.ColumnHeaders(4).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(5).Text = "Invoice Date"
    lstMCOList.ColumnHeaders(5).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(6).Text = "Prin. IA"
    lstMCOList.ColumnHeaders(6).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(7).Text = "Prin. IC"
    lstMCOList.ColumnHeaders(7).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(8).Text = "Prin. F"
    lstMCOList.ColumnHeaders(8).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(9).Text = "Prin. GA"
    lstMCOList.ColumnHeaders(9).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(10).Text = "Prin. GC"
    lstMCOList.ColumnHeaders(10).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(11).Text = "Prin. H"
    lstMCOList.ColumnHeaders(11).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(12).Text = "Print. Total"
    lstMCOList.ColumnHeaders(12).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(13).Text = "Mem. IA"
    lstMCOList.ColumnHeaders(13).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(14).Text = "Mem. IC"
    lstMCOList.ColumnHeaders(14).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(15).Text = "Mem. F"
    lstMCOList.ColumnHeaders(15).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(16).Text = "Mem. GA"
    lstMCOList.ColumnHeaders(16).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(17).Text = "Mem. GC"
    lstMCOList.ColumnHeaders(17).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(18).Text = "Mem. H"
    lstMCOList.ColumnHeaders(18).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(19).Text = "Mem. Total"
    lstMCOList.ColumnHeaders(19).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(20).Text = "MCO IA"
    lstMCOList.ColumnHeaders(20).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(21).Text = "MCO IC"
    lstMCOList.ColumnHeaders(21).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(22).Text = "MCO F"
    lstMCOList.ColumnHeaders(22).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(23).Text = "MCO GA"
    lstMCOList.ColumnHeaders(23).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(24).Text = "MCO GC"
    lstMCOList.ColumnHeaders(24).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(25).Text = "MCO H"
    lstMCOList.ColumnHeaders(25).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(26).Text = "MCO Total"
    lstMCOList.ColumnHeaders(26).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader5()
    lstMCOList.ColumnHeaders(1).Text = "Member No"
    lstMCOList.ColumnHeaders(2).Text = "Member Name"
    lstMCOList.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(3).Text = "Credit Note"
    lstMCOList.ColumnHeaders(3).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(4).Text = "CN Date"
    lstMCOList.ColumnHeaders(4).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(5).Text = "End. Bordx Date"
    lstMCOList.ColumnHeaders(5).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(6).Text = "End. Bordx Batch"
    lstMCOList.ColumnHeaders(6).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(7).Text = "Cancel Date"
    lstMCOList.ColumnHeaders(7).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(8).Text = "Bordx Date"
    lstMCOList.ColumnHeaders(8).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(9).Text = "Bordx Batch"
    lstMCOList.ColumnHeaders(9).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(10).Text = "Invoice No"
    lstMCOList.ColumnHeaders(10).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(11).Text = "Invoice Date"
    lstMCOList.ColumnHeaders(11).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(12).Text = "HLT Type"
    lstMCOList.ColumnHeaders(12).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(13).Text = "Adult/Child"
    lstMCOList.ColumnHeaders(13).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(14).Text = "Eff. Date"
    lstMCOList.ColumnHeaders(14).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(15).Text = "Exp. Date"
    lstMCOList.ColumnHeaders(15).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(16).Text = "MCO Basic"
    lstMCOList.ColumnHeaders(16).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(17).Text = "MCO Gross"
    lstMCOList.ColumnHeaders(17).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(18).Text = "MCO Net"
    lstMCOList.ColumnHeaders(18).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(19).Text = "MCO Refund"
    lstMCOList.ColumnHeaders(19).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader6()
    lstMCOList.ColumnHeaders(1).Text = "Payor"
    lstMCOList.ColumnHeaders(2).Text = "End. Bordx Date"
    lstMCOList.ColumnHeaders(2).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(3).Text = "Batch No"
    lstMCOList.ColumnHeaders(3).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(4).Text = "Credit Note"
    lstMCOList.ColumnHeaders(4).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(5).Text = "CN Date"
    lstMCOList.ColumnHeaders(5).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(6).Text = "MCO Gross"
    lstMCOList.ColumnHeaders(6).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(7).Text = "MCO Net"
    lstMCOList.ColumnHeaders(7).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(8).Text = "MCO Refund"
    lstMCOList.ColumnHeaders(8).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader7()
    lstMCOList.ColumnHeaders(1).Text = "Payor"
    lstMCOList.ColumnHeaders(2).Text = "Credit Note"
    lstMCOList.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(3).Text = "CN Date"
    lstMCOList.ColumnHeaders(3).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(4).Text = "End. Bordx Date"
    lstMCOList.ColumnHeaders(4).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(5).Text = "Batch No"
    lstMCOList.ColumnHeaders(5).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(6).Text = "MCO Gross"
    lstMCOList.ColumnHeaders(6).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(7).Text = "MCO Net"
    lstMCOList.ColumnHeaders(7).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(8).Text = "MCO Refund"
    lstMCOList.ColumnHeaders(8).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader8()
    lstMCOList.ColumnHeaders(1).Text = "Payor"
    lstMCOList.ColumnHeaders(2).Text = "Credit Note"
    lstMCOList.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(3).Text = "CN Date"
    lstMCOList.ColumnHeaders(3).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(4).Text = "Endt. Bordx Date"
    lstMCOList.ColumnHeaders(4).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(5).Text = "Batch No"
    lstMCOList.ColumnHeaders(5).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(6).Text = "Prin. IA"
    lstMCOList.ColumnHeaders(6).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(7).Text = "Prin. IC"
    lstMCOList.ColumnHeaders(7).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(8).Text = "Prin. F"
    lstMCOList.ColumnHeaders(8).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(9).Text = "Prin. GA"
    lstMCOList.ColumnHeaders(9).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(10).Text = "Prin. GC"
    lstMCOList.ColumnHeaders(10).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(11).Text = "Prin. H"
    lstMCOList.ColumnHeaders(11).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(12).Text = "Prin. Total"
    lstMCOList.ColumnHeaders(12).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(13).Text = "Mem. IA"
    lstMCOList.ColumnHeaders(13).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(14).Text = "Mem. IC"
    lstMCOList.ColumnHeaders(14).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(15).Text = "Mem. F"
    lstMCOList.ColumnHeaders(15).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(16).Text = "Mem. GA"
    lstMCOList.ColumnHeaders(16).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(17).Text = "Mem. GC"
    lstMCOList.ColumnHeaders(17).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(18).Text = "Mem. H"
    lstMCOList.ColumnHeaders(18).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(19).Text = "Mem. Total"
    lstMCOList.ColumnHeaders(19).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(20).Text = "MCO Refd. IA"
    lstMCOList.ColumnHeaders(20).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(21).Text = "MCO Refd. IC"
    lstMCOList.ColumnHeaders(21).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(22).Text = "MCO Refd. F"
    lstMCOList.ColumnHeaders(22).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(23).Text = "MCO Refd. GA"
    lstMCOList.ColumnHeaders(23).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(24).Text = "MCO Refd. GC"
    lstMCOList.ColumnHeaders(24).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(25).Text = "MCO Refd. H"
    lstMCOList.ColumnHeaders(25).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(26).Text = "MCO Refd. Total"
    lstMCOList.ColumnHeaders(26).Alignment = lvwColumnRight
End Sub

Private Sub ListViewHeader9()
    lstMCOList.ColumnHeaders(1).Text = "Payor"
    lstMCOList.ColumnHeaders(2).Text = "Credit Note"
    lstMCOList.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(3).Text = "CN Date"
    lstMCOList.ColumnHeaders(3).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(4).Text = "Endt. Bordx Date"
    lstMCOList.ColumnHeaders(4).Alignment = lvwColumnCenter
    lstMCOList.ColumnHeaders(5).Text = "Batch No"
    lstMCOList.ColumnHeaders(5).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(6).Text = "Prin. IA"
    lstMCOList.ColumnHeaders(6).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(7).Text = "Prin. IC"
    lstMCOList.ColumnHeaders(7).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(8).Text = "Prin. F"
    lstMCOList.ColumnHeaders(8).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(9).Text = "Prin. GA"
    lstMCOList.ColumnHeaders(9).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(10).Text = "Prin. GC"
    lstMCOList.ColumnHeaders(10).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(11).Text = "Prin. H"
    lstMCOList.ColumnHeaders(11).Alignment = lvwColumnLeft
    lstMCOList.ColumnHeaders(12).Text = "Prin. Total"
    lstMCOList.ColumnHeaders(12).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(13).Text = "Mem. IA"
    lstMCOList.ColumnHeaders(13).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(14).Text = "Mem. IC"
    lstMCOList.ColumnHeaders(14).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(15).Text = "Mem. F"
    lstMCOList.ColumnHeaders(15).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(16).Text = "Mem. GA"
    lstMCOList.ColumnHeaders(16).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(17).Text = "Mem. GC"
    lstMCOList.ColumnHeaders(17).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(18).Text = "Mem. H"
    lstMCOList.ColumnHeaders(18).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(19).Text = "Mem. Total"
    lstMCOList.ColumnHeaders(19).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(20).Text = "MCO Net IA"
    lstMCOList.ColumnHeaders(20).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(21).Text = "MCO Net IC"
    lstMCOList.ColumnHeaders(21).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(22).Text = "MCO Net F"
    lstMCOList.ColumnHeaders(22).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(23).Text = "MCO Net GA"
    lstMCOList.ColumnHeaders(23).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(24).Text = "MCO Net GC"
    lstMCOList.ColumnHeaders(24).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(25).Text = "MCO Net H"
    lstMCOList.ColumnHeaders(25).Alignment = lvwColumnRight
    lstMCOList.ColumnHeaders(26).Text = "MCO Net Total"
    lstMCOList.ColumnHeaders(26).Alignment = lvwColumnRight
End Sub
' ================= ListView Header ============================

