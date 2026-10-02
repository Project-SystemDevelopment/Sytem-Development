VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAllocMedixBankChrg 
   Caption         =   "Medix Bank Charges"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   80
      TabIndex        =   40
      Top             =   360
      Width           =   11775
      Begin VB.TextBox txtPVNo2 
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
         Left            =   7200
         MaxLength       =   2015
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtPVNo1 
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
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtPVDate1 
         Height          =   285
         Left            =   4800
         TabIndex        =   4
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
      Begin MSMask.MaskEdBox txtPVDate2 
         Height          =   285
         Left            =   7200
         TabIndex        =   5
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
      Begin VB.TextBox txtPVCasNo 
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
         Left            =   960
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "10100"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtAllocNumber 
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
         Left            =   960
         MaxLength       =   8
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Alloc. No"
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
         TabIndex        =   45
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "PV Number"
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
         Left            =   3720
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "PV Date"
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
         Left            =   3720
         TabIndex        =   43
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
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
         Left            =   6720
         TabIndex        =   42
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label7 
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
         Left            =   6720
         TabIndex        =   41
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame10 
      Height          =   585
      Left            =   80
      TabIndex        =   37
      Top             =   6840
      Width           =   11775
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1080
         TabIndex        =   50
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   615
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
         Left            =   5160
         TabIndex        =   20
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Ex&port"
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
         TabIndex        =   23
         Top             =   200
         Width           =   840
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
         Height          =   315
         Left            =   10440
         TabIndex        =   25
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
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
         TabIndex        =   24
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdDelete 
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
         Left            =   6600
         TabIndex        =   22
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdEdit 
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
         Left            =   3720
         TabIndex        =   18
         Top             =   200
         Width           =   720
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
         Height          =   315
         Left            =   3000
         TabIndex        =   17
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdSearch 
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
         Left            =   4440
         TabIndex        =   19
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdStop 
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
         Height          =   315
         Left            =   5880
         TabIndex        =   21
         Top             =   200
         Width           =   720
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   39
         Top             =   200
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Records"
         Height          =   255
         Left            =   1920
         TabIndex        =   38
         Top             =   200
         Width           =   735
      End
   End
   Begin VB.Frame FrameDet 
      Enabled         =   0   'False
      Height          =   975
      Left            =   80
      TabIndex        =   26
      Top             =   5880
      Width           =   11775
      Begin VB.TextBox txtChqNo 
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
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtPVNo 
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
         Left            =   2760
         MaxLength       =   15
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   9600
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   240
         Width           =   2040
      End
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
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
      Begin MSMask.MaskEdBox txtChqDate 
         Height          =   285
         Left            =   4800
         TabIndex        =   12
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.TextBox txtCasNo 
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
         Left            =   840
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "10100"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAllocNo 
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
         Left            =   840
         MaxLength       =   8
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtChqAmt 
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
         Height          =   285
         Left            =   6840
         MaxLength       =   12
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox CboACPeriodYr 
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
         ItemData        =   "FrmAllocMedixBankChrg.frx":0000
         Left            =   8040
         List            =   "FrmAllocMedixBankChrg.frx":0002
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox CboACPeriodMth 
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
         ItemData        =   "FrmAllocMedixBankChrg.frx":0004
         Left            =   6840
         List            =   "FrmAllocMedixBankChrg.frx":0006
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
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
         Left            =   8880
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Year"
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
         Left            =   7640
         TabIndex        =   35
         Top             =   560
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "A/C Month"
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
         Left            =   6000
         TabIndex        =   34
         Top             =   500
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Chq. Amt"
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
         Left            =   6000
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Chq. No."
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
         Left            =   3960
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "PV Date"
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
         Left            =   2040
         TabIndex        =   31
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "PV No."
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
         Left            =   2040
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Alloc. No"
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
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label19 
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
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
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
         Left            =   3960
         TabIndex        =   27
         Top             =   480
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstBankCharges 
      Height          =   4455
      Left            =   80
      TabIndex        =   6
      Top             =   1440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Case No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Alloc No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "PV No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "PV Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Chq. No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Chq Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "BankChrgs"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Acc Period"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "AllocBy"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Remark"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Medix Bank Charges"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   11895
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
      TabIndex        =   47
      Top             =   7440
      Width           =   1575
   End
End
Attribute VB_Name = "FrmAllocMedixBankChrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Integer
Dim EditFlag As Boolean
Private Sub ComboListing()
Dim monthstart
Dim yearstart
    For monthstart = 1 To 12
        With CboACPeriodMth
            .AddItem Format(monthstart, "00")
        End With
    Next monthstart
    For yearstart = 1999 To 2500
        With CboACPeriodYr
            .AddItem yearstart
        End With
    Next yearstart
End Sub

Private Sub CboACPeriodMth_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboACPeriodMth.Text, CboACPeriodMth.SelStart) + Chr(KeyAscii) + Mid(CboACPeriodMth.Text, CboACPeriodMth.SelStart + CboACPeriodMth.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboACPeriodMth.ListCount - 1
        If NewText = LCase(Left(CboACPeriodMth.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboACPeriodMth.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboACPeriodMth.Text = ValidValue
               CboACPeriodMth.SelStart = 0
               CboACPeriodMth.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboACPeriodYr_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboACPeriodYr.Text, CboACPeriodYr.SelStart) + Chr(KeyAscii) + Mid(CboACPeriodYr.Text, CboACPeriodYr.SelStart + CboACPeriodYr.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To CboACPeriodYr.ListCount - 1
        If NewText = LCase(Left(CboACPeriodYr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboACPeriodYr.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
           CboACPeriodYr.Text = ValidValue
           CboACPeriodYr.SelStart = 0
           CboACPeriodYr.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CmdAdd_Click()
    EditFlag = False
    Call ClearDet
    FrameDet.Enabled = True
    CmdSave.Enabled = True
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    txtPVNo.SetFocus
End Sub

Private Sub CmdClear_Click()
Dim TmpCaseNO As String
    TmpCaseNO = txtCasNo
    Call ClearForm(Me)
    txtPVCasNo = TmpCaseNO
    txtCasNo = TmpCaseNO
    lstBankCharges.ListItems.Clear
    CboACPeriodMth.Text = ""
    CboACPeriodYr.Text = ""
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    lblCurrent = ""
    lblTotal = ""
    txtAllocNumber.SetFocus
End Sub

Private Sub CmdDelete_Click()
On Error Resume Next
Dim rst2 As New ADODB.Recordset
Dim sql2 As String
Dim RecordDel
    CmdSearch.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
    If lstBankCharges.ListItems.Count > 0 Then
        If Trim(lstBankCharges.SelectedItem.SubItems(1)) <> "" Then
            RecordDel = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
            If RecordDel = vbYes Then
                Screen.MousePointer = vbHourglass
                medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
                    sql2 = ""
                    sql2 = "Delete From PaymentMedixBankCharges Where PVAllocNo = '" & Trim(lstBankCharges.SelectedItem.SubItems(1)) & "'"
                    rst2.Open sql2, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                    MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
                    Call ClearDet
                medixmasterconnPayment.Close
                Set medixmasterconnPayment = Nothing
                Screen.MousePointer = vbDefault
            End If
            lstBankCharges.ListItems.Remove (lstBankCharges.SelectedItem.Index)
        End If
    End If
    CmdSearch.Enabled = True
    CmdDelete.Enabled = True
    CmdSave.Enabled = True
End Sub

Private Sub CmdEdit_Click()
    If Trim(txtAllocNo) <> "" Then
        EditFlag = True
        CmdSave.Caption = "&Update"
        CmdSave.Enabled = True
        FrameDet.Enabled = True
        txtPVNo.SetFocus
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
    If lstBankCharges.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstBankCharges)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSave_Click()
On Error Resume Next
Dim UpdateRec
Dim BankChargesRec As New ADODB.Recordset
Dim BankChargesAllocSeq As String
Dim BankChargesSql As String
Dim tmpAllocNo As String
    CmdSearch.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    If Trim(txtPVNo) = "" Then
        MsgBox "Please Keyin PV Number ! ", vbOKOnly
        txtPVNo.SetFocus
    ElseIf txtPVDate = "__/__/____" Then
        MsgBox "Please Keyin PV Date!", vbCritical
        txtPVDate.SetFocus
    ElseIf Trim(txtChqNo) = "" Then
        MsgBox "Please Keyin Cheque Number ! ", vbCritical
        txtChqNo.SetFocus
    ElseIf txtChqDate = "__/__/____" Then
        MsgBox "Please Keyin Cheque Date!", vbCritical
        txtChqDate.SetFocus
    ElseIf Trim(txtChqAmt) = "" Or IsNumeric(txtChqAmt) = False Then
        MsgBox "Please Keyin Cheque Amt!", vbCritical
        txtChqAmt.SetFocus
    ElseIf Trim(CboACPeriodMth) = "" Then
        MsgBox "Please Keyin Account Period - Month !", vbCritical
        CboACPeriodMth.SetFocus
    ElseIf Trim(CboACPeriodYr) = "" Then
        MsgBox "Please Keyin Account Period - Year !", vbCritical
        CboACPeriodYr.SetFocus
    Else
        UpdateRec = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        If UpdateRec = vbYes Then
            If EditFlag = False Then
                BankChargesAllocSeq = GenerateAllocNo("AB")
                txtAllocNo = ""
                txtAllocNo = BankChargesAllocSeq
            End If
            BankChargesSql = "Select * from PaymentMEdixBankCharges WHERE PVAllocno='" & Trim(txtAllocNo) & "'"
            BankChargesRec.Open BankChargesSql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            If EditFlag = False Then
                BankChargesRec.AddNew
            End If
            With BankChargesRec
                !PVAllocNo = Trim(txtAllocNo)
                !PVCasNo = Trim(txtCasNo)
                !PayPVNo = Trim(txtPVNo)
                !PayPVDate = Trim(txtPVDate)
                !PVChqNo = Trim(txtChqNo)
                !PVChqDate = Trim(txtChqDate)
                !PVChqAmt = Trim(txtChqAmt)
                !PayAccPeriodMth = Format(Trim(CboACPeriodMth), "00")
                !PayAccPeriodYr = Trim(CboACPeriodYr)
                !PayRemarks = IIf(Len(Trim(txtRemark)), Trim(txtRemark), "")
                !PAYLastUpdateUser = Logon
                !PayLastUpdateDate = Date
            End With
            BankChargesRec.Update
            BankChargesRec.Close
            Set BankChargesRec = Nothing
            If EditFlag = True Then
                MsgBox "Record Updated..... ! "
            Else
                MsgBox "Record Saved..... ! "
            End If
            CmdSave.Caption = "Sa&ve"
            CmdSave.Enabled = False
            CmdEdit.Enabled = False
            tmpAllocNo = txtAllocNo
            Call CmdClear_Click
            Call MedixBankChargesListing(" And PVAllocNo = '" & Trim(tmpAllocNo) & "'")
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    CmdSearch.Enabled = True
    CmdDelete.Enabled = True
    CmdSave.Enabled = True
End Sub

Private Sub CmdSearch_Click()
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        Call SearchRecord
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    CmdSearch.Enabled = True
    CmdDelete.Enabled = True
    CmdSave.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
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
    EditFlag = False
    Call ComboListing
End Sub

Private Sub lstBankCharges_Click()
On Error Resume Next
    Dim LstREc As ListItem
    FrameDet.Enabled = False
    If lstBankCharges.ListItems.Count > 0 Then
        Call ClearDet
        CmdEdit.Enabled = True
        CmdDelete.Enabled = True
        txtCasNo = lstBankCharges.SelectedItem.Text
        txtAllocNo = lstBankCharges.SelectedItem.SubItems(1)
        txtPVNo = lstBankCharges.SelectedItem.SubItems(2)
        If IsDate(lstBankCharges.SelectedItem.SubItems(3)) Then
            txtPVDate = lstBankCharges.SelectedItem.SubItems(3)
        End If
        txtChqNo = lstBankCharges.SelectedItem.SubItems(4)
        If IsDate(lstBankCharges.SelectedItem.SubItems(5)) Then
            txtChqDate = lstBankCharges.SelectedItem.SubItems(5)
        End If
        txtChqAmt = IIf(IsNumeric(lstBankCharges.SelectedItem.SubItems(6)), lstBankCharges.SelectedItem.SubItems(6), 0)
        If Len(lstBankCharges.SelectedItem.SubItems(7)) Then
            CboACPeriodMth = Mid(lstBankCharges.SelectedItem.SubItems(7), 1, 2)
            CboACPeriodYr = Mid(lstBankCharges.SelectedItem.SubItems(7), 4, 4)
        End If
        txtRemark = lstBankCharges.SelectedItem.SubItems(9)
    End If
End Sub

Private Sub txtAllocNo_LostFocus()
    txtAllocNo = ChkStr(txtAllocNo)
End Sub

Private Sub txtAllocNumber_LostFocus()
    txtAllocNumber = ChkStr(txtAllocNumber)
End Sub

Private Sub txtChqAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtChqNo_LostFocus()
    txtChqNo = ChkStr(txtChqNo)
End Sub

Private Sub txtChqDate_LostFocus()
    If IsDate(txtChqDate) Then
    Else
        If txtChqDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtChqDate.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDate_LostFocus()
    If IsDate(txtPVDate) Then
        CboACPeriodMth.Text = Format(txtPVDate, "mm")
        CboACPeriodYr.Text = Format(txtPVDate, "yyyy")
    Else
        If txtPVDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDate1_LostFocus()
    If IsDate(txtPVDate1) Then
    Else
        If txtPVDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDate2_LostFocus()
    If IsDate(txtPVDate2) Then
    Else
        If txtPVDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtPVNo_LostFocus()
    txtPVNo = ChkStr(txtPVNo)
End Sub

Private Sub txtPVNo1_LostFocus()
    txtPVNo1 = ChkStr(txtPVNo1)
End Sub

Private Sub txtPVNo2_LostFocus()
    txtPVNo2 = ChkStr(txtPVNo2)
End Sub

Private Sub txtRemark_LostFocus()
    txtRemark = ChkStr(txtRemark)
End Sub

Private Sub SearchRecord()
On Error Resume Next
    Dim strAppend As String
    strAppend = ""
    If Len(Trim(txtPVCasNo)) Then
        strAppend = strAppend + " And PVCasNo = '" & Trim(txtPVCasNo) & "'"
    End If
    If Len(Trim(txtAllocNumber)) Then
        strAppend = strAppend + " And PVAllocNo = '" & Trim(txtAllocNumber) & "'"
    End If
    If Len(Trim(txtPVNo1)) Then
        strAppend = strAppend + " And PAYPVNo >= '" & Trim(txtPVNo1) & "'"
    End If
    If Len(txtPVNo2) Then
        strAppend = strAppend + " and PAYPVNo <= '" & Trim(txtPVNo2) & "'"
    End If
    If Len(Trim(txtPVDate1)) > 0 And IsDate(txtPVDate1) = True Then
        strAppend = strAppend + " And PAYPVDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
    End If
    If Len(Trim(txtPVDate2)) > 0 And IsDate(txtPVDate2) = True Then
        strAppend = strAppend + " And PAYPVDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
    End If
    Call MedixBankChargesListing(strAppend)
End Sub
Private Sub MedixBankChargesListing(Optional strAppend As String)
On Error Resume Next
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim totalREc As String
Dim CurrentRec As String
Dim Check As Boolean
Dim TmpType As String
    totalREc = 0
    CurrentRec = 0
    lstBankCharges.ListItems.Clear
    Sql = "Select * FROM PaymentMedixBankCharges where 0 = 0 "
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
    rst2.Open Sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
    Do
        Do Until rst2.EOF
        totalREc = rst2.recordCount
        lblTotal = totalREc
            With rst2
                Set Record = lstBankCharges.ListItems.Add(, , !PVCasNo)
                Record.SubItems(1) = !PVAllocNo
                Record.SubItems(2) = IIf(Len(!PayPVNo), !PayPVNo, "")
                If Len(!PayPVDate) Then
                    Record.SubItems(3) = !PayPVDate
                End If
                Record.SubItems(4) = IIf(Len(!PVChqNo), !PVChqNo, "")
                If Len(!PVChqDate) Then
                    Record.SubItems(5) = Trim(!PVChqDate)
                End If
                Record.SubItems(6) = Format(IIf(IsNumeric(!PVChqAmt), !PVChqAmt, 0), "#,##0.00")
                If Len(Trim(!PayAccPeriodMth) & Trim(!PayAccPeriodYr)) Then
                    Record.SubItems(7) = Format(!PayAccPeriodMth & "/" & !PayAccPeriodYr, "MM/YYYY")
                End If
                Record.SubItems(8) = IIf(Len(!PAYLastUpdateUser), !PAYLastUpdateUser, "")
                Record.SubItems(9) = IIf(Len(!PayRemarks), !PayRemarks, "")
                CurrentRec = CurrentRec + 1
                lblCurrent = CurrentRec
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
    
End Sub

Private Sub ClearDet()
    txtAllocNo = ""
    txtPVNo = ""
    txtPVDate = "__/__/____"
    txtChqNo = ""
    txtChqDate = "__/__/____"
    txtChqAmt = ""
    CboACPeriodMth.Text = ""
    CboACPeriodYr.Text = ""
    txtRemark = ""
End Sub

