VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRecHospAdj 
   Caption         =   "Hospital Adj"
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
   Begin VB.Frame FramePv 
      Height          =   1080
      Left            =   0
      TabIndex        =   55
      Top             =   6000
      Width           =   11745
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
         ItemData        =   "FrmRecHospAdj.frx":0000
         Left            =   7200
         List            =   "FrmRecHospAdj.frx":0002
         TabIndex        =   22
         Top             =   180
         Width           =   975
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
         ItemData        =   "FrmRecHospAdj.frx":0004
         Left            =   5760
         List            =   "FrmRecHospAdj.frx":0006
         TabIndex        =   21
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtChequeNo 
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
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtAllocNo 
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
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtRemark 
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
         Left            =   5760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   470
         Width           =   2415
      End
      Begin VB.TextBox TxtReceiptNo 
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
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtChequeDate 
         Height          =   285
         Left            =   3360
         TabIndex        =   19
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
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
      Begin VB.TextBox txtChequeAmt 
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
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Receipt No"
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
         TabIndex        =   64
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Receipt Date"
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
         TabIndex        =   63
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Chq. No"
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
         Left            =   2520
         TabIndex        =   62
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Chq Amt"
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
         Left            =   2520
         TabIndex        =   61
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label24 
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
         Left            =   6840
         TabIndex        =   60
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "A/C Mth"
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
         Left            =   4800
         TabIndex        =   59
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label22 
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
         Left            =   2520
         TabIndex        =   58
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
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
         Left            =   4800
         TabIndex        =   57
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Alloc No"
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
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   0
      TabIndex        =   48
      Top             =   7080
      Width           =   11775
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Height          =   330
         Left            =   7560
         TabIndex        =   28
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   8160
         TabIndex        =   29
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   8880
         TabIndex        =   30
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10440
         TabIndex        =   32
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   11040
         TabIndex        =   33
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Height          =   330
         Left            =   6240
         TabIndex        =   26
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   330
         Left            =   6840
         TabIndex        =   27
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9480
         TabIndex        =   31
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   5040
         TabIndex        =   24
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   330
         Left            =   5640
         TabIndex        =   25
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   49
         Top             =   210
         Width           =   645
      End
      Begin VB.Label LblAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3480
         TabIndex        =   54
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   53
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   52
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   210
         Width           =   645
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   1800
         TabIndex        =   50
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11775
      Begin VB.TextBox TxtAllocNo1 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   1
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox TxtAllocNo2 
         Height          =   285
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   2
         Top             =   180
         Width           =   1335
      End
      Begin MSMask.MaskEdBox TxtChequeDate1 
         Height          =   315
         Left            =   6480
         TabIndex        =   9
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtChequeDate2 
         Height          =   315
         Left            =   8160
         TabIndex        =   10
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtReceiptNofr 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   3
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox txtReceiptNoto 
         Height          =   285
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   4
         Top             =   450
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtReceiptDatefr 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   705
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtReceiptDateto 
         Height          =   315
         Left            =   3120
         TabIndex        =   6
         Top             =   705
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtChequeNo1 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   7
         Top             =   945
         Width           =   1215
      End
      Begin VB.TextBox TxtChequeNo2 
         Height          =   285
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   8
         Top             =   945
         Width           =   1335
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmRecHospAdj.frx":0008
         Left            =   6480
         List            =   "FrmRecHospAdj.frx":000A
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmRecHospAdj.frx":000C
         Left            =   8160
         List            =   "FrmRecHospAdj.frx":000E
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmRecHospAdj.frx":0010
         Left            =   6480
         List            =   "FrmRecHospAdj.frx":0012
         TabIndex        =   13
         Top             =   765
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmRecHospAdj.frx":0014
         Left            =   8160
         List            =   "FrmRecHospAdj.frx":0016
         TabIndex        =   14
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "A/C Period Yr"
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
         Left            =   5040
         TabIndex        =   47
         Top             =   855
         Width           =   1455
      End
      Begin VB.Label Label14 
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
         Left            =   7800
         TabIndex        =   46
         Top             =   855
         Width           =   375
      End
      Begin VB.Label Label12 
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
         Left            =   7800
         TabIndex        =   45
         Top             =   570
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "A/C Period Mth"
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
         Left            =   5040
         TabIndex        =   44
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Chq. No"
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
         Left            =   240
         TabIndex        =   43
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   2760
         TabIndex        =   42
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Chq. Date"
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
         Left            =   5040
         TabIndex        =   41
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Left            =   7800
         TabIndex        =   40
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Receipt No"
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
         Left            =   240
         TabIndex        =   39
         Top             =   510
         Width           =   1095
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
         Left            =   2760
         TabIndex        =   38
         Top             =   510
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Receipt Date"
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
         Left            =   240
         TabIndex        =   37
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label19 
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
         Left            =   2760
         TabIndex        =   36
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label20 
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
         Left            =   2760
         TabIndex        =   35
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "Alloc No"
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
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstHospAdj 
      Height          =   4380
      Left            =   0
      TabIndex        =   65
      Top             =   1680
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7726
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Receipt No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Receipt Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Adj Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Alloc No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Alloc By"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Bank Contra"
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
      TabIndex        =   67
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
      TabIndex        =   66
      Top             =   7680
      Width           =   1575
   End
End
Attribute VB_Name = "FrmRecHospAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim one As Boolean
Private m_blnDirection(1 To 5) As Boolean
Dim AllocNo As String

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
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblAmt = 0
End Sub

Private Sub cmdAdd_Click()
    TxtReceiptNo.SetFocus
    Call ClearForm(Me)
End Sub

Private Sub CmdSave_Click()
Dim RecordSaved
Dim SaveHospAdj As New ADODB.Recordset
Dim strsql As String
    CmdSave.Enabled = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    If TxtReceiptNo = "" Then
        MsgBox "Please Enter Receipt No!", vbOKOnly + vbCritical, DialogBoxTitle
        TxtReceiptNo.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Receitp Date!", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    ElseIf txtChequeNo = "" Then
        MsgBox "Please Enter Cheque No!", vbOKOnly + vbCritical, DialogBoxTitle
        txtChequeNo.SetFocus
    ElseIf txtChequeDate = "__/__/____" Then
        MsgBox "Please Enter Cheque Date!", vbOKOnly + vbCritical, DialogBoxTitle
        txtChequeDate.SetFocus
    ElseIf txtChequeAmt = "" Then
        MsgBox "Please Enter Cheque Amount!", vbOKOnly + vbCritical, DialogBoxTitle
        txtChequeAmt.SetFocus
    Else
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            SaveHospAdj.Open "RecHospAdj", medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            SaveHospAdj.AddNew
            
            AllocNo = GenerateAllocNo("AH")
            TxtAllocNo = ""
            TxtAllocNo = AllocNo
            
            With SaveHospAdj
                !HospAdjAllocNo = Trim(AllocNo)
                !HospAdjCaseNo = "10400"
                !HospAdjRecNo = ChkStr(TxtReceiptNo)
                !HospAdjRecDate = Format(txtReceiptDate, "dd/mm/yyyy")
                !HospAdjChqNo = ChkStr(txtChequeNo)
                !HospAdjChqAmt = Format(txtChequeAmt, "#,##0.00")
                !HospAdjChqDate = Format(txtChequeDate, "dd/mm/yyyy")
                If CboACPeriodMth <> "" Then
                    !HospAdjACPeriodMth = Trim(CboACPeriodMth)
                End If
                If CboACPeriodYr <> "" Then
                    !HospAdjACPeriodYR = Trim(CboACPeriodYr)
                End If
                !HospAdjRemarks = Trim(ChkStr(TxtRemark))
                !HospAdjLastUpdateDate = Date
                !HospAdjLastUpdateUser = Logon
                !HospAdjAllocBy = Logon
                
            End With
            SaveHospAdj.Update
            Call ClearForm(Me)
            Call HospAdjListing
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                  
            SaveHospAdj.Close
            Set SaveHospAdj = Nothing
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    CmdSave.Enabled = True
End Sub

Private Sub CmdEdit_Click()
    one = True
    CmdSave.Enabled = False
    CmdUpdate.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
Dim RecordSaved
Dim HospAdj As New ADODB.Recordset
Dim strsql As String
Dim Index As String
    CmdUpdate.Enabled = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

    If one = False Then
        ElseIf lstHospAdj.listitems.count = 0 Then
            MsgBox "Please select the record", vbOKOnly + vbCritical
        ElseIf TxtReceiptNo = "" Then
            MsgBox "Please Enter Receipt No!", vbCritical
            TxtReceiptNo.SetFocus
        ElseIf txtReceiptDate = "__/__/____" Then
            MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtReceiptDate.SetFocus
        ElseIf txtChequeNo = "" Then
            MsgBox "Please Enter Cheque No!", vbCritical
            txtChequeNo.SetFocus
        ElseIf txtChequeDate = "__/__/____" Then
            MsgBox "Please Enter Cheque Date!", vbCritical
            txtChequeDate.SetFocus
        ElseIf txtChequeAmt = "" Then
            MsgBox "Please Enter Receipt Amount!", vbCritical
            txtChequeAmt.SetFocus
        Else
            
            Index = Trim(lstHospAdj.SelectedItem.SubItems(3))
                        
            strsql = "Select * From RecHospAdj Where HospAdjAllocNo = '" & Trim(Index) & "'"
            HospAdj.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If HospAdj.recordCount = 0 Then
                    HospAdj.AddNew
                End If
                
                With HospAdj
                    !HospAdjCaseNo = "10400"
                    !HospAdjRecNo = ChkStr(TxtReceiptNo)
                    !HospAdjRecDate = Format(txtReceiptDate, "dd/mm/yyyy")
                    !HospAdjChqAmt = Format(txtChequeAmt, "#,##0.00")
                    !HospAdjChqNo = Trim(ChkStr(txtChequeNo))
                    !HospAdjChqDate = Format(txtChequeDate, "dd/mm/yyyy")
                    !HospAdjRemarks = Trim(ChkStr(TxtRemark))
                    !HospAdjLastUpdateDate = Date
                    !HospAdjLastUpdateUser = Logon
                    If CboACPeriodMth <> "" Then
                        !HospAdjACPeriodMth = CboACPeriodMth
                    End If
                    If CboACPeriodYr <> "" Then
                        !HospAdjACPeriodYR = CboACPeriodYr
                    End If
                    .Update
                End With
                
                MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
          
                HospAdj.Close
                Set HospAdj = Nothing
                
                TxtAllocNo.Text = ""
                TxtReceiptNo.Text = ""
                txtReceiptDate.Text = "__/__/____"
                txtChequeNo.Text = ""
                txtChequeAmt.Text = ""
                txtChequeDate.Text = "__/__/____"
                TxtRemark.Text = ""
                CboACPeriodMth.Text = ""
                CboACPeriodYr.Text = ""
                Call SearchRecord
                one = False
                CmdSave.Enabled = True
                CmdUpdate.Enabled = False
            End If
        CmdSave.Enabled = True
        CmdUpdate.Enabled = False
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    CmdUpdate.Enabled = True
End Sub

Private Sub CmdDelete_Click()
Dim DeleteHospAdj As New ADODB.Recordset
Dim strsql As String
Dim DeleteRecord
    CmdDelete.Enabled = False
    If lstHospAdj.listitems.count = 0 Then
       MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
       DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
         If DeleteRecord = vbYes Then
            medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
            strsql = "Delete From RecBankContra Where HospAdjAllocNo = '" & ChkStr(lstHospAdj.SelectedItem.SubItems(3)) & "'"
            DeleteHospAdj.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
            Call ClearForm(Me)
            Call HospAdjListing
            medixmasterconnPayment.Close
            Set medixmasterconnPayment = Nothing
         End If
    End If
    CmdDelete.Enabled = True
End Sub

Private Sub cmdSearch_Click()
    Screen.MousePointer = vbHourglass
        CmdSearch.Enabled = False
        medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        Call SearchRecord
        medixmasterconnPayment.Close
        Set medixmasterconnPayment = Nothing
        CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    Check = False
    CmdExit.Enabled = True
End Sub

Private Sub CmdPrint_Click()
    'Screen.MousePointer = vbHourglass
        'Call ExportToExcelAnalysis1(lstHospAdj)
    'Screen.MousePointer = vbDefault
    
    Screen.MousePointer = vbHourglass
        
    If lstHospAdj.listitems.count <> 0 Then
        Call ExportListViewtoExcel(lstHospAdj)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CmdClear_Click()
    one = False
    CmdSave.Enabled = True
    CmdUpdate.Enabled = False
    TxtAllocNo.Text = ""
    TxtReceiptNo.Text = ""
    txtReceiptDate.Text = "__/__/____"
    txtChequeNo.Text = ""
    txtChequeAmt.Text = ""
    txtChequeDate.Text = "__/__/____"
    TxtRemark.Text = ""
    CboACPeriodMth.Text = ""
    CboACPeriodYr.Text = ""
    lstHospAdj.listitems.Clear
    LblAmt = ""
    lblCurrent = ""
    lblTotal = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub lstHospAdj_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
Dim Index As String

    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    
    If one = True Then
    
        TxtAllocNo.Text = ""
        TxtReceiptNo.Text = ""
        txtReceiptDate.Text = "__/__/____"
        txtChequeNo.Text = ""
        txtChequeDate.Text = "__/__/____"
        txtChequeAmt.Text = ""
        CboACPeriodMth.Text = ""
        CboACPeriodYr.Text = ""
        TxtRemark.Text = ""
        
        CmdSave.Enabled = False
    
        If lstHospAdj.listitems.count Then
        
            Index = Trim(lstHospAdj.SelectedItem.SubItems(3))
                
            strsql = " SELECT * From RecHospAdj Where HospAdjAllocNo = '" & Trim(Index) & "'"
                    
            rst3.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            
                If rst3.recordCount = 0 Then
                    MsgBox "No Record Found!", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    TxtReceiptNo.SetFocus
                    If Len(rst3!HospAdjAllocNo) Then
                        TxtAllocNo = Trim(rst3!HospAdjAllocNo)
                    End If
                    If Len(rst3!HospAdjRecNo) Then
                        TxtReceiptNo = Trim(rst3!HospAdjRecNo)
                    End If
                    If Len(rst3!HospAdjRecDate) Then
                        txtReceiptDate = Format(rst3!HospAdjRecDate, "dd/mm/yyyy")
                    End If
                    If Len(rst3!HospAdjChqNo) Then
                        txtChequeNo = Trim(rst3!HospAdjChqNo)
                    End If
                    If Len(rst3!HospAdjChqDate) Then
                        txtChequeDate = Format(rst3!HospAdjChqDate, "dd/mm/yyyy")
                    End If
                    If Len(rst3!HospAdjChqAmt) Then
                        txtChequeAmt = Format(rst3!HospAdjChqAmt, "#,##0.00")
                    End If
                    If Len(rst3!HospAdjRemarks) Then
                        TxtRemark = Trim(rst3!HospAdjRemarks)
                    End If
                    If Len(rst3!HospAdjACPeriodMth) Then
                        CboACPeriodMth = rst3!HospAdjACPeriodMth
                    End If
                    If Len(rst3!HospAdjACPeriodYR) Then
                        CboACPeriodYr = rst3!HospAdjACPeriodYR
                    End If
            
                    CmdSave.Enabled = False
                    rst3.Close
                    Set rst3 = Nothing
                End If
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
End Sub

Private Sub txtChequeAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
        CboACPeriodMth.Text = Format(txtReceiptDate, "mm")
        CboACPeriodYr.Text = Format(txtReceiptDate, "yyyy")
    Else
        If txtReceiptDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDate.SetFocus
        End If
    End If
End Sub

Private Sub txtReceiptDatefr_LostFocus()
    If IsDate(txtReceiptDatefr) Then
    Else
        If txtReceiptDatefr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDatefr.SetFocus
        End If
    End If
End Sub

Private Sub txtReceiptDateto_LostFocus()
    If IsDate(txtReceiptDateto) Then
    Else
        If txtReceiptDateto <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDateto.SetFocus
        End If
    End If
End Sub

Private Sub txtChequeDate_LostFocus()
    If IsDate(txtChequeDate) Then
    Else
        If txtChequeDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtChequeDate.SetFocus
        End If
    End If
End Sub

Private Sub TxtChequeDate1_LostFocus()
    If IsDate(TxtChequeDate1) Then
    Else
        If TxtChequeDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtChequeDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtChequeDate2_LostFocus()
    If IsDate(TxtChequeDate2) Then
    Else
        If TxtChequeDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtChequeDate2.SetFocus
        End If
    End If
End Sub


Private Sub CboACPeriodMth_KeyPress(KeyAscii As Integer)
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

Private Function SearchRecord()
Dim strAppend As String
Dim Sql As String
    
    strAppend = ""
    
    
    If Len(TxtAllocNo1) Then
        strAppend = strAppend + " and HospAdjAllocNo >= '" & Trim(TxtAllocNo1) & "'"
    End If
        
    If Len(TxtAllocNo1) Then
        strAppend = strAppend + " and HospAdjAllocNo <= '" & Trim(TxtAllocNo1) & "'"
    End If
    
    If Len(txtReceiptNofr) Then
        strAppend = strAppend + " and HospAdjRecNo >= '" & Trim(txtReceiptNofr) & "'"
    End If
        
    If Len(txtReceiptNoto) Then
        strAppend = strAppend + " and HospAdjRecNo <= '" & Trim(txtReceiptNoto) & "'"
    End If
    
    If Len(Trim(txtReceiptDatefr)) > 0 And IsDate(txtReceiptDatefr) = True Then
        strAppend = strAppend + " And HospAdjRecDate >= '" & Format(txtReceiptDatefr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtReceiptDateto)) > 0 And IsDate(txtReceiptDateto) = True Then
        strAppend = strAppend + " And HospAdjRecDate <= '" & Format(txtReceiptDateto, "mm/dd/yyyy") & "'"
    End If
    
    If Len(TxtChequeNo1) Then
        strAppend = strAppend + " and HospAdjChqNo >= '" & Trim(TxtChequeNo1) & "'"
    End If
        
    If Len(TxtChequeNo2) Then
        strAppend = strAppend + " and HospAdjChqNo <= '" & Trim(TxtChequeNo2) & "'"
    End If
        
    If Len(cboACPeriodMthFr) Then
        strAppend = strAppend + " and HospAdjACPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
        
    If Len(cboACPeriodMthTo) Then
        strAppend = strAppend + " and HospAdjACPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If Len(cboACPeriodYRTo) Then
        strAppend = strAppend + " and HospAdjACPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
        
    If Len(cboACPeriodYRTo) Then
        strAppend = strAppend + " and HospAdjACPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    Call HospAdjListing(strAppend)
   
End Function

Private Sub HospAdjListing(Optional strAppend As String)
Dim Sql As String
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim CurrRec As String
   
   total = 0
   CurrRec = 0

    lstHospAdj.listitems.Clear

    Sql = " Select * FROM RecHospAdj where 0 = 0 "

    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
    
    rst2.Open Sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
   
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        lblTotal = rst2.recordCount
        Do
            Do Until rst2.EOF
            
                With rst2
                    Set Record = lstHospAdj.listitems.Add(, , !HospAdjRecNo)
                    
                    If Len(!HospAdjRecDate) Then
                        Record.SubItems(1) = Format(!HospAdjRecDate, "dd/mm/yyyy")
                    End If
                    If Len(!HospAdjChqAmt) Then
                        Record.SubItems(2) = Format(!HospAdjChqAmt, "#,##0.00")
                    End If
                    If Len(!HospAdjAllocNo) Then
                        Record.SubItems(3) = Trim(!HospAdjAllocNo)
                    End If
                    
                    If Len(!HospAdjAllocBy) Then
                        Record.SubItems(4) = Trim(!HospAdjAllocBy)
                    End If
                                    
                    CurrRec = CurrRec + 1
                    lblCurrent = CurrRec
                    
                    total = total + IIf(IsNumeric(!HospAdjChqAmt) = True, !HospAdjChqAmt, 0)
                    LblAmt = Format(total, "#,##,0.00")
                                
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


Private Sub cboACPeriodMthFr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
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
End Sub

Private Sub cboACPeriodMthTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
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
End Sub

Private Sub cboACPeriodYRFr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
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
End Sub

Private Sub cboACPeriodYRTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
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
End Sub

Private Sub ComboListing()
Dim MthStart As Integer
Dim YrStart As Integer

    For MthStart = 1 To 12
        CboACPeriodMth.AddItem Format(MthStart, "00")
        cboACPeriodMthFr.AddItem Format(MthStart, "00")
        cboACPeriodMthTo.AddItem Format(MthStart, "00")
    Next MthStart
    
    For YrStart = 1999 To 3000
        CboACPeriodYr.AddItem YrStart
        cboACPeriodYRFr.AddItem YrStart
        cboACPeriodYRTo.AddItem YrStart
    Next YrStart
    
End Sub


Private Sub lstHospAdj_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstHospAdj, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstHospAdj, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
    Case 3
        SortListView lstHospAdj, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstHospAdj, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstHospAdj, ColumnHeader.Index, ldtString, m_blnDirection(5)
    End Select
End Sub


