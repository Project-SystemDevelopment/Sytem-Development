VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateWS2Bordx 
   Caption         =   "Update Bordx"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   10545
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkBordxNo 
      Height          =   255
      Left            =   4800
      TabIndex        =   79
      Top             =   1200
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.TextBox txtRECNo 
      Height          =   285
      Left            =   4800
      TabIndex        =   14
      Top             =   6840
      Width           =   975
   End
   Begin VB.ComboBox cboAccPeriodMth 
      Height          =   315
      ItemData        =   "frmUpdateWS2Bordx.frx":0000
      Left            =   10800
      List            =   "frmUpdateWS2Bordx.frx":0028
      TabIndex        =   20
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtBordxNo 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   6840
      Width           =   975
   End
   Begin MSComctlLib.ListView lstWorksheetTrackList 
      Height          =   4935
      Left            =   120
      TabIndex        =   52
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Case No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Claim Version"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Wks No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Actual Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "App Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Hosp Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Member Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "DOA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "DOD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Patient"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Hospital"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Grp Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "Plan Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "ICD Code"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   120
      TabIndex        =   30
      Top             =   240
      Width           =   11655
      Begin VB.TextBox txtBordxFr 
         Height          =   310
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   3
         Top             =   840
         Width           =   1560
      End
      Begin VB.TextBox txtBordxTo 
         Height          =   310
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   4
         Top             =   840
         Width           =   1560
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   310
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   2
         Top             =   480
         Width           =   1560
      End
      Begin VB.TextBox TxtCaseNo 
         Height          =   310
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   1
         Top             =   480
         Width           =   1560
      End
      Begin VB.ComboBox CboAmt 
         Height          =   315
         ItemData        =   "frmUpdateWS2Bordx.frx":0053
         Left            =   6000
         List            =   "frmUpdateWS2Bordx.frx":0060
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox CboSequence 
         Height          =   315
         Left            =   5400
         TabIndex        =   38
         Top             =   1800
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox CboPln2 
         Height          =   315
         Left            =   3360
         Sorted          =   -1  'True
         TabIndex        =   37
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox CboPln1 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   140
         Width           =   3495
      End
      Begin VB.ComboBox CboHos 
         Height          =   315
         Left            =   5400
         TabIndex        =   35
         Top             =   2160
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox CboCaseStatus 
         Height          =   315
         Left            =   1560
         TabIndex        =   33
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox CboWorksheet 
         Height          =   315
         Left            =   5400
         TabIndex        =   32
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox CboAmount 
         Height          =   315
         Left            =   1680
         TabIndex        =   31
         Top             =   2040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSMask.MaskEdBox txtDOADate1 
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   255
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate1 
         Height          =   255
         Left            =   6000
         TabIndex        =   7
         Top             =   540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate2 
         Height          =   255
         Left            =   7440
         TabIndex        =   6
         Top             =   255
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate2 
         Height          =   255
         Left            =   7440
         TabIndex        =   8
         Top             =   540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label37 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   83
         Top             =   915
         Width           =   495
      End
      Begin VB.Label Label36 
         Caption         =   "Bordx No Fr."
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Case No Fr."
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   63
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Sequence by"
         Height          =   255
         Left            =   4080
         TabIndex        =   51
         Top             =   1890
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   50
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Plan Code from"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   4080
         TabIndex        =   47
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Group Company"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Case Status"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Bordx Status"
         Height          =   255
         Left            =   3960
         TabIndex        =   44
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "DOA from"
         Height          =   255
         Left            =   5160
         TabIndex        =   43
         ToolTipText     =   "Date of Admission"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "DOD from"
         Height          =   255
         Left            =   5160
         TabIndex        =   42
         ToolTipText     =   "Date of Discharged"
         Top             =   560
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2100
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   7200
         TabIndex        =   40
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   7200
         TabIndex        =   39
         Top             =   520
         Width           =   495
      End
      Begin VB.Label Label7 
         Height          =   15
         Left            =   600
         TabIndex        =   53
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Amount"
         Height          =   255
         Left            =   5160
         TabIndex        =   62
         Top             =   840
         Width           =   735
      End
   End
   Begin MSMask.MaskEdBox txtBordxDate 
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtToPAYDate 
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFRPAYDate 
      Height          =   255
      Left            =   8520
      TabIndex        =   17
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox ToMNRBDate 
      Height          =   255
      Left            =   6600
      TabIndex        =   18
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtFRMNRBDate 
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtChequeNo 
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      Top             =   6840
      Width           =   1095
   End
   Begin VB.ComboBox cboAccPeriodYR 
      Height          =   315
      ItemData        =   "frmUpdateWS2Bordx.frx":007A
      Left            =   10800
      List            =   "frmUpdateWS2Bordx.frx":0102
      TabIndex        =   21
      Top             =   7080
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   7320
      Width           =   11655
      Begin VB.TextBox txtNoofDay 
         Height          =   285
         Left            =   4320
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdBordx 
         Caption         =   "&Update"
         Height          =   375
         Left            =   9360
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7920
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10800
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4800
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10080
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   8640
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblItemChecked 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   4080
         TabIndex        =   65
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Amt Checked"
         Height          =   255
         Left            =   4080
         TabIndex        =   68
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblBordxNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   5520
         TabIndex        =   67
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   5520
         TabIndex        =   66
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   54
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   2640
         TabIndex        =   57
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   59
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1935
         TabIndex        =   56
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "of"
         Height          =   255
         Left            =   720
         TabIndex        =   58
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.TextBox txtChequeAmt 
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Top             =   7080
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtRECDate 
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Update Bordx"
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
      Height          =   220
      Left            =   0
      TabIndex        =   84
      Top             =   0
      Width           =   11805
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
      Left            =   120
      TabIndex        =   81
      Top             =   6600
      Width           =   375
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
      Left            =   9480
      TabIndex        =   80
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label35 
      Caption         =   "Rec No"
      Height          =   255
      Left            =   4080
      TabIndex        =   78
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label34 
      Caption         =   "Rec Date"
      Height          =   255
      Left            =   4065
      TabIndex        =   77
      Top             =   7110
      Width           =   735
   End
   Begin VB.Label Label33 
      Caption         =   "Cheque Amt"
      Height          =   255
      Left            =   1965
      TabIndex        =   76
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label32 
      Caption         =   "A/C Period Mth"
      Height          =   255
      Left            =   9600
      TabIndex        =   75
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label31 
      Caption         =   "A/C Period Yr"
      Height          =   255
      Left            =   9600
      TabIndex        =   74
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label30 
      Caption         =   "Cheque No"
      Height          =   255
      Left            =   2040
      TabIndex        =   73
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label27 
      Caption         =   "FR MNRB"
      Height          =   255
      Left            =   7680
      TabIndex        =   72
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label26 
      Caption         =   "To MNRB"
      Height          =   255
      Left            =   5880
      TabIndex        =   71
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label24 
      Caption         =   "FR Payor"
      Height          =   255
      Left            =   7680
      TabIndex        =   70
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "To Payor"
      Height          =   255
      Left            =   5880
      TabIndex        =   69
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "Bordx Date"
      Height          =   255
      Left            =   120
      TabIndex        =   61
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Bordx No"
      Height          =   255
      Left            =   120
      TabIndex        =   60
      Top             =   6840
      Width           =   735
   End
End
Attribute VB_Name = "frmUpdateWS2Bordx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim CaseNoArray() As String
Dim CLMBordxNo As String
Dim CASNo As String
Dim CASVer As String
Dim TotalApp As Variant
Dim TotCLM As Integer
Dim searchCriteria As String
Private m_blnDirection(1 To 14) As Boolean
'****************** Start WorksheetTracking Listing *****************************

Private Sub WorksheetTrackingListing(Optional StrAppend As String, Optional Noprompt As Boolean)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim ClaimAdmin As Boolean
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim rstCount

    Total = 0
    Current = 0
    
    lstWorksheetTrackList.ListItems.Clear

    sql = "Select CASNumber, ClaimVersion, PAYCode, CLMTotalApproved, " & _
          "CasDateAdmitted, CASDischargeDate, ClaimStatus, " & _
          "MBMPatientCovID, MBMCName, HOSName, CASTokenDate, " & _
          "MBMName, ClaimBordxStatus, MBMGroupCompany, PLNCode, ICDCode1, " & _
          "CLMTotalActual, CLMPayableByMedix2H,CLMPayableByMedix2M, BordxRefNo " & _
          " From VIEW_WS_PendingBordx where 0 = 0 "
    
    strCount = "select count(*) as totalRecord From VIEW_WS_PendingBordx where 0 = 0 "
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
        strCount = strCount + StrAppend
    End If

    Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rstCount!TotalRecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdBordx.Enabled = True
        CmdSearch.Enabled = True
        lblCurrent = 0
        lblTotal = 0
        LblTotalAmt = 0
    Else
         
    Do
        Do Until rst2.EOF
        Total = rstCount!TotalRecord
        lblTotal = Total
        
            With rst2
                Set Record = lstWorksheetTrackList.ListItems.Add(, , "")
        
                If Len(rst2!PAYCode) Then
                    Record.SubItems(1) = rst2!PAYCode
                End If
                If Len(rst2!CASNumber) Then
                    Record.SubItems(2) = rst2!CASNumber
                End If
                If Len(rst2!claimversion) Then
                    Record.SubItems(3) = rst2!claimversion
                End If
                If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                    Record.SubItems(4) = (rst2!CASNumber & "-" & rst2!claimversion)
                End If
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(5) = (rst2!CLMTotalActual)
                End If
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(6) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                If Len(rst2!CLMPayableByMedix2H) Then
                    Record.SubItems(7) = (rst2!CLMPayableByMedix2H)
                End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(8) = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(9) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(10) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                If rst2!MBMPatientCovID = "00" Then
                    If Len(rst2!MBMName) Then
                        Record.SubItems(11) = rst2!MBMName
                    End If
                Else
                    If Len(rst2!MBMCName) Then
                        Record.SubItems(11) = rst2!MBMCName
                    End If
                End If
                If Len(rst2!HOSName) Then
                    Record.SubItems(12) = rst2!HOSName
                End If
                If Len(rst2!MBMGroupCompany) Then
                    Record.SubItems(13) = rst2!MBMGroupCompany
                End If
                If Len(rst2!PLNCode) Then
                    Record.SubItems(14) = rst2!PLNCode
                End If
                If Len(rst2!ICDCode1) Then
                    Record.SubItems(15) = rst2!ICDCode1
                End If
                txtNoofDay = rst2!ClaimStatus
                'If Len(rst2!ClaimBordxStatus) Then
                    Text1 = rst2!ClaimBordxStatus
                'End If
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!CLMTotalApproved) Then
                    TotalAmt = TotalAmt + rst2!CLMTotalApproved
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
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
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdBordx.Enabled = True
    CmdSearch.Enabled = True
End Sub
'****************** End WorksheetTracking Listing *****************************

'****************** Start SaveCLMBordx *****************************

Private Sub SaveCLMBordx()
Dim strsqlCLMBordx As String
Dim ClaimBordx As New ADODB.Recordset
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String
    
    strsqlCLMBordx = "Select * from Bordx where BordxRefNo = '" & Trim(CLMBordxNo) & "'"
    ClaimBordx.Open strsqlCLMBordx, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    strsqlCLM = "Select * from Claim where BordxRefNo = '" & Trim(CLMBordxNo) & "'"
    CLM.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If ClaimBordx.recordCount = 0 Then
         ClaimBordx.AddNew
    'End If
        With ClaimBordx
            !BordxPAYCode = Mid(ChkStr(CboPayor), 1, 2)
            !BordxRefNo = ChkStr(CLMBordxNo)
            !BordxDate = Format(txtBordxDate, "dd/mm/yyyy")
            !BordxBy = Logon
            !BordxPrintedDate = Date
            !BordxPrintedBy = Logon
            !BordxClaimAmt = TotalApp
            !BordxNoClaims = TotCLM
            If Len(cboAccPeriodMth) Then
                !PAYToACPeriodMth = cboAccPeriodMth
            End If
            If Len(cboAccPeriodYR) Then
                !PAYToACPeriodYR = cboAccPeriodYR
            End If
            If txtToPAYDate <> "__/__/____" Then
                !BordxPAYSentDate = Format(txtToPAYDate, "dd/mm/yyyy")
                !BordxPayToStatus = "1"
            End If
            If txtFRPAYDate <> "__/__/____" Then
                !BordxPAYRecDocDate = Format(txtFRPAYDate, "dd/mm/yyyy")
                !BordxPayDocFRStatus = "1"
            End If
            If ToMNRBDate <> "__/__/____" Then
                !BordxMNRBSentDate = Format(ToMNRBDate, "dd/mm/yyyy")
                !BordxMNRBToStatus = "1"
            End If
            If txtFRMNRBDate <> "__/__/____" Then
                !BordxMNRBRecDocDate = Format(txtFRMNRBDate, "dd/mm/yyyy")
                !BordxMNRBDOCFRStatus = "1"
            End If



        End With
        ClaimBordx.Update
    Else
    
        With ClaimBordx
            If CLM.recordCount = 0 Then
                !BordxDate = Date
                !BordxPrintedDate = Date
                !BordxPrintedBy = Logon
                !BordxClaimAmt = CDbl(!BordxClaimAmt) + CDbl(TotalApp)
                !BordxNoClaims = CDbl(!BordxNoClaims) + CDbl(TotCLM)
            End If
                If Len(cboAccPeriodMth) Then
                    !PAYToACPeriodMth = cboAccPeriodMth
                End If
            If Len(cboAccPeriodYR) Then
                !PAYToACPeriodYR = cboAccPeriodYR
            End If
            If txtToPAYDate <> "__/__/____" Then
                !BordxPAYSentDate = Format(txtToPAYDate, "dd/mm/yyyy")
                !BordxPayToStatus = "1"
            End If
            If txtFRPAYDate <> "__/__/____" Then
                !BordxPAYRecDocDate = Format(txtFRPAYDate, "dd/mm/yyyy")
                !BordxPayDocFRStatus = "1"
            End If
            If ToMNRBDate <> "__/__/____" Then
                !BordxMNRBSentDate = Format(ToMNRBDate, "dd/mm/yyyy")
                !BordxMNRBToStatus = "1"
            End If
            If txtFRMNRBDate <> "__/__/____" Then
                !BordxMNRBRecDocDate = Format(txtFRMNRBDate, "dd/mm/yyyy")
                !BordxMNRBDOCFRStatus = "1"
            End If

            
        End With
        ClaimBordx.Update
    'End If
       
        
    End If
        
        ClaimBordx.Close
        CLM.Close
        Set ClaimBordx = Nothing
        Set CLM = Nothing
End Sub
'****************** Start lstWorksheetTrackListItemCheck *****************************

Private Sub CmdBordx_Click()

On Error GoTo errSave
Dim PAYCode As String
Dim ii As Integer
Dim Cnt As Integer
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim Update
Dim PayorIndex As Integer
Dim startTrans
    
    Cnt = 0
    RecordLeftOver = 0
    startTrans = False
    
    TotalApp = 0
    TotCLM = 0
    
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
      '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        
        If lstWorksheetTrackList.ListItems.Count = 0 Then
            MsgBox "Please select a record before proceeding!", vbOKOnly + vbCritical
        ElseIf UBound(CaseNoArray) = 0 Then
            MsgBox "Please select at least one record to be updated! ", vbCritical
        ElseIf txtBordxNo = "" Then
            MsgBox "Please Enter Bordx No! ", vbCritical
        ElseIf txtBordxDate = "__/__/____" Then
            MsgBox "Please Enter Bordx Date! ", vbCritical
        Else
            Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                startTrans = True
                medixmasterconnWS.beginTrans
                'PAYCode = Mid(ChkStr(CboPayor), 1, 2)
                CLMBordxNo = txtBordxNo
                
                For ii = 1 To UBound(CaseNoArray)
                 If Trim(CaseNoArray(ii)) <> "" Then
                    CASNo = Trim(Mid(CaseNoArray(ii), 1, IIf(InStr(1, CaseNoArray(ii), "-") = 0, 10, InStr(1, CaseNoArray(ii), "-") - 1)))
                    CASVer = Trim(Mid(CaseNoArray(ii), InStr(1, CaseNoArray(ii), "-") + 1, Len(CaseNoArray(ii)) - InStr(1, CaseNoArray(ii), "-")))
                    
                    strsqlCLM = "Select casNumber , ClaimVersion ,CLMTotalApproved From Claim where casNumber ='" & Trim(CASNo) & "' And ClaimVersion = '" & Trim(CASVer) & "'"
                    CLM.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
                    TotalApp = TotalApp + CLM!CLMTotalApproved
                    TotCLM = TotCLM + 1
                    CLM.Close
                    Set CLM = Nothing
                    Call UpdateRecord
                    Call UpdateMNRB
                    recordFound = True
                 End If
               Next ii
               Erase CaseNoArray
               Arraycnt = 0
               ReDim CaseNoArray(0)

               Call SaveCLMBordx
                
                If startTrans = True Then
                    medixmasterconnWS.commitTrans
                End If
                
                lstWorksheetTrackList.ListItems.Clear
                LblTotalAmt = 0
                lblTotal = 0
                lblCurrent = 0
                lblItemChecked = 0
                'Call WorksheetTrackingListing(searchCriteria, True)
            'End If
                MsgBox "Record Saved", vbOKOnly + vbInformation, DialogBoxTitle
                TxtCaseNo.SetFocus
                'lblBordxNo = CLMBordxNo
            End If
        End If
'    medixmasterconnWS.Close
'    Set medixmasterconnWS = Nothing
    
 '   medixmasterconnPayment.Close
 '   Set medixmasterconnPayment = Nothing
    
    Exit Sub
errSave:
    MsgBox Err.Description
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
    If startTrans = True Then
        medixmasterconnWS.RollbackTrans
    End If
    TxtCaseNo.SetFocus
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    CboPayor.Clear
    lstWorksheetTrackList.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    lblItemChecked = 0
    lblBordxNo = ""
    ReDim CaseNoArray(0)
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
    Call ComboListing
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End Sub

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub CmdSearch_Click()
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
    CmdBordx.Enabled = False
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    ReDim CaseNoArray(0)
    If CboPayor = "" Then
        MsgBox "Please select a Payor", vbOKOnly + vbCritical, DialogBoxTitle
        CboPayor.SetFocus
    Else
        CmdClear.Enabled = False
        CmdExit.Enabled = False
        CmdBordx.Enabled = False
        CmdSearch.Enabled = False
        Call SearchRecord
    End If
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    CmdBordx.Enabled = True
    
    Screen.MousePointer = Default

End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdBordx.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub Form_Load()
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    ReDim CaseNoArray(0)
    lblItemChecked = 0
End Sub


Private Sub lstWorksheetTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'**********************Start Sorted ListView *******************************
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        'SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        'SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        'SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(14)
    End Select
End Sub
'**********************End Sorted ListView *******************************


Private Sub lstWorksheetTrackList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Static Arraycnt As Integer
Dim temp
Dim cmd As New ADODB.Command

 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute

    If Item.Checked = True Then

            Arraycnt = IIf(UBound(CaseNoArray) = 0, 1, UBound(CaseNoArray) + 1)
            ReDim Preserve CaseNoArray(Arraycnt)
            CaseNoArray(Arraycnt) = Item.SubItems(4)
            If IsNumeric(Item.SubItems(6)) = False Then
                Item.SubItems(6) = 0
            End If
            lblItemChecked = CDbl(lblItemChecked) + CDbl(Item.SubItems(6))
            'End If
    Else
    
        ' Remove from the array
        newbound = 0
    
        If UBound(CaseNoArray) <> 0 Then
            For i = 0 To UBound(CaseNoArray)
                If (Item.SubItems(2) & "-" & Item.SubItems(3)) <> CaseNoArray(i) Then
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
            lblItemChecked = CDbl(lblItemChecked) - CDbl(Item.SubItems(6))
        Else
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            Arraycnt = 0
        End If
    End If
    PAY.Close
    Set PAY = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub
'****************** End lstWorksheetTrackListItemCheck *****************************

'****************** Start UpdateRecord *****************************

Private Sub UpdateRecord()
Dim Update
Dim strsqlCLMBordx As String
Dim ClaimBordx As New ADODB.Recordset
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim strsqlINV As String
Dim INV As New ADODB.Recordset
Dim strsqlMNRB As String
Dim StrSql As String
Dim INVSearch As New ADODB.Recordset

    strsqlCLM = " update Claim  set " & _
                " BordxRefNo =  '" & CLMBordxNo & "'," & _
                " ClaimStatus = 'C' ," & _
                " ClaimBordxStatus = 1 " & _
                " where casNumber ='" & Trim(CASNo) & "' And ClaimVersion = '" & Trim(CASVer) & "'"
            
    strsqlINV = "Select * from ClaimInvoice where CASNumber = '" & Trim(CASNo) & "' And INVWSversion = '" & Trim(CASVer) & "'"
    INVSearch.Open strsqlINV, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If INVSearch.recordCount Then
        strsqlINV = " update ClaimInvoice  set " & _
                    " INVBordxRefNo =  '" & CLMBordxNo & "'," & _
                    " INVBordxed = 1 " & _
                    " where CASNumber = '" & Trim(CASNo) & "' And INVWSversion = '" & Trim(CASVer) & "'"
        medixmasterconnWS.Execute strsqlINV
    Else
        INVSearch.AddNew
        With INVSearch
            !CASNumber = CASNo
            !INVWSversion = CASVer
            !INVBordxed = 1
            !INVBordxRefNo = CLMBordxNo
            !INVNo = ""
        End With
        INVSearch.Update
        'INVSearch.Update
        INVSearch.Close
        Set INVSearch = Nothing
    End If
    medixmasterconnWS.Execute strsqlCLM
    
End Sub
'****************** End UpdateRecord *****************************
'****************** Start Search Record *****************************

Private Function SearchRecord()
Dim StrSql As String
Dim StrAppend As String
Dim sql As String
Dim sql2 As String
Dim Claim As New ADODB.Recordset
    
    
    StrAppend = ""
    
    If Len(Trim(CboPayor)) Then
        StrAppend = StrAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If CboAmt.ListIndex = 0 Then
        StrAppend = StrAppend + " AND CLMTotalApproved <= 2000"
    End If
    
    If CboAmt.ListIndex = 1 Then
        StrAppend = StrAppend + " AND CLMTotalApproved > 2000"
    End If
    
    If chkBordxNo.Value = 1 Then
        StrAppend = StrAppend + " AND (BordxRefNo Is null OR BordxRefNo = '')"
    End If
    
    If Len(Trim(txtDOADate1)) And IsDate(txtDOADate1) = True Then
        StrAppend = StrAppend + " And CasDateAdmitted >= '" & Format(txtDOADate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOADate2)) And IsDate(txtDOADate2) = True Then
        StrAppend = StrAppend + " And CasDateAdmitted <= '" & Format(txtDOADate2, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(txtDODDate1)) And IsDate(txtDODDate1) = True Then
        StrAppend = StrAppend + " And CasDischargeDate >= '" & Format(txtDODDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDODDate2)) And IsDate(txtDODDate2) = True Then
        StrAppend = StrAppend + " And CasDischargeDate <= '" & Format(txtDODDate2, "mm/dd/yyyy") & "'"
    End If

    
    If Len(TxtCaseNo) Then
        StrAppend = StrAppend + " AND CASNumber >= '" & Trim(TxtCaseNo) & "'"
    End If

    If Len(TxtCaseNo2) Then
        StrAppend = StrAppend + " AND CASNumber <= '" & Trim(TxtCaseNo2) & "'"
    End If

    If txtBordxFr <> "" Then
        StrAppend = StrAppend + " AND BordxRefNo >= '" & Trim(txtBordxFr) & "'"
    End If

    If txtBordxTo <> "" Then
        StrAppend = StrAppend + " AND BordxRefNo <= '" & Trim(txtBordxTo) & "'"
    End If

    searchCriteria = StrAppend
    Call WorksheetTrackingListing(StrAppend)

End Function
'****************** End Search Record *****************************

    
Private Sub ComboListing()
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
    
    Do Until PAY.EOF
        With PAY
            CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
End Sub
    

Private Sub ToMNRBDate_LostFocus()
    If IsDate(ToMNRBDate) Then
    Else
        If ToMNRBDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            ToMNRBDate.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDate_LostFocus()
    If IsDate(txtBordxDate) Then
    Else
        If txtBordxDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate.SetFocus
        End If
    End If

End Sub

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

'**********************Start Update all Record functions*******************************

Private Sub UpdateMNRB()
Dim Update
Dim StrSql As String
Dim ToMNRB As New ADODB.Recordset

   StrSql = "Select * From ReceiptFrMNRB where BordxNo ='" & Trim(txtBordxNo) & "'"
      
   ToMNRB.Open StrSql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
       
    If ToMNRB.recordCount = 0 Then
        ToMNRB.AddNew
    End If
      With ToMNRB
        If Len(txtRECNo) Then
            !RECEIPTNO = txtRECNo
        End If
        If txtRECDate <> "__/__/____" Then
            !ReceiptDate = Format(txtRECDate, "DD/MM/YYYY")
        End If
        If Len(txtChequeAmt) Then
            !ReceiptAmt = txtChequeAmt
        End If
        If Len(txtRECNo) Then
            !ChequeNo = txtRECNo
        End If
        If Len(txtBordxNo) Then
            !BordxNo = txtBordxNo
        End If
        If Len(CASNo) Then
            !CASNumber = CASNo
        End If
        If Len(CASVer) Then
            !claimversion = CASVer
        End If
        If Len(lblItemChecked) Then
            !RECMNRBAmt = lblItemChecked
        End If
        
        '!BordxMNRBFrStatus = "F"
        !MNRBLastUpdateDate = Date
        !MNRBLastUpdateUser = Logon
        
      End With
      ToMNRB.Update
      ToMNRB.Close
      Set ToMNRB = Nothing

End Sub
'**********************End Update all Record functions*******************************


'Private Sub txtFRMNRBDate_LostFocus()
'    If IsDate(txtFRMNRBDate) Then
'    Else
'        If txtFRMNRBDate <> "__/__/____" Then
'            MsgBox "Invalid Date Format! ", vbCritical
'            txtFRMNRBDate.SetFocus
'        End If
'    End If
'End Sub

Private Sub txtFRPAYDate_LostFocus()
    If IsDate(txtFRPAYDate) Then
    Else
        If txtFRPAYDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtFRPAYDate.SetFocus
        End If
    End If
End Sub

Private Sub txtToPAYDate_LostFocus()
    If IsDate(txtToPAYDate) Then
    Else
        If txtToPAYDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtToPAYDate.SetFocus
        End If
    End If
End Sub
