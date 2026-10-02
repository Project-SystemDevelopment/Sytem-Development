VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEnquiryPaymentHospital 
   Caption         =   "Enquiry Payment to Hospital"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   9030
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   0
      TabIndex        =   41
      Top             =   7200
      Width           =   11775
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9240
         TabIndex        =   46
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10800
         TabIndex        =   45
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10200
         TabIndex        =   44
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   8640
         TabIndex        =   43
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   7920
         TabIndex        =   42
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label LblAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3600
         TabIndex        =   52
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   1680
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   49
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label29 
         Caption         =   "Pay Hosp Amt"
         Height          =   255
         Left            =   2520
         TabIndex        =   48
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1995
      Left            =   60
      TabIndex        =   23
      Top             =   240
      Width           =   11715
      Begin VB.CheckBox ChkSunNo 
         Height          =   255
         Left            =   11280
         TabIndex        =   59
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox ChkSunDate 
         Height          =   255
         Left            =   11280
         TabIndex        =   58
         Top             =   1590
         Width           =   375
      End
      Begin VB.CheckBox chkPayor 
         Height          =   255
         Left            =   5240
         TabIndex        =   56
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox ChkHospital 
         Height          =   255
         Left            =   5240
         TabIndex        =   55
         Top             =   1040
         Width           =   375
      End
      Begin VB.CheckBox chkPVNo 
         Height          =   255
         Left            =   11280
         TabIndex        =   15
         Top             =   810
         Width           =   375
      End
      Begin VB.CheckBox chkPVDate 
         Height          =   255
         Left            =   11280
         TabIndex        =   22
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox ChkInvNo 
         Height          =   255
         Left            =   11280
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox ChkInvDate 
         Height          =   255
         Left            =   11280
         TabIndex        =   12
         Top             =   510
         Width           =   375
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   180
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox TxtWSNoTo 
         Height          =   310
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   2
         Top             =   160
         Width           =   1560
      End
      Begin VB.TextBox TxtWSNo 
         Height          =   310
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   0
         Top             =   160
         Width           =   1560
      End
      Begin VB.TextBox TxtInvNo 
         Height          =   285
         Left            =   8160
         MaxLength       =   20
         TabIndex        =   8
         Top             =   220
         Width           =   3015
      End
      Begin VB.ComboBox cboWSStatus 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   440
         Width           =   3615
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   3615
      End
      Begin VB.ComboBox CboHos 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1000
         Width           =   3615
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   1280
         Width           =   3615
      End
      Begin VB.ComboBox CboPayment 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   3615
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   8160
         TabIndex        =   10
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPVNo1 
         Height          =   315
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   13
         Top             =   760
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtPVDate1 
         Height          =   315
         Left            =   8160
         TabIndex        =   16
         Top             =   1035
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   11
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPVNo2 
         Height          =   315
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   14
         Top             =   760
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDate2 
         Height          =   315
         Left            =   9840
         TabIndex        =   17
         Top             =   1035
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtSunNo1 
         Height          =   285
         Left            =   8160
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtSUNDate1 
         Height          =   285
         Left            =   8160
         TabIndex        =   20
         Top             =   1575
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtSunNo2 
         Height          =   285
         Left            =   9840
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtSUNDate2 
         Height          =   285
         Left            =   9840
         TabIndex        =   21
         Top             =   1575
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "SUN PV No"
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
         Left            =   6720
         TabIndex        =   62
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "SUN PV Date"
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
         Left            =   6720
         TabIndex        =   61
         Top             =   1620
         Width           =   1455
      End
      Begin VB.Label Label32 
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
         Left            =   9480
         TabIndex        =   60
         Top             =   1350
         Width           =   375
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5260
         TabIndex        =   57
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "WS Claimability"
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
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Hospital"
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
         Left            =   120
         TabIndex        =   37
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Payor"
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
         Left            =   120
         TabIndex        =   36
         Top             =   1300
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Inv No"
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
         Left            =   6720
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Pymt Status"
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
         Left            =   120
         TabIndex        =   34
         Top             =   1600
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "PV No"
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
         Left            =   6720
         TabIndex        =   33
         Top             =   840
         Width           =   855
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
         Left            =   9480
         TabIndex        =   32
         Top             =   825
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "PV Date"
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
         Left            =   6720
         TabIndex        =   31
         Top             =   1125
         Width           =   855
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
         Left            =   9480
         TabIndex        =   30
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "Inv Date"
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
         Left            =   6720
         TabIndex        =   29
         Top             =   530
         Width           =   975
      End
      Begin VB.Label Label21 
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
         Left            =   9480
         TabIndex        =   28
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label12 
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
         Left            =   11280
         TabIndex        =   27
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label25 
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
         Left            =   3240
         TabIndex        =   26
         Top             =   160
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Clm No"
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
         Left            =   120
         TabIndex        =   25
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label35 
         Caption         =   "Health Type"
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
         Left            =   120
         TabIndex        =   24
         Top             =   760
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lstEnquiryPaymentHos 
      Height          =   4935
      Left            =   0
      TabIndex        =   38
      Top             =   2280
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8705
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
      NumItems        =   26
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pymt Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hosp Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Inv No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "INV Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Pay2HOS"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amt Paid"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "OS Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Patient Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Clm No"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "AR Status"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "DOA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "DOD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "DOP (Payment)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Entry By"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "PV HOS Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "PV No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Cheque No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   18
         Text            =   "Cheque Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "A/C Period"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "Bank Charges"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Receipt No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Text            =   "Receipt Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   23
         Text            =   "WS Claimability"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "SUN PV"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "SUN PVDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Enquiry Payment to Hospital"
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
      TabIndex        =   53
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
      TabIndex        =   40
      Top             =   7800
      Width           =   1575
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
      Height          =   255
      Left            =   2520
      TabIndex        =   39
      Top             =   7800
      Width           =   2295
   End
End
Attribute VB_Name = "FrmEnquiryPaymentHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Source As Integer
Public intMBMAccount As Integer
Public intExit As Integer
Dim CaseNoArray() As String
Dim one As Boolean
Dim searchCriteria As String
Private HOspitals As String
Dim ScanCode As String
Dim Check As Boolean
Dim listloop1 As Integer
Dim PVNumber As String
Private m_blnDirection(1 To 23) As Boolean

Private Sub cboHLTcode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHLTCode.Text, cboHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode.Text, cboHLTCode.SelStart + cboHLTCode.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboHLTCode.ListCount - 1
       
        If NewText = LCase(Left(cboHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHLTCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboHLTCode.Text = ValidValue
               cboHLTCode.SelStart = 0
               cboHLTCode.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub ChkHospital_Click()
Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String
    CboHos.Clear
    If ChkHospital.Value = 1 Then
        tmpSelField = "SELECT HOSCode as tmpCode, HOSName as tmpName "
        tmpTable = "HOS"
        tmpSelCriteria = "0=0"
        Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, CboHos)
    End If
End Sub

Private Sub chkPayor_Click()
Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String

    CboPayor.Clear
    If chkPayor.Value = 1 Then
        tmpSelField = "SELECT PAYCode as tmpCode, PAYCompanyName as tmpName "
        tmpTable = "PAY"
        tmpSelCriteria = "0=0"
        If Trim(cboHLTCode) <> "" Then
            tmpSelCriteria = " HLTCode='" & Mid(cboHLTCode, 1, 1) & "'"
        End If
        Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, CboPayor)
    End If
End Sub

Private Sub ChkScan_Click()
    lstEnquiryPaymentHos.ListItems.Clear
    If ChkScan.Value = 1 Then
        Label25.Visible = False
        TxtWSNoTo.Visible = False
        ChkScan.Left = 3240
        TxtWSNo = ""
        TxtWSNo.SetFocus
    Else
        Label25.Visible = True
        TxtWSNoTo.Visible = True
        ChkScan.Left = 5280
        TxtWSNo.SetFocus
    End If
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

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        'Call ExportToExcelAnalysis1(lstEnquiryPaymentHos)
    If lstEnquiryPaymentHos.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstEnquiryPaymentHos)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Public Sub cmdSearch_Click()
Dim listrecord As Integer
    CmdSearch.Enabled = False
    'medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

    ScanCode = TxtWSNo
    If Mid(ScanCode, 1, 2) = "LO" Then
        TxtWSNo = ""
        Call SearchForm(ScanCode)
    Else
        
        Screen.MousePointer = vbHourglass
        LblAmt = 0
        one = False
        Erase CaseNoArray
        ReDim CaseNoArray(0)
        If ChkScan.Value = 1 Then
            If TxtWSNo = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtWSNo.SetFocus
            Else
                listrecord = lstEnquiryPaymentHos.ListItems.Count
                If listrecord <> 0 Then
                    
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                    If Check = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    ElseIf Check = False Then
                        Call SearchRecord
                    End If
                    Check = False
                Screen.MousePointer = Default
                TxtWSNo = ""
                TxtWSNo.SetFocus
            End If
            TxtWSNo.SetFocus
        Else
            Call SearchRecord
        End If
    End If
    'medixmasterconnPayment.Close
    'Set medixmasterconnPayment = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    Check = False
    TxtWSNo.SetFocus
    CmdExit.Enabled = True
End Sub

Private Sub Form_Load()
    ReDim CaseNoArray(0)
    Call ComboListing
    CboPayment.ListIndex = 0
    intMBMAccount = 1
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblAmt = 0
    Label25.Visible = False
    TxtWSNoTo.Visible = False
    Check = False
    one = False
    'TxtWSNo.SetFocus
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
'    CmdSave.Enabled = True
    one = False
    ReDim CaseNoArray(0)
    lstEnquiryPaymentHos.ListItems.Clear
    CboPayment.ListIndex = 0
    lblCurrent = 0
    lblTotal = 0
    LblAmt = 0
    Check = False
    ChkScan.Value = 1
    TxtWSNo.SetFocus
End Sub


Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

'************************Start ComboListing**************************

Private Sub ComboListing()
Dim HOS As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim HLT As New ADODB.Recordset
Dim cmd As New ADODB.Command

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHLTType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HLT.CursorLocation = adUseClient
    HLT.CursorType = adOpenForwardOnly
    HLT.CacheSize = 100
    Set HLT = cmd.Execute
    
    Do Until HLT.EOF
        With HLT
            cboHLTCode.AddItem ChkStr(!HLTCode) & " - " & ChkStr(!HLTDescription)
            HLT.MoveNext
        End With
    Loop
    HLT.Close
    Set HLT = Nothing

    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
    With CboPayment
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
        .AddItem "3 - Partial Paid"
        .AddItem "4 - Paid & Partial Paid"
        .AddItem "5 - Unpaid & Partial Paid"
    End With
    

    cboWSStatus.Clear
    cboWSStatus.AddItem "AA" & " - " & "Accepted"
    cboWSStatus.AddItem "RR" & " - " & "Repudiated"
    cboWSStatus.AddItem "CC" & " - " & "Cancelled"
    cboWSStatus.AddItem "SA" & " - " & "Special Approval"
    cboWSStatus.AddItem "PP" & " - " & "Pay & Claim"
    cboWSStatus.AddItem "NN" & " - " & "Not Decided"
    cboWSStatus.AddItem "NA" & " - " & "Not Decided/Accepted"
    cboWSStatus.AddItem "NR" & " - " & "Not Decided/Repudiated"
    cboWSStatus.AddItem "PA" & " - " & "Pay&Claim/Accepted"
    cboWSStatus.AddItem "PR" & " - " & "Pay&Claim/Rejected"
    cboWSStatus.AddItem "AR" & " - " & "Accepted/Rejected"
    cboWSStatus.AddItem "RA" & " - " & "Rejected/Accepted"
    cboWSStatus.AddItem "BR" & " - " & "BTBG/Rejected"

End Sub
'************************Start ComboListing**************************

Private Function SearchRecord()
Dim strSql As String
Dim strAppend As String
Dim sql As String
    
    strAppend = ""
    HOspitals = ""
            
    If Len(Trim(CboHos)) Then
        HOspitals = Replace(CboHos, "'", "''")
        strAppend = strAppend + " AND HOSName = '" & Trim(HOspitals) & "'"
    End If
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    If Len(Trim(cboWSStatus)) Then
        strAppend = strAppend + " and CLMWSStatus = '" & Mid(cboWSStatus, 1, 2) & "'"
    End If
    If Trim(cboHLTCode) <> "" Then
        strAppend = strAppend + " and HLTCode = '" & Mid(cboHLTCode, 1, 1) & "'"
    End If
    If Mid(CboPayment, 1, 1) = 1 Then
        strAppend = strAppend + " and PVPayHOSStatus = 'F'"
    End If
 
    If Mid(CboPayment, 1, 1) = 2 Then
        strAppend = strAppend + " and (PVPayHOSStatus = '' or PVPayHOSStatus is null) "
    End If
 
    If Mid(CboPayment, 1, 1) = 3 Then
        strAppend = strAppend + " and PVPayHOSStatus = 'P'"
    End If
 
    If Mid(CboPayment, 1, 1) = 4 Then
        strAppend = strAppend + " and (PVPayHOSStatus = 'P' or PVPayHOSStatus = 'F') "
    End If
            
    If Mid(CboPayment, 1, 1) = 5 Then
        strAppend = strAppend + " and (PVPayHOSStatus = '' or PVPayHOSStatus is null or PVPayHOSStatus = 'P') "
    End If
            
    If ChkScan.Value = 1 Then
        strAppend = strAppend + " AND CASNumber = '" & Trim(TxtWSNo) & "'"
    Else
        If Len(TxtWSNo) Then
            strAppend = strAppend + " AND CASNumber >= '" & Trim(TxtWSNo) & "'"
        End If
    
        If Len(TxtWSNoTo) Then
            strAppend = strAppend + " AND CASNumber <= '" & Trim(TxtWSNoTo) & "'"
        End If
    End If
    
    If Len(TxtInvNo) Then
        strAppend = strAppend + " and InvNo = '" & Trim(TxtInvNo) & "'"
    End If
    
    If ChkInvNo.Value = 1 Then
        sql = ""
        sql = "AND (InvNo Is null OR InvNo = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If (txtInvDate1) <> "__/__/____" And IsDate(txtInvDate1) = True _
        And (txtInvDate2) <> "__/__/____" And IsDate(txtInvDate2) = True Then
        sql = ""
        sql = " AND InvDate >= '" & Format(Trim(txtInvDate1), "mm/dd/yyyy") & _
              "' And  InvDate <= '" & Format(Trim(txtInvDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
        strAppend = strAppend + " And InvDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDate2)) > 0 And IsDate(txtInvDate2) = True Then
        strAppend = strAppend + " And InvDate <= '" & Format(txtInvDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkInvDate.Value = 1 Then
        sql = ""
        sql = "AND (PVDate Is null OR PVDate = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If Len(txtPVNo1) Then
        strAppend = strAppend + " and PVHOSNo >= '" & Trim(txtPVNo1) & "'"
    End If
        
    If Len(txtPVNo2) Then
        strAppend = strAppend + " and PVHOSNo <= '" & Trim(txtPVNo2) & "'"
    End If
    
    If ChkPVNo.Value = 1 Then
        sql = ""
        sql = "AND (PVNo Is null OR PVNo = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If (txtPVDate1) <> "__/__/____" And IsDate(txtPVDate1) = True _
        And (txtPVDate2) <> "__/__/____" And IsDate(txtPVDate2) = True Then
        sql = ""
        sql = " AND PVHOSDate >= '" & Format(Trim(txtPVDate1), "mm/dd/yyyy") & _
              "' And PVHOSDate <= '" & Format(Trim(txtPVDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtPVDate1)) > 0 And IsDate(txtPVDate1) = True Then
        strAppend = strAppend + " And PVHOSDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDate2)) > 0 And IsDate(txtPVDate2) = True Then
        strAppend = strAppend + " And PVHOSDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkPVDate.Value = 1 Then
        sql = ""
        sql = "AND (PVHOSDate Is null OR PVHOSDate = '') " & ""
        strAppend = strAppend + sql
    End If
        
    If Len(txtSunNo1) Then
        strAppend = strAppend & " and PVSunHOSNo >= '" & Trim(txtSunNo1) & "'"
    End If
    If Len(txtSunNo2) Then
        strAppend = strAppend & " and PVSunHOSNo <= '" & Trim(txtSunNo2) & "'"
    End If
    If ChkSunNo.Value = 1 Then
        sql = ""
        sql = "AND (PVSunHOSNo Is null OR PVSunHOSNo = '') " & ""
        strAppend = strAppend & sql
    End If
    
    If Len(Trim(txtSUNDate1)) > 0 And IsDate(txtSUNDate1) = True Then
        strAppend = strAppend & " And PVSUNHOSDate >= '" & Format(txtSUNDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtSUNDate2)) > 0 And IsDate(txtSUNDate2) = True Then
        strAppend = strAppend & " And PVSUNHOSDate <= '" & Format(txtSUNDate2, "mm/dd/yyyy") & "'"
    End If
    
    searchCriteria = strAppend
    Call Payment2HOSListing(strAppend)

End Function


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

Private Sub CboPayment_Change()
    If InStr(CboPayment.Text, "null") Then
       CboPayment.Enabled = False
    End If
End Sub

Private Sub CboPayment_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(CboPayment.Text, CboPayment.SelStart) + Chr(KeyAscii) + Mid(CboPayment.Text, CboPayment.SelStart + CboPayment.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayment.ListCount - 1
 
        If NewText = LCase(Left(CboPayment.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayment.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayment.Text = ValidValue
                CboPayment.SelStart = 0
                CboPayment.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub Payment2HOSListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim INVAmt As Variant
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim Check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim CASNumber As String
Dim WSVersion As String
Dim cmd1 As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter
Dim RECFound As Boolean
Dim tmpName As String
Dim tmpReceiptNo As String
Dim tmpReceiptAmt As Double

    total = 0
    Current = 0
    If ChkScan.Value = 0 Then
         lstEnquiryPaymentHos.ListItems.Clear
    End If
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
'    sql = " Select CASNumber, ClaimVersion, MBMName, PAYCode,HLTCode,PAYCheckAmt,PAYReceiptNo, " & _
'          " BordxRefNo, PVPayHOSStatus, HOSName, MBMCName,InvDate, CASARStatus, " & _
'          " INVAmount, MBMPatientCovID, CASDischargeDate,CLMPayableByMedix2H, ReceiptNo," & _
'          " CASDateAdmitted, HOSPaidAmt, InvNo, ClaimType, PVHOSNo, PVHOSChequeDate,PVBankCharges," & _
'          " PVHOSDate, PVHOSChequeNo, INVEnteredBy, PVHOSIndex, HOSAccPeriodMth, HOSAccPeriodYr,CheckAmt " & _
'          "  From VIEW_Payment2HOS where 0 = 0 "
'    rst2.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
    RECFound = False
    Set cmd1.ActiveConnection = medixmasterconnPayment
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_Payment"
    Set prm1 = cmd1.CreateParameter("PaymentType", adVarChar, adParamInput, 2000, "Pay2HospEnq")
    cmd1.Parameters.Append prm1
    Set prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, strAppend)
    cmd1.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd1.Execute
    Do
        Do Until rst2.EOF
        RECFound = True
        txtPVStatus = ""
        If Len(rst2!PVPayHOSStatus) Then
            txtPVStatus = rst2!PVPayHOSStatus
        End If
        With rst2
            If Trim(txtPVStatus) = "F" Then
                Set Record = lstEnquiryPaymentHos.ListItems.Add(, , "Full")
            ElseIf Trim(txtPVStatus) = "P" Then
                Set Record = lstEnquiryPaymentHos.ListItems.Add(, , "Partial")
            Else
                Set Record = lstEnquiryPaymentHos.ListItems.Add(, , "Unpaid")
            End If
            Record.SubItems(1) = IIf(Trim(rst2!HosName) <> "", Trim(rst2!HosName), "")
            Record.SubItems(2) = IIf(Len(rst2!INVNo), rst2!INVNo, "")
            Record.SubItems(3) = IIf(Len(rst2!InvDate), rst2!InvDate, "")
            If Len(rst2!CLMPayableByMedix2H) Then
                Record.SubItems(4) = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
            End If
            If Len(rst2!HOSPaidAmt) Then
                Record.SubItems(5) = Format(rst2!HOSPaidAmt, "#,##0.00")
            End If
            tmpName = IIf(Len(rst2!MBMName), Trim(rst2!MBMName), "")
            If rst2!MBMPatientCovID <> "00" Then
                tmpName = IIf(Len(rst2!MBMCName), Trim(rst2!MBMCName), "")
            End If
            Record.SubItems(7) = tmpName
            If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                Record.SubItems(8) = (rst2!CASNumber & "-" & rst2!claimversion)
            End If
            CASNumber = rst2!CASNumber
            If Len(rst2!claimversion) Then
                WSVersion = Trim(rst2!claimversion)
            End If
            Record.SubItems(9) = IIf(Len(rst2!BordxRefNo), Trim(rst2!BordxRefNo), "")
            Record.SubItems(10) = IIf(Len(rst2!CASARStatus), rst2!CASARStatus, "")
            If Len(rst2!CASDateAdmitted) Then
                Record.SubItems(11) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
            End If
            If Len(rst2!CASDischargeDate) Then
                Record.SubItems(12) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
            End If
            If Len(rst2!PVHOSDate) Then
                Record.SubItems(13) = Format(rst2!PVHOSDate, "dd/mm/yyyy")
            End If
            If Len(rst2!INVEnteredBy) Then
                Record.SubItems(14) = Trim(rst2!INVEnteredBy)
            End If
                
            If Len(rst2!PVHOSIndex) Then
                Record.SubItems(15) = rst2!PVHOSIndex
            End If
            
            If Len(rst2!PVHOSNo) Then
                Record.SubItems(16) = Trim(rst2!PVHOSNo)
            End If
            
            If Len(rst2!PVHOSChequeNo) Then
                Record.SubItems(17) = Trim(rst2!PVHOSChequeNo)
            End If
            
            If Len(!PVHOSChequeDate) Then
                Record.SubItems(18) = Trim(rst2!PVHOSChequeDate)
            End If
            
            If Len(rst2!HOSAccPeriodMth & "/" & rst2!HOSAccPeriodYr) Then
                Record.SubItems(19) = (rst2!HOSAccPeriodMth & "/" & rst2!HOSAccPeriodYr)
            End If
            
            If Len(rst2!PVBankCharges) Then
                Record.SubItems(20) = (rst2!PVBankCharges)
            End If
            tmpReceiptNo = ""
            tmpReceiptAmt = 0
            Call REceiptInfo(!CASNumber, !claimversion, tmpReceiptNo, tmpReceiptAmt)
            Record.SubItems(21) = tmpReceiptNo
            Record.SubItems(22) = tmpReceiptAmt
            Record.SubItems(23) = IIf(Len(rst2!CLMWSStatus), rst2!CLMWSStatus, "")
            Record.SubItems(24) = IIf(Len(!PVSUNHOSNo), Trim(!PVSUNHOSNo), "")
            Record.SubItems(25) = IIf(Len(!PVSUNHOSDate), Trim(!PVSUNHOSDate), "")
 
            strsqlPAY = "Select * from Payment2HOS where CASNumber = '" & Trim(CASNumber) & "' and ClaimVersion = '" & Trim(WSVersion) & "'"
            PAY.Open strsqlPAY, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            INVAmt = 0
            If Trim(txtPVStatus) = "P" Then
                Do Until PAY.EOF
                    INVAmt = INVAmt + PAY!HOSPaidAmt
                PAY.MoveNext
                Loop
                If Len(INVAmt) Then
                    Record.SubItems(5) = INVAmt
                    Record.SubItems(6) = rst2!CLMPayableByMedix2H - INVAmt
                End If
            ElseIf Trim(txtPVStatus) = "" Then
                If Len(rst2!CLMPayableByMedix2H) Then
                    Record.SubItems(6) = rst2!CLMPayableByMedix2H
                End If
            End If
            PAY.Close
            Set PAY = Nothing
            Current = Current + 1
            lblCurrent = Current
            TotalAmt = TotalAmt + rst2!CLMPayableByMedix2H
            LblAmt = Format(TotalAmt, "#,##0.00")
        End With
        rst2.MoveNext
        DoEvents
        If intExit = 1 Then
           Check = False
           Exit Do
        End If
        Loop
    Loop Until Check = False
    If RECFound = False Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub txtInvAmt1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtInvAmt2_KeyPress(KeyAscii As Integer)
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

Private Sub lstEnquiryPaymentHos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtDatetime, m_blnDirection(4)
    Case 5
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtDatetime, m_blnDirection(12)
    Case 13
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtDatetime, m_blnDirection(13)
    Case 14
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtDatetime, m_blnDirection(14)
    Case 15
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtDatetime, m_blnDirection(19)
    Case 20
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
    Case 22
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtString, m_blnDirection(22)
    Case 23
        SortListView lstEnquiryPaymentHos, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
    End Select
End Sub

Private Sub cboPayor_Change()
If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If

End Sub

Private Sub cboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
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
End Sub

Private Sub TxtWSNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtWSNo_LostFocus()
    TxtWSNo = ChkStr(TxtWSNo)
    CmdSearch.Default = False
End Sub

Private Sub TxtWSNoTo_LostFocus()
TxtWSNoTo = ChkStr(TxtWSNoTo)
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtWSNo)) = Trim(Mid(lstEnquiryPaymentHos.ListItems(listloop1).SubItems(8), 1, IIf(InStr(1, lstEnquiryPaymentHos.ListItems(listloop1).SubItems(8), "-") = 0, 10, InStr(1, lstEnquiryPaymentHos.ListItems(listloop1).SubItems(8), "-") - 1))) Then
        Check = True
    End If
End Function
Private Sub LoadComboListing(tmpSelField As String, tmpSelCriteria As String, tmpTableName As String, tmpComboBox As ComboBox)
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As New ADODB.Parameter
Dim prm2 As New ADODB.Parameter

    tmpComboBox.Clear
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelField)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, tmpSelCriteria)
    cmd.Parameters.Append Prm3
    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    
    Do Until TmpRec.EOF
        With TmpRec
        If tmpTableName = "HOS" Then
            tmpComboBox.AddItem !tmpName
        ElseIf tmpTableName = "PAY" Then
            tmpComboBox.AddItem !tmpCode & " - " & !tmpName
        End If
            TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

