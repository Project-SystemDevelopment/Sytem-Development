VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBordxTracking 
   Caption         =   "Bordx & Worksheet"
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   30
      Top             =   7080
      Width           =   11655
      Begin VB.TextBox LblTotalAmt3 
         Height          =   285
         Left            =   5400
         TabIndex        =   64
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox LblTotalAmt2 
         Height          =   285
         Left            =   4080
         TabIndex        =   63
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   2880
         TabIndex        =   62
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   8400
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9120
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10800
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10080
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6840
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7680
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Balance O/S"
         Height          =   255
         Left            =   5400
         TabIndex        =   51
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "From Payor "
         Height          =   255
         Left            =   4080
         TabIndex        =   50
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Worksheet"
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin MSComctlLib.ListView lstBordxTrackList 
      Height          =   4695
      Left            =   120
      TabIndex        =   34
      Top             =   2400
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8281
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
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Payor"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Bordx Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Bordx Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Bordx Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Wks No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Wks Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Received Amt (MNRB)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Received Amt (Payor)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Wks Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "DFMNRB (Money)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Receipt No (MNRB)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "DFPayor (Money)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Receipt No (Payor)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Approval Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Approval Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   35
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox chkBordxDate 
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
         Left            =   10920
         TabIndex        =   23
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox chkPayAppDate 
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
         Left            =   10920
         TabIndex        =   20
         Top             =   680
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6960
         TabIndex        =   52
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox ChkDFP 
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
         Left            =   10920
         TabIndex        =   14
         Top             =   160
         Width           =   495
      End
      Begin VB.CheckBox ChkDFMNRB 
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
         Left            =   10920
         TabIndex        =   17
         Top             =   420
         Width           =   495
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   200
         Width           =   3615
      End
      Begin MSMask.MaskEdBox TxtDFP2 
         Height          =   285
         Left            =   9240
         TabIndex        =   13
         Top             =   165
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDFMNRB2 
         Height          =   285
         Left            =   9240
         TabIndex        =   16
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDFP1 
         Height          =   285
         Left            =   7320
         TabIndex        =   12
         Top             =   165
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDFMNRB1 
         Height          =   285
         Left            =   7320
         TabIndex        =   15
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TxtClaimNo2 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   4
         Top             =   735
         Width           =   1575
      End
      Begin VB.TextBox TxtRecPAY2 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TxtRecMNRB2 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox TxtClaimNo1 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   3
         Top             =   735
         Width           =   1575
      End
      Begin VB.TextBox TxtRecPAY1 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TxtRecMNRB1 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtPAYAppCodeFr 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtPAYAppCodeTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox cboPayAppStatus 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtPAYAppDateTo 
         Height          =   285
         Left            =   9240
         TabIndex        =   19
         Top             =   675
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPAYAppDateFr 
         Height          =   285
         Left            =   7320
         TabIndex        =   18
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   285
         Left            =   9240
         TabIndex        =   22
         Top             =   915
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDatefr 
         Height          =   285
         Left            =   7320
         TabIndex        =   21
         Top             =   915
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         Caption         =   "To"
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
         TabIndex        =   60
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Bordx Date"
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
         Left            =   6000
         TabIndex        =   59
         Top             =   915
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Payor App Date"
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
         Left            =   6000
         TabIndex        =   58
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "To"
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
         Left            =   8880
         TabIndex        =   57
         Top             =   675
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Payor App Status"
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
         Left            =   240
         TabIndex        =   56
         Top             =   1730
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Payor App Code fr"
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
         Left            =   240
         TabIndex        =   55
         Top             =   1450
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   54
         Top             =   1450
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Rec. No (Pay) fr."
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
         TabIndex        =   48
         Top             =   980
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   47
         Top             =   980
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   46
         Top             =   1220
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Rec. No (MNRB)fr."
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
         Left            =   240
         TabIndex        =   45
         Top             =   1220
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "To"
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
         TabIndex        =   44
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Date Fr Payor"
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
         TabIndex        =   43
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "To"
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
         Left            =   8880
         TabIndex        =   42
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Date Fr MNRB"
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
         Left            =   6000
         TabIndex        =   41
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Claim No from"
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
         TabIndex        =   40
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   39
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   38
         Top             =   490
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No from"
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
         TabIndex        =   37
         Top             =   470
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
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
         TabIndex        =   36
         Top             =   200
         Width           =   1215
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Bordx and Worksheet"
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
      TabIndex        =   53
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmBordxTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Public searchCriteria As String
Private m_blnDirection(1 To 19) As Boolean

Private Sub CboPayor_Click()
Dim HLT As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    If Len(CboPayor) Then
        
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'strsql = "Select HLTCode from PAY where PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
        'HLT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "1"
        strAppend = " and PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "BordxWorksheet"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        HLT.CursorLocation = adUseClient
        HLT.CursorType = adOpenForwardOnly
        HLT.CacheSize = 100
        Set HLT = cmd.Execute
        
        Text1 = Trim(HLT!HLTCode)
        
        If Text1 = "S" Then
            Label16.Caption = "From MNRB"
        Else
            Label16.Caption = "From Payor"
        End If
            
        HLT.Close
        Set HLT = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
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
            
    If lstBordxTrackList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstBordxTrackList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
        
    'If lstBordxTrackList.listitems.count <> 0 Then
        'RptBordxWks.show
    'Else
        'MsgBox "Report cannot be printed without any record.", vbInformation
    'End If
    
End Sub

Private Sub CmdPrintFile_Click()
    Screen.MousePointer = vbHourglass
        'Call ExportToExcelBordx(lstBordxTrackList)
    If lstBordxTrackList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstBordxTrackList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    LblTotalAmt = 0
    LblTotalAmt2 = 0
    LblTotalAmt3 = 0
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call PayorComboListing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

'**********************Start Command Button *******************************

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lstBordxTrackList.ListItems.Clear
    lblCurrent = 0
    LblTotalAmt = 0
    LblTotalAmt2 = 0
    LblTotalAmt3 = 0
    CboPayor.Clear
    CboPayor.SetFocus
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call PayorComboListing
   ' medixmasterconn.Close
End Sub

Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub CmdSearch_Click()
On Error GoTo errRun

    CmdClear.Enabled = False
    cmdPrint.Enabled = False
    CmdPrintFile.Enabled = False
    CmdExit.Enabled = False
    CmdSearch.Enabled = False
    
    Screen.MousePointer = vbHourglass
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Call SearchRecord
  '  medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
    Screen.MousePointer = Default
    CmdSearch.Enabled = True
    Exit Sub
    
errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstBordxTrackList.ListItems.Clear
    lblCurrent = ""
    cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
End Sub
'**********************End Command Button *******************************

'**********************Start Search Record *******************************

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim CAS As New ADODB.Recordset
    
    strAppend = ""
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtBordx1)) Then
        strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
    End If
    
    If Len(Trim(TxtBordx2)) Then
        strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
    End If
    
    If Len(Trim(TxtClaimNo1)) Then
        strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtClaimNo1)) & "'"
    End If
    
    If Len(Trim(TxtClaimNo2)) Then
        strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtClaimNo2)) & "'"
    End If
    
    If Len(Trim(TxtRecPAY1)) Then
        strAppend = strAppend + " AND ReceiptNoPAY >= '" & RemStr(Trim(TxtRecPAY1)) & "'"
    End If
    
    If Len(Trim(TxtRecPAY2)) Then
        strAppend = strAppend + " AND ReceiptNoPAY <= '" & RemStr(Trim(TxtRecPAY2)) & "'"
    End If
    
    If Len(Trim(TxtRecMNRB1)) Then
        strAppend = strAppend + " AND ReceiptNoMNRB >= '" & RemStr(Trim(TxtRecMNRB1)) & "'"
    End If
    
    If Len(Trim(TxtRecMNRB2)) Then
        strAppend = strAppend + " AND ReceiptNoMNRB <= '" & RemStr(Trim(TxtRecMNRB2)) & "'"
    End If
          
    If IsDate(TxtDFP1) = True Then
        strAppend = strAppend + " And BordxPAYRecDate >= '" & Format(TxtDFP1, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(TxtDFP2) = True Then
        strAppend = strAppend + " And BordxPAYRecDate <= '" & Format(TxtDFP2, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(TxtDFMNRB1) = True Then
        strAppend = strAppend + " And BordxMNRBRecDate >= '" & Format(TxtDFMNRB1, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(TxtDFMNRB2) = True Then
        strAppend = strAppend + " And BordxMNRBRecDate <= '" & Format(TxtDFMNRB2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkDFP.Value = 1 Then
        sql = ""
        sql = " AND (BordxPAYRecDate Is null OR BordxPAYRecDate = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If ChkDFMNRB.Value = 1 Then
        sql = ""
        sql = " AND (BordxMNRBRecDate Is null OR BordxMNRBRecDate = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtPAYAppCodeFr)) Then
        strAppend = strAppend + " AND CLMPayorAppCode >= '" & RemStr(Trim(txtPAYAppCodeFr)) & "'"
    End If
    
    If Len(Trim(txtPAYAppCodeTo)) Then
        strAppend = strAppend + " AND CLMPayorAppCode <= '" & RemStr(Trim(txtPAYAppCodeTo)) & "'"
    End If
    
    If IsDate(txtPAYAppDateFr) = True Then
        strAppend = strAppend + " AND CLMPayorAppDate >= '" & Format(txtPAYAppDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtPAYAppDateTo) Then
        strAppend = strAppend + " AND CLMPayorAppDate <= '" & Format(txtPAYAppDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(cboPayAppStatus)) Then
        If Mid(cboPayAppStatus, 1, 1) = "Y" Then
            strAppend = strAppend + " AND CLMPayorAppStatus = 'Y'"
        Else
            strAppend = strAppend + " AND (CLMPayorAppStatus Is null OR CLMPayorAppStatus = '' OR CLMPayorAppStatus = 'N') " & ""
        End If
    End If
    
    If IsDate(txtBordxDatefr) Then
        strAppend = strAppend + " AND BordxDate >= '" & Format(txtBordxDatefr, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtBordxDateTo) Then
        strAppend = strAppend + " AND BordxDate <= '" & Format(txtBordxDateTo, "mm/dd/yyyy") & "'"
    End If
        
    If chkPayAppDate.Value = 1 Then
        sql = ""
        sql = " AND (CLMPayorAppDate Is null OR CLMPayorAppDate = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If chkBordxDate.Value = 1 Then
        sql = ""
        sql = " AND (BordxDate Is null OR BordxDate = '') " & ""
        strAppend = strAppend + sql
    End If
    
    searchCriteria = strAppend
    Call BordxTrackingListing(strAppend)
    
End Function
'********************** End Search Record *******************************

'**********************Start BordxTracking Listing *******************************

Private Sub BordxTrackingListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim ChkClaimNo As String
Dim TotalAmt As Variant
Dim TotalAmt2 As Variant
Dim TotalAmt3 As Variant
Dim tmpMgmtFees As Variant
    'Total = 0
    Current = 0
    ChkClaimNo = ""
    
    lstBordxTrackList.ListItems.Clear
            
    'sql = " Select CASNumber, BordxPAYCode, BordxClaimAmt, " & _
          " CheckPayorAmt, CheckMNRBAmt, BordxDate, " & _
          " BordxRefNo, CLMTotalApproved, ClaimOpendate, " & _
          " BordxStatus, ClaimVersion, BordxPAYRecDate, " & _
          " BordxMNRBRecDate, ReceiptNoMNRB, ReceiptNoPAY, " & _
          " MBMBTBG, CLMPayableByMedix2H " & _
          " From VIEW_Bordx_Wks2 where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxWorksheet"
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
        
    Do
        Do Until rst2.EOF
                        
            With rst2
                tmpMgmtFees = IIf(IsNumeric(rst2!CLMCaseMgmtFees), rst2!CLMCaseMgmtFees, 0)

                If ChkClaimNo <> rst2!CasNumber & "-" & rst2!claimversion Then
                    If Trim(rst2!MBMBTBG) = "Y" Then
                        TotalAmt = TotalAmt + IIf(IsNumeric(rst2!CLMPayableByMedix2H) = True, rst2!CLMPayableByMedix2H + tmpMgmtFees, 0)
                    Else
                        TotalAmt = TotalAmt + IIf(IsNumeric(rst2!CLMTotalApproved) = True, rst2!CLMTotalApproved + tmpMgmtFees, 0)
                    End If
                    
                    LblTotalAmt = Format(TotalAmt, "#,##0.00")
                Else
                    LblTotalAmt = Format(TotalAmt, "#,##0.00")
                End If
                
                If Len(rst2!BordxPAYCode) Then
                    Set Record = lstBordxTrackList.ListItems.Add(, , rst2!BordxPAYCode)
                Else
                    Set Record = lstBordxTrackList.ListItems.Add(, , "")
                End If
                
                If Len(rst2!BordxRefNo) Then
                        Record.SubItems(1) = rst2!BordxRefNo
                End If
                
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(2) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!BordxDate) Then
                   Record.SubItems(3) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!Bordxstatus) Then
                    Record.SubItems(4) = rst2!Bordxstatus
                End If
                If Len(rst2!CasNumber & "-" & rst2!claimversion) Then
                    Record.SubItems(5) = rst2!CasNumber & "-" & rst2!claimversion
                    ChkClaimNo = Trim(Record.SubItems(5))
                End If
                If Trim(rst2!MBMBTBG) = "Y" Then
                    If Len(rst2!CLMPayableByMedix2H) Then
                        Record.SubItems(6) = Format(rst2!CLMPayableByMedix2H + tmpMgmtFees, "#,##0.00")
                    End If
                Else
                    If Len(rst2!CLMTotalApproved) Then
                        Record.SubItems(6) = Format(rst2!CLMTotalApproved + tmpMgmtFees, "#,##0.00")
                    End If
                End If
                If Len(rst2!CheckMNRBAmt) Then
                    Record.SubItems(7) = Format(rst2!CheckMNRBAmt, "#,##0.00")
                End If
                If Len(rst2!CheckPayorAmt) Then
                    Record.SubItems(8) = Format(rst2!CheckPayorAmt, "#,##0.00")
                End If
                If Len(rst2!ClaimOpendate) Then
                    Record.SubItems(9) = Format(rst2!ClaimOpendate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxMNRBRecDate) Then
                    Record.SubItems(10) = Format(rst2!BordxMNRBRecDate, "dd/mm/yyyy")
                End If
                If Len(rst2!ReceiptNoMNRB) Then
                    Record.SubItems(11) = Trim(rst2!ReceiptNoMNRB)
                End If
                If Len(rst2!BordxPAYRecDate) Then
                    Record.SubItems(12) = Format(rst2!BordxPAYRecDate, "dd/mm/yyyy")
                End If
                If Len(rst2!ReceiptNoPAY) Then
                    Record.SubItems(13) = Trim(rst2!ReceiptNoPAY)
                End If
                
                If Len(rst2!MBMPolicyNo) Then
                    Record.SubItems(14) = rst2!MBMPolicyNo
                End If
                If Len(rst2!MBMGroupCompany) Then
                    Record.SubItems(15) = Format(rst2!MBMGroupCompany, "dd/mm/yyyy")
                End If
                If Trim(rst2!MBMPatientCovID) = "00" Then
                    Record.SubItems(16) = rst2!MBMName
                Else
                    Record.SubItems(16) = rst2!MBMCName
                End If
                If Len(rst2!CLMPayorAppCode) Then
                    Record.SubItems(17) = Trim(rst2!CLMPayorAppCode)
                End If
                If Len(rst2!CLMPayorAppDate) Then
                    Record.SubItems(18) = Format(rst2!CLMPayorAppDate, "dd/mm/yyyy")
                End If
                
                
                Current = Current + 1
                lblCurrent = Current
                
                If ChkClaimNo <> Record.SubItems(5) Then
                    If Trim(rst2!MBMBTBG) = "Y" Then
                        TotalAmt = TotalAmt + IIf(IsNumeric(rst2!CLMPayableByMedix2H) = True, rst2!CLMPayableByMedix2H + tmpMgmtFees, 0)
                    Else
                        TotalAmt = TotalAmt + IIf(IsNumeric(rst2!CLMTotalApproved) = True, rst2!CLMTotalApproved + tmpMgmtFees, 0)
                    End If
                    LblTotalAmt = Format(TotalAmt, "#,##0.00")
                Else
                    LblTotalAmt = Format(TotalAmt, "#,##0.00")
                End If
                
                If Mid(CboPayor, 1, 2) <> "" Then
                    If Text1 = "N" Then
                    
                        TotalAmt2 = TotalAmt2 + IIf(IsNumeric(rst2!CheckPayorAmt) = True, rst2!CheckPayorAmt, 0)
                        LblTotalAmt2 = Format(TotalAmt2, "#,##0.00")
                    Else
                        TotalAmt2 = TotalAmt2 + IIf(IsNumeric(rst2!CheckMNRBAmt) = True, rst2!CheckMNRBAmt, 0)
                        LblTotalAmt2 = Format(TotalAmt2, "#,##0.00")
                    End If
                End If
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If Mid(CboPayor, 1, 2) <> "" Then
        LblTotalAmt3 = (CDbl(LblTotalAmt) - CDbl(LblTotalAmt2))
    End If
    
    If lstBordxTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        cmdPrint.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
End Sub
'**********************End BordxTracking Listing *******************************

'**********************Start ComboListing *******************************

Private Sub PayorComboListing()
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
            CboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
    
    cboPayAppStatus.AddItem "Y" & " - " & "Yes"
    cboPayAppStatus.AddItem "N" & " - " & "No"
End Sub
'**********************End ComboListing *******************************

'**********************Start ComboBox Change/KeyPress *******************************

Private Sub CboPayor_Change()
Dim strsql As String
Dim HLT As New ADODB.Recordset
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cnn As New ADODB.Command
Dim strAppendCase As String
Dim strAppend As String

If Len(CboPayor) Then
    
    'strsql = "Select HLTCode from PAY where PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
    'HLT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "1"
    strAppend = " AND PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "BordxWorksheet"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        HLT.CursorLocation = adUseClient
        HLT.CursorType = adOpenForwardOnly
        HLT.CacheSize = 100
        Set HLT = cmd.Execute
        
    Text1 = Trim(HLT!HLTCode)
    
    If Text1 = "S" Then
        Label16.Caption = "From MNRB"
    Else
        Label16.Caption = "From Payor"
    End If
        
    HLT.Close
    Set HLT = Nothing
    
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If

If Len(Mid(CboPayor, 1, 2)) Then '.Text = "" Then

    CboGrpCom.Clear

    'strsql = "select MBMGroupCompany from View_PAYGrpCom where PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
    'PAYGRP.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "2"
    strAppend = " AND PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
    Set cnn.ActiveConnection = medixmasterconn
        cnn.CommandText = "BordxWorksheet"
        cnn.CommandType = adCmdStoredProc
        cnn.CommandTimeout = 300
        Set Prm1 = cnn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cnn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cnn.Parameters.Append Prm2
        PAYGRP.CursorLocation = adUseClient
        PAYGRP.CursorType = adOpenForwardOnly
        PAYGRP.CacheSize = 100
        Set PAYGRP = cnn.Execute
    
    If PAYGRP.recordCount Then
        Do Until PAYGRP.EOF
           CboGrpCom.AddItem Trim(PAYGRP!MBMGroupCompany)
           PAYGRP.MoveNext
        Loop
    End If
    
End If
End Sub


'**********************End ComboBox Change/KeyPress *******************************

'**********************Start Sorted ListView *******************************
Private Sub lstBordxTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
    Case 5
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(12)
    Case 13
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lstBordxTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(19)
        
    End Select
End Sub
'**********************End Sorted ListView *******************************

