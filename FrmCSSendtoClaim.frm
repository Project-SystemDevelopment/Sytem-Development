VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCSSendtoClaim 
   Caption         =   "Send to Bordx Unit"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   44
      Top             =   7080
      Width           =   11655
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   69
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Amend"
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
         Left            =   7440
         TabIndex        =   68
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txtBatchNo 
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
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   615
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
         Height          =   375
         Left            =   10800
         TabIndex        =   37
         Top             =   180
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
         Height          =   375
         Left            =   10200
         TabIndex        =   36
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
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
         Left            =   9240
         TabIndex        =   35
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdPass 
         Caption         =   "&Pass to BU"
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
         Left            =   8160
         TabIndex        =   31
         Top             =   180
         Width           =   1095
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
         Height          =   375
         Left            =   9120
         TabIndex        =   34
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
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
         Left            =   8520
         TabIndex        =   33
         Top             =   180
         Width           =   615
      End
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
         Height          =   375
         Left            =   6840
         TabIndex        =   30
         Top             =   180
         Width           =   615
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
         Height          =   375
         Left            =   6120
         TabIndex        =   28
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
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
         Left            =   7680
         TabIndex        =   32
         Top             =   180
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtUpdateDPO 
         Height          =   285
         Left            =   3720
         TabIndex        =   66
         Top             =   240
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label21 
         Caption         =   "DPO"
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
         Left            =   3240
         TabIndex        =   67
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label21 
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
         Index           =   0
         Left            =   1680
         TabIndex        =   45
         Top             =   255
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selection Criteria"
      Height          =   1995
      Left            =   120
      TabIndex        =   29
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox chkDOB 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   10485
         TabIndex        =   27
         Top             =   1260
         Width           =   495
      End
      Begin VB.CheckBox ChkDPO 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   10485
         TabIndex        =   24
         Top             =   1000
         Width           =   495
      End
      Begin VB.CheckBox ChkDOD 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   10485
         TabIndex        =   21
         Top             =   750
         Width           =   495
      End
      Begin VB.CheckBox ChkDOA 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   10485
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkDOR 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   10485
         TabIndex        =   15
         Top             =   200
         Width           =   495
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
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
         Left            =   5200
         TabIndex        =   2
         Top             =   240
         Width           =   1050
      End
      Begin VB.TextBox TxtCaseNo2 
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
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   1
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TxtCaseNo 
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   1560
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5100
         TabIndex        =   12
         Top             =   1630
         Width           =   255
      End
      Begin VB.TextBox txtBordxfr 
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
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   3
         Top             =   500
         Width           =   1560
      End
      Begin MSMask.MaskEdBox txtDORDate1 
         Height          =   285
         Left            =   7680
         TabIndex        =   13
         Top             =   240
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate1 
         Height          =   285
         Left            =   7680
         TabIndex        =   16
         Top             =   500
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate1 
         Height          =   285
         Left            =   7680
         TabIndex        =   19
         Top             =   750
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDPODate1 
         Height          =   285
         Left            =   7680
         TabIndex        =   22
         Top             =   1000
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBordxTo 
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
         Left            =   3480
         MaxLength       =   1
         TabIndex        =   4
         Top             =   500
         Width           =   1560
      End
      Begin MSMask.MaskEdBox txtDORDate2 
         Height          =   285
         Left            =   9360
         TabIndex        =   14
         Top             =   240
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate2 
         Height          =   285
         Left            =   9360
         TabIndex        =   17
         Top             =   500
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate2 
         Height          =   285
         Left            =   9360
         TabIndex        =   20
         Top             =   750
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDPODate2 
         Height          =   285
         Left            =   9360
         TabIndex        =   23
         Top             =   1000
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboPayor 
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
         TabIndex        =   5
         Top             =   750
         Width           =   3855
      End
      Begin VB.ComboBox CboCaseStatus 
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
         TabIndex        =   6
         Top             =   1030
         Width           =   1335
      End
      Begin VB.ComboBox CboClaimability 
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
         TabIndex        =   7
         Top             =   1030
         Width           =   1575
      End
      Begin VB.ComboBox cboStayType 
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
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cboCashType 
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
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox CboDept 
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
         TabIndex        =   10
         Top             =   1600
         Width           =   1335
      End
      Begin VB.ComboBox CboPass 
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
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   1600
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtBordxdatefr 
         Height          =   285
         Left            =   7680
         TabIndex        =   25
         Top             =   1260
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxdateTo 
         Height          =   285
         Left            =   9360
         TabIndex        =   26
         Top             =   1260
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label19 
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
         TabIndex        =   64
         Top             =   1300
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "DOB from"
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
         Left            =   6720
         TabIndex        =   63
         ToolTipText     =   "Date of Bordx"
         Top             =   1300
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Cash Type"
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
         Left            =   2640
         TabIndex        =   62
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Stay Type"
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
         TabIndex        =   61
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label3 
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
         TabIndex        =   60
         Top             =   850
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Left            =   3000
         TabIndex        =   59
         Top             =   520
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Bordx"
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
         TabIndex        =   58
         Top             =   520
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "DPO from"
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
         Left            =   6720
         TabIndex        =   55
         ToolTipText     =   "Date Pass On"
         Top             =   1030
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "DOR from"
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
         Left            =   6720
         TabIndex        =   54
         ToolTipText     =   "Date of Registered"
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "DOA from"
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
         Left            =   6720
         TabIndex        =   53
         ToolTipText     =   "Date of Admission"
         Top             =   550
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "DOD from"
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
         Left            =   6720
         TabIndex        =   52
         ToolTipText     =   "Date of Discharged"
         Top             =   800
         Width           =   1215
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
         Left            =   8880
         TabIndex        =   51
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label15 
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
         TabIndex        =   50
         Top             =   550
         Width           =   495
      End
      Begin VB.Label Label16 
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
         TabIndex        =   49
         Top             =   800
         Width           =   495
      End
      Begin VB.Label Label17 
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
         TabIndex        =   48
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Claimability"
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
         Left            =   2640
         TabIndex        =   43
         Top             =   1100
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Dept. Stage"
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
         TabIndex        =   42
         Top             =   1630
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Case Status"
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
         TabIndex        =   41
         Top             =   1100
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Passed by"
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
         Left            =   2640
         TabIndex        =   40
         Top             =   1630
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Claim No"
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
         TabIndex        =   39
         Top             =   280
         Width           =   855
      End
      Begin VB.Label Label18 
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
         Left            =   3000
         TabIndex        =   38
         Top             =   280
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lstCaseTrackList 
      Height          =   4815
      Left            =   120
      TabIndex        =   46
      Top             =   2280
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8493
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
         Text            =   "Pass-On"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Claim No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Mem'ship No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "DOR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "DOA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "DOD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "DPO"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "No of Days (DPO-DOD)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Batch No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Passed By"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Reg. By"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Hospital Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "S.Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "C.Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "App.Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Text            =   "Claimability"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Dept Stage"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CheckBox chkCheckAll 
      Caption         =   "Use Scan"
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
      Left            =   360
      TabIndex        =   65
      Top             =   2400
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Send to Bordx Unit"
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
      TabIndex        =   57
      Top             =   0
      Width           =   11805
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   7800
      Width           =   2055
   End
End
Attribute VB_Name = "FrmCSSendtoClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intExit As Integer
Dim listloop1 As Integer
Dim Check As Boolean
Dim Current2 As String
Dim Month As String
Dim ScanCode As String
Dim PayorCode As String
Dim DPO As String
Dim CASNo As String
Dim strAppend As String
Dim tmpCaseNumber As String
Dim tmpVerNumber As String
Dim tmpDeptStage As String
Dim m_blnDirection(1 To 18) As Boolean

Private Sub cboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayor.Text, cboPayor.SelStart) + Chr(KeyAscii) + Mid(cboPayor.Text, cboPayor.SelStart + cboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayor.ListCount - 1
 
        If NewText = LCase(Left(cboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayor.Text = ValidValue
                cboPayor.SelStart = 0
                cboPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub chkCheckAll_Click()
    If chkCheckAll.Value = "1" Then
        SetCheckAllItems True
    Else
        SetCheckAllItems False
    End If

End Sub

Private Sub cmdUpdate_Click()
Dim strsql As String
Dim Claim As New ADODB.Recordset
    
    If Trim(tmpDeptStage) = "1" Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        strsql = ""
        strsql = "Select CASTokenDate, CASTokenBatch,CASClaimAdmin,CASUserToken from ClaimOthersTwo where CASNumber = '" & tmpCaseNumber & "' And ClaimVersion = '" & tmpVerNumber & "'"
        Claim.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        If Claim.recordCount > 0 Then
            If txtUpdateDPO = "__/__/____" Then
                Claim!CASTokenDate = Null
                Claim!CASTokenBatch = ""
                Claim!CASClaimAdmin = 0
                Claim!CASUserToken = ""
            Else
                Claim!CASTokenDate = Format(txtUpdateDPO, "dd/mm/yyyy")
            End If
            Claim.Update
            strAppend = ""
            
            If ChkScan.Value = 1 Then
                If Len(Trim(tmpCaseNumber)) Then
                    strAppend = strAppend + " AND CASNumber = '" & Trim(tmpCaseNumber) & "'"
                End If
            Else
                If Len(Trim(tmpCaseNumber)) Then
                    strAppend = strAppend + " AND CASNumber >= '" & Trim(tmpCaseNumber) & "'"
                End If
                
                If Len(Trim(tmpCaseNumber)) Then
                    strAppend = strAppend + " AND CASNumber <= '" & Trim(tmpCaseNumber) & "'"
                End If
            End If

            If ChkScan.Value = 0 Then
                Call DischargeListing(strAppend)
            ElseIf ChkScan.Value = 1 Then
                Call DischargeListing2(strAppend)
            End If
            
        End If
        Claim.Close
        Set Claim = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
     '   medixmasterconnWS.Close
      '  Set medixmasterconnWS = Nothing
        MsgBox "Record updated successfully!", vbOKOnly + vbCritical, DialogBoxTitle
        tmpCaseNumber = ""
        tmpVerNumber = ""
        txtUpdateDPO = "__/__/____"
        tmpDeptStage = ""
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

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    Current2 = 0
    Check = False
    ChkScan.Value = 1
    Call ComboListing
End Sub

'**********************Start Command Button *******************************

Private Sub CmdBack_Click()
    Frame1.Visible = True
    TxtCaseNo.SetFocus
    lstCaseTrackList.Top = 2160
    lstCaseTrackList.Width = 11655
    lstCaseTrackList.Left = 120
    lstCaseTrackList.Height = 4935
End Sub

Private Sub CmdView_Click()
    Frame1.Visible = False
    lstCaseTrackList.Top = 480
    lstCaseTrackList.Width = 11655
    lstCaseTrackList.Left = 120
    lstCaseTrackList.Height = 6615
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
    CmdPass.Enabled = True
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstCaseTrackList.ListItems.Clear
    lblCurrent = 0
    Current2 = 0
    Check = False
End Sub

Private Sub CmdRefresh_Click()
Dim User3 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

CboPass.Clear

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

'User3.Open "Select Distinct CASUserToken from CAS", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    strAppend = ""
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Send2Claim"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    User3.CursorLocation = adUseClient
    User3.CursorType = adOpenForwardOnly
    User3.CacheSize = 100
    Set User3 = cmd.Execute

    Do Until User3.EOF
        If Len(User3!CASUserToken) Then
            CboPass.AddItem User3!CASUserToken
        End If
        User3.MoveNext
    Loop
    User3.Close
    Set User3 = Nothing
    
'medixmasterconn.Close
'Set medixmasterconn = Nothing

End Sub

Private Sub CmdCheck_Click()
    SetCheckAllItems True
End Sub

Public Sub cmdSearch_Click()
    tmpCaseNumber = ""
    tmpVerNumber = ""
    txtUpdateDPO = "__/__/____"
    tmpDeptStage = ""

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    ScanCode = TxtCaseNo
    CmdSearch.Enabled = False
    If Mid(ScanCode, 1, 2) = "LO" Then
        TxtCaseNo = ""
        Call SearchForm(ScanCode)
    Else
            
        CmdClear.Enabled = False
        CmdExit.Enabled = False
        CmdPrintFile.Enabled = False
        CmdPass.Enabled = False
            
        If ChkScan.Value = 1 Then
        
            If TxtCaseNo = "" Then
                MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtCaseNo.SetFocus
                CmdClear.Enabled = True
                CmdExit.Enabled = True
                'CmdPrint.Enabled = True
                CmdPrintFile.Enabled = True
                CmdPass.Enabled = True
            Else
                Screen.MousePointer = vbHourglass
            
                listrecord = lstCaseTrackList.ListItems.Count
                If listrecord <> 0 Then
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                
                If Check = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                ElseIf Check = False Then
                    CmdClear.Enabled = True
                    CmdExit.Enabled = True
                    CmdPrintFile.Enabled = True
                    CmdPass.Enabled = True
                    Call SearchRecord
                End If
                Check = False
                Screen.MousePointer = Default
                TxtCaseNo = ""
                TxtCaseNo.SetFocus
            End If
            CmdClear.Enabled = True
            CmdExit.Enabled = True
            CmdPrintFile.Enabled = True
            CmdPass.Enabled = True
            TxtCaseNo.SetFocus
        Else
            
            Screen.MousePointer = vbHourglass
                
                Call SearchRecord
            
            Screen.MousePointer = Default
            
            CmdClear.Enabled = True
            CmdExit.Enabled = True
            CmdPrintFile.Enabled = True
            CmdPass.Enabled = True
        End If
    End If
    CmdSearch.Enabled = True
    
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPass_Click()
Dim i As Integer
Dim ItemChecked As Boolean
Dim CaseNo As String
Dim Version As String
Dim Update
Dim startTrans
  
    startTrans = False
    
    ItemChecked = False
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
    If lstCaseTrackList.ListItems.Count = 0 Then
        MsgBox "Please select a payor before proceeding!", vbOKOnly + vbCritical
    Else
        
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                
                startTrans = True
                medixmasterconnMaint.beginTrans
                medixmasterconnWS.beginTrans
                DPO = Date
                txtBatchNo = GenerateCasePass(DPO)
                
                For i = 1 To lstCaseTrackList.ListItems.Count
                    If lstCaseTrackList.ListItems(i).Checked = True Then
                        ItemChecked = True
                        Exit For
                    End If
                Next i
                    
                If ItemChecked = True Then
                
                    For i = 1 To lstCaseTrackList.ListItems.Count
                        If lstCaseTrackList.ListItems(i).Checked = True Then
                            
                            CaseNo = lstCaseTrackList.ListItems(i).ListSubItems(1).Text
                            Version = lstCaseTrackList.ListItems(i).ListSubItems(2).Text
                            Call UpdateRecord(CaseNo, Version)
                        End If
                    Next i
                    
                    If startTrans = True Then
                        medixmasterconnMaint.commitTrans
                        medixmasterconnWS.commitTrans
                    End If
                    
                    MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    
                    lstCaseTrackList.ListItems.Clear
                    
                    'If ChkScan.value = 0 Then
                        Call DischargeListing(strAppend)
                    'End If
                Else
                    medixmasterconnMaint.commitTrans
                    medixmasterconnWS.commitTrans
                End If
            End If
    End If
    
'    medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    
    
Exit Sub
errSave:
    MsgBox Err.Description
    If startTrans = True Then
        medixmasterconnMaint.RollbackTrans
        medixmasterconnWS.RollbackTrans
    End If
End Sub

Private Sub CmdPrintFile_Click()
    
    Screen.MousePointer = vbHourglass
    
    If lstCaseTrackList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstCaseTrackList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
  
End Sub
'**********************End Command Button *******************************

'**********************Start Update Record*******************************
Private Sub UpdateRecord(CaseNo As String, Version As String)
Dim REC As New ADODB.Recordset
Dim str As String
Dim strsql As String
Dim strsqlREC As String
Dim TokenDate As String

    TokenDate = Format(Date, "YYYY/MM/DD")
    str = Trim(CaseNo) & "," & Trim(Version)
    strsql = "EXEC Case_Send2Claim '" & Trim(str) & "',4"
    REC.Open strsql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    
    If REC.EOF Then
        strsqlREC = "EXEC Case_Send2InsertRec '" & Trim(CaseNo) & "', '" & Trim(Version) & "', '" & TokenDate & "', '" & txtBatchNo & "', '1', '" & Logon & "'"
    Else
        strsqlREC = "EXEC Case_Send2UpdateRec '" & Trim(CaseNo) & "', '" & Trim(Version) & "', '" & TokenDate & "', '" & txtBatchNo & "', '1', '" & Logon & "'"
    End If
    medixmasterconnWS.Execute strsqlREC
    REC.Close
    Set REC = Nothing
End Sub
'**********************End Update Record*******************************

'**********************End Search Record*******************************
Private Function SearchRecord()
Dim strsql As String
Dim sql As String
Dim CAS As New ADODB.Recordset
    
    strAppend = ""
        
        If ChkScan.Value = 1 Then
            If Len(Trim(TxtCaseNo)) Then
                strAppend = strAppend + " AND CASNumber = '" & Trim(TxtCaseNo) & "'"
            End If
        Else
            If Len(Trim(TxtCaseNo)) Then
                strAppend = strAppend + " AND CASNumber >= '" & Trim(TxtCaseNo) & "'"
            End If
            
            If Len(Trim(TxtCaseNo2)) Then
                strAppend = strAppend + " AND CASNumber <= '" & Trim(TxtCaseNo2) & "'"
            End If
        End If
                
        If Len(Trim(txtBordxfr)) Then
            strAppend = strAppend + " AND BordxRefNo >= '" & Trim(txtBordxfr) & "'"
        End If
        
        If Len(Trim(txtBordxTo)) Then
            strAppend = strAppend + " AND BordxRefNo <= '" & Trim(txtBordxTo) & "'"
        End If

        
        If Mid(CboCaseStatus, 1, 1) = 0 Then
            strAppend = strAppend + " And CASStatus = 'O'"
        ElseIf Mid(CboCaseStatus, 1, 1) = 1 Then
            strAppend = strAppend + " And CASStatus = 'D'"
        ElseIf Mid(CboCaseStatus, 1, 1) = 2 Then
            strAppend = strAppend + " And (CASStatus = 'O' OR CASStatus = 'C')"
        End If
        
        'If Mid(CboClaimability, 1, 1) = "R" Then
        
        If Len(CboClaimability) Then
            If Trim(Mid(CboClaimability, 1, 2)) = "1A" Then
                strAppend = strAppend & " AND (CLMWSStatus = 'AA' OR CLMWSStatus = 'SA' OR CLMWSStatus = 'NA' OR CLMWSStatus = 'PA' OR CLMWSStatus = 'RA')"
            ElseIf Trim(Mid(CboClaimability, 1, 2)) = "1R" Then
                strAppend = strAppend & " AND (CLMWSStatus = 'RR' OR CLMWSStatus = 'NR' OR CLMWSStatus = 'PR' OR CLMWSStatus = 'AR' OR CLMWSStatus = 'BR' OR CLMWSStatus = 'XR'"
            Else
                strAppend = strAppend & " AND CLMWSStatus = '" & Trim(Mid(CboClaimability, 1, 2)) & "'"
            End If
        End If
        
        '    strAppend = strAppend + " AND CASClaimability = 'R'"
        'ElseIf Mid(CboClaimability, 1, 1) = "A" Then
        '    strAppend = strAppend + " AND CASClaimability = 'A'"
        'End If
            
        If Mid(CboDept, 1, 1) = 0 Or Mid(CboDept, 1, 1) = 1 Then
           sql = ""
           sql = " AND  CASClaimAdmin = '" & Trim(Mid(CboDept, 1, 1)) & "'"
           strAppend = strAppend + sql
        End If
        
        If Len(Trim(CboPass)) Then
            strAppend = strAppend + " AND CASUserToken = '" & Trim(CboPass) & "'"
        End If
        
        If Len(Trim(txtBordxfr)) Then
            strAppend = strAppend + " AND ClaimVersion >= '" & Trim(txtBordxfr) & "'"
        End If
        
        If Len(Trim(txtVersion2)) Then
            strAppend = strAppend + " AND ClaimVersion <= '" & Trim(txtVersion2) & "'"
        End If
                   
        If ChkDOA.Value = 1 Then
            sql = ""
            sql = "AND (CasDateAdmitted Is null OR CasDateAdmitted = '')" & ""
            strAppend = strAppend + sql
        End If
        
        If chkDOR.Value = 1 Then
            sql = ""
            sql = "AND (CaseRegDate Is null OR CaseRegDate = '')" & ""
            strAppend = strAppend + sql
        End If
        
        If ChkDOD.Value = 1 Then
            sql = ""
            sql = "AND (CasDischargeDate Is null OR CasDischargeDate = '')" & ""
            strAppend = strAppend + sql
        End If
        
        If ChkDPO.Value = 1 Then
            sql = ""
            sql = "AND (CASTokenDate Is null OR CASTokenDate = '')" & ""
            strAppend = strAppend + sql
        End If
        
        If Len(Trim(txtDORDate1)) > 0 And IsDate(txtDORDate1) = True Then
            strAppend = strAppend + " And CaseRegDate >= '" & Format(txtDORDate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDORDate2)) > 0 And IsDate(txtDORDate2) = True Then
            strAppend = strAppend + " And CaseRegDate <= '" & Format(txtDORDate2, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDOADate1)) > 0 And IsDate(txtDOADate1) = True Then
            strAppend = strAppend + " And CasDateAdmitted >= '" & Format(txtDOADate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDOADate2)) > 0 And IsDate(txtDOADate2) = True Then
            strAppend = strAppend + " And CasDateAdmitted <= '" & Format(txtDOADate2, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDODDate1)) > 0 And IsDate(txtDODDate1) = True Then
            strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDODDate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDODDate2)) > 0 And IsDate(txtDODDate2) = True Then
            strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDODDate2, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDPODate1)) > 0 And IsDate(txtDPODate1) = True Then
            strAppend = strAppend + " And CASTokenDate >= '" & Format(txtDPODate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDPODate2)) > 0 And IsDate(txtDPODate2) = True Then
            strAppend = strAppend + " And CASTokenDate <= '" & Format(txtDPODate2, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtBordxdatefr)) > 0 And IsDate(txtBordxdatefr) = True Then
            strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxdatefr, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtBordxdateTo)) > 0 And IsDate(txtBordxdateTo) = True Then
            strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxdateTo, "mm/dd/yyyy") & "'"
        End If
        
        If Len(cboPayor) Then
            strAppend = strAppend + " And PayCode = '" & Mid(cboPayor, 1, 2) & "'"
        End If
        
        If Mid(cboStayType, 1, 1) = 0 Then
            strAppend = strAppend & " AND CLMStayType = 'I'"
        ElseIf Mid(cboStayType, 1, 1) = 1 Then
            strAppend = strAppend & " AND CLMStayType = 'O'"
        ElseIf Mid(cboStayType, 1, 1) = 2 Then
            strAppend = strAppend & " AND CLMStayType = 'F'"
        ElseIf Mid(cboStayType, 1, 1) = 3 Then
            strAppend = strAppend & " AND CLMStayType = 'P'"
        ElseIf Mid(cboStayType, 1, 1) = 4 Then
            strAppend = strAppend & " AND (CLMStayType = 'I' OR CLMStayType = 'O')"
        ElseIf Mid(cboStayType, 1, 1) = 5 Then
            strAppend = strAppend & " AND (CLMStayType = 'F' OR CLMStayType = 'P')"
        End If
        
        If Len(cboCashType) Then
            strAppend = strAppend + " AND CLMCashType = '" & Mid(cboCashType, 1, 1) & "'"
        End If
        
        
    If ChkScan.Value = 0 Then
        Call DischargeListing(strAppend)
    ElseIf ChkScan.Value = 1 Then
        Call DischargeListing2(strAppend)
    End If
End Function
'**********************End Search Record*******************************

'**********************Start List Record*******************************

Private Sub DischargeListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim ClaimAdmin As Boolean
Dim total As String
Dim Current As String
Dim strCount
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String

    Current = 0
    
    lstCaseTrackList.ListItems.Clear

    'sql = " Select CASNumber, MBMNumber, PatientName, " & _
          " CasDateAdmitted, CASStayType, CASPayType, CASClaimAdmin, " & _
          " CASDischargeDate, CaseRegDate, CASHOSCode, " & _
          " CASUserToken, CASClaimability, ActualAmt, " & _
          " CASTokenDate, CASTokenBatch, NoofDay, CaseOpenBy " & _
          " From VIEW_CAS_2_CLM where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If

    'rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "3"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Send2Claim"
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
                If !CASClaimAdmin = "1" Then
                    Set Record = lstCaseTrackList.ListItems.Add(, , "Passed")
                Else
                    Set Record = lstCaseTrackList.ListItems.Add(, , "Not")
                End If
                If Len(rst2!CasNumber) Then
                    Record.SubItems(1) = rst2!CasNumber
                End If
                If Len(rst2!claimversion) Then
                    Record.SubItems(2) = rst2!claimversion
                Else
                    Record.SubItems(2) = "1"
                End If
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(3) = rst2!MBMNumber
                End If
                If Len(rst2!PatientName) Then
                    Record.SubItems(4) = Trim(rst2!PatientName)
                End If
                If Len(rst2!CaseRegDate) Then
                    Record.SubItems(5) = Format(rst2!CaseRegDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(6) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(7) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CASTokenDate) Then
                    Record.SubItems(8) = Format(rst2!CASTokenDate, "dd/mm/yyyy")
                End If
                If rst2!NoofDay = "0" Then
                    Record.SubItems(9) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(9) = rst2!NoofDay
                End If
                If Len(rst2!CASTokenBatch) Then
                    Record.SubItems(10) = rst2!CASTokenBatch
                End If
                If Len(rst2!CASUserToken) Then
                    Record.SubItems(11) = Trim(rst2!CASUserToken)
                End If
                If Len(rst2!CaseOpenBy) Then
                    Record.SubItems(12) = Trim(rst2!CaseOpenBy)
                End If
                If Len(rst2!CASHOSCode) Then
                    Record.SubItems(13) = rst2!CASHOSCode
                End If
                If Len(rst2!CLMStayType) Then
                    Record.SubItems(14) = Trim(rst2!CLMStayType)
                End If
                If Len(rst2!CLMCashType) Then
                    Record.SubItems(15) = Trim(rst2!CLMCashType)
                End If
                If Len(rst2!ActualAmt) Then
                    Record.SubItems(16) = rst2!ActualAmt
                End If
                If Len(rst2!CLMWSStatus) Then
                    Record.SubItems(17) = rst2!CLMWSStatus
                End If
                If Len(rst2!CASClaimAdmin) Then
                    Record.SubItems(18) = rst2!CASClaimAdmin
                End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(19) = Trim(rst2!BordxRefNo)
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(20) = rst2!BordxDate
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
    
    If lstCaseTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdPrintFile.Enabled = True
        CmdPass.Enabled = False
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
    CmdPass.Enabled = True
End Sub

Private Sub DischargeListing2(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim ClaimAdmin As Boolean
Dim strCount
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String

    'sql = " Select CASNumber, MBMNumber, PatientName, " & _
          " CasDateAdmitted, CASStayType, CASPayType, CASClaimAdmin, " & _
          " CASDischargeDate, CaseRegDate, CASHOSCode, " & _
          " CASUserToken, CASClaimability, ActualAmt, " & _
          " CASTokenDate, CASTokenBatch, NoofDay, CaseOpenBy " & _
          " From VIEW_CAS_2_CLM where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If

    'rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "3"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Send2Claim"
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
                If !CASClaimAdmin = "1" Then
                    Set Record = lstCaseTrackList.ListItems.Add(, , "Passed")
                Else
                    Set Record = lstCaseTrackList.ListItems.Add(, , "Not")
                End If
                If Len(rst2!CasNumber) Then
                    Record.SubItems(1) = rst2!CasNumber
                End If
                If Len(rst2!claimversion) Then
                    Record.SubItems(2) = rst2!claimversion
                Else
                    Record.SubItems(2) = "1"
                End If
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(3) = rst2!MBMNumber
                End If
                If Len(rst2!PatientName) Then
                    Record.SubItems(4) = Trim(rst2!PatientName)
                End If
                If Len(rst2!CaseRegDate) Then
                    Record.SubItems(5) = Format(rst2!CaseRegDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(6) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(7) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASTokenDate) Then
                    Record.SubItems(8) = Format(rst2!CASTokenDate, "dd/mm/yyyy")
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(9) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(9) = rst2!NoofDay
                End If
                
                If Len(rst2!CASTokenBatch) Then
                    Record.SubItems(10) = rst2!CASTokenBatch
                End If
                
                If Len(rst2!CASUserToken) Then
                    Record.SubItems(11) = Trim(rst2!CASUserToken)
                End If
                
                If Len(rst2!CaseOpenBy) Then
                    Record.SubItems(12) = Trim(rst2!CaseOpenBy)
                End If
                
                If Len(rst2!CASHOSCode) Then
                    Record.SubItems(13) = rst2!CASHOSCode
                End If
                If Len(rst2!CLMStayType) Then
                    Record.SubItems(14) = Trim(rst2!CLMStayType)
                End If
                If Len(rst2!CLMCashType) Then
                    Record.SubItems(15) = Trim(rst2!CLMCashType)
                End If
                If Len(rst2!ActualAmt) Then
                    Record.SubItems(16) = rst2!ActualAmt
                End If
                If Len(rst2!CLMWSStatus) Then
                    Record.SubItems(17) = rst2!CLMWSStatus
                End If
                If Len(rst2!CASClaimAdmin) Then
                    Record.SubItems(18) = rst2!CASClaimAdmin
                End If
                                                            
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(19) = Trim(rst2!BordxRefNo)
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(20) = rst2!BordxDate
                End If
                                                            
                                                            
                Current2 = Current2 + 1
                lblCurrent = Current2
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If lstCaseTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdPrintFile.Enabled = True
        CmdPass.Enabled = False
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
    CmdPass.Enabled = True
End Sub
'**********************End List Record*******************************

Private Sub lstCaseTrackList_Click()
    tmpCaseNumber = ""
    tmpVerNumber = ""
    txtUpdateDPO = "__/__/____"
    tmpDeptStage = ""
    If lstCaseTrackList.ListItems.Count > 0 Then
        If lstCaseTrackList.SelectedItem.SubItems(18) = "1" Then
            If lstCaseTrackList.SelectedItem.SubItems(8) <> "" Then
                txtUpdateDPO = lstCaseTrackList.SelectedItem.SubItems(8)
                tmpCaseNumber = lstCaseTrackList.SelectedItem.SubItems(1)
                tmpVerNumber = lstCaseTrackList.SelectedItem.SubItems(2)
                tmpDeptStage = lstCaseTrackList.SelectedItem.SubItems(18)
            End If
        End If
    End If
End Sub

'**********************Start Sort ListView*******************************

Private Sub lstCaseTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
    Case 16
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(18)
    End Select
End Sub
'**********************Start Sort ListView*******************************


'**********************Start Check ListView Record*******************************
Private Sub lstCaseTrackList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    If Item.Checked = True Then
        If Trim(Item.SubItems(5)) = "" Then
            MsgBox "Date of Register is Empty.", vbCritical
            Item.Checked = False
        ElseIf Trim(Item.SubItems(6)) = "" Then
            MsgBox "Date of Admitted is Empty.", vbCritical
            Item.Checked = False
        ElseIf Trim(Item.SubItems(7)) = "" Then
            MsgBox "Date of Discharge cannot be empty!", vbCritical
            Item.Checked = False
        ElseIf Trim(Item.SubItems(13)) = "" Then
            MsgBox "Hospital Name is Empty.", vbCritical
            Item.Checked = False
        ElseIf Trim(Item.SubItems(18)) = "1" Then
            MsgBox "File has been send out.", vbCritical
            Item.Checked = False
        ElseIf Trim(Item.SubItems(19)) = "" Then
            MsgBox "Bordx is not generated", vbCritical
            Item.Checked = False
        End If
    End If
End Sub
'**********************End Check ListView Record*******************************

'**********************Start Combobox Listing *******************************
Private Sub ComboListing()
Dim User3 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim PAY As New ADODB.Recordset

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchCASToken"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    User3.CursorLocation = adUseClient
    User3.CursorType = adOpenForwardOnly
    User3.CacheSize = 100
    Set User3 = cmd.Execute
    
    Do Until User3.EOF
    If Len(User3!CASUserToken) Then
        CboPass.AddItem User3!CASUserToken
    End If
    User3.MoveNext
    Loop
    User3.Close
    Set User3 = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    
    cboPayor.Clear
    
    Do Until PAY.EOF
        With PAY
            cboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
    
    With CboDept
        .AddItem "0 - Admission"
        .AddItem "1 - Claim"
        .AddItem "2 - Both"
    End With
    
    With CboCaseStatus
        .Text = "2 - Open & Close"
        .AddItem "0 - Open"
        .AddItem "1 - Delete"
        .AddItem "2 - Open & Close"
    End With
    
    With CboClaimability
        .AddItem "1A" & " - " & "All Accepted"
        .AddItem "1R" & " - " & "All Repudiated"
        .AddItem "AA" & " - " & "Accepted"
        .AddItem "RR" & " - " & "Repudiated"
        .AddItem "CC" & " - " & "Cancelled"
        .AddItem "SA" & " - " & "Special Approval"
        .AddItem "PP" & " - " & "Pay & Claim"
        .AddItem "NN" & " - " & "Not Decided"
        .AddItem "NA" & " - " & "Not Decided/Accepted"
        .AddItem "NR" & " - " & "Not Decided/Repudiated"
        .AddItem "PA" & " - " & "Pay&Claim/Accepted"
        .AddItem "PR" & " - " & "Pay&Claim/Rejected"
        .AddItem "AR" & " - " & "Accepted/Rejected"
        .AddItem "RA" & " - " & "Rejected/Accepted"
        .AddItem "BR" & " - " & "BTBG/Rejected"
        .AddItem "AI" & " - " & "Accepted/Insurance"
    End With
    
    With cboStayType
        .AddItem "0 - IN-PATIENT"
        .AddItem "1 - OUT-PATIENT"
        .AddItem "2 - Follow-Up"
        .AddItem "3 - Pre-hosp."
        .AddItem "4 - IP + OP"
        .AddItem "5 - Fol. Up + Pre-Hosp"
    End With
    
    With cboCashType
        .AddItem "C - CASHLESS"
        .AddItem "R - REIMBURSEMENT"
    End With
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing

End Sub
'**********************Start Combobox Listing *******************************

'**********************Start Combobox Listed *******************************
Private Sub CboCaseStatus_Change()
    If InStr(CboCaseStatus.Text, "null") Then
       CboCaseStatus.Enabled = False
    End If
End Sub

Private Sub CboCaseStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCaseStatus.Text, CboCaseStatus.SelStart) + Chr(KeyAscii) + Mid(CboCaseStatus.Text, CboCaseStatus.SelStart + CboCaseStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCaseStatus.ListCount - 1
 
        If NewText = LCase(Left(CboCaseStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCaseStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCaseStatus.Text = ValidValue
                CboCaseStatus.SelStart = 0
                CboCaseStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClaimability_Change()
    If InStr(CboClaimability.Text, "null") Then
       CboClaimability.Enabled = False
    End If
End Sub

Private Sub CboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClaimability.Text, CboClaimability.SelStart) + Chr(KeyAscii) + Mid(CboClaimability.Text, CboClaimability.SelStart + CboClaimability.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClaimability.ListCount - 1
 
        If NewText = LCase(Left(CboClaimability.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClaimability.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClaimability.Text = ValidValue
                CboClaimability.SelStart = 0
                CboClaimability.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPass_Change()
    If InStr(CboPass.Text, "null") Then
       CboPass.Enabled = False
    End If
End Sub

Private Sub CboPass_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPass.Text, CboPass.SelStart) + Chr(KeyAscii) + Mid(CboPass.Text, CboPass.SelStart + CboPass.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPass.ListCount - 1
 
        If NewText = LCase(Left(CboPass.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPass.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPass.Text = ValidValue
                CboPass.SelStart = 0
                CboPass.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub CboDept_Change()
    If InStr(CboDept.Text, "null") Then
       CboDept.Enabled = False
    End If
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboDept.Text, CboDept.SelStart) + Chr(KeyAscii) + Mid(CboDept.Text, CboDept.SelStart + CboDept.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboDept.ListCount - 1
 
        If NewText = LCase(Left(CboDept.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboDept.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboDept.Text = ValidValue
                CboDept.SelStart = 0
                CboDept.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub txtCaseNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub txtCaseNo_LostFocus()
    TxtCaseNo = ChkStr(TxtCaseNo)
    CmdSearch.Default = False
End Sub

'**********************End Combobox Listed *******************************

Private Sub txtDORDate1_LostFocus()
    If IsDate(txtDORDate1) Then
    Else
        If txtDORDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDORDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDORDate2_LostFocus()
    If IsDate(txtDORDate2) Then
    Else
        If txtDORDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDORDate2.SetFocus
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

Private Sub txtDPODate1_LostFocus()
    If IsDate(txtDPODate1) Then
    Else
        If txtDPODate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDPODate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDPODate2_LostFocus()
    If IsDate(txtDPODate2) Then
    Else
        If txtDPODate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDPODate2.SetFocus
        End If
    End If
End Sub

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        lstCaseTrackList.ListItems.Clear
        Label18.Visible = False
        TxtCaseNo2.Visible = False
        lblCurrent = 0
        TxtCaseNo = ""
        TxtCaseNo2 = ""
    Else
        lstCaseTrackList.ListItems.Clear
        Label18.Visible = True
        TxtCaseNo2.Visible = True
        lblCurrent = 0
        TxtCaseNo = ""
        TxtCaseNo2 = ""
    End If
    'TxtCaseNo.SetFocus
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtCaseNo)) = ChkStr(lstCaseTrackList.ListItems(listloop1).SubItems(1)) Then
        Check = True
    End If
End Function

Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstCaseTrackList.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstCaseTrackList.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Public Sub SetCheck(hwnd As Long, lItemIndex As Long, bState As Boolean)
Dim LV As LVITEM
Dim r As Long
    With LV
    .mask = LVIF_STATE
    .state = IIf(bState, &H2000, &H1000)
    .stateMask = LVIS_STATEIMAGEMASK
    End With
    r = SendMessageAny(hwnd, LVM_SETITEMSTATE, lItemIndex, LV)
End Sub

Private Sub txtUpdateDPO_LostFocus()
    If IsDate(txtUpdateDPO) Then
    Else
        If txtUpdateDPO <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtUpdateDPO.SetFocus
        End If
    End If
End Sub
