VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReturnFrPayor 
   Caption         =   "Return from Payor"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkCheckAll 
      Caption         =   "Check All"
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   2140
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   6360
      Width           =   11655
      Begin VB.TextBox txtCLMAppCode 
         Height          =   285
         Left            =   9720
         MaxLength       =   15
         TabIndex        =   22
         Top             =   195
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cboREVBordxInd 
         Height          =   315
         Left            =   9720
         TabIndex        =   23
         Top             =   435
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSMask.MaskEdBox TxtDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   195
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
      Begin MSMask.MaskEdBox txtPAYAppDate 
         Height          =   285
         Left            =   7320
         TabIndex        =   20
         Top             =   195
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboBordxStatus 
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
         ItemData        =   "frmReturnFrPayor.frx":0000
         Left            =   3840
         List            =   "frmReturnFrPayor.frx":0002
         TabIndex        =   19
         Top             =   195
         Width           =   1215
      End
      Begin VB.ComboBox cboPayAppStatus 
         Height          =   315
         ItemData        =   "frmReturnFrPayor.frx":0004
         Left            =   7320
         List            =   "frmReturnFrPayor.frx":0006
         TabIndex        =   21
         Top             =   435
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label65 
         Caption         =   "App Code"
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
         Index           =   118
         Left            =   8760
         TabIndex        =   61
         Top             =   195
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label65 
         Caption         =   "Rev Bordx"
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
         Index           =   119
         Left            =   5760
         TabIndex        =   60
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Pay App Date"
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
         Index           =   1
         Left            =   5880
         TabIndex        =   54
         Top             =   195
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Pay App Status"
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
         Index           =   1
         Left            =   5880
         TabIndex        =   53
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Return Status"
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
         Left            =   2640
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Return Date"
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
         Left            =   120
         TabIndex        =   28
         Top             =   200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   120
      TabIndex        =   39
      Top             =   240
      Width           =   11655
      Begin VB.OptionButton OptBordx 
         Caption         =   "By Bordx"
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
         Left            =   1080
         TabIndex        =   0
         Top             =   200
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptWks 
         Caption         =   "By Worksheet"
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
         Left            =   2400
         TabIndex        =   1
         Top             =   200
         Width           =   1575
      End
      Begin VB.ComboBox CboPayor 
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
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox TxtBordx2 
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
         MaxLength       =   20
         TabIndex        =   4
         Top             =   765
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   4560
         TabIndex        =   24
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox TxtBordx1 
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
         MaxLength       =   20
         TabIndex        =   3
         Top             =   765
         Width           =   1335
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
         Left            =   7200
         TabIndex        =   10
         Top             =   465
         Width           =   2775
      End
      Begin VB.TextBox TxtBordxAmt1 
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
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox TxtBordxAmt2 
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
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1020
         Width           =   1695
      End
      Begin VB.TextBox txtCLMNoTo 
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
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1280
         Width           =   1695
      End
      Begin VB.TextBox txtCLMNoFr 
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
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1280
         Width           =   1335
      End
      Begin VB.ComboBox cboAppStatus 
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
         Left            =   7200
         TabIndex        =   11
         Top             =   720
         Width           =   2775
      End
      Begin MSMask.MaskEdBox txtReturnDateFr 
         Height          =   285
         Left            =   7200
         TabIndex        =   12
         Top             =   975
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
      Begin MSMask.MaskEdBox txtReturnDateTo 
         Height          =   285
         Left            =   8760
         TabIndex        =   13
         Top             =   975
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
      Begin MSMask.MaskEdBox txtPayAppDateFr 
         Height          =   285
         Left            =   7200
         TabIndex        =   14
         Top             =   1200
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
      Begin MSMask.MaskEdBox txtPayAppDateTo 
         Height          =   285
         Left            =   8760
         TabIndex        =   15
         Top             =   1200
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
      Begin MSMask.MaskEdBox txtBordxDate1 
         Height          =   285
         Left            =   7200
         TabIndex        =   16
         Top             =   1440
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
      Begin MSMask.MaskEdBox txtBordxDate2 
         Height          =   285
         Left            =   8760
         TabIndex        =   17
         Top             =   1440
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
      Begin VB.ComboBox cboCashType 
         Height          =   315
         ItemData        =   "frmReturnFrPayor.frx":0008
         Left            =   1080
         List            =   "frmReturnFrPayor.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1500
         Width           =   3375
      End
      Begin VB.Label Label19 
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
         Index           =   7
         Left            =   120
         TabIndex        =   62
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label10 
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
         Index           =   2
         Left            =   5880
         TabIndex        =   59
         Top             =   1260
         Width           =   1215
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
         Index           =   2
         Left            =   8520
         TabIndex        =   58
         Top             =   1275
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
         Index           =   1
         Left            =   8520
         TabIndex        =   57
         Top             =   1050
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Returned Date"
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
         Left            =   5880
         TabIndex        =   56
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         Index           =   1
         Left            =   5880
         TabIndex        =   55
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   1095
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
         Left            =   2520
         TabIndex        =   51
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Returned Status"
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
         Left            =   5880
         TabIndex        =   47
         Top             =   510
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx Amt"
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
         TabIndex        =   46
         Top             =   1035
         Width           =   1095
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
         Left            =   2520
         TabIndex        =   45
         Top             =   1035
         Width           =   255
      End
      Begin VB.Label Label10 
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
         Index           =   0
         Left            =   5880
         TabIndex        =   44
         Top             =   1520
         Width           =   1215
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
         Index           =   0
         Left            =   8520
         TabIndex        =   43
         Top             =   1520
         Width           =   495
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
         Left            =   120
         TabIndex        =   42
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No"
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
         TabIndex        =   41
         Top             =   800
         Width           =   1095
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
         Left            =   2520
         TabIndex        =   40
         Top             =   800
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   29
      Top             =   6960
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3480
         TabIndex        =   64
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   975
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
         Left            =   10920
         TabIndex        =   33
         Top             =   150
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
         Left            =   10320
         TabIndex        =   32
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdPrint 
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
         TabIndex        =   30
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton CmdToPayor 
         Caption         =   "&Update"
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
         Left            =   8400
         TabIndex        =   25
         Top             =   150
         Width           =   855
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
         Left            =   7680
         TabIndex        =   31
         Top             =   150
         Width           =   735
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
         Height          =   375
         Left            =   6960
         TabIndex        =   26
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
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
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1680
         TabIndex        =   36
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
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
         Left            =   1440
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lstBordxDelivery 
      Height          =   3975
      Left            =   120
      TabIndex        =   38
      Top             =   2400
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7011
      SortKey         =   2
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
      NumItems        =   11
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
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Bordx Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Date From Payor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Returned From Payor"
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
      TabIndex        =   48
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmReturnFrPayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intExit As Integer
Dim BordxNoArray() As String
Dim searchCriteria As String
Dim ItemChecked As Boolean
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim Check2 As Boolean
Dim Current2 As String
Dim TotalAmt2 As String
Private m_blnDirection(1 To 9) As Boolean


Private Sub cboAppStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboAppStatus.Text, cboAppStatus.SelStart) + Chr(KeyAscii) + Mid(cboAppStatus.Text, cboAppStatus.SelStart + cboAppStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboAppStatus.ListCount - 1
 
        If NewText = LCase(Left(cboAppStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboAppStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboAppStatus.Text = ValidValue
                cboAppStatus.SelStart = 0
                cboAppStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboBordxStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboBordxStatus.Text, CboBordxStatus.SelStart) + Chr(KeyAscii) + Mid(CboBordxStatus.Text, CboBordxStatus.SelStart + CboBordxStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboBordxStatus.ListCount - 1
 
        If NewText = LCase(Left(CboBordxStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboBordxStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboBordxStatus.Text = ValidValue
                CboBordxStatus.SelStart = 0
                CboBordxStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

'Private Sub cboPayAppStatus_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
'    If KeyAscii >= 32 And KeyAscii <> 127 Then
'    NewText = LCase(Left(cboBordxStatus.Text, cboBordxStatus.SelStart) + Chr(KeyAscii) + Mid(cboBordxStatus.Text, cboBordxStatus.SelStart + cboBordxStatus.SelLength + 1))
'
'    ValidCount = 0
'    ValidValue = ""
 
'    For i = 0 To cboBordxStatus.ListCount - 1
 
'        If NewText = LCase(Left(cboBordxStatus.List(i), Len(NewText))) Then
'                ValidCount = ValidCount + 1
'                ValidValue = cboBordxStatus.List(i)
'        End If
'        Next
'        If ValidCount <= 1 Then KeyAscii = 0
'            If ValidCount = 1 Then
'                cboBordxStatus.Text = ValidValue
'                cboBordxStatus.SelStart = 0
'                cboBordxStatus.SelLength = Len(ValidValue)
'            End If
'        End If
'End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
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

'Private Sub cboREVBordxInd_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
'    If KeyAscii >= 32 And KeyAscii <> 127 Then
'    NewText = LCase(Left(cboREVBordxInd.Text, cboREVBordxInd.SelStart) + Chr(KeyAscii) + Mid(cboREVBordxInd.Text, cboREVBordxInd.SelStart + cboREVBordxInd.SelLength + 1))
 
'    ValidCount = 0
'    ValidValue = ""
 
'    For i = 0 To cboREVBordxInd.ListCount - 1
 
'        If NewText = LCase(Left(cboREVBordxInd.List(i), Len(NewText))) Then
'                ValidCount = ValidCount + 1
'                ValidValue = cboREVBordxInd.List(i)
'        End If
'        Next
'        If ValidCount <= 1 Then KeyAscii = 0
'            If ValidCount = 1 Then
'                 cboREVBordxInd.Text = ValidValue
'                 cboREVBordxInd.SelStart = 0
'                 cboREVBordxInd.SelLength = Len(ValidValue)
'            End If
'        End If

'End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboStatus.Text, cboStatus.SelStart) + Chr(KeyAscii) + Mid(cboStatus.Text, cboStatus.SelStart + cboStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboStatus.ListCount - 1
 
        If NewText = LCase(Left(cboStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboStatus.Text = ValidValue
                cboStatus.SelStart = 0
                cboStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub chkCheckAll_Click()
Dim lvCount As Long
Dim lvIndex As Long
Dim lvCountS As Long

    lvCountS = 0
    If lstBordxDelivery.ListItems.Count > 0 Then
    
        If chkCheckAll.Value = 1 Then
            lvCount = lstBordxDelivery.ListItems.Count
            Do
                lvCountS = lvCountS + 1
                lstBordxDelivery.ListItems(lvCountS).Checked = True
    
                lvIndex = lvIndex + 1
            Loop Until lvIndex = lvCount
        Else
            lvCount = lstBordxDelivery.ListItems.Count
            Do
                lvCountS = lvCountS + 1
                    lstBordxDelivery.ListItems(lvCountS).Checked = False
                lvIndex = lvIndex + 1
            Loop Until lvIndex = lvCount
        End If
    Else
        MsgBox "No record found on listing! ", vbOKOnly + vbCritical, DialogBoxTitle
        chkCheckAll.Value = 0
    End If
End Sub

Private Sub ChkScan_Click()
    If OptWks.Value = False Then
        If ChkScan.Value = 1 Then
            lstBordxDelivery.ListItems.Clear
            Label14.Visible = False
            TxtBordx2.Visible = False
            TxtBordx1.Text = ""
            TxtBordx2.Text = ""
            TxtBordx1.SetFocus
        Else
            lstBordxDelivery.ListItems.Clear
            Label14.Visible = True
            TxtBordx2.Visible = True
            TxtBordx1.Text = ""
            TxtBordx2.Text = ""
            TxtBordx1.SetFocus
        End If
    Else
        'If ChkScan.Value = 1 Then
        '    lstBordxDelivery.ListItems.Clear
        '    Label6.Visible = False
        '    txtCLMNoTo.Visible = False
        '    txtCLMNoTo.Text = ""
        '    txtCLMNoFr.Text = ""
            'txtCLMNoFr.SetFocus
        'Else
        '    lstBordxDelivery.ListItems.Clear
        '    Label6.Visible = True
        '    txtCLMNoTo.Visible = True
        '    txtCLMNoTo.Text = ""
        '    txtCLMNoFr.Text = ""
            'txtCLMNoFr.SetFocus
        'End If
    End If
End Sub


Private Sub CmdClear_Click()
    Call ClearForm(Me)
    ChkScan.Value = 1
    lstBordxDelivery.ListItems.Clear
    Label14.Visible = False
    TxtBordx2.Visible = False
    TxtBordx1.SetFocus
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        
    If lstBordxDelivery.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstBordxDelivery)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
Dim listrecord As Integer
'On Error GoTo errRun

    ItemChecked = False
    LblTotalAmt = ""
    lblTotal = ""
    lblCurrent = ""
    ReDim BordxNoArray(0)
    
    CmdSearch.Enabled = False
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    If ChkScan.Value = 1 Then
            If CboPayor = "" Then
                MsgBox "Please select a Payor", vbOKOnly + vbCritical, DialogBoxTitle
                CboPayor.SetFocus
            ElseIf TxtBordx1 = "" And OptBordx.Value = True Then
                    MsgBox "Please Enter Bordx No", vbOKOnly + vbCritical, DialogBoxTitle
                    TxtBordx1.SetFocus
            ElseIf txtCLMNoFr = "" And OptWks.Value = True Then
                    MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
                    txtCLMNoFr.SetFocus
            'End If
                
                CmdClear.Enabled = True
                CmdToPayor.Enabled = True
                CmdPrint.Enabled = True
                CmdExit.Enabled = True
                CmdSearch.Enabled = True
            Else
            
                Screen.MousePointer = vbHourglass
                
                listrecord = lstBordxDelivery.ListItems.Count
                If listrecord <> 0 Then
                    
                    For listloop1 = 1 To listrecord
                        Call BordxNoCheck
                    Next listloop1
                End If
                    If Check = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    ElseIf Check = False Then
                        CmdClear.Enabled = False
                        CmdToPayor.Enabled = False
                        CmdPrint.Enabled = False
                        CmdExit.Enabled = False
                        CmdSearch.Enabled = False
                        Call SearchRecord
                    End If
                    Check = False
                    CmdClear.Enabled = True
                    CmdToPayor.Enabled = True
                    CmdPrint.Enabled = True
                    CmdExit.Enabled = True
                    CmdSearch.Enabled = True
                Screen.MousePointer = Default
                If OptBordx.Value = True Then
                    TxtBordx1 = ""
                    TxtBordx1.SetFocus
                Else
                    txtCLMNoFr = ""
                    txtCLMNoFr.SetFocus
                End If
            End If
            'TxtBordx1.SetFocus
    Else
        
            Screen.MousePointer = vbHourglass
            Call SearchRecord
            Screen.MousePointer = Default
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        CmdPrint.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    End If
    'medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Exit Sub
    
'errRun:
'    MsgBox Err.Description
'    Screen.MousePointer = vbDefault
    lstBordxDelivery.ListItems.Clear
    lblCurrent = ""
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    'medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub CmdToPayor_Click()
Dim Msg

    If lstBordxDelivery.ListItems.Count <= 0 Then
        MsgBox "Please select record(s)", vbOKOnly + vbCritical
        Exit Sub
    'ElseIf Trim(TxtDate) = "__/__/____" And txtPAYAppDate = "__/__/____" Then
    ElseIf Trim(TxtDate) = "__/__/____" Then
        MsgBox "Please Enter Bordx Return Date or Payor Approval Date! ", vbOKOnly + vbCritical, DialogBoxTitle
        TxtDate.SetFocus
        Exit Sub
    'ElseIf cboBordxStatus = "" And cboPayAppStatus = "" Then
    ElseIf CboBordxStatus = "" Then
        MsgBox "Please select bordx status or Payor Approval Date! ", vbOKOnly + vbCritical, DialogBoxTitle
        CboBordxStatus.SetFocus
        Exit Sub
    'ElseIf Trim(txtCLMAppCode) <> "" And Len(Trim(txtCLMAppCode)) < 12 Then
    '    MsgBox "Approval Code must be 12 digits", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtCLMAppCode.SetFocus
        Exit Sub
        
    Else
        Msg = MsgBox("Do you want to update the record(s) ?", vbYesNo + vbExclamation)
        If Msg = vbYes Then
            'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                Call UpdateReturnPayor
           ' medixmasterconnWS.Close
           ' Set medixmasterconnWS = Nothing
        End If
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
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
        Call ListViewHeader
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub


Private Sub OptBordx_Click()
    'If OptBordx.value = True Then
        txtCLMNoFr.Enabled = False
        txtCLMNoTo.Enabled = False
        TxtBordx2.Visible = False
        txtCLMNoTo.Visible = True
        cboStatus.Enabled = False
        cboAppStatus.Enabled = False
        txtReturnDateFr.Enabled = False
        txtPayAppDateFr.Enabled = False
        txtPayAppDateTo.Enabled = False
        txtReturnDateTo.Enabled = False
        ChkScan.Visible = True
        Label6.Visible = True
        Label14.Visible = False
        ChkScan.Top = 795
        ChkScan.Value = 1
    'Else
    '    txtCLMNoFr.Enabled = True
    '    txtCLMNoTo.Enabled = True
    '    TxtBordx2.Visible = True
    '    txtCLMNoTo.Visible = False
    '    Label6.Visible = True
    '    Label14.Visible = False
    '    ChkScan.Top = 1050
    'End If
    lstBordxDelivery.ListItems.Clear
    Call ListViewHeader

End Sub

Private Sub OptWks_Click()
    If OptWks.Value = True Then
        'txtCLMNoFr.Enabled = True
        'txtCLMNoTo.Enabled = True
        TxtBordx2.Visible = True
        'txtCLMNoTo.Visible = False
        cboStatus.Enabled = True
        cboAppStatus.Enabled = True
        txtReturnDateFr.Enabled = True
        txtPayAppDateFr.Enabled = True
        txtPayAppDateTo.Enabled = True
        txtReturnDateTo.Enabled = True
        'Label6.Visible = False.
        ChkScan.Visible = False
        Label14.Visible = True
        ChkScan.Value = 0
        'ChkScan.Top = 1350
        'ChkScan.Value = 1
    'Else
    '    txtCLMNoFr.Enabled = False
    '    txtCLMNoTo.Enabled = False
    '    TxtBordx2.Visible = False
    '    txtCLMNoTo.Visible = True
    '    ChkScan.Top = 795
    End If
  
    lstBordxDelivery.ListItems.Clear
    Call ListViewHeader
End Sub

Private Sub ListViewHeader()
Dim AA As Integer

    If OptBordx.Value = True Then
        lstBordxDelivery.ColumnHeaders(1).Text = "Check"
        lstBordxDelivery.ColumnHeaders(2).Text = "Payor"
        lstBordxDelivery.ColumnHeaders(3).Text = "Bordx No"
        lstBordxDelivery.ColumnHeaders(4).Text = "Bordx Date"
        lstBordxDelivery.ColumnHeaders(5).Text = "BordxAmount"
        lstBordxDelivery.ColumnHeaders(5).Alignment = lvwColumnRight
        lstBordxDelivery.ColumnHeaders(6).Text = "Bordx Status"
        'lstBordxDelivery.ColumnHeaders(7).Text = "Returned fr Payor Status"
        lstBordxDelivery.ColumnHeaders(7).Text = "Date from Payor"
        lstBordxDelivery.ColumnHeaders(8).Text = "Payor Approval Date"
        lstBordxDelivery.ColumnHeaders(9).Text = "Payor Approval Code"
        lstBordxDelivery.ColumnHeaders(10).Text = ""
        lstBordxDelivery.ColumnHeaders(11).Text = ""
    Else
        lstBordxDelivery.ColumnHeaders(1).Text = "Check"
        lstBordxDelivery.ColumnHeaders(2).Text = "Payor"
        lstBordxDelivery.ColumnHeaders(3).Text = "Bordx No"
        lstBordxDelivery.ColumnHeaders(4).Text = "Bordx Date"
        lstBordxDelivery.ColumnHeaders(5).Text = "BordxAmount"
        lstBordxDelivery.ColumnHeaders(6).Text = "Claim No"
        lstBordxDelivery.ColumnHeaders(7).Text = "ClaimAmount"
        'lstBordxDelivery.ColumnHeaders(6).Alignment = lvwColumnRight
        lstBordxDelivery.ColumnHeaders(8).Text = "Bordx Status"
        'lstBordxDelivery.ColumnHeaders(8).Text = "Returned fr Payor Status"
        lstBordxDelivery.ColumnHeaders(9).Text = "Date from Payor"
        lstBordxDelivery.ColumnHeaders(10).Text = "Payor Approval Date"
        lstBordxDelivery.ColumnHeaders(11).Text = "Payor Approval Code"
        'lstBordxDelivery.ColumnHeaders(10).Text = "A/C Period"
    End If
End Sub

Private Function BordxNoCheck()

    If OptBordx.Value = True Then
        If Trim(ChkStr(TxtBordx1)) = ChkStr(Trim(lstBordxDelivery.ListItems(listloop1).SubItems(2))) Then
            Check = True
        End If
    Else
        If Trim(ChkStr(TxtBordx1)) = ChkStr(Trim(Mid(lstBordxDelivery.ListItems(listloop1).SubItems(2), 1, 7))) Then
            Check = True
        End If
    End If
End Function

'**********************Start BordxDelivery Listing  *******************************
Private Sub BordxDeliveryListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim TempBordxAmt As String
Dim TempBordxNo As String
    total = 0
    Current = 0
    
    lstBordxDelivery.ListItems.Clear

    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "Bordx_ReturnPayor"
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
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        CmdPrint.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        
            With rst2

                Set Record = lstBordxDelivery.ListItems.Add(, "")
                
                If Len(!BordxPAYCode) Then
                    Record.SubItems(1) = !BordxPAYCode
                End If
                If Len(!BordxRefNo) Then
                    
                    Record.SubItems(2) = !BordxRefNo
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(3) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                
                If Trim(TempBordxNo) <> Trim(Record.SubItems(2)) Then
                    If Len(!BordxClaimAmt) Then
                        Record.SubItems(4) = Format(!BordxClaimAmt, "#,##0.00")
                    End If
                Else
                    Record.SubItems(4) = "D"
                End If
                
                If Len(!CLMAppStatus) Then
                    Record.SubItems(5) = "Returned"
                End If
                If Len(!BordxPAYRecDocDate) Then
                    Record.SubItems(6) = Format(!BordxPAYRecDocDate, "dd/mm/yyyy")
                End If
                
                If Len(!CLMPayorAppDate) Then
                    Record.SubItems(7) = Format(!CLMPayorAppDate, "dd/mm/yyyy")
                End If
                If Len(!CLMPayorAppCode) Then
                    Record.SubItems(8) = !CLMPayorAppCode
                End If

                TempBordxNo = IIf(Len(Record.SubItems(2)), Record.SubItems(2), "")

                Current = Current + 1
                lblCurrent = Current
                If IsNumeric(rst2!BordxClaimAmt) Then
                    TotalAmt = TotalAmt + rst2!BordxClaimAmt
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
    CmdToPayor.Enabled = True
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub BordxDeliveryListing2(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sqlCount As String
Dim ClaimAdmin As Boolean

    Current2 = 0
    TotalAmt2 = 0
    strAppendCase = "2"
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "Bordx_ReturnPayor"
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
    'lstBordxDelivery.ListItems.Clear
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        CmdPrint.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
                
            With rst2
                Set Record = lstBordxDelivery.ListItems.Add(, "")
                
                If Len(rst2!BordxPAYCode) Then
                    Record.SubItems(1) = rst2!BordxPAYCode
                End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(2) = rst2!BordxRefNo
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(3) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(4) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!CLMAppStatus) Then
                    Record.SubItems(5) = "Returned"
                End If
                'If Len(rst2!BordxPAYDOCFRStatus) Then
                '   Record.SubItems(6) = rst2!BordxPAYDOCFRStatus
                'End If
                If Len(rst2!BordxPAYRecDocDate) Then
                    Record.SubItems(6) = Format(rst2!BordxPAYRecDocDate, "dd/mm/yyyy")
                End If
                
                If Len(!CLMPayorAppDate) Then
                    Record.SubItems(7) = Format(!CLMPayorAppDate, "dd/mm/yyyy")
                End If
                
                If Len(!CLMPayorAppCode) Then
                    Record.SubItems(8) = !CLMPayorAppCode
                End If
                
                Current2 = Current2 + 1
                lblTotal = Current2
                If Len(rst2!BordxClaimAmt) Then
                    TotalAmt2 = TotalAmt2 + rst2!BordxClaimAmt
                End If
                LblTotalAmt = Format(TotalAmt2, "#,##0.00")
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
    CmdToPayor.Enabled = True
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
End Sub
'**********************End BordxDelivery Listing *******************************
Private Sub ClaimListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
    
    total = 0
    Current = 0
    TotalAmt = 0
    lstBordxDelivery.ListItems.Clear

    strAppendCase = "3"
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "Bordx_ReturnPayor"
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
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        CmdPrint.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        
            With rst2
                Set Record = lstBordxDelivery.ListItems.Add(, "")
                
                If Len(rst2!BordxPAYCode) Then
                    Record.SubItems(1) = rst2!BordxPAYCode
                End If
                'If Len(rst2!BordxRefNo) Then
                'End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(2) = rst2!BordxRefNo
                End If
                
                If Len(rst2!BordxDate) Then
                    Record.SubItems(3) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(4) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                
                Record.SubItems(5) = rst2!CASNumber & "-" & rst2!claimversion
                
                If Len(rst2!CLMBordxAmt) Then
                   Record.SubItems(6) = rst2!CLMBordxAmt
                End If

                If Len(rst2!CLMAppStatus) Then
                    Record.SubItems(7) = "Returned"
                End If
                If Len(rst2!BordxPAYRecDocDate) Then
                    Record.SubItems(8) = Format(rst2!BordxPAYRecDocDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CLMPayorAppDate) Then
                    Record.SubItems(9) = Format(rst2!CLMPayorAppDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CLMPayorAppCode) Then
                    Record.SubItems(10) = rst2!CLMPayorAppCode
                End If

                'If Len(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR) Then
                '    Record.SubItems(9) = Trim(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR)
                'End If
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!BordxClaimAmt) Then
                    TotalAmt = TotalAmt + rst2!BordxClaimAmt
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
    CmdToPayor.Enabled = True
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim CAS As New ADODB.Recordset

    strAppend = ""
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    'If OptBordx.value = True Then
        If ChkScan.Value = 1 Then
            If Len(Trim(TxtBordx1)) Then
                strAppend = strAppend + " AND BordxRefNo = '" & Trim(TxtBordx1) & "'"
            End If
        Else
            If Len(Trim(TxtBordx1)) Then
                strAppend = strAppend + " AND BordxRefNo >= '" & Trim(TxtBordx1) & "'"
            End If
            
            If Len(Trim(TxtBordx2)) Then
                strAppend = strAppend + " AND BordxRefNo <= '" & Trim(TxtBordx2) & "'"
            End If
        End If
    'Else
        If ChkScan.Value = 1 Then
            If Len(Trim(txtCLMNoFr)) Then
                strAppend = strAppend + " AND CASNumber = '" & Trim(txtCLMNoFr) & "'"
            End If
        Else
            If Len(Trim(txtCLMNoFr)) Then
                strAppend = strAppend + " AND CASNumber >= '" & Trim(txtCLMNoFr) & "'"
            End If
            
            If Len(Trim(txtCLMNoTo)) Then
                strAppend = strAppend + " AND CASNumber <= '" & Trim(txtCLMNoTo) & "'"
            End If
        End If
    'End If
    If IsDate(txtBordxDate1) = True Then
        strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(txtBordxDate2) = True Then
        strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxAmt1)) Then
        strAppend = strAppend + " AND BordxClaimAmt >= '" & Trim(TxtBordxAmt1) & "'"
    End If
    
    If Len(Trim(TxtBordxAmt2)) Then
        strAppend = strAppend + " AND BordxClaimAmt <= '" & Trim(TxtBordxAmt2) & "'"
    End If
    
    If Len(cboStatus) Then
        If Mid(cboStatus, 1, 1) = "W" Then
            strAppend = strAppend + " AND (CLMAppStatus IS NULL OR CLMAppStatus = '')"
        Else
            strAppend = strAppend + " AND CLMAppStatus = '" & Mid(cboStatus, 1, 1) & "'"
        End If
    End If
    
    If Len(cboAppStatus) Then
        If Mid(cboAppStatus, 1, 1) = "N" Then
            strAppend = strAppend + " AND (CLMAppStatus IS NULL OR CLMAppStatus = '' OR CLMAppStatus = '" & Mid(cboAppStatus, 1, 1) & "')"
        Else
            strAppend = strAppend + " AND CLMAppStatus = '" & Mid(cboAppStatus, 1, 1) & "'"
        End If
    End If
    
    If IsDate(txtPayAppDateFr) = True Then
        strAppend = strAppend + " And CLMPayorAppDate >= '" & Format(txtPayAppDateFr, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(txtPayAppDateTo) = True Then
        strAppend = strAppend + " And CLMPayorAppDate <= '" & Format(txtPayAppDateTo, "mm/dd/yyyy") & "'"
    End If

    If OptWks.Value = True Then
        If IsDate(txtReturnDateFr) = True Then
            strAppend = strAppend + " And BordxPAYRecDocDate >= '" & Format(txtReturnDateFr, "mm/dd/yyyy") & "'"
        End If
      
        If IsDate(txtReturnDateTo) = True Then
            strAppend = strAppend + " And BordxPAYRecDocDate <= '" & Format(txtReturnDateTo, "mm/dd/yyyy") & "'"
        End If
    End If
    
    If Len(cboCashType) Then
        strAppend = strAppend + " And INVPayTo = '" & Mid(cboCashType, 1, 1) & "'"
    End If

    
    
    searchCriteria = strAppend
    
    If OptBordx.Value = True Then
        If ChkScan.Value = 0 Then
            Call BordxDeliveryListing(strAppend)
        Else
            Call BordxDeliveryListing2(strAppend)
        End If
    Else
        Call ClaimListing(strAppend)
    End If
End Function

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim MonthStart
Dim YearStart
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

    With CboBordxStatus
        .AddItem "A - Approved"
        .AddItem "R - Rejected"
        .AddItem "K - KIV"
        .AddItem "P - Partial Approved"
        .Text = "A - Approved"
    End With
    
    With cboStatus
        .AddItem "A - Approved"
        .AddItem "R - Rejected"
        .AddItem "K - KIV"
        .AddItem "P - Partial Approved"
        .AddItem "W - Without Status"
    End With
    
    With cboCashType
        .AddItem "H - Hopital"
        .AddItem "M - Member"
    End With
    
    With cboAppStatus
        .AddItem "N - No"
        .AddItem "Y - Yes"
    End With
    
    'cboREVBordxInd.AddItem "N - No"
    'cboREVBordxInd.AddItem "Y - Yes"
    
End Sub
'**********************End Combo Listing*******************************

Private Function UpdateReturnPayor()
Dim strsql As String
Dim strsql2 As String
'Dim FRPayor As New ADODB.Recordset
Dim RecDocDate As String
Dim ii As Integer
Dim TempBodrxNo As String
Dim TempCASNumber As String
Dim TempCLMVersion As String
Dim CLM As New ADODB.Recordset
Dim strsqlRem1 As String
Dim strsqlRem2 As String
Dim strsqlRem3 As String
Dim REM1 As New ADODB.Recordset
Dim REM2 As New ADODB.Recordset
Dim REM3 As New ADODB.Recordset
Dim RemLen1 As Integer
Dim RemLen2 As Integer
Dim RemLen3 As Integer
Dim RemLen4 As Integer
Dim RemLen5 As Integer
Dim RemLen6 As Integer
Dim RemLen7 As Integer
Dim Remarks1 As String
Dim Remarks2 As String
Dim Remarks3 As String
Dim Remarks4 As String
Dim Remarks5 As String
Dim Remarks6 As String
Dim Remarks7 As String
Dim AddRemark As String
Dim strsqlREM As String

    ' update bordx table
    RecDocDate = Format(TxtDate, "YYYY/MM/DD")
    For ii = 1 To lstBordxDelivery.ListItems.Count
       If lstBordxDelivery.ListItems(ii).Checked = True Then
       
            RemLen1 = 0
            RemLen2 = 0
            RemLen3 = 0
            RemLen4 = 0
            RemLen5 = 0
            RemLen6 = 0
            RemLen7 = 0
            
            Remarks1 = ""
            Remarks2 = ""
            Remarks3 = ""
            Remarks4 = ""
            Remarks5 = ""
            Remarks6 = ""
            Remarks7 = ""
        
        TempBodrxNo = lstBordxDelivery.ListItems(ii).ListSubItems(2)
        AddRemark = vbNewLine & Date & "   Received Returned Bordx  -  " & TempBodrxNo & " on " & TxtDate & "                           " & Logon & " "
        AddRemark = ChkStr(Replace(AddRemark, Chr(13), "~"))
        
        'strsql = " update Bordx set " & _
        '         " BordxPAYRecDocDate = '" & RecDocDate & "', " & _
        '         " BordxPayDOCFRStatus = '1', " & _
        '         " BordxPAYRecDocBy = '" & Logon & "' " & _
        '         " Where BordxRefNo = '" & Trim(TempBodrxNo) & "'"
        
        'medixmasterconnWS.Execute strsql
        
        If OptBordx.Value = True Then
            'update claim table
            strsql2 = ""
            'If TxtDate <> "__/__/____" And txtPAYAppDate <> "__/__/____" Then
           
            '    strsql2 = " update Claim set " & _
            '            " BordxPAYRecDocDate = '" & RecDocDate & "', " & _
            '            " CLMAppStatus = '" & Mid(cboBordxStatus, 1, 1) & "', " & _
            '            " CLMPayorAppStatus = '" & Mid(cboPayAppStatus, 1, 1) & "', " & _
            '            " CLMPayorAppDate = '" & Format(txtPAYAppDate, "yyyy/mm/dd") & "', " & _
            '            " CLMPayorAppCode = '" & IIf(Len(txtCLMAppCode), txtCLMAppCode, "") & "' ,CLMBordxRevInd = '" & IIf(Len(cboREVBordxInd), Mid(cboREVBordxInd, 1, 1), "") & "'" & _
            '            " Where BordxRefNo = '" & Trim(TempBodrxNo) & "'"
             If TxtDate <> "__/__/____" Then
             strsql2 = " update Claim set " & _
                        " BordxPAYRecDocDate = '" & RecDocDate & "', " & _
                        " CLMAppStatus = '" & Mid(CboBordxStatus, 1, 1) & "' " & _
                        " Where BordxRefNo = '" & Trim(TempBodrxNo) & "'"
                        
            'ElseIf TxtDate <> "__/__/____" And txtPAYAppDate = "__/__/____" Then
            '    strsql2 = " update Claim set " & _
            '            " BordxPAYRecDocDate = '" & RecDocDate & "', " & _
            '            " CLMAppStatus = '" & Mid(cboBordxStatus, 1, 1) & "', " & _
            '            " CLMPayorAppCode = '" & IIf(Len(txtCLMAppCode), txtCLMAppCode, "") & "' ,CLMBordxRevInd = '" & IIf(Len(cboREVBordxInd), Mid(cboREVBordxInd, 1, 1), "") & "'" & _
            '            " Where BordxRefNo = '" & Trim(TempBodrxNo) & "'"
            'ElseIf TxtDate = "__/__/____" And txtPAYAppDate <> "__/__/____" Then
            ' strsql2 = " update Claim set " & _
            '            " CLMPayorAppStatus = '" & Mid(cboPayAppStatus, 1, 1) & "', " & _
            '            " CLMPayorAppDate = '" & Format(txtPAYAppDate, "yyyy/mm/dd") & "', " & _
            '            " CLMPayorAppCode = '" & IIf(Len(txtCLMAppCode), txtCLMAppCode, "") & "' ,CLMBordxRevInd = '" & IIf(Len(cboREVBordxInd), Mid(cboREVBordxInd, 1, 1), "") & "'" & _
            '            " Where BordxRefNo = '" & Trim(TempBodrxNo) & "'"
            
            End If
        Else
            TempCASNumber = Trim(Mid(lstBordxDelivery.ListItems(ii).ListSubItems(5).Text, 1, IIf(InStr(1, lstBordxDelivery.ListItems(ii).ListSubItems(5).Text, "-") = 0, 10, InStr(1, lstBordxDelivery.ListItems(ii).ListSubItems(5).Text, "-") - 1)))
            TempCLMVersion = Trim(Mid(lstBordxDelivery.ListItems(ii).ListSubItems(5).Text, InStr(1, lstBordxDelivery.ListItems(ii).ListSubItems(5).Text, "-") + 1, Len(lstBordxDelivery.ListItems(ii).ListSubItems(5).Text) - InStr(1, lstBordxDelivery.ListItems(ii).ListSubItems(5).Text, "-")))
            strsql2 = ""

            'If TxtDate <> "__/__/____" And txtPAYAppDate <> "__/__/____" Then
            '    strsql2 = " update Claim set " & _
            '              " BordxPAYRecDocDate = '" & RecDocDate & "', " & _
            '              " CLMAppStatus = '" & Mid(cboBordxStatus, 1, 1) & "', " & _
            '              " CLMPayorAppStatus = '" & Mid(cboPayAppStatus, 1, 1) & "'," & _
            '              " CLMPayorAppDate = '" & Format(txtPAYAppDate, "yyyy/mm/dd") & "', " & _
            '              " CLMPayorAppCode = '" & IIf(Len(txtCLMAppCode), txtCLMAppCode, "") & "' , CLMBordxRevInd = '" & IIf(Len(cboREVBordxInd), Mid(cboREVBordxInd, 1, 1), "") & "'" & _
            '              " Where CASNumber = '" & Trim(TempCASNumber) & "' And ClaimVersion = '" & Trim(TempCLMVersion) & "'"
            If TxtDate <> "__/__/____" Then
                strsql2 = " update Claim set " & _
                          " BordxPAYRecDocDate = '" & RecDocDate & "', " & _
                          " CLMAppStatus = '" & Mid(CboBordxStatus, 1, 1) & "' " & _
                          " Where CASNumber = '" & Trim(TempCASNumber) & "' And ClaimVersion = '" & Trim(TempCLMVersion) & "'"

            'ElseIf TxtDate <> "__/__/____" And txtPAYAppDate = "__/__/____" Then
                'strsql2 = " update Claim set " & _
                          '" BordxPAYRecDocDate = '" & RecDocDate & "', " & _
                          '" CLMAppStatus = '" & Mid(cboBordxStatus, 1, 1) & "', " & _
                          '" CLMPayorAppCode = '" & IIf(Len(txtCLMAppCode), txtCLMAppCode, "") & "' , CLMBordxRevInd = '" & IIf(Len(cboREVBordxInd), Mid(cboREVBordxInd, 1, 1), "") & "'" & _
                          '" Where CASNumber = '" & Trim(TempCASNumber) & "' And ClaimVersion = '" & Trim(TempCLMVersion) & "'"
            'ElseIf TxtDate = "__/__/____" And txtPAYAppDate <> "__/__/____" Then
            '    strsql2 = " update Claim set " & _
            '              " CLMPayorAppStatus = '" & Mid(cboPayAppStatus, 1, 1) & "'," & _
            '              " CLMPayorAppDate = '" & Format(txtPAYAppDate, "yyyy/mm/dd") & "', " & _
            '              " CLMPayorAppCode = '" & IIf(Len(txtCLMAppCode), txtCLMAppCode, "") & "' , CLMBordxRevInd = '" & IIf(Len(cboREVBordxInd), Mid(cboREVBordxInd, 1, 1), "") & "'" & _
            '              " Where CASNumber = '" & Trim(TempCASNumber) & "' And ClaimVersion = '" & Trim(TempCLMVersion) & "'"
            End If
        End If
        medixmasterconnWS.Execute strsql2
        
        strsql2 = ""
        strsql2 = "Select CASNumber from Claim Where BordxRefNo = '" & Trim(TempBodrxNo) & "'"
        CLM.Open strsql2, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        If CLM.recordCount Then
            TempCASNumber = CLM!CASNumber
        End If
        CLM.Close
        Set CLM = Nothing
        
        If Len(TempCASNumber) Then
         '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

                        strsqlRem1 = "Select * from CASRemarks where CASNumber = '" & Trim(TempCASNumber) & "'"
                        REM1.Open strsqlRem1, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If REM1.recordCount > 0 Then
                                If REM1!CASORemarks <> "" Then
                                    RemLen1 = Len(REM1!CASORemarks)
                                    Remarks1 = REM1!CASORemarks & AddRemark
                                    Remarks1 = ChkStr(Replace(Remarks1, "'", "''"))
                                    If REM1!CASORemarksTwo <> "" Then
                                        RemLen2 = Len(REM1!CASORemarksTwo)
                                        Remarks2 = REM1!CASORemarksTwo & AddRemark
                                        Remarks2 = ChkStr(Replace(Remarks2, "'", "''"))
                                    End If
                                Else
                                    Remarks1 = AddRemark
                                End If
                            Else
                                Remarks1 = AddRemark
                            End If
                        REM1.Close
                        Set REM1 = Nothing
                        
                        strsqlRem2 = "Select * from CASRemarksTwo where CASNumber= '" & Trim(TempCASNumber) & "'"
                        REM2.Open strsqlRem2, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If REM2.recordCount Then
                                If REM2!CASORemarks <> "" Then
                                    RemLen3 = Len(REM2!CASORemarks)
                                    Remarks3 = REM2!CASORemarks & AddRemark
                                    Remarks3 = ChkStr(Replace(Remarks3, "'", "''"))
                                    RemLen4 = Len(REM2!CASORemarksTwo)
                                    Remarks4 = REM2!CASORemarksTwo & AddRemark
                                    Remarks4 = ChkStr(Replace(Remarks4, "'", "''"))
                                End If
                            End If
                        REM2.Close
                        Set REM2 = Nothing
                        'REM3.Close
                        'Set REM3 = Nothing
                        strsqlRem3 = ""
                        strsqlRem3 = "Select * from CASRemarkThree where CASNumber= '" & Trim(TempCASNumber) & "'"
                        REM3.Open strsqlRem3, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If REM3.recordCount Then
                            RemLen5 = Len(REM3!CasRemarks)
                            Remarks5 = REM3!CasRemarks & AddRemark
                            Remarks5 = ChkStr(Replace(Remarks5, "'", "''"))
                            RemLen6 = Len(REM3!CASRemarksFour)
                            Remarks6 = REM3!CASRemarksFour & AddRemark
                            Remarks6 = ChkStr(Replace(Remarks6, "'", "''"))
                            RemLen7 = Len(REM3!CasRemarksFive)
                            Remarks7 = REM3!CasRemarksFive & AddRemark
                            Remarks7 = ChkStr(Replace(Remarks7, "'", "''"))
                        Else
                            RemLen5 = 0
                            Remarks5 = ""
                            RemLen6 = 0
                            Remarks6 = ""
                            RemLen7 = 0
                            Remarks7 = ""
                        End If
                        REM3.Close
                        Set REM3 = Nothing
                        
                        strsqlREM = ""
                        If (RemLen7 <> 0 And RemLen7 < 3700) Or RemLen6 > 3700 Then
                            strsqlREM = " update CASRemarkThree set " & _
                                     " CASRemarksFive = '" & Remarks7 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf (RemLen6 <> 0 And RemLen6 < 3700) Or RemLen5 > 3700 Then
                            strsqlREM = " update CASRemarkThree set " & _
                                     " CASRemarksFour = '" & Remarks6 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf (RemLen5 <> 0 And RemLen5 < 3700) Or RemLen4 > 1800 Then
                            strsqlREM = " update CASRemarkThree set " & _
                                     " CasRemarks = '" & Remarks5 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf (RemLen4 <> 0 And RemLen4 < 1800) Or RemLen3 > 1800 Then
                            strsqlREM = " update CASRemarksTwo set " & _
                                     " CASORemarksTwo = '" & Remarks4 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf (RemLen3 <> 0 And RemLen3 < 1800) Or RemLen2 > 1800 Then
                            strsqlREM = " update CASRemarksTwo set " & _
                                     " CASORemarks = '" & Remarks3 & "' " & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf (RemLen2 <> 0 And RemLen2 < 1800) Or RemLen1 > 1800 Then
                            strsqlREM = " update CasRemarks set " & _
                                     " CASORemarksTwo = '" & Remarks2 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        ElseIf RemLen1 <> 0 And RemLen1 < 1800 Then
                            strsqlREM = " update CasRemarks set " & _
                                     " CASORemarks = '" & Remarks1 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"
                        Else
                            strsqlREM = " update CasRemarks set " & _
                                     " CASORemarks = '" & Remarks1 & "'" & _
                                     " where CASNumber = '" & Trim(TempCASNumber) & "'"

                        End If
                    medixmasterconn.Execute strsqlREM
                    'medixmasterconn.Close
                    'Set medixmasterconn = Nothing
                End If
            End If
    
    
    Next ii
        searchCriteria = ""
        searchCriteria = searchCriteria + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
        searchCriteria = searchCriteria + " And BordxPAYRecDocDate = '" & Format(TxtDate, "mm/dd/yyyy") & "'"
        
        If OptBordx.Value = True Then
            'If ChkScan.Value = 0 Then
                Call BordxDeliveryListing(searchCriteria)
            'Else
            '    Call BordxDeliveryListing2(searchCriteria)
            'End If
        Else
            Call ClaimListing(searchCriteria)
        End If
        
End Function

Private Sub txtBordxDate1_LostFocus()
    If IsDate(txtBordxDate1) Then
    Else
        If txtBordxDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDate2_LostFocus()
    If IsDate(txtBordxDate2) Then
    Else
        If txtBordxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate2.SetFocus
        End If
    End If

End Sub



Private Sub txtDate_LostFocus()
    If IsDate(TxtDate) Then
    Else
        If TxtDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDate.SetFocus
        End If
    End If
End Sub


'Private Sub txtPAYAppDate_LostFocus()
'    If IsDate(txtPAYAppDate) Then
'    Else
'        If txtPAYAppDate <> "__/__/____" Then
'            MsgBox "Invalid Date Format! ", vbCritical
'            txtPAYAppDate.SetFocus
'        End If
'    End If

'End Sub

Private Sub txtPayAppDateFr_LostFocus()
    If IsDate(txtPayAppDateFr) Then
    Else
        If txtPayAppDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPayAppDateFr.SetFocus
        End If
    End If

End Sub
Private Sub txtPayAppDateTo_LostFocus()
    If IsDate(txtPayAppDateTo) Then
    Else
        If txtPayAppDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPayAppDateTo.SetFocus
        End If
    End If

End Sub

Private Sub txtReturnDateFr_LostFocus()
    If IsDate(txtReturnDateFr) Then
    Else
        If txtReturnDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReturnDateFr.SetFocus
        End If
    End If

End Sub

Private Sub txtReturnDateTo_LostFocus()
    If IsDate(txtReturnDateTo) Then
    Else
        If txtReturnDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReturnDateTo.SetFocus
        End If
    End If

End Sub




