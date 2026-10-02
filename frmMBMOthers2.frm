VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMBMOthers2 
   Caption         =   "Member Others"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   60
      Top             =   5160
      Width           =   11340
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   10440
         TabIndex        =   62
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   9720
         TabIndex        =   61
         Top             =   165
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Member No:"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   63
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Height          =   4785
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   11340
      Begin VB.TextBox txtMBMAllergic 
         Height          =   1365
         Left            =   8400
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   3360
         Width           =   2850
      End
      Begin VB.Frame Frame14 
         Height          =   735
         Left            =   9480
         TabIndex        =   38
         Top             =   1245
         Width           =   1695
         Begin VB.TextBox txtPLNCode 
            Height          =   285
            Left            =   480
            MaxLength       =   4
            TabIndex        =   39
            Top             =   150
            Width           =   1100
         End
         Begin MSMask.MaskEdBox txtNewPolEffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   40
            Top             =   405
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   476
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            Caption         =   "PLN"
            Height          =   255
            Left            =   100
            TabIndex        =   42
            Top             =   160
            Width           =   405
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            Caption         =   "Eff"
            Height          =   255
            Left            =   100
            TabIndex        =   41
            Top             =   420
            Width           =   285
         End
      End
      Begin VB.TextBox txtRemarks 
         Height          =   1365
         Left            =   4920
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   3360
         Width           =   3255
      End
      Begin VB.Frame Frame12 
         Height          =   755
         Left            =   9480
         TabIndex        =   32
         Top             =   300
         Width           =   1695
         Begin MSMask.MaskEdBox TxtClmNonPayFr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   555
            TabIndex        =   33
            Top             =   150
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtClmNonPayTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   555
            TabIndex        =   34
            Top             =   390
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   80
            TabIndex        =   36
            Top             =   435
            Width           =   285
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   80
            TabIndex        =   35
            Top             =   200
            Width           =   405
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1335
         Left            =   4920
         TabIndex        =   17
         Top             =   300
         Width           =   4455
         Begin VB.TextBox TxtGenWP 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   21
            Top             =   240
            Width           =   705
         End
         Begin VB.TextBox txtPREWP 
            Height          =   285
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   20
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox TxtSpeWP 
            Height          =   300
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   19
            Top             =   720
            Width           =   705
         End
         Begin VB.ComboBox cboStaffDisc 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   960
            Width           =   705
         End
         Begin MSMask.MaskEdBox txtGENWPEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   22
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPREWPEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   23
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSPEWPEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   24
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label99 
            Caption         =   "Staff Disc Req"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label100 
            Caption         =   "End of Period"
            Height          =   255
            Left            =   2160
            TabIndex        =   30
            Top             =   765
            Width           =   1005
         End
         Begin VB.Label Label98 
            Caption         =   "End of Period"
            Height          =   255
            Left            =   2160
            TabIndex        =   29
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label97 
            Caption         =   "End of Period"
            Height          =   255
            Left            =   2160
            TabIndex        =   28
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label96 
            Caption         =   "Gen W. P."
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label Label95 
            Caption         =   "Pre W. P."
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label Label94 
            Caption         =   "Spe. Ill. W. P."
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   2025
         End
      End
      Begin VB.TextBox txtMBMTelNoH 
         Height          =   315
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   16
         Top             =   165
         Width           =   3105
      End
      Begin VB.TextBox txtMBMTelNoM 
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   15
         Top             =   450
         Width           =   3105
      End
      Begin VB.TextBox txtPOLSubNo 
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   14
         Top             =   720
         Width           =   3105
      End
      Begin VB.TextBox txtMBMExclusion 
         Height          =   1065
         Left            =   4920
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2040
         Width           =   6300
      End
      Begin VB.ComboBox txtMBMClinical 
         Height          =   315
         Left            =   5760
         TabIndex        =   12
         Top             =   5160
         Width           =   1290
      End
      Begin VB.TextBox txtMBMPrevPlan 
         Height          =   330
         Left            =   1545
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox txtMBMRBAmt 
         Height          =   330
         Left            =   3720
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox txtMBMPrevINSCom 
         Height          =   330
         Left            =   1560
         MaxLength       =   60
         TabIndex        =   9
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtMBMPAYRemarks 
         Height          =   300
         Left            =   1560
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1620
         Width           =   3135
      End
      Begin VB.ComboBox cboCOPayRB 
         Height          =   315
         ItemData        =   "frmMBMOthers2.frx":0000
         Left            =   1560
         List            =   "frmMBMOthers2.frx":0007
         TabIndex        =   7
         Top             =   1890
         Width           =   1095
      End
      Begin VB.ComboBox cboBTBG 
         Height          =   315
         ItemData        =   "frmMBMOthers2.frx":000F
         Left            =   1560
         List            =   "frmMBMOthers2.frx":0011
         TabIndex        =   6
         Top             =   2180
         Width           =   1095
      End
      Begin VB.TextBox txtPayPrem 
         Height          =   285
         Left            =   3630
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1890
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   3630
         TabIndex        =   4
         Top             =   2120
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "Claim Non-Pay Date"
         Height          =   375
         Left            =   9480
         TabIndex        =   55
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label Label41 
         Caption         =   "Allergic"
         Height          =   255
         Left            =   8400
         TabIndex        =   66
         Top             =   3120
         Width           =   1590
      End
      Begin VB.Label Label110 
         Caption         =   "Payor Prem"
         Height          =   255
         Left            =   2760
         TabIndex        =   59
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label Label106 
         Caption         =   "BTB Gurantee"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label105 
         Caption         =   "New Plan Code"
         Height          =   255
         Left            =   9480
         TabIndex        =   57
         Top             =   1080
         Width           =   1725
      End
      Begin VB.Label Label84 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   4920
         TabIndex        =   56
         Top             =   3120
         Width           =   1590
      End
      Begin VB.Label Label102 
         Caption         =   "Principal Limitation"
         Height          =   255
         Left            =   4920
         TabIndex        =   54
         Top             =   120
         Width           =   1590
      End
      Begin VB.Label Label93 
         Caption         =   "Co-Payment RB"
         Height          =   240
         Left            =   120
         TabIndex        =   53
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Label Label7 
         Caption         =   "Lbl Status"
         Height          =   255
         Left            =   2760
         TabIndex        =   52
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label101 
         Caption         =   "Payor Remarks"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label88 
         Caption         =   "T/O Pln Code"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label87 
         Caption         =   "Prev Ins Com"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1380
         Width           =   1140
      End
      Begin VB.Label Label86 
         Caption         =   "R&&B Amount"
         Height          =   255
         Left            =   2760
         TabIndex        =   48
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label28 
         Caption         =   "Tel Num (H)"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   165
         Width           =   1005
      End
      Begin VB.Label Label32 
         Caption         =   "H/P Num"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   450
         Width           =   1005
      End
      Begin VB.Label Label33 
         Caption         =   "Policy Sub No"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label40 
         Caption         =   "Principal Warranty "
         Height          =   255
         Left            =   4920
         TabIndex        =   44
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Label Label46 
         Caption         =   "Clininal (Y/N)"
         Height          =   255
         Left            =   4800
         TabIndex        =   43
         Top             =   5160
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12120
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Member Others"
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmMBMOthers2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
