VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmTempMBMReg 
   Caption         =   "Temporary Membership Registration"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   600
      Left            =   80
      TabIndex        =   184
      Top             =   240
      Width           =   11775
      Begin VB.CommandButton CmdBordx 
         Caption         =   "Clear"
         Height          =   285
         Left            =   10920
         TabIndex        =   192
         Top             =   200
         Width           =   600
      End
      Begin VB.ComboBox CboHLTCode 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   200
         Width           =   690
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   3480
         TabIndex        =   1
         Top             =   200
         Width           =   2535
      End
      Begin VB.TextBox txtMBMBatchNo 
         Height          =   285
         Left            =   9360
         MaxLength       =   2
         TabIndex        =   3
         Top             =   200
         Width           =   315
      End
      Begin VB.TextBox txtPAYCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6000
         TabIndex        =   185
         Top             =   200
         Visible         =   0   'False
         Width           =   105
      End
      Begin MSMask.MaskEdBox txtMBMBordxDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   2
         Top             =   195
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label44 
         Caption         =   "Last Batch"
         Height          =   255
         Index           =   1
         Left            =   9720
         TabIndex        =   191
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "Batch No."
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
         Index           =   0
         Left            =   8400
         TabIndex        =   190
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Bordx. Date"
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
         Left            =   6120
         TabIndex        =   189
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblLastBatch 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10560
         TabIndex        =   188
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label64 
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
         TabIndex        =   187
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label55 
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
         Left            =   2280
         TabIndex        =   186
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame FrameMBM1 
      Enabled         =   0   'False
      Height          =   600
      Left            =   80
      TabIndex        =   78
      Top             =   760
      Width           =   11775
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Enabled         =   0   'False
         Height          =   300
         Left            =   6840
         TabIndex        =   80
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Enabled         =   0   'False
         Height          =   300
         Left            =   6120
         TabIndex        =   79
         Top             =   200
         Width           =   735
      End
      Begin VB.TextBox txtMBMPrevMEMNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   5
         Top             =   200
         Width           =   2535
      End
      Begin VB.ComboBox cboMBMRenewal 
         Height          =   315
         ItemData        =   "FrmTempMBMReg.frx":0000
         Left            =   1320
         List            =   "FrmTempMBMReg.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "N"
         Top             =   200
         Width           =   690
      End
      Begin VB.Label txtMBMPrevPolNo 
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   9240
         TabIndex        =   84
         Top             =   195
         Width           =   2295
      End
      Begin VB.Label Label68 
         Caption         =   "Prev. Pol. Num"
         Height          =   255
         Left            =   7920
         TabIndex        =   83
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label67 
         Caption         =   "Prev Mem No"
         Height          =   240
         Left            =   2280
         TabIndex        =   82
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label66 
         Caption         =   "Renewal (Y/N)"
         Height          =   240
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame FrameMBM2 
      Enabled         =   0   'False
      Height          =   855
      Left            =   80
      TabIndex        =   60
      Top             =   1280
      Width           =   11775
      Begin VB.CommandButton cmdClinic 
         BackColor       =   &H00808080&
         Caption         =   "Clinical >"
         Height          =   255
         Left            =   9960
         MaskColor       =   &H00808080&
         TabIndex        =   10
         Top             =   520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboMBMSalutation 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Text            =   "MR"
         Top             =   160
         Width           =   690
      End
      Begin VB.TextBox txtMBMName 
         Height          =   330
         Left            =   3480
         MaxLength       =   60
         TabIndex        =   7
         Top             =   160
         Width           =   4245
      End
      Begin VB.TextBox txtMBMPolicyNo 
         Height          =   330
         Left            =   9240
         MaxLength       =   25
         TabIndex        =   9
         Top             =   160
         Width           =   2325
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   285
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   8
         Top             =   460
         Width           =   4245
      End
      Begin VB.Label Label65 
         Caption         =   "Mem Name"
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
         Left            =   2280
         TabIndex        =   77
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label54 
         Caption         =   "Policy Num"
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
         Left            =   7920
         TabIndex        =   76
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label8 
         Caption         =   "IC Num"
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
         Left            =   2280
         TabIndex        =   75
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Salutation"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   1140
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5445
      Left            =   120
      TabIndex        =   85
      Top             =   2160
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   9604
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Principal &Details"
      TabPicture(0)   =   "FrmTempMBMReg.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameMBM3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Supplementary"
      TabPicture(1)   =   "FrmTempMBMReg.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "lstCovAdd"
      Tab(1).ControlCount=   2
      Begin VB.Frame FrameMBM3 
         Enabled         =   0   'False
         Height          =   4965
         Left            =   120
         TabIndex        =   131
         Top             =   360
         Width           =   11535
         Begin VB.ComboBox cboINSCode 
            Height          =   315
            ItemData        =   "FrmTempMBMReg.frx":003C
            Left            =   7080
            List            =   "FrmTempMBMReg.frx":003E
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   150
            Width           =   1215
         End
         Begin VB.TextBox txtLifeTimeStatus 
            Height          =   285
            Left            =   10920
            TabIndex        =   147
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtRemarks 
            Height          =   1575
            Left            =   5880
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   2760
            Width           =   5130
         End
         Begin VB.ComboBox cboPLNCode 
            Height          =   315
            Left            =   9720
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   120
            Width           =   1140
         End
         Begin MSMask.MaskEdBox txtMBMPayorExpDate 
            Height          =   315
            Left            =   9720
            TabIndex        =   29
            Top             =   410
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMPayorEffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   7080
            TabIndex        =   28
            Top             =   430
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMAdd1 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   11
            Top             =   150
            Width           =   4095
         End
         Begin VB.TextBox txtMBMAdd2 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   12
            Top             =   400
            Width           =   4095
         End
         Begin VB.TextBox txtMBMAdd3 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   13
            Top             =   660
            Width           =   4095
         End
         Begin VB.TextBox txtMBMPostCode 
            Height          =   315
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   14
            Top             =   910
            Width           =   1215
         End
         Begin VB.ComboBox cboCTYCode 
            Height          =   315
            ItemData        =   "FrmTempMBMReg.frx":0040
            Left            =   3480
            List            =   "FrmTempMBMReg.frx":0042
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   910
            Width           =   1680
         End
         Begin VB.ComboBox cboSTACode 
            Height          =   315
            Left            =   1080
            TabIndex        =   16
            Top             =   1200
            Width           =   4095
         End
         Begin VB.ComboBox cboMBMTakeover 
            Height          =   315
            Left            =   3525
            TabIndex        =   18
            Text            =   "N"
            Top             =   1480
            Width           =   1650
         End
         Begin VB.ComboBox cboMBMMaritalStatus 
            Height          =   315
            Left            =   3525
            TabIndex        =   20
            Text            =   "M"
            Top             =   1760
            Width           =   1650
         End
         Begin VB.ComboBox cboMBMSex 
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   17
            Tag             =   "M"
            Text            =   "M"
            Top             =   1480
            Width           =   1215
         End
         Begin VB.ComboBox cboNATCode 
            Height          =   315
            Left            =   3525
            Sorted          =   -1  'True
            TabIndex        =   22
            Text            =   "MY-Malaysia"
            Top             =   2040
            Width           =   1650
         End
         Begin MSMask.MaskEdBox txtMBMBirthdate 
            Height          =   315
            Left            =   1080
            TabIndex        =   19
            Top             =   1760
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboRACCode 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   21
            Text            =   "C - Chinese"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox cboBRNCode 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   2320
            Width           =   4095
         End
         Begin VB.TextBox txtMBMAgentCode 
            Height          =   285
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   24
            Top             =   2600
            Width           =   4095
         End
         Begin MSMask.MaskEdBox txtMBMDateJoin 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   7080
            TabIndex        =   30
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMEmployeeNo 
            Height          =   315
            Left            =   7080
            MaxLength       =   15
            TabIndex        =   32
            Top             =   990
            Width           =   3825
         End
         Begin VB.TextBox txtMBMDepartment 
            Height          =   315
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   33
            Top             =   1275
            Width           =   3825
         End
         Begin VB.TextBox txtMBMGroupCompany 
            Height          =   315
            Left            =   7080
            MaxLength       =   150
            TabIndex        =   34
            Top             =   1560
            Width           =   3825
         End
         Begin VB.ComboBox CboInsMode 
            Height          =   315
            Left            =   9720
            TabIndex        =   36
            Text            =   "A - Annual"
            Top             =   1845
            Width           =   1185
         End
         Begin VB.ComboBox CboGuaRenewal 
            Height          =   315
            ItemData        =   "FrmTempMBMReg.frx":0044
            Left            =   7080
            List            =   "FrmTempMBMReg.frx":0046
            TabIndex        =   37
            Top             =   2130
            Width           =   3855
         End
         Begin VB.TextBox txtMBMExclusion 
            Height          =   1530
            Left            =   1080
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   2840
            Width           =   4095
         End
         Begin VB.Frame Frame4 
            Height          =   585
            Left            =   120
            TabIndex        =   132
            Top             =   4320
            Width           =   11055
            Begin VB.Label Label75 
               Caption         =   "Cancer Treat. Lmt"
               Height          =   255
               Left            =   2640
               TabIndex        =   146
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label74 
               Caption         =   "Home Nursing"
               Height          =   255
               Left            =   8520
               TabIndex        =   145
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label73 
               Caption         =   "Kidney Dialysis Lmt"
               Height          =   255
               Left            =   5400
               TabIndex        =   144
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblKidneyLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
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
               Left            =   6840
               TabIndex        =   143
               Top             =   240
               Width           =   1170
            End
            Begin VB.Label lblCancerLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
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
               Left            =   4080
               TabIndex        =   142
               Top             =   240
               Width           =   1170
            End
            Begin VB.Label lblHomeNurseLmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
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
               Left            =   9675
               TabIndex        =   141
               Top             =   240
               Width           =   1050
            End
            Begin VB.Label txtAnnualLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
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
               Left            =   960
               TabIndex        =   140
               Top             =   240
               Width           =   1200
            End
            Begin VB.Label txtMBMPremium 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
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
               Left            =   960
               TabIndex        =   139
               Top             =   600
               Width           =   1200
            End
            Begin VB.Label Label25 
               Caption         =   "Premium"
               Height          =   255
               Left            =   120
               TabIndex        =   138
               Top             =   600
               Width           =   600
            End
            Begin VB.Label Label26 
               Caption         =   "MCO Fees"
               Height          =   240
               Left            =   2640
               TabIndex        =   137
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label27 
               Caption         =   "Ann Limit"
               Height          =   255
               Left            =   120
               TabIndex        =   136
               Top             =   240
               Width           =   720
            End
            Begin VB.Label txtMCOFees 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
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
               Left            =   3600
               TabIndex        =   135
               Top             =   600
               Width           =   1200
            End
            Begin VB.Label txtBPrem 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
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
               Left            =   2640
               TabIndex        =   134
               Top             =   600
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label txtBMco 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
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
               Left            =   3600
               TabIndex        =   133
               Top             =   600
               Visible         =   0   'False
               Width           =   600
            End
         End
         Begin VB.Label Label4 
            Caption         =   "Guaranty Renew"
            Height          =   255
            Left            =   5880
            TabIndex        =   173
            Top             =   2175
            Width           =   1260
         End
         Begin VB.Label Label40 
            Caption         =   "Princ Warr"
            Height          =   255
            Left            =   180
            TabIndex        =   172
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label76 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   5880
            TabIndex        =   171
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label70 
            Caption         =   "Instl mode"
            Height          =   255
            Left            =   8925
            TabIndex        =   170
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label59 
            Caption         =   "Branch"
            Height          =   255
            Left            =   180
            TabIndex        =   169
            Top             =   2320
            Width           =   600
         End
         Begin VB.Label Label10 
            Caption         =   "Race"
            Height          =   255
            Left            =   180
            TabIndex        =   168
            Top             =   2040
            Width           =   600
         End
         Begin VB.Label Label57 
            Caption         =   "Dept. Name"
            Height          =   255
            Left            =   5880
            TabIndex        =   167
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "E'yee Num"
            Height          =   255
            Left            =   5880
            TabIndex        =   166
            Top             =   1020
            Width           =   1545
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Code"
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
            Left            =   8640
            TabIndex        =   165
            Top             =   210
            Width           =   945
         End
         Begin VB.Label Label16 
            Caption         =   "Exp Date"
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
            Left            =   8640
            TabIndex        =   164
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label Label63 
            Caption         =   "Take Over"
            Height          =   255
            Left            =   2520
            TabIndex        =   163
            Top             =   1480
            Width           =   825
         End
         Begin VB.Label Label24 
            Caption         =   "Marital Status"
            Height          =   255
            Left            =   2520
            TabIndex        =   162
            Top             =   1760
            Width           =   1095
         End
         Begin VB.Label Label42 
            Caption         =   "Date Join"
            Height          =   255
            Left            =   5880
            TabIndex        =   161
            Top             =   750
            Width           =   780
         End
         Begin VB.Label Label7 
            Caption         =   "State"
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
            Left            =   180
            TabIndex        =   160
            Top             =   1200
            Width           =   660
         End
         Begin VB.Label Label20 
            Caption         =   "Ins Type"
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
            Left            =   5880
            TabIndex        =   159
            Top             =   150
            Width           =   960
         End
         Begin VB.Label Label12 
            Caption         =   "PostCode"
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
            Left            =   180
            TabIndex        =   158
            Top             =   910
            Width           =   900
         End
         Begin VB.Label Label9 
            Caption         =   "Birthdate"
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
            Left            =   180
            TabIndex        =   157
            Top             =   1760
            Width           =   765
         End
         Begin VB.Label Label5 
            Caption         =   "Sex (M/F)"
            Height          =   255
            Left            =   180
            TabIndex        =   156
            Top             =   1480
            Width           =   825
         End
         Begin VB.Label Label2 
            Caption         =   "Address"
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
            Left            =   180
            TabIndex        =   155
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label15 
            Caption         =   "Eff Date"
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
            Left            =   5880
            TabIndex        =   154
            Top             =   440
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "City/Town"
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
            Left            =   2520
            TabIndex        =   153
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label11 
            Caption         =   "Nationality"
            Height          =   255
            Left            =   2520
            TabIndex        =   152
            Top             =   2040
            Width           =   960
         End
         Begin VB.Label Label14 
            Caption         =   "Agent "
            Height          =   255
            Left            =   180
            TabIndex        =   151
            Top             =   2600
            Width           =   660
         End
         Begin VB.Label Label43 
            Caption         =   "Pol Year"
            Height          =   255
            Left            =   8640
            TabIndex        =   150
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label Label18 
            Caption         =   "Grp Company"
            Height          =   255
            Left            =   5880
            TabIndex        =   149
            Top             =   1605
            Width           =   1545
         End
         Begin VB.Label txtAGECode 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7080
            TabIndex        =   35
            Top             =   1875
            Width           =   1785
         End
         Begin VB.Label Label29 
            Caption         =   "Age Band"
            Height          =   255
            Left            =   5880
            TabIndex        =   148
            Top             =   1905
            Width           =   1095
         End
         Begin VB.Label txtMBMPolicyYear 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "1"
            Height          =   315
            Left            =   9720
            TabIndex        =   31
            Top             =   720
            Width           =   1140
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   86
         Top             =   360
         Width           =   11580
         Begin VB.TextBox txtCovIcBcPp 
            Height          =   315
            Left            =   7440
            MaxLength       =   30
            TabIndex        =   41
            Top             =   180
            Width           =   2895
         End
         Begin VB.TextBox txtCovName 
            Height          =   315
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   40
            Top             =   180
            Width           =   3220
         End
         Begin VB.Frame Frame9 
            Height          =   1695
            Left            =   10440
            TabIndex        =   87
            Top             =   120
            Width           =   1065
            Begin VB.CommandButton CmdClearCov 
               Caption         =   "Clear"
               Height          =   315
               Left            =   120
               MousePointer    =   8  'Size NW SE
               TabIndex        =   91
               Top             =   1320
               Width           =   840
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               MousePointer    =   7  'Size N S
               TabIndex        =   90
               Top             =   960
               Width           =   840
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   315
               Left            =   120
               MousePointer    =   5  'Size
               TabIndex        =   89
               Top             =   240
               Width           =   840
            End
            Begin VB.CommandButton cmdRemove 
               Caption         =   "&Remove"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               MousePointer    =   6  'Size NE SW
               TabIndex        =   88
               Top             =   600
               Width           =   840
            End
         End
         Begin VB.TextBox txtCovExclusion 
            Height          =   915
            Left            =   6360
            MaxLength       =   2500
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   2160
            Width           =   3135
         End
         Begin VB.TextBox txtCLNCovExclusion 
            Enabled         =   0   'False
            Height          =   915
            Left            =   9600
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox txtCovAllergic 
            Height          =   915
            Left            =   4680
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   72
            Top             =   2160
            Width           =   1575
         End
         Begin MSMask.MaskEdBox TxtSuppClmNonPayFr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   9240
            TabIndex        =   92
            Top             =   1245
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtSuppClmNonPayTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   9240
            TabIndex        =   93
            Top             =   1510
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCovBirthdate 
            Height          =   315
            Left            =   3480
            TabIndex        =   43
            Top             =   460
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSuppEffDate 
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
            Left            =   7440
            TabIndex        =   45
            Top             =   465
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSuppExpDate 
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
            Left            =   9240
            TabIndex        =   46
            Top             =   465
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboCovSalutation 
            Height          =   315
            Left            =   1320
            TabIndex        =   39
            Top             =   195
            Width           =   975
         End
         Begin VB.ComboBox cboCovRelationship 
            Height          =   315
            Left            =   1320
            TabIndex        =   42
            Text            =   "W -Wife"
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cboCovAdultChild 
            Height          =   315
            Left            =   1320
            TabIndex        =   47
            Text            =   "A"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtAvailableLimit 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   50
            Top             =   1010
            Width           =   975
         End
         Begin VB.TextBox txtRBAmt 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   54
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox txtSuppPayPremGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   58
            Top             =   1545
            Width           =   975
         End
         Begin VB.TextBox txtSuppPayMCOGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   61
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtRenewalDisc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   64
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox txtPolicyDiscAmt 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   66
            Top             =   2295
            Width           =   975
         End
         Begin VB.TextBox txtLoadingPrem 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   68
            Top             =   2550
            Width           =   975
         End
         Begin VB.TextBox txtPrevPLNCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   70
            Top             =   2800
            Width           =   975
         End
         Begin VB.ComboBox cboCovSex 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            TabIndex        =   48
            Text            =   "M"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCovAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   51
            Top             =   1005
            Width           =   1095
         End
         Begin VB.ComboBox cboSuppStaffDisc 
            Height          =   315
            Left            =   3480
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   1260
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayPremNet 
            Height          =   285
            Left            =   3480
            TabIndex        =   59
            Top             =   1550
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayMCONet 
            Height          =   285
            Left            =   3480
            TabIndex        =   63
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtPolicyDisc 
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   65
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox txtRenewalDiscAmt 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   67
            Top             =   2280
            Width           =   1095
         End
         Begin VB.TextBox txtCovUndExcess 
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   69
            Top             =   2535
            Width           =   1095
         End
         Begin VB.TextBox txtPrevINSCom 
            Height          =   330
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   71
            Top             =   2760
            Width           =   1095
         End
         Begin VB.ComboBox CboSuppPln 
            Height          =   315
            Left            =   5640
            TabIndex        =   44
            Top             =   460
            Width           =   1065
         End
         Begin VB.TextBox txtCOVPayRemarks 
            Height          =   285
            Left            =   5640
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   740
            Width           =   4695
         End
         Begin VB.TextBox txtPayMBMNo 
            Height          =   285
            Left            =   5640
            MaxLength       =   8
            TabIndex        =   52
            Top             =   990
            Width           =   1095
         End
         Begin VB.TextBox txtClientNo 
            Height          =   285
            Left            =   5640
            MaxLength       =   8
            TabIndex        =   56
            Top             =   1240
            Width           =   1095
         End
         Begin VB.TextBox txtPaySeqID 
            Height          =   285
            Left            =   7440
            MaxLength       =   2
            TabIndex        =   53
            Top             =   990
            Width           =   1095
         End
         Begin VB.TextBox txtClientID 
            Height          =   285
            Left            =   7440
            MaxLength       =   2
            TabIndex        =   57
            Top             =   1240
            Width           =   1095
         End
         Begin VB.Label Label56 
            Caption         =   "Bal. Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label Label47 
            Caption         =   "Supplementary Warranty "
            Height          =   255
            Left            =   6360
            TabIndex        =   129
            Top             =   1920
            Width           =   2100
         End
         Begin VB.Label Label52 
            Caption         =   "R'Ship"
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
            TabIndex        =   128
            Top             =   525
            Width           =   1005
         End
         Begin VB.Label Label46 
            Caption         =   "Supp Lmt"
            Height          =   255
            Left            =   2355
            TabIndex        =   127
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label Label49 
            Caption         =   "DOB"
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
            Left            =   2355
            TabIndex        =   126
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label Label41 
            Caption         =   "IC No."
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
            Left            =   6760
            TabIndex        =   125
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label39 
            Caption         =   "Salutation"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label50 
            Caption         =   "Sex "
            Height          =   255
            Left            =   2355
            TabIndex        =   123
            Top             =   750
            Width           =   825
         End
         Begin VB.Label Label51 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   795
            Width           =   825
         End
         Begin VB.Label Label38 
            Caption         =   "Name"
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
            Left            =   2355
            TabIndex        =   121
            Top             =   195
            Width           =   1140
         End
         Begin VB.Label Label82 
            Caption         =   "P. Pln Code"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   2840
            Width           =   1125
         End
         Begin VB.Label Label90 
            Caption         =   "R&&B Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   1300
            Width           =   1140
         End
         Begin VB.Label Label91 
            Caption         =   "T/O Ins Com"
            Height          =   255
            Left            =   2400
            TabIndex        =   118
            Top             =   2820
            Width           =   1140
         End
         Begin VB.Label Label36 
            Caption         =   "Disc Req"
            Height          =   255
            Left            =   2355
            TabIndex        =   117
            Top             =   1300
            Width           =   1065
         End
         Begin VB.Label Label37 
            Caption         =   "Clm Non-Pay Date"
            Height          =   255
            Left            =   8580
            TabIndex        =   116
            Top             =   1050
            Width           =   1365
         End
         Begin VB.Label Label111 
            Caption         =   "Pay Prem Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   8580
            TabIndex        =   114
            Top             =   1290
            Width           =   405
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   8580
            TabIndex        =   113
            Top             =   1575
            Width           =   285
         End
         Begin VB.Label Label31 
            Caption         =   "Pay Prem Net"
            Height          =   255
            Left            =   2355
            TabIndex        =   112
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label13 
            Caption         =   "Pay MCO Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label23 
            Caption         =   "Pay MCO Net"
            Height          =   255
            Left            =   2355
            TabIndex        =   110
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label28 
            Caption         =   "Renewal Disc."
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   2055
            Width           =   1485
         End
         Begin VB.Label Label32 
            Caption         =   "Policy Disc."
            Height          =   255
            Left            =   2355
            TabIndex        =   108
            Top             =   2040
            Width           =   1365
         End
         Begin VB.Label Label108 
            Caption         =   "Supp. Clinic Warranty"
            Height          =   255
            Left            =   9600
            TabIndex        =   107
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label48 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   4680
            TabIndex        =   106
            Top             =   1920
            Width           =   600
         End
         Begin VB.Label Label58 
            Caption         =   "Loading Prem"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   2550
            Width           =   1365
         End
         Begin VB.Label Label60 
            Caption         =   "Pol. Disc. Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   2325
            Width           =   1365
         End
         Begin VB.Label Label61 
            Caption         =   "Ren. Disc. Amt"
            Height          =   255
            Left            =   2355
            TabIndex        =   103
            Top             =   2295
            Width           =   1125
         End
         Begin VB.Label Label62 
            Caption         =   "PLN Code"
            Height          =   255
            Index           =   0
            Left            =   4680
            TabIndex        =   102
            Top             =   560
            Width           =   885
         End
         Begin VB.Label Label22 
            Caption         =   "Und Excess"
            Height          =   255
            Left            =   2400
            TabIndex        =   101
            Top             =   2565
            Width           =   1005
         End
         Begin VB.Label Label62 
            Caption         =   "EffDate"
            Height          =   255
            Index           =   1
            Left            =   6760
            TabIndex        =   100
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label62 
            Caption         =   "ExpDate"
            Height          =   255
            Index           =   2
            Left            =   8580
            TabIndex        =   99
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label92 
            Caption         =   "Pay Remark."
            Height          =   255
            Left            =   4680
            TabIndex        =   98
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label45 
            Caption         =   "Pay MBM No"
            Height          =   255
            Index           =   0
            Left            =   4680
            TabIndex        =   97
            Top             =   1040
            Width           =   975
         End
         Begin VB.Label Label45 
            Caption         =   "Pay Seq"
            Height          =   255
            Index           =   1
            Left            =   6760
            TabIndex        =   96
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label45 
            Caption         =   "Client No"
            Height          =   255
            Index           =   2
            Left            =   4680
            TabIndex        =   95
            Top             =   1280
            Width           =   975
         End
         Begin VB.Label Label45 
            Caption         =   "Client ID"
            Height          =   255
            Index           =   3
            Left            =   6760
            TabIndex        =   94
            Top             =   1320
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   1755
         Left            =   -74880
         TabIndex        =   174
         Top             =   3600
         Width           =   11520
         _ExtentX        =   20320
         _ExtentY        =   3096
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   49
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mem Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Salutation"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "IC No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Birthdate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sex"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Adult/Child"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Relationship"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "CovID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Allergic"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Exclusion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Covered ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Annual Limit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Occupation"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Work Nature"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Smoking"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Alcohol"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Eff Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Exp Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "PrevPlnCode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "RBAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "PAY Remarks"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Prev Insurance Com"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Gen W.P"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "EOP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Pre W.P."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "EOP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Spe Ill W. P."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "EOP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Staff DR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   29
            Text            =   "Clm Non-Payment Fr"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   30
            Text            =   "Clm Non-Payment To"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "Clinic Warranty"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "PayorPrem"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "PayorPremNet"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "PayorMCOGross"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "PayorMCONet"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "RenewalDisc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "PolDisc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "PolDiscAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "RenDiscAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "LoadingAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "PLNCode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "AvailLimit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "UndExcess"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   44
            Text            =   "PayMBMNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   45
            Text            =   "PaySeqID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   46
            Text            =   "ClientNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   47
            Text            =   "ClientID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   48
            Text            =   "2ndIC"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   120
      TabIndex        =   176
      Top             =   7560
      Width           =   11715
      Begin VB.CommandButton CmdMBMOthers 
         Caption         =   "Others Details"
         Height          =   315
         Left            =   7200
         TabIndex        =   183
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtcount 
         Height          =   285
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   182
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtINSFlag 
         Height          =   285
         Left            =   3120
         TabIndex        =   181
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtMBMAdultChild 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   180
         Text            =   "A"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   9435
         TabIndex        =   179
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10440
         TabIndex        =   178
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   8445
         TabIndex        =   177
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Temporary Member Registration"
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
      TabIndex        =   175
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "FrmTempMBMReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordSaved As Boolean
Dim startTrans As Boolean
Dim bExit As Boolean
Dim tabIndexDone As Boolean
Dim MBMPayorEXPDate As Integer
Dim tempSPremium As String
Dim tempSAnnLmt As String
Dim txtTempProvider As Double
Dim txtTempPrem As Double
Dim txtTmpBasicMCO As Double
Dim txtTempExpDate  As String
Dim txtTempEffDate As String
Dim txtFirstMeMNo As String
Dim txtFirstPolNo As String
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim Disc As Boolean
Public PolDics As String
Public RenDisc As String
Dim ADOrstMBM  As New ADODB.Recordset
Dim sql As String
Dim MembershipNo As String
Dim completeMBM, completeMBMOthers, completeMBMC, completeMBMCLNOthers As Boolean
Dim completeMBMTwo, completeMBMOthersTwo, completeMBMCRef As Boolean
Dim TempPayEffDate As String
Dim TempPayExpDate As String
Dim WLifeStatus As Boolean
Dim PrevLifeMBMNo As String
Dim TempLifeTimeStatus As String
Dim tPLNPrmIND As String
Dim tPLNMcoIND As String
Dim tPLNAnnLmtIND As String
Dim tPLNAgeLimit As String
Dim tmpPLNProRate As String
Dim RoundUpStatus As String
Dim TmpTopupNo As String
Dim CISDCStatus As String
Dim CISDCPayor As String
Dim CISPDentStatus As String
Dim CISSDentStatus As String
Dim CISPSpecStatus As String
Dim CISSSpecStatus As String
Dim CISPMepStatus As String
Dim CISSMepStatus As String
Dim OverAge As Boolean
Dim PlnAgeLmt As Integer
Dim tmpPremSexStatus As String
Dim TmpChkDigits As String

Private Sub CboGuaRenewal_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboGuaRenewal, KeyAscii)
End Sub

Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboHLTCode, KeyAscii)
End Sub

Private Sub CboHLTCode_LostFocus()
Dim SelFields As String
Dim SelCriteria As String
    If Trim(CboHLTCode) <> "" Then
        CboPayor.Clear
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        SelFields = "SELECT PAYCode as tmpCode, PAYCompanyName as tmpName "
        SelCriteria = "  HLTCode ='" & Trim(CboHLTCode) & "'"
        Call ComboHISDBLoad(SelFields, "PAY", SelCriteria, CboPayor, "3")
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
    Call EnableFrame
End Sub

Private Sub cboINSCode_LostFocus()
Dim SelFields As String
Dim SelCriteria As String

    If Len(Trim(cboINSCode)) Then
        If Trim(cboINSCode) = "F" Or Trim(cboINSCode) = "H" Or Trim(frmMBMClinical.cboCisINS) = "F" Or Trim(frmMBMClinical.cboCisINS) = "H" Then
            cboMBMMaritalStatus = "M"
            Call EnableSupplementary(True)
        Else
            cboMBMMaritalStatus = "S"
            Call EnableSupplementary(False)
        End If
    End If
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If Len(Trim(cboPLNCode)) = 0 Then
        SelFields = "Select PLNCode as tmpCode,PLNDescription as tmpName "
        SelCriteria = ""
        If Trim(CboHLTCode) = "N" Then
            SelCriteria = SelCriteria & "  PAYCode= '" & Trim(txtPAYCode) & "'"  'need to reverify
            SelCriteria = SelCriteria & " AND HLTCode = 'N'"
        Else
            SelCriteria = SelCriteria & "  HLTCode = 'S'"
        End If
        Call ComboHISDBLoad(SelFields, "PLN", SelCriteria, cboPLNCode, "1")
        Call ComboHISDBLoad(SelFields, "PLN", SelCriteria, CboSuppPln, "1")

    End If
    If Trim(cboPLNCode) <> "" Then
        Call CallLimit
        Call MemberCallFees
    End If
   ' medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub cboMBMSalutation_LostFocus()
    cboMBMSalutation = ChkStr(cboMBMSalutation)
End Sub

Private Sub cboMBMSex_LostFocus()
    cboMBMSex.Text = ChkStr(cboMBMSex)
End Sub

Private Sub cboMBMTakeover_LostFocus()
    cboMBMTakeover.Text = ChkStr(cboMBMTakeover.Text)
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboPayor, KeyAscii)
End Sub

Private Sub CboPayor_LostFocus()
Dim AA As Integer
Dim SelFields As String
Dim SelCriteria As String
Dim PayRec As New ADODB.Recordset
    txtPAYCode = ""
    If Trim(CboPayor) <> "" Then
        AA = InStr(1, CboPayor, "-")
        txtPAYCode = Trim(Mid(CboPayor.Text, 1, AA - 1))
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        SelFields = "SELECT HLTCode From Pay Where PayCode='" & txtPAYCode & "'"
        PayRec.Open SelFields, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not PayRec.EOF
            CboHLTCode.Text = IIf(Len(PayRec!HLTCode), Trim(PayRec!HLTCode), "")
            PayRec.MoveNext
        Loop
        PayRec.Close
        Set PayRec = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
    End If
    Call EnableFrame
End Sub

Private Sub cboPLNCode_GotFocus()
Dim SelFields As String
Dim SelCriteria As String
    If Trim(CboHLTCode) <> "" Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        SelFields = "Select PLNCode as tmpCode,PLNDescription as tmpName "
        SelCriteria = ""
        SelCriteria = SelCriteria & "  HLTCode = '" & Trim(CboHLTCode) & "'"
        Call ComboHISDBLoad(SelFields, "PLN", SelCriteria, cboPLNCode, "1")
       ' 'medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
End Sub

Private Sub cboPLNCode_LostFocus()

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call PLNLostFocus
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    If Trim(cboPLNCode) <> "" Then
        CboSuppPln.Text = cboPLNCode
    End If
End Sub

Private Sub CboSuppPln_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboSuppPln, KeyAscii)
End Sub

Private Sub CboSuppPln_LostFocus()
Dim TempSuppExpDate As String
    CboSuppPln = IIf(Trim(CboSuppPln) <> "", CboSuppPln, cboPLNCode)
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If Trim(CboSuppPln) <> "" And IsDate(txtCovBirthdate) = True And IsDate(txtSuppEffDate) And IsDate(txtSuppExpDate) Then
        TempSuppExpDate = GetSuppExpDate(txtCovBirthdate, CboSuppPln, txtSuppEffDate, txtSuppExpDate)
        txtSuppExpDate = TempSuppExpDate
    End If
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    If OverAge = True Then
        MsgBox "Supplementary age is above " & PlnAgeLmt
    End If
End Sub


Private Sub CmdBordx_Click()
    CboHLTCode.Text = ""
    CboPayor.Text = ""
    txtMBMBatchNo = ""
    txtMBMBordxDate.Text = "__/__/____"
    lblLastBatch = ""
    Call EnableFrame
End Sub

Private Sub CmdClearCov_Click()
    Call ClearCoveredEntry
End Sub

Private Sub cmdClinic_Click()
    frmMBMClinical.Show
End Sub

Private Sub CmdMBMOthers_Click()
    frmNewMBMOthers.Source = 5
    
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    frmNewMBMOthers.Show
    
    frmNewMBMOthers.Frame5.Enabled = True
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

Private Sub CboInsMode_GotFocus()
    KeyPreview = True
End Sub

Private Sub CboInsMode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SSTab1.Tab = 1
        tabIndexDone = True
    End If
End Sub

Private Sub cboPLNCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 0
        Shift = 0
        Clipboard.Clear
    End If
End Sub

Private Sub Form_Load()
    Call ComboListing
    cboMBMRenewal.SelStart = 0
    bExit = False
    tabIndexDone = False
    txtTempProvider = 0
    txtTempPrem = 0
    PolDics = 0
    RenDisc = 0
    If frmNewMBMOthers.Visible = False Then
        frmNewMBMOthers.Source = 2
        Load frmNewMBMOthers
    End If
    Call ClearForm(frmNewMBMOthers)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub ComboListing()
Dim SelFields As String
Dim SelCriteria As String
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    SelFields = "SELECT HLTCode as tmpCode, HLTDescription as tmpName "
    SelCriteria = "  HLTCode is not null "
    Call ComboHISDBLoad(SelFields, "HLT", SelCriteria, CboHLTCode, "1")
    
    SelFields = "SELECT RACCode as tmpCode, RACName as tmpName "
    SelCriteria = "  RACCode is not null "
    Call ComboHISDBLoad(SelFields, "RAC", SelCriteria, cboRACCode, "3")

    SelFields = "SELECT RACCode as tmpCode, RACName as tmpName "
    SelCriteria = "  RACCode is not null "
    Call ComboHISDBLoad(SelFields, "RAC", SelCriteria, cboRACCode, "3")

    SelFields = "SELECT NATCode as tmpCode, NATName as tmpName "
    SelCriteria = "  NATCode is not null "
    Call ComboHISDBLoad(SelFields, "NAT", SelCriteria, cboNATCode, "3")
    
    SelFields = "SELECT RELCODE as tmpCode, RELDescription as tmpName "
    SelCriteria = "  RELCODE is not null "
    Call ComboHISDBLoad(SelFields, "REL", SelCriteria, cboCovRelationship, "3")
    
    SelFields = "SELECT CTYCode as tmpCode, CTYDescription as tmpName "
    SelCriteria = "  CTYCode is not null "
    Call ComboMaintLoad(SelFields, "CTY", SelCriteria, cboBRNCode, "2")
    Call ComboMaintLoad(SelFields, "CTY", SelCriteria, cboCTYCode, "2")
    
    SelFields = "SELECT STACode as tmpCode, STADescription as tmpName "
    SelCriteria = "  STACode is not null "
    Call ComboHISDBLoad(SelFields, "STA", SelCriteria, cboSTACode, "2")
    
    SelFields = "Select INSCode as tmpCode , INSDescription  as tmpName "
    SelCriteria = " 0=0 order by INSCode "
    Call ComboHISDBLoad(SelFields, "INS", SelCriteria, cboINSCode, "1")
    
    SelFields = "Select PLNCode as tmpCode,PLNDescription as tmpName "
    SelCriteria = ""
    If Trim(CboHLTCode) = "N" Then
        SelCriteria = SelCriteria & "  PAYCode= '" & Trim(txtPAYCode) & "'"  'need to reverify
        SelCriteria = SelCriteria & " AND HLTCode = 'N'"
    Else
        SelCriteria = SelCriteria & "  HLTCode = 'S'"
    End If
    Call ComboHISDBLoad(SelFields, "PLN", SelCriteria, cboPLNCode, "1")
    Call ComboHISDBLoad(SelFields, "PLN", SelCriteria, CboSuppPln, "1")
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    CboInsMode.AddItem "A - Annually"
    CboInsMode.AddItem "Q - Quarterly"
    CboInsMode.AddItem "S - Semi Annual"
    CboInsMode.AddItem "M - Monthly"
        
    cboMBMSex.AddItem "M"
    cboMBMSex.AddItem "F"
    cboMBMSex.AddItem ""
    
    cboCovSex.AddItem "M"
    cboCovSex.AddItem "F"
    cboCovSex.AddItem ""
    
    cboMBMMaritalStatus.AddItem "S"
    cboMBMMaritalStatus.AddItem "M"
    
    cboCovAdultChild.AddItem "A"
    cboCovAdultChild.AddItem "C"
    
    cboMBMSalutation.AddItem "  "
    cboMBMSalutation.AddItem "MR"
    cboMBMSalutation.AddItem "MS"
    cboMBMSalutation.AddItem "MDM"
    cboMBMSalutation.AddItem "MRS"
    
    cboCovSalutation.AddItem "  "
    cboCovSalutation.AddItem "MR"
    cboCovSalutation.AddItem "MS"
    cboCovSalutation.AddItem "MDM"
    cboCovSalutation.AddItem "MRS"

    cboMBMTakeover.AddItem "N - None"
    cboMBMTakeover.AddItem "Y - Take Over"
    cboMBMTakeover.AddItem "C - Conversion"
    
    cboSuppStaffDisc.AddItem "N - No"
    cboSuppStaffDisc.AddItem "Y - Yes"
        
    cboMBMRenewal.AddItem "N"
    cboMBMRenewal.AddItem "R"
    cboMBMRenewal.AddItem "T"
    cboMBMRenewal.AddItem "P"
    
    CboGuaRenewal.Clear
    CboGuaRenewal.AddItem "  "
    CboGuaRenewal.AddItem "1 - No Life Time Limit"
    CboGuaRenewal.AddItem "3 - With Life Time Limit"
End Sub

Private Sub showMBMCov()
Dim MemberConveredPerson As ListItem

    lstCovAdd.ListItems.Clear
    Do While Not ADOrstMBM.EOF
        With ADOrstMBM
            Set MemberConveredPerson = lstCovAdd.ListItems.Add(, , IIf(!MBMCName <> "", !MBMCName, ""))
                MemberConveredPerson.SubItems(1) = IIf(Len(!MBMCSalutation), Trim(!MBMCSalutation), "")
                MemberConveredPerson.SubItems(2) = IIf(Len(!MBMCICBCPPNo), Trim(!MBMCICBCPPNo), "")
                MemberConveredPerson.SubItems(3) = !MBMCDOB
                MemberConveredPerson.SubItems(4) = IIf(Len(!MBMCSex), Trim(!MBMCSex), "")
                MemberConveredPerson.SubItems(5) = IIf(Len(!MBMCAdultChild), Trim(!MBMCAdultChild), "")
                MemberConveredPerson.SubItems(6) = IIf(Len(!MBMCRELCode), Trim(!MBMCRELCode), "")
                MemberConveredPerson.SubItems(7) = IIf(Len(!MBMCCoverID), Trim(!MBMCCoverID), "")
                MemberConveredPerson.SubItems(8) = IIf(Len(!MBMCAllergic), Trim(!MBMCAllergic), "")
                MemberConveredPerson.SubItems(9) = IIf(Len(!MBMCExclusion), Trim(!MBMCExclusion), "")
                MemberConveredPerson.SubItems(10) = IIf(Len(!MBMCCoverID), Trim(!MBMCCoverID), "")
                MemberConveredPerson.SubItems(11) = IIf(IsNumeric(!MBMCAnnualLimit), !MBMCAnnualLimit, 0)
                MemberConveredPerson.SubItems(12) = IIf(Len(!MBMCOccupation), Trim(!MBMCOccupation), "")
                MemberConveredPerson.SubItems(13) = IIf(Len(!MBMCWorkNature), Trim(!MBMCWorkNature), "")
                MemberConveredPerson.SubItems(14) = IIf(Len(!MBMCSmoking), Trim(!MBMCSmoking), "")
                MemberConveredPerson.SubItems(15) = IIf(Len(!MBMCAlcohol), Trim(!MBMCAlcohol), "")
                MemberConveredPerson.SubItems(36) = IIf(IsNumeric(!MBMCRenewalDisc), !MBMCRenewalDisc, 0)
                MemberConveredPerson.SubItems(37) = IIf(IsNumeric(!MBMCPolicyDisc), !MBMCPolicyDisc, 0)
                MemberConveredPerson.SubItems(32) = IIf(IsNumeric(!MBMCPayorPremGross), !MBMCPayorPremGross, 0)
                MemberConveredPerson.SubItems(33) = IIf(IsNumeric(!MBMCPayorPremNet), !MBMCPayorPremNet, 0)
                MemberConveredPerson.SubItems(34) = IIf(IsNumeric(!MBMCPayorMCOGross), !MBMCPayorMCOGross, 0)
                MemberConveredPerson.SubItems(35) = IIf(IsNumeric(!MBMCPayorMCONet), !MBMCPayorMCONet, 0)
                MemberConveredPerson.SubItems(41) = IIf(Len(!MBMCPLNCode), Trim(!MBMCPLNCode), cboPLNCode.Text)
                MemberConveredPerson.SubItems(42) = IIf(IsNumeric(!MBMCAvailableLimit), !MBMCAvailableLimit, 0)
                MemberConveredPerson.SubItems(44) = IIf(Len(!MBMCPayNo), Trim(!MBMCPayNo), "")
                MemberConveredPerson.SubItems(45) = IIf(Len(!MBMCPaySeqNo), Trim(!MBMCPaySeqNo), "")
                MemberConveredPerson.SubItems(46) = IIf(Len(!MBMCClientNo), Trim(!MBMCClientNo), "")
                MemberConveredPerson.SubItems(47) = IIf(Len(!MBMCClientID), Trim(!MBMCClientID), "")
                MemberConveredPerson.SubItems(48) = IIf(Len(!MBMCIcBcPp2nd), Trim(!MBMCIcBcPp2nd), "")
            
            End With
            ADOrstMBM.MoveNext
        Loop
End Sub


Private Sub EnableEntries(status As Boolean)
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = status
        End If
    Next ctl
    
    cboMBMSex.Enabled = False
    'txtPAYCode.Enabled = False
    'txtMBMBordxDate.Enabled = False
    'CboHLTCode.Enabled = False
    txtMBMAdultChild.Enabled = False
    'txtCovAnnualLimit.Enabled = False
End Sub

Private Sub ClearEntries()
    On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Enabled = True Then
               If ctl.Name <> "txtMBMBatchNo" Then
                    ctl.Text = ""
               End If
            End If
        ElseIf TypeOf ctl Is ComboBox Then
            If ctl.Enabled = True Then
                If ctl.Name <> "CboHLTCode" And ctl.Name <> "CboPayor" Then
                    ctl.Text = ""
                End If
            End If
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is MaskEdBox Then
            If ctl.Enabled = True Then
                If ctl.Name <> "txtMBMBordxDate" Then
                    ctl.mask = ""
                    ctl.Text = ""
                    ctl.mask = "##/##/####"
                End If
            End If
        End If
    Next ctl
    cboCovRelationship.Text = "W - Wife"
    cboCovSalutation.Text = "MR"
    cboCovSex.Text = "M"
    cboCovAdultChild.Text = "A"
    txtCovAnnualLimit = 0
    cboMBMSalutation.Text = "MR"
    cboMBMRenewal = "N"
    cboMBMSex = "M"
    cboMBMTakeover = "N - None"
    cboMBMMaritalStatus = "M"
    cboRACCode.Text = "C - Chinese"
    cboNATCode.Text = "MY - Malaysia"
    cboINSCode = " "
    lstCovAdd.ListItems.Clear
    txtAGECode = ""
    txtMBMPremium = ""
    txtMCOFees = ""
    txtAnnualLimit = ""
    lblKidneyLmt = ""
    lblHomeNurseLmt = ""
    lblCancerLmt = ""
    txtSuppExpDate.mask = ""
    txtSuppExpDate.Text = ""
    txtSuppExpDate.mask = "##/##/####"
    frmNewMBMOthers.txtGENWPEnd.mask = ""
    frmNewMBMOthers.txtGENWPEnd.Text = ""
    frmNewMBMOthers.txtGENWPEnd.mask = "##/##/####"
    frmNewMBMOthers.txtPREWPEnd.mask = ""
    frmNewMBMOthers.txtPREWPEnd.Text = ""
    frmNewMBMOthers.txtPREWPEnd.mask = "##/##/####"
    frmNewMBMOthers.txtSPEWPEnd.mask = ""
    frmNewMBMOthers.txtSPEWPEnd.Text = ""
    frmNewMBMOthers.txtSPEWPEnd.mask = "##/##/####"
    frmNewMBMOthers.TxtClmNonPayFr.Text = ""
    frmNewMBMOthers.TxtClmNonPayFr.mask = "##/##/####"
    frmNewMBMOthers.TxtClmNonPayTo.Text = ""
    frmNewMBMOthers.TxtClmNonPayTo.mask = "##/##/####"
    frmNewMBMOthers.cboBTBG = "N - No"
    frmNewMBMOthers.txtUndExcess = ""
    frmMBMClinical.cboCisINS.Text = ""
    frmMBMClinical.cboCisPln.Text = ""
    frmMBMClinical.CboCardType.Text = ""
    frmMBMClinical.txtCLNEffDate = "__/__/____"
    frmMBMClinical.txtClnExpDate = "__/__/____"
    frmMBMClinical.txtAGECode = ""
    frmMBMClinical.lblAnnualLimit = ""
    frmMBMClinical.lblMBMPremium = ""
    frmMBMClinical.LblVisitLimit = ""
    frmMBMClinical.lblMCOFees = ""
    TempLifeTimeStatus = ""
    MembershipNo = ""
End Sub

Private Sub lstCovAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstCovAdd.ListItems.Count Then
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovName = lstCovAdd.SelectedItem.Text
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(11), "###,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(18)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(19)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(20)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(21)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(29)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(30)
        txtCLNCovExclusion = lstCovAdd.SelectedItem.SubItems(31)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(32)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(33)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(34)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(35)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(36)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(37)
        txtPolicyDiscAmt = lstCovAdd.SelectedItem.SubItems(38)
        txtRenewalDiscAmt = lstCovAdd.SelectedItem.SubItems(39)
        txtLoadingPrem = lstCovAdd.SelectedItem.SubItems(40)
        CboSuppPln = lstCovAdd.SelectedItem.SubItems(41)
        txtAvailableLimit = Format(lstCovAdd.SelectedItem.SubItems(42), "###,##0.00")

        Call EnableCovered(False)
        cmdUpdate.Enabled = True
        cmdRemove.Enabled = True
        cmdAdd.Enabled = False
    End If
    Call EnableCovered(True)
End Sub

Private Sub lstCovAdd_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstCovAdd.ListItems.Count Then
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovName = lstCovAdd.SelectedItem.Text
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(11), "###,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(18)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(19)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(20)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(21)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(29)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(30)
        txtCLNCovExclusion = lstCovAdd.SelectedItem.SubItems(31)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(32)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(33)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(34)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(35)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(36)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(37)
        txtPolicyDiscAmt = lstCovAdd.SelectedItem.SubItems(38)
        txtRenewalDiscAmt = lstCovAdd.SelectedItem.SubItems(39)
        txtLoadingPrem = lstCovAdd.SelectedItem.SubItems(40)
        CboSuppPln = lstCovAdd.SelectedItem.SubItems(41)
        txtAvailableLimit = Format(lstCovAdd.SelectedItem.SubItems(42), "###,##0.00")

        Call EnableCovered(False)
        cmdUpdate.Enabled = True
        cmdRemove.Enabled = True
        cmdAdd.Enabled = False
    End If
    Call EnableCovered(True)
End Sub

'  Contorl of Tab index  *********************************88

Private Sub SSTab1_Click(PreviousTab As Integer)
    txtCLNCovExclusion.Enabled = False
    If SSTab1.Tab = 1 Then
'        If Trim(Mid(cboINSCode, 1, 1)) = "I" Or Trim(Mid(cboINSCode, 1, 1)) = "G" Then
        If Trim(Mid(cboINSCode, 1, 1)) = "F" Or Trim(Mid(cboINSCode, 1, 1)) = "H" Or _
            Trim(Mid(frmMBMClinical.cboCisINS, 1, 1)) = "F" Or Trim(Mid(frmMBMClinical.cboCisINS, 1, 1)) = "H" Then
            cboCovSalutation.Enabled = True
            If Trim(Mid(frmMBMClinical.cboCisINS, 1, 1)) = "F" Or Trim(Mid(frmMBMClinical.cboCisINS, 1, 1)) = "H" Then
                txtCLNCovExclusion.Enabled = True
            End If
            cboCovSalutation.SetFocus
        Else
            MsgBox "No Covered Person Needed For This Insured Type", vbOKOnly + vbCritical, DialogBoxTitle
            cboINSCode.SetFocus
            SSTab1.Tab = 0
        End If
    ElseIf SSTab1.Tab = 0 Then
        If txtMBMName.Enabled = True Then
            cboMBMRenewal.SetFocus
        End If
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub

Private Sub cmdExit_LostFocus()
    If Me.ActiveControl.TabIndex = 39 Then
        cboMBMRenewal.SetFocus
    ElseIf Me.ActiveControl.TabIndex = 53 Then
        cboMBMRenewal.SetFocus
    End If
End Sub

'Private Sub SSTab1_GotFocus()
'    If SSTab1.Tab = 0 Then
'        txtMBMAdd1.SetFocus
'    End If
'End Sub

'Private Sub SSTab1_LostFocus()
'    If SSTab1.Tab = 0 Then
'        txtMBMAdd1.SetFocus
'    ElseIf SSTab1.Tab = 1 Then
'        cboCovSalutation.SetFocus
'    End If
'End Sub

Private Sub txtCLNCovExclusion_LostFocus()
    txtCLNCovExclusion = ChkStr(txtCLNCovExclusion)
End Sub

Private Sub txtCovAllergic_Click()
    KeyPreview = False
End Sub

Private Sub txtCovAllergic_LostFocus()
    KeyPreview = True
    txtCovAllergic = ChkStr(txtCovAllergic)
End Sub

Private Sub txtCovBirthdate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
Dim TempSuppExpDate As String
Dim tmpDOBStart As String
Dim tmpDOBEnd As String
    
    Dim strsql As String
    If txtCovBirthdate <> "__/__/____" Then
        If IsDate(txtCovBirthdate) Then
           tmpDOBStart = ""
           tmpDOBEnd = ""
           'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
           strsql = "Select ValidID ,MBMDOBStart , MBMDOBEnd  From MBMValidDateRange "
           MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If MBMValidDate.recordCount Then
                tmpDOBStart = MBMValidDate!MBMDOBStart
                tmpDOBEnd = MBMValidDate!MBMDOBEnd
            End If
            MBMValidDate.Close
            Set MBMValidDate = Nothing
            If Trim(CboSuppPln) <> "" And IsDate(txtCovBirthdate) = True And IsDate(txtSuppEffDate) And IsDate(txtSuppExpDate) Then
                TempSuppExpDate = GetSuppExpDate(txtCovBirthdate, CboSuppPln, txtSuppEffDate, txtMBMPayorExpDate)
                txtSuppExpDate = TempSuppExpDate
            End If
            'medixmasterconn.Close
           ' Set medixmasterconn = Nothing
            If OverAge = True Then
                MsgBox "Supplementary age is above " & PlnAgeLmt
            End If
            
            If Len(tmpDOBStart) And Len(tmpDOBEnd) Then
                If DateDiff("d", tmpDOBStart, txtCovBirthdate) < 0 Or DateDiff("d", tmpDOBEnd, txtCovBirthdate) > 0 Then
                    MsgBox "Birthdate is out of acceptable range! " & vbNewLine & _
                            " It must be in between " & tmpDOBStart & " and " & tmpDOBEnd, vbCritical + vbOKOnly
                    txtCovBirthdate.SetFocus
                End If
            End If
        Else
            MsgBox "Invalid date format! ", vbCritical + vbOKCancel
            txtCovBirthdate.SetFocus
        End If
        
    End If
End Sub

Private Sub txtCovExclusion_Click()
    KeyPreview = False
End Sub

Private Sub txtCovExclusion_LostFocus()
    KeyPreview = True
    txtCovExclusion = ChkStr(txtCovExclusion)
End Sub

Private Sub txtCovIcBcPp_LostFocus()
    txtCovIcBcPp = ChkStr(txtCovIcBcPp)
End Sub

Private Sub txtCovName_LostFocus()
    txtCovName = ChkStr(txtCovName)
End Sub

Private Sub txtCOVPayRemarks_LostFocus()
    txtCOVPayRemarks = ChkStr(txtCOVPayRemarks)
End Sub

Private Sub txtMBMAdd1_LostFocus()
    txtMBMAdd1 = ChkStr(txtMBMAdd1)
End Sub

Private Sub txtMBMAdd2_LostFocus()
    txtMBMAdd2 = ChkStr(txtMBMAdd2)
End Sub

Private Sub txtMBMAdd3_LostFocus()
    txtMBMAdd3 = ChkStr(txtMBMAdd3)
End Sub

Private Sub txtMBMAgentCode_LostFocus()
    txtMBMAgentCode = ChkStr(txtMBMAgentCode)
End Sub

Private Sub txtMBMBatchNo_LostFocus()
    Call EnableFrame
End Sub

Private Sub txtMBMBordxDate_LostFocus()
Dim BAT As New ADODB.Recordset
Dim strsql As String
    
    If txtMBMBordxDate = "__/__/____" Then
        Exit Sub
    End If
    If IsDate(txtMBMBordxDate) = False Then
        MsgBox "Invalid Date Format! Please key in 'dd/mm/yyyy' format!", vbCritical
        txtMBMBordxDate.SetFocus
        txtMBMBordxDate.SelStart = 0
    Else
        If Trim(CboPayor) = "" Then
            MsgBox "Please Keyin Payor Code !", vbCritical + vbOKOnly
            CboPayor.SetFocus
        Else
            lblLastBatch = 0
          '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            strsql = "Select  BATSequence , BordxType From BAT " & _
                    " where BATBordxDate = '" & Format(txtMBMBordxDate, "mm/dd/yyyy") & "'"
            strsql = strsql & " AND PAYCode ='" & Trim(Mid(CboPayor, 1, InStr(1, CboPayor, "-") - 1)) & "'"
            BAT.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            Do While Not BAT.EOF
                If BAT!BordxType = "N" Then
                    lblLastBatch = BAT!BATSequence
                End If
                BAT.MoveNext
            Loop
            BAT.Close
            Set BAT = Nothing
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        End If
    End If
    Call EnableFrame
End Sub

Private Sub txtMBMDepartment_LostFocus()
    txtMBMDepartment = ChkStr(txtMBMDepartment)
End Sub

Private Sub txtMBMEmployeeNo_LostFocus()
    txtMBMEmployeeNo = ChkStr(txtMBMEmployeeNo)
End Sub

Private Sub txtMBMGroupCompany_LostFocus()
    txtMBMGroupCompany = ChkStr(txtMBMGroupCompany)
End Sub

Private Sub txtMBMName_LostFocus()
    txtMBMName = ChkStr(txtMBMName)
End Sub

Private Sub txtMBMPolicyNo_LostFocus()
    txtMBMPolicyNo = ChkStr(txtMBMPolicyNo)
End Sub

Private Sub txtMBMPrevMEMNo_GotFocus()
    cmdShow.Default = True
End Sub

Private Sub txtMBMPrevMEMNo_LostFocus()
    txtMBMPrevMEMNo = ChkStr(txtMBMPrevMEMNo)
    cmdShow.Default = False
End Sub

Private Sub txtPrevINSCom_LostFocus()
    txtPrevINSCom = ChkStr(txtPrevINSCom)
End Sub

Private Sub TxtSuppClmNonPayFr_LostFocus()
    If TxtSuppClmNonPayFr <> "__/__/____" Then
        If IsDate(TxtSuppClmNonPayFr) = False Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtSuppClmNonPayFr.SetFocus
        End If
    End If
End Sub

Private Sub TxtSuppClmNonPayTo_LostFocus()
    If TxtSuppClmNonPayTo <> "__/__/____" Then
        If IsDate(TxtSuppClmNonPayTo) = False Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtSuppClmNonPayTo.SetFocus
        End If
    End If
End Sub


Private Sub EnterToTab(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        If KeyAscii = 13 Then   'Turn off the beep
            KeyAscii = 0
        End If
        If tabIndexDone = False Then
            SendKeys "{tab}"
        End If
    End If
    tabIndexDone = False
End Sub

Private Sub txtMBMDateJoin_LostFocus()
    If txtMBMDateJoin <> "__/__/____" Then
        If IsDate(txtMBMDateJoin) = False Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMBMDateJoin.SetFocus
        End If
    End If
End Sub

Private Sub txtMBMExclusion_Click()
    KeyPreview = False
End Sub

Private Sub txtMBMExclusion_LostFocus()
    KeyPreview = True
    txtMBMExclusion = ChkStr(txtMBMExclusion)
End Sub

' ******************************

' ######################### END INITIALIZATION ###########

' ######################### Start Checking  ###########
Private Sub txtMBMIcBcPP_LostFocus()
    Dim CheckRecord
    Dim msgstr As String
    Dim RecordConfirm
    Dim CheckFlag As Boolean
    Dim DupIc As Boolean
    Dim tmpMBMNo As String
    Dim tmpIc As String
    
    txtMBMIcBcPP = ChkStr(txtMBMIcBcPP)

    If Len(Trim(txtMBMIcBcPP)) Then
        If CheckFlag = False Then
            DupIc = False
            tmpIc = txtMBMIcBcPP
            Call CheckIC(tmpIc, DupIc, tmpMBMNo)
            If DupIc = True Then
                msgstr = "Duplicate IC/BC/PP In The Membership" & Chr(13) & "Membership No: " & tmpMBMNo & vbNewLine & _
                        " Do you want to accept it? "
                RecordConfirm = MsgBox(msgstr, vbCritical + vbYesNoCancel, DialogBoxTitle)
                If RecordConfirm = vbNo Then
                    txtMBMIcBcPP.SetFocus
                    txtMBMIcBcPP.SelStart = 0
                    CheckFlag = False
                ElseIf RecordConfirm = vbCancel Then
                    txtMBMIcBcPP = ""
                    txtMBMIcBcPP.SetFocus
                    CheckFlag = False
                Else
                    CheckFlag = True
                    If Mid(cboMBMRenewal, 1, 1) = "R" Then
                        txtMBMPrevMEMNo = tmpMBMNo
                        Call cmdShow_Click
                    Else
                        txtMBMPolicyNo.SetFocus
                    End If
                End If
            End If
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cboMBMSalutation_Click()
  On Error Resume Next
    If ChkStr(cboMBMSalutation) = "MR" Then
        cboMBMSex.ListIndex = 2
    ElseIf ChkStr(cboMBMSalutation) = "" Then
        cboMBMSex.ListIndex = 0
    Else
        cboMBMSex.ListIndex = 1
    End If
End Sub

Private Sub cboMBMSalutation_CHANGE()
    If ChkStr(cboMBMSalutation) = "MR" Then
        cboMBMSex.ListIndex = 2
    ElseIf ChkStr(cboMBMSalutation) = "" Then
        cboMBMSex.ListIndex = 0
    Else
        cboMBMSex.ListIndex = 1
    End If
End Sub


Private Sub cboCovSalutation_Click()
    If ChkStr(cboCovSalutation) = "MR" Then
        cboCovSex = "M"
    ElseIf ChkStr(cboCovSalutation) = "" Then
        cboCovSex = ""
    Else
        cboCovSex = "F"
    End If
End Sub

Private Sub cboCovSalutation_Change()
    If ChkStr(cboCovSalutation) = "MR" Then
        cboCovSex = "M"
    ElseIf ChkStr(cboCovSalutation) = "" Then
        cboCovSex = ""
    Else
        cboCovSex = "F"
    End If
End Sub

Private Sub txtMBMBirthdate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
Dim strsql As String
Dim DateExit
Dim TmpEffDate As String
Dim tmpDOb As String
Dim tmpHltCode As String
Dim tmpEffStartDate As String
Dim tmpEffEndDate As String

    If IsDate(txtMBMBirthdate) = False Then
        MsgBox "Invalid date format ! ", vbCritical + vbOKOnly
        txtMBMBirthdate.SetFocus
    Else
        TmpEffDate = txtMBMPayorEffDate
        tmpDOb = txtMBMBirthdate
        tmpHltCode = Trim(CboHLTCode)
        tmpEffStartDate = ""
        tmpEffEndDate = ""
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        strsql = "Select ValidID ,MBMDOBStart , MBMDOBEnd  From MBMValidDateRange "
        MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMValidDate.recordCount Then
            tmpEffStartDate = IIf(Len(MBMValidDate!MBMDOBStart), MBMValidDate!MBMDOBStart, "")
            tmpEffEndDate = IIf(Len(MBMValidDate!MBMDOBEnd), MBMValidDate!MBMDOBEnd, "")
        End If
        MBMValidDate.Close
        Set MBMValidDate = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        If tmpEffStartDate <> "" And tmpEffEndDate <> "" Then
            If DateDiff("d", tmpEffStartDate, txtMBMBirthdate) >= 0 And DateDiff("d", tmpEffEndDate, txtMBMBirthdate) <= 0 Then
                If Trim(cboPLNCode) <> "" Then
                    If txtMBMPayorEffDate <> "__/__/____" Then
                        TmpEffDate = txtMBMPayorEffDate
                    End If
                ElseIf Trim(frmMBMClinical.cboCisPln) <> "" Then
                    If frmMBMClinical.txtCLNEffDate <> "__/__/____" Then
                        TmpEffDate = frmMBMClinical.txtCLNEffDate
                    End If
                End If
            Else
                DateExit = MsgBox("Uncertain Birthdate, please select OK to continue " & vbNewLine & _
                            " Otherwise press Cancel to make changes! ", vbExclamation + vbOKCancel + vbDefaultButton2)
                If DateExit = vbCancel Then
                    txtMBMBirthdate.SetFocus
                End If
            End If
            If IsDate(TmpEffDate) Then
            '    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                Call CallLimit
                Call MemberCallFees
               ' medixmasterconn.Close
              '  Set medixmasterconn = Nothing
            End If
        End If
        frmMBMClinical.txtMBMBirthdate = txtMBMBirthdate
    End If
End Sub

Private Sub cboMBMRenewal_GotFocus()
    Me.KeyPreview = True
End Sub

Private Sub cboMBMRenewal_Click()
    If ChkStr(cboMBMRenewal) = "R" Then
        cmdSearch.Enabled = True
        cmdShow.Enabled = True
        txtMBMPrevMEMNo.Enabled = True
    Else
        cmdSearch.Enabled = False
        cmdShow.Enabled = False
        txtMBMPrevMEMNo.Enabled = False
        txtMBMPrevMEMNo = ""
        txtMBMPrevPolNo = ""
    End If
End Sub

Private Sub cboMBMRenewal_Change()
    cmdSearch.Enabled = False
    cmdShow.Enabled = False
    If ChkStr(cboMBMRenewal) = "R" Then
        cmdSearch.Enabled = True
        cmdShow.Enabled = True
        txtMBMPrevMEMNo.Enabled = True
    Else
        txtMBMPrevMEMNo = ""
        txtMBMPrevPolNo = ""
        txtMBMPrevMEMNo.Enabled = False
    End If
End Sub

Private Sub CmdSearch_Click()
    Call ClearEntries
    cboMBMRenewal = "R"
    frmSearchMBM.Source = 5
    frmSearchMBM.Show vbModal
End Sub


Private Sub txtMBMPayorEffDate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
Dim strsql As String
Dim TempDate As Date
Dim answer
Dim TmpEffDate As String
Dim tmpDOb As String
Dim TmpValidEffdate As String

    tmpDOb = txtMBMBirthdate

    If (txtMBMPayorEffDate) = "__/__/____" Then
        Exit Sub
    End If
    If IsDate(txtMBMPayorEffDate) = False Then
        answer = MsgBox("Invalid Date Format!", vbCritical + vbOKCancel)
        If answer = vbOK Then
            txtMBMPayorEffDate.SetFocus
        Else
            txtMBMPayorExpDate.SetFocus
        End If
    Else
        If ChkStr(cboMBMRenewal) = "N" Then
            txtMBMDateJoin = txtMBMPayorEffDate
        Else
            If txtMBMDateJoin = "__/__/____" Then
                txtMBMDateJoin = txtMBMPayorEffDate
            End If
        End If
        TempDate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
        txtMBMPayorExpDate = Format(DateAdd("d", -1, TempDate), "dd/mm/yyyy")
        TmpEffDate = txtMBMPayorExpDate

       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        strsql = "Select ValidID , MBMEffStart, MBMEffEnd From MBMValidDateRange "
        MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
        If MBMValidDate.recordCount Then
             If DateDiff("d", MBMValidDate!MBMEffStart, txtMBMPayorEffDate) >= 0 And DateDiff("d", MBMValidDate!MBMEffEnd, txtMBMPayorEffDate) <= 0 Then
                 TempDate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
                 txtMBMPayorExpDate = Format(DateAdd("d", -1, TempDate), "dd/mm/yyyy")
                 'txtMBMEmployeeNo.SetFocus
                 TmpEffDate = txtMBMPayorExpDate
             Else
                 MsgBox "Policy Effective Date is out of acceptable range! " & vbNewLine & _
                         "It must be in between " & MBMValidDate!MBMEffStart & " and " & MBMValidDate!MBMEffEnd & "! ", vbCritical + vbOKOnly
                 'txtMBMPayorEffDate.SetFocus
             End If
        Else
        End If
        If txtMBMBirthdate <> "__/__/____" Then
            Call MemberCallFees
            Call CallLimit
        End If
        MBMValidDate.Close
        Set MBMValidDate = Nothing
    '    medixmasterconn.Close
     '   Set medixmasterconn = Nothing
        txtTempEffDate = txtMBMPayorEffDate
        txtTempExpDate = txtMBMPayorExpDate
        txtSuppEffDate = txtTempEffDate
        txtSuppExpDate = txtTempExpDate
    End If
End Sub

Private Sub txtMBMPayorExpDate_KeyPress(KeyAscii As Integer)
    MBMPayorEXPDate = 0
    MBMPayorEXPDate = 1
End Sub

Private Sub txtMBMPayorExpDate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
Dim strsql As String

    If txtMBMPayorExpDate = "__/__/____" Then
        Exit Sub
    End If
    
    If IsDate(txtMBMPayorExpDate) = False Then
        MsgBox "Invalid Date Format!", vbCritical + vbOKOnly
         txtMBMPayorExpDate.SetFocus
    Else
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        strsql = "Select ValidID , MBMExpStart, MBMExpEnd From MBMValidDateRange "
        MBMValidDate.MaxRecords = 1
        MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                 
        If MBMValidDate.recordCount Then
            If DateDiff("d", MBMValidDate!MBMExpStart, txtMBMPayorExpDate) >= 0 And DateDiff("d", MBMValidDate!MBMExpEnd, txtMBMPayorExpDate) <= 0 Then
            Else
                MsgBox "Policy Expiry Date is out of acceptable range! " & vbNewLine & _
                        " It must be in between " & MBMValidDate!MBMExpStart & " and " & MBMValidDate!MBMExpEnd, vbCritical + vbOKOnly
            End If
        End If
        MBMValidDate.Close
        Set MBMValidDate = Nothing
        Call CallLimit
        Call MemberCallFees
    End If
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub


Private Sub cboCovRelationship_cLICK()
    If Mid(ChkStr(cboCovRelationship), 1, 1) = "W" Or Mid(ChkStr(cboCovRelationship), 1, 1) = "H" Then
        cboCovAdultChild = "A"
    Else
        cboCovAdultChild = "C"
    End If
End Sub

Private Sub cboCovRelationship_CHANGE()
    If Mid(ChkStr(cboCovRelationship), 1, 1) = "W" Or Mid(ChkStr(cboCovRelationship), 1, 1) = "H" Then
        cboCovAdultChild = "A"
    Else
        cboCovAdultChild = "C"
    End If
End Sub


Private Sub MemberCLNCallFees()
Dim chkPLN As New ADODB.Recordset
Dim strsql As String
Dim tmpPlnCode As String
Dim tmpINSType As String
Dim tmpPremINd As String
Dim tmpMCOInd As String
Dim tmpRoundUp As String
Dim tmpProRate As String
Dim XPos As Integer
Dim tmpMCO As String
    
    tmpPlnCode = frmMBMClinical.cboCisPln
    tmpINSType = Mid(frmMBMClinical.cboCisINS, 1, 1)

    If Trim(txtAGECode) <> "" And Trim(tmpPlnCode) <> "" Then
    
        frmMBMClinical.lblMBMPremium = 0
        frmMBMClinical.lblMCOFees = 0
        frmMBMClinical.txtTempPrem = 0
        frmMBMClinical.txtTempMCOFees = 0
        'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

        strsql = " Select PLNCode, MCOInd, PRMInd, PLNRoundUpStatus, PLNProRate " & _
                 " From PLN Where PLNCode = '" & Trim(tmpPlnCode) & "'"
        chkPLN.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic

        tmpPremINd = "N"
        tmpMCOInd = "N"
        tmpRoundUp = "Y"
        tmpProRate = "Y"
        If chkPLN.recordCount Then
            tmpPremINd = IIf(Len(Trim(chkPLN!PRMInd)), Trim(chkPLN!PRMInd), "N")
            tmpMCOInd = IIf(Len(Trim(chkPLN!MCOInd)), Trim(chkPLN!MCOInd), "N")
            tmpRoundUp = IIf(Len(Trim(chkPLN!PLNRoundUpStatus)), Trim(chkPLN!PLNRoundUpStatus), "Y")
            tmpProRate = IIf(Len(Trim(chkPLN!PLNProRate)), Trim(chkPLN!PLNProRate), "Y")
        End If
        chkPLN.Close
        Set chkPLN = Nothing
        
        If tmpPremINd = "Y" Then
            frmMBMClinical.lblMBMPremium = Format(GetCLNPremium(tmpINSType, Trim(txtPAYCode), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), _
            Trim(tmpPlnCode), Trim(CboHLTCode), CDate(frmMBMClinical.txtCLNEffDate)), "#,##0.00")
            frmMBMClinical.txtTempPrem = Format(GetCLNPremium(tmpINSType, Trim(txtPAYCode), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), _
            Trim(tmpPlnCode), Trim(CboHLTCode), CDate(frmMBMClinical.txtCLNEffDate)), "#,##0.00")
        End If
        If tmpMCOInd = "Y" Then
            tmpMCO = Format(GetCLNMCO(tmpINSType, Trim(CboHLTCode), txtMBMAdultChild, CDate(frmMBMClinical.txtCLNEffDate), tmpPlnCode), "#,##0.00")
            If tmpProRate <> "N" Then
                tmpMCO = InsertMCOFees(CDate(frmMBMClinical.txtClnExpDate), CDate(frmMBMClinical.txtCLNEffDate), Format(tmpMCO, "#,##0.00"))
            End If
            tmpMCO = Format(tmpMCO, "#,##0.00")
            If tmpRoundUp <> "N" Then
                tmpMCO = Round(tmpMCO)
            End If
            frmMBMClinical.txtTempMCOFees = tmpMCO
            frmMBMClinical.lblMCOFees = tmpMCO
        End If
        'MediConnCISDB.Close
        'Set MediConnCISDB = Nothing
    End If
End Sub

Private Sub CallLimit()
Dim tmpKidneyLmt As Double
Dim tmpCancerLmt As Double
Dim tmpHomeNurseLmt As Double
Dim tmpSuppBNFStatus As String
    tmpKidneyLmt = 0
    tmpCancerLmt = 0
    tmpHomeNurseLmt = 0
    txtAnnualLimit = ""
    If Len(Trim(cboPLNCode)) And IsDate(txtMBMPayorEffDate) Then
        txtAnnualLimit = Format(GetAnnualLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(CboHLTCode), CDate(txtMBMPayorEffDate)), "#,###,##0.00")
        If TempLifeTimeStatus = "Y" Then
            Call GetBenefitLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(CboHLTCode), CDate(txtMBMPayorEffDate), tmpKidneyLmt, tmpCancerLmt, tmpHomeNurseLmt, tmpSuppBNFStatus)
            lblKidneyLmt = Format(tmpKidneyLmt, "#,##0.00")
            lblCancerLmt = Format(tmpCancerLmt, "#,##0.00")
            lblHomeNurseLmt = Format(tmpHomeNurseLmt, "#,##0.00")
        End If
        Call PrincipalLifeTimeChecking
    End If

End Sub

' ######################### END CALCULATIOn###########

' ######################### BUTTON  CLICKING ###########
Public Sub cmdShow_Click()
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim sqlstrMBMTwo As String
Dim MBMTwo As New ADODB.Recordset
Dim TmpEffDate As String

    Screen.MousePointer = vbHourglass
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    sql = "Exec SearchMBMCovPersonsProc '" & Trim(txtMBMPrevMEMNo) & "'"
    Set ADOrstMBM = New ADODB.Recordset
    ADOrstMBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If Not ADOrstMBM.EOF Then
        With ADOrstMBM
            txtMBMName = IIf(Len(!MBMName), Trim(!MBMName), "")
            txtMBMIcBcPP = IIf(Len(!MBMIcBcPp), Trim(!MBMIcBcPp), "")
            txtMBMPrevPolNo = IIf(Len(!MBMPolicyNo), Trim(!MBMPolicyNo), "")
            cboINSCode = IIf(Len(!INSCode), Trim(!INSCode), "")
            If Trim(CboHLTCode) <> Trim(!HLTCode) Then
                cboPLNCode.Text = ""
            Else
                cboPLNCode = IIf(Len(!PLNCode), Trim(!PLNCode), "")
            End If
            txtMBMAdd1 = IIf(Len(!MBMAddress1), Trim(!MBMAddress1), "")
            txtMBMAdd2 = IIf(Len(!MBMAddress2), Trim(!MBMAddress2), "")
            txtMBMAdd3 = IIf(Len(!MBMAddress3), Trim(!MBMAddress3), "")
            txtMBMPostCode = IIf(Len(!MBMPostCode), Trim(!MBMPostCode), "")
            cboCTYCode = IIf(Len(!CTYCode), Trim(!CTYCode), "")
            cboMBMSex = IIf(Len(!MBMSex), Trim(!MBMSex), "")
            cboBRNCode = IIf(Len(!MBMBranch), Trim(!MBMBranch), "")
            If Len(Trim(!MBMDOB)) Then
                txtMBMBirthdate.mask = ""
                txtMBMBirthdate = Trim(!MBMDOB)
                txtMBMBirthdate.mask = "##/##/####"
            End If
            cboRACCode = IIf(Len(!RACCode), Trim(!RACCode), "")
            cboSTACode = IIf(Len(!STACode), Trim(!STACode), "")
            If Len(Trim(!MBMTakeover)) Then
                cboMBMTakeover.Enabled = True
                cboMBMTakeover = Trim(!MBMTakeover)
            End If
            
            If !MBMPolicyYear <> "" Then
                txtMBMPolicyYear = IIf(Trim(!MBMPolicyYear) = "", 1, !MBMPolicyYear + 1)
            Else
                txtMBMPolicyYear = 1
            End If
            If Trim(!HSCLNInd) = "C" Or Trim(!HSCLNInd) = "M" Or Trim(!HSCLNInd) = "B" Then
                Call showClinicalInfo
            End If
        End With
        Call PLNLostFocus
        Call showMBMTwo
        Call showMBMOthers
        Call showMBMOthersTwo
        If Trim(Mid(cboINSCode, 1, 1)) = "F" Or Trim(Mid(cboINSCode, 1, 1)) = "H" Then
            If Len(Trim(ADOrstMBM!MBMCCoverID)) And Trim(ADOrstMBM!MBMCCoverID) <> "00" Then
                Call showMBMCov
            End If
        End If
        Call CallLimit
        Call PLNLostFocus
        Call MemberCallFees
    Else
        MsgBox "No record found!", vbCritical + vbOKOnly
    End If
    ADOrstMBM.Close
    Set ADOrstMBM = Nothing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    Screen.MousePointer = Default
End Sub

Private Sub txtCovBirthdate_gotfocus()
    txtCovBirthdate.SelStart = 0
End Sub

Private Sub txtMBMBirthdate_gotfocus()
    txtMBMBirthdate.SelStart = 0
End Sub

Private Sub txtMBMPayorEffDate_gotfocus()
    txtMBMPayorEffDate.SelStart = 0
End Sub

Private Sub cmdClear_Click()
On Error Resume Next
Dim RecordSaved
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Call ClearEntries
            Call EnableEntries(True)
            cboMBMRenewal.SetFocus
            cboMBMRenewal = "N"
            txtMBMPolicyYear = 1
            
        End If
End Sub

Private Sub cmdExit_Click()
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    If frmMBMClinical.Visible = True Then
        Unload frmMBMClinical
    End If
    Call ClearForm(frmNewMBMOthers)
    Unload FrmTempMBMReg
    Set FrmTempMBMReg = Nothing

End Sub


Private Sub CmdSave_Click()
Dim strsqlPLN As String
Dim PLNRec As New ADODB.Recordset
Dim PLNFound As Boolean
Dim RecordSaved

    PLNFound = False
    If Trim(cboPLNCode.Text) <> "" Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        strsqlPLN = "select PLNCode " & _
            " from PLN where PLNCode = '" & Trim(cboPLNCode) & "' and HLTCode = '" & Trim(CboHLTCode) & "'"
        PLNRec.Open strsqlPLN, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        PLNFound = False
        Do While Not PLNRec.EOF
            PLNFound = True
            PLNRec.MoveNext
        Loop
        PLNRec.Close
        Set PLNRec = Nothing
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
    End If
    If cboMBMRenewal = "R" And _
       Len(Trim(txtMBMPrevMEMNo)) = 0 And _
       Len(Trim(txtMBMPrevPolNo)) = 0 Then
        MsgBox "Please Enter Previous Membership No or Previous Policy No", vbOKOnly + vbCritical, DialogBoxTitle
        cmdShow.SetFocus
    Else
        If Len(Trim(txtMBMName)) = 0 Then
            MsgBox "Please Enter Membership Name", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMName.SetFocus
        ElseIf Len(Trim(txtMBMIcBcPP)) = 0 Then
            MsgBox "Please Enter IC/BC/PP", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMIcBcPP.SetFocus
        ElseIf Len(Trim(CboHLTCode)) = 0 Then
            MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
            CboHLTCode.SetFocus
        ElseIf Len(Trim(txtPAYCode)) = 0 Then
            MsgBox "Please Enter Payor", vbOKOnly + vbCritical, DialogBoxTitle
            CboPayor.SetFocus
        ElseIf Len(Trim(txtMBMAdd1)) = 0 Then
            MsgBox "Please Enter Address", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMAdd1.SetFocus
        ElseIf Len(Trim(txtMBMPostCode)) = 0 Then
            MsgBox "Please Enter Post Code", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPostCode.SetFocus
        ElseIf Len(Trim(cboSTACode)) = 0 Then
            MsgBox "Please Enter State", vbOKOnly + vbCritical, DialogBoxTitle
            cboSTACode.SetFocus
        ElseIf txtMBMBirthdate = "__/__/____" Then
            MsgBox "Please Enter Birthdate", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMBirthdate.SetFocus
        ElseIf IsDate(txtMBMBirthdate) = False Then
            MsgBox "Invalid BirthDate", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMBirthdate = "__/__/____"
            txtMBMBirthdate.SetFocus
        ElseIf Trim(cboPLNCode.Text) = "" And Trim(frmMBMClinical.cboCisPln) = "" Then
            MsgBox "Please Enter Plan Code or Clinical Plan", vbOKOnly + vbCritical, DialogBoxTitle
            cboPLNCode.SetFocus
        ElseIf Len(Trim(txtAGECode)) = 0 Then
            MsgBox "Invalid Age Group", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf Trim(TempLifeTimeStatus) = "Y" And Trim(CboGuaRenewal) = "" Then
            MsgBox "Please Enter Guarantee Renewal for 1st year Policy of Life Time ", vbOKOnly + vbCritical, DialogBoxTitle
' ------------ HIS Verification
        ElseIf Trim(cboPLNCode.Text) <> "" And Len(Trim(txtMBMPolicyNo)) = 0 Then
            MsgBox "Please Enter Policy No", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPolicyNo.SetFocus
        ElseIf Trim(cboPLNCode.Text) <> "" And IsDate(txtMBMPayorEffDate) = False Then
            MsgBox "Please Enter Payor Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPayorEffDate.SetFocus
        ElseIf Trim(cboPLNCode.Text) <> "" And IsDate(txtMBMPayorExpDate) = False Then
            MsgBox "Please Enter Payor Expiry Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPayorExpDate.SetFocus
        ElseIf Trim(cboPLNCode.Text) <> "" And Trim(cboINSCode.Text) = "" Then
            MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
            cboINSCode.SetFocus
        ElseIf Trim(cboINSCode.Text) = "" And Trim(frmMBMClinical.cboCisINS.Text) = "" Then
            MsgBox "Please Enter Insured Code OR Clinical Insured Code ", vbOKOnly + vbCritical, DialogBoxTitle
            cboINSCode.SetFocus
        ElseIf PLNFound = False Then
            MsgBox "No Such Plan Code! ", vbOKOnly + vbCritical, DialogBoxTitle
            cboPLNCode.SetFocus
        Else
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
              If RecordSaved = vbYes Then
                  Screen.MousePointer = vbHourglass
                   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

                    Call NewRegistrationSub
                    
                   ' medixmasterconn.Close
                   ' Set medixmasterconn = Nothing
                   ' medixmasterconnMaint.Close
                   ' Set medixmasterconnMaint = Nothing
                  txtMBMPolicyYear = 1
                  Screen.MousePointer = vbDefault
              End If
        End If
    End If
    

End Sub

' ######### To save the membership data ---> cater for new registration ############3
Private Sub NewRegistrationSub()
On Error GoTo ErrorSave
Dim StrSqlLife As String
Dim MBM As New ADODB.Recordset
Dim MBM1 As New ADODB.Recordset
Dim MBMTwo As New ADODB.Recordset
Dim MBMCRef As New ADODB.Recordset
Dim MBMLifeRec As New ADODB.Recordset
Dim CLNPln As New ADODB.Recordset
Dim AnnLmt As New ADODB.Recordset

Dim strsqlMBM As String
Dim strsqlMBMTwo As String
Dim strsqlMBMCRef As String
Dim CLNPlnSql As String
Dim StrSqlAnnLmt As String

Dim MBMAvailLmt As String
Dim TempDOB As String
Dim TempBordxdate As String
Dim TempValueDate As String
Dim TempMBMAdd2 As String
Dim TempMBMName As String
Dim SuppLmtStatus As String
Dim HisCisEffDate, HISCISExpDate As String
Dim HISCisPLN, HISCiSIns As String
Dim TmpMemberType, TmpHSClnInd, TmpDCStatus As String
Dim tmpKidneyAnnLmt As Double
Dim tmpCancerAnnLmt As Double
Dim tmpHomeNursAnnLmt As Double
Dim tmpSuppBNFStatus  As String
Dim tmpMBMStatus As String
    
    strsqlPLN = "select * from DiscPlan where PLNCode = '" & Trim(Mid(cboPLNCode, 1, 4)) & "'"
    PLN.Open strsqlPLN, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    If PLN.recordCount Then
        Disc = True
    End If
    
    PLN.Close
    Set PLN = Nothing
    
    ' INsert into Membership table
    ' Begin transaction
    MBMAvailLmt = CDbl(IIf(IsNumeric(txtAnnualLimit) = True, txtAnnualLimit, 0))
    TempDOB = Format(txtMBMBirthdate, "YYYY/MM/DD")
    TempBordxdate = Format(txtMBMBordxDate, "YYYY/MM/DD")
    TempValueDate = Format(Date, "YYYY/MM/DD")
    TempPayEffDate = IIf(IsDate(txtMBMPayorEffDate), Format(txtMBMPayorEffDate, "YYYY/MM/DD"), "")
    TempPayExpDate = IIf(IsDate(txtMBMPayorExpDate), Format(txtMBMPayorExpDate, "YYYY/MM/DD"), "")
    TempMBMAdd2 = Replace(txtMBMAdd2, "'", "''")
    TempMBMName = Replace(txtMBMName, "'", "''")
    TmpDCStatus = ""
    CISDCStatus = ""
    CISDCPayor = Mid(txtPAYCode, 1, 2)
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

    If Trim(frmMBMClinical.cboCisPln) <> "" Then

        CLNPlnSql = "SELECT ANNLmtInd, DCStatus, PAYCode, PDentalStatus, SDentalStatus," & _
                    "PSpecStatus, SSpecStatus, PMepStatus, SMepStatus " & _
                    " from PLN where PLNCode='" & Mid(frmMBMClinical.cboCisPln, 1, 4) & _
                    "' AND HLTCode='" & Trim(CboHLTCode) & "'"
        CLNPln.Open CLNPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        If CLNPln.recordCount Then
            TmpDCStatus = IIf(Len(Trim(CLNPln!DCStatus)), Trim(CLNPln!DCStatus), "")
            CISDCStatus = IIf(Len(Trim(CLNPln!DCStatus)), Trim(CLNPln!DCStatus), "")
            CISDCPayor = IIf(Len(Trim(CLNPln!PAYCode)), Trim(CLNPln!PAYCode), "")
            CISPDentStatus = IIf(Len(Trim(CLNPln!PDentalStatus)), Trim(CLNPln!PDentalStatus), "")
            CISSDentStatus = IIf(Len(Trim(CLNPln!SDentalStatus)), Trim(CLNPln!SDentalStatus), "")
            CISPSpecStatus = IIf(Len(Trim(CLNPln!PSpecStatus)), Trim(CLNPln!PSpecStatus), "")
            CISSSpecStatus = IIf(Len(Trim(CLNPln!SSpecStatus)), Trim(CLNPln!SSpecStatus), "")
            CISPMepStatus = IIf(Len(Trim(CLNPln!PMepStatus)), Trim(CLNPln!PMepStatus), "")
            CISSMepStatus = IIf(Len(Trim(CLNPln!SMepStatus)), Trim(CLNPln!SMepStatus), "")

        End If
        CLNPln.Close
        Set CLNPln = Nothing
    End If
    
    TmpHSClnInd = "H"
    If Trim(cboPLNCode.Text) <> "" And Trim(frmMBMClinical.cboCisPln) = "" Then
        TmpHSClnInd = "H"
    End If
    If Trim(cboPLNCode.Text) = "" And Trim(frmMBMClinical.cboCisPln) <> "" Then
        TmpHSClnInd = "C"
    End If
    If Trim(cboPLNCode.Text) <> "" And Trim(frmMBMClinical.cboCisPln) <> "" Then
        If Trim(TmpDCStatus) = "Y" Then
            TmpHSClnInd = "M"
        Else
            TmpHSClnInd = "B"
        End If
    End If
    CISDCStatus = TmpDCStatus
    HISCisPLN = Trim(cboPLNCode.Text)
    HISCiSIns = Trim(cboINSCode)
    HisCisEffDate = txtMBMPayorEffDate
    HISCISExpDate = txtMBMPayorExpDate
        
    If Trim(cboPLNCode.Text) = "" Then
        HisCisEffDate = frmMBMClinical.txtCLNEffDate
        HISCISExpDate = frmMBMClinical.txtClnExpDate
        HISCisPLN = Trim(frmMBMClinical.cboCisPln)
        HISCiSIns = Trim(frmMBMClinical.cboCisINS)
    End If

    SuppLmtStatus = "A"
    TmpMemberType = "P"
    If Trim(cboPLNCode.Text) <> "" Then
        StrSqlAnnLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(HISCisPLN) & _
                        "' AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(HISCiSIns) & "'"
        AnnLmt.Open StrSqlAnnLmt, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        StrSqlAnnLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(HISCisPLN) & _
                       "' AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(HISCiSIns) & "'"
        AnnLmt.Open StrSqlAnnLmt, MediConnCISDB, adOpenKeyset, adLockOptimistic
    End If
    If AnnLmt.recordCount > 0 Then
        If Trim(AnnLmt!SuppLimitStatus) <> "" Then
            SuppLmtStatus = Trim(AnnLmt!SuppLimitStatus)
        End If
    End If
    AnnLmt.Close
    Set AnnLmt = Nothing
    If Trim(SuppLmtStatus) <> "A" Then
        TmpMemberType = "Q"
    End If
  '  MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
    startTrans = True
    medixmasterconn.beginTrans
    medixmasterconnMaint.beginTrans
    If Trim(cboPLNCode) = "" And Trim(frmMBMClinical.cboCisPln) <> "" Then
        'MembershipNo = Trim(txtPAYCode) & Trim(Mid(frmMBMClinical.cboCisINS, 1, 1)) & GenerateMembershipNo("DC", Trim(Mid(frmMBMClinical.cboCisPln, 1, IIf(InStr(1, frmMBMClinical.cboCisPln, "-") = 0, 4, InStr(1, frmMBMClinical.cboCisPln, "-") - 1))), Trim(Mid(frmMBMClinical.cboCisINS, 1, 1)))
        MembershipNo = Trim(txtPAYCode) & Trim(Mid(frmMBMClinical.cboCisINS, 1, 1)) & GenerateMembershipNo("DC", Trim(Mid(frmMBMClinical.cboCisINS, 1, 1)))
        TmpChkDigits = GenerateChkDigits
        MembershipNo = MembershipNo & TmpChkDigits
    ElseIf Trim(cboPLNCode) <> "" And Trim(frmMBMClinical.cboCisPln) <> "" Then
        MembershipNo = Trim(txtPAYCode) & GenerateMBMNo05(Trim(txtPAYCode), Trim(cboINSCode), CDate(txtMBMBirthdate), Trim(txtMBMName), Trim(cboMBMRenewal))
        'TmpChkDigits = GenerateChkDigits
        'MembershipNo = MembershipNo & TmpChkDigits
    Else
        MembershipNo = Trim(txtPAYCode) & GenerateMBMNo05(Trim(txtPAYCode), Trim(cboINSCode), CDate(txtMBMBirthdate), Trim(txtMBMName), Trim(cboMBMRenewal))
        'TmpChkDigits = GenerateChkDigits
        'MembershipNo = Mid(MembershipNo, 1, 9) & TmpChkDigits
    End If
    
  '  If Trim(cboPLNCode) = "" And Trim(frmMBMClinical.cboCisPln) <> "" Then
   '     MembershipNo = Trim(txtPAYCode) & Trim(Mid(frmMBMClinical.cboCisPln, 1, IIf(InStr(1, frmMBMClinical.cboCisPln, "-") = 0, 4, InStr(1, frmMBMClinical.cboCisPln, "-") - 1))) & Trim(Mid(frmMBMClinical.cboCisINS, 1, 1)) & GenerateMembershipNo("DC", Trim(Mid(frmMBMClinical.cboCisPln, 1, IIf(InStr(1, frmMBMClinical.cboCisPln, "-") = 0, 4, InStr(1, frmMBMClinical.cboCisPln, "-") - 1))), Trim(Mid(frmMBMClinical.cboCisINS, 1, 1)))
   ' Else
   '     MembershipNo = Trim(txtPAYCode) & GenerateMBMNo05(Trim(txtPAYCode), Trim(cboINSCode), CDate(txtMBMBirthdate), Trim(txtMBMName), Trim(cboMBMRenewal))
   ' End If
    MembershipNo = ChkStr(MembershipNo)
    TmpTopupNo = ""
    WLifeStatus = False
    PrevLifeMBMNo = ""
 
    tmpKidneyAnnLmt = 0
    tmpCancerAnnLmt = 0
    tmpHomeNursAnnLmt = 0
    If Trim(TempLifeTimeStatus) = "Y" Then
        Call GetBenefitLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(CboHLTCode), CDate(txtMBMPayorEffDate), tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt, tmpSuppBNFStatus)
        Call SaveMBMBnfAnnLmt(MembershipNo, "00", "A", tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
    End If
    Call PrincipalLifeTimeChecking 'LifeTimeLimit & BenefitLimit
    If WLifeStatus = True Then
        Call SaveMBMLifeTimeLmt(TmpTopupNo, "00", "A")
    End If
    tmpMBMStatus = IIf(Trim(cboPLNCode) <> "", "A", "")
    strsqlMBM = "EXEC Upload_MBM_HIS '" & Mid(txtPAYCode, 1, 2) & "', '" & Mid(CboHLTCode, 1, 1) & "', '" & MembershipNo & "', '" & txtMBMPolicyNo & "', '00', " & _
            "'" & Trim(txtMBMIcBcPP) & "', '" & TempMBMName & "', '" & TempDOB & "','" & cboMBMSex & "','" & Mid(cboRACCode, 1, 1) & "', " & _
            "'" & tmpMBMStatus & "','A','" & TmpMemberType & "', " & _
            "'" & Trim(txtMBMAdd1) & "', '" & Trim(txtMBMAdd2) & "', '" & Trim(txtMBMAdd3) & "', " & _
            "'" & cboCTYCode & "', '" & txtMBMPostCode & "', '" & Trim(cboSTACode) & "', " & _
            "'N', '" & TempValueDate & "', '" & TempPayEffDate & "', '" & TempPayExpDate & "', " & _
            "'" & Mid(cboMBMTakeover, 1, 1) & "', '" & cboMBMRenewal & "', '" & TempBordxdate & "', '" & Trim(txtMBMBatchNo) & "', " & _
            "'" & Trim(cboPLNCode) & "', '" & Trim(Mid(cboINSCode, 1, 1)) & "', '" & Trim(Mid(txtAGECode, 1, 2)) & "', '" & txtMBMAdultChild & "', " & MBMAvailLmt & ", " & _
            "'" & TmpHSClnInd & "', '" & Trim(TmpTopupNo) & "', 'N'"
    
    medixmasterconn.Execute strsqlMBM

    completeMBM = True
    strsqlMBMTwo = "insert into MBMTwo (MBMNumber, MBMSalutation, MBMTelnoH, " & _
                "MBMTelNoM, MBMTelNoO, MBMEmail, " & _
                "NATCode, SRVCodeList, LGNCode, MBMRegUser, MBMLastUpdateUser, MBMUploadStatus) " & _
                "Values('" & MembershipNo & "', '" & cboMBMSalutation & "' , '" & frmNewMBMOthers.txtMBMTelNoH & "', " & _
                "'" & frmNewMBMOthers.txtMBMTelNoM & "', '" & frmNewMBMOthers.txtMBMTelNoO & "', '" & frmNewMBMOthers.txtMBMEmail & "', " & _
                "'" & Trim(Mid(cboNATCode, 1, 2)) & "', 'ME', '" & Trim(Mid(frmNewMBMOthers.cboLGNCode, 1, 1)) & "', '" & Logon & "', '" & Logon & "','M')"
    medixmasterconn.Execute strsqlMBMTwo

    completeMBMTwo = True
    If Trim(cboPLNCode.Text) <> "" Then
        strsqlMBMCRef = "insert into MBMCrossReference (MBMNumber, MBMCOVID, MBMPolicyNo, " & _
                    "MBMIcBcPp, MBMName, HLTCode, PAYCode, MBMDOB," & _
                    "MBMStatus, MBMDataStatus,MBMBordxDate,AgeCode," & _
                    "MBMPayorEffDate, MBMPayorEXPDate, MBMBatchNo, " & _
                    "INSCode, PLNCode, MBMTopupNumber) " & _
                    "Values('" & MembershipNo & "', '00', '" & txtMBMPolicyNo & "', " & _
                    "'" & txtMBMIcBcPP & "', '" & TempMBMName & "', '" & Trim(CboHLTCode) & "', '" & txtPAYCode & "','" & TempDOB & "', " & _
                    "'" & tmpMBMStatus & "','A', '" & TempBordxdate & "', '" & Trim(Mid(txtAGECode, 1, 2)) & "', " & _
                    "'" & TempPayEffDate & "', '" & TempPayExpDate & "', '" & txtMBMBatchNo & "', " & _
                    "'" & Trim(Mid(cboINSCode, 1, 1)) & "', '" & Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") > 0, InStr(1, cboPLNCode, "-") - 1, 4))) & "'," & _
                    "'" & Trim(TmpTopupNo) & "')"
                                
    Else
        strsqlMBMCRef = "insert into MBMCrossReference (MBMNumber, MBMCOVID, MBMPolicyNo, MBMBatchNo," & _
                    "MBMIcBcPp, MBMName, HLTCode, PAYCode, MBMDOB," & _
                    "MBMStatus, MBMDataStatus,MBMBordxDate,AgeCode, MBMTopupNumber)" & _
                    "Values('" & MembershipNo & "', '00', '" & txtMBMPolicyNo & "', '" & txtMBMBatchNo & "', " & _
                    "'" & txtMBMIcBcPP & "', '" & TempMBMName & "', '" & Trim(CboHLTCode) & "', '" & txtPAYCode & "','" & TempDOB & "', " & _
                    "'" & tmpMBMStatus & "','A', '" & TempBordxdate & "', '" & Trim(Mid(txtAGECode, 1, 2)) & "'," & _
                    "'" & Trim(TmpTopupNo) & "')"
    End If
    medixmasterconnMaint.Execute strsqlMBMCRef
    completeMBMCRef = True
    
    Call SaveMBMOthers(MembershipNo)
    
    Call SaveMBMOthersTwo(MembershipNo)
    
    If cboINSCode = "F" Or cboINSCode = "H" Or Mid(frmMBMClinical.cboCisINS, 1, 1) = "F" _
        Or Mid(frmMBMClinical.cboCisINS, 1, 1) = "H" Then
            Call saveCoverPersons(MembershipNo)
        End If
    If Trim(frmMBMClinical.cboCisPln) <> "" Then
        Call SaveMBMCLNOthers(MembershipNo)
    End If
    
    '========Clinical=======================
    
    Call SaveBatch
    
    medixmasterconn.commitTrans
    medixmasterconnMaint.commitTrans
    startTrans = False
     
     cboMBMRenewal.Enabled = True
     MsgBox "Record Saved, Your Membership Number is : " & MembershipNo, vbOKOnly + vbInformation, DialogBoxTitle
     ' Clering previous record
     txtMBMPayorExpDate.Enabled = True
     Call ClearEntries
   
     SSTab1.Tab = 0
     cboMBMRenewal.SetFocus
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnMaint.RollbackTrans
    End If
    MsgBox Err.Description
End Sub

Private Sub SearchSuppLmt()
Dim StrSqlAnnLmt As String
Dim AnnLmt As New ADODB.Recordset
Dim strsqlPrem As String
Dim Prem As New ADODB.Recordset
    If txtMBMPayorEffDate <> "__/__/____" Then
    
    StrSqlAnnLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(cboPLNCode) & "'"
    AnnLmt.Open StrSqlAnnLmt, medixmasterconn, adOpenKeyset, adLockOptimistic

    strsqlPrem = "Select SuppPrmStatus, PLNCode, SuppPrmAmt from PRM where PLNCode = '" & Trim(cboPLNCode) & "' And PAYCode = '" & Trim(txtPAYCode) & "' And INSCode = '" & Trim(cboINSCode) & "' And HLTCode = '" & Trim(CboHLTCode) & "'"
    Prem.Open strsqlPrem, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'Cater for SuppLimitStatus is "C"
    If AnnLmt.recordCount Then
        If AnnLmt!SuppLimitStatus <> "B" Then
            If AnnLmt!SuppLimitStatus = "C" Then
                tempSAnnLmt = Format(GetSuppLimit(cboINSCode, cboPLNCode, Trim(CboHLTCode), CDate(txtMBMPayorEffDate)), "#,##0.00")
                txtCovAnnualLimit = tempSAnnLmt
                tempSPremium = Format(IIf(IsNull(Prem!SuppPrmAmt), 0, Prem!SuppPrmAmt), "#,##0.00")
            Else
                tempSAnnLmt = txtAnnualLimit
                txtCovAnnualLimit = txtAnnualLimit
            End If
        End If
    End If
    AnnLmt.Close
    Prem.Close
    Set AnnLmt = Nothing
    Set Prem = Nothing

    End If

End Sub

Private Sub saveCoverPersons(MembershipNo As String)   ' , MBMtype As String)
Dim MBMCoveredPersons As New ADODB.Recordset
Dim MBMCoveredPersonsTwo As New ADODB.Recordset
Dim listrecord As Integer
Dim listloop As Integer
Dim StrSqlAnnLmt As String
Dim strsqlPrem As String
Dim strsqlMBMCov As String
Dim strsqlMBMCovTwo As String
Dim Prem As New ADODB.Recordset
Dim AnnLmt As New ADODB.Recordset
Dim MemberDOB As String
Dim SelectAge As String
Dim AgeValue As String
Dim tempSPremium As String
Dim tempDPremium As String
Dim tempSAnnLmt As String
Dim CoverID As String
Dim CovMBMType As String
Dim MBMCovAnnLmt As Double
Dim MBMCovAvaLmt As Double
Dim TmpMBMCRBAmt As Double
Dim TmpMBMCPayPremGross As Double
Dim TmpMBMCPayPremNet As Double
Dim TmpMBMCPayMCOGross As Double
Dim TmpMBMCPayMCONet As Double
Dim TmpMBMCFamDiscAmt As Double
Dim TmpMBMCRenewDiscAmt As Double
Dim TmpMBMCLoadingPrem As Double
Dim tempMBMCPolicyDisc As Double
Dim tempMBMCRenewalDisc  As Double
Dim TempMBMCDOB As String
Dim TempMBMCBordxDate As String
Dim TempMBMCExlcusion As String
Dim TempMBMCName As String
Dim SuppPlanCode As String
Dim SptLmtInd As String
Dim TmpPlnRec As New ADODB.Recordset
Dim TmpPlnSql As String
Dim tmpPlnCode As String
Dim tmpPLNLifeTimeStatus, TmpSuppPrmStatus As String
Dim TmpSuppLmt, TmpSuppPrmAmt As String
Dim TmpSql2 As String
Dim tmpMBM As New ADODB.Recordset
Dim strSptLmt As String
Dim SptPlnLmt As New ADODB.Recordset
Dim strSptMCOSql As String
Dim SptMCOInd As String
Dim McoCovIDChrg As String
Dim SptPlnMCO As New ADODB.Recordset
Dim tmpUndExcess As String
Dim TmpBasicMCO, tmpMCO As String
Dim tmpINSType As String
Dim tmpAdultChild As String
Dim tmpKidneyAnnLmt As Double
Dim tmpCancerAnnLmt As Double
Dim tmpHomeNursAnnLmt As Double
Dim tmpSuppBNFStatus  As String
Dim tmpPLNRoundUpStatus As String
Dim tmpProRate As String
Dim tmpCSex, TmpCAdultChild, tmpCRelCode, tmpCIC As String
Dim tmpCRegDate As String
Dim tmpCEffDate As String
Dim TmpCExpDate As String
Dim tmpCSal, tmpCCovID, tmpCAllergies, tmpCOccup, tmpCWN As String
Dim tmpCSmking, tmpCAlcohol, tmpCPrevPln, tmpCPayREm, tmpCPrevINsCom As String
Dim tmpCPayMBMNo, tmpCPaySeqNo, tmpCClientNo, tmpCClientID As String
Dim tmpCNonPayFr As String
Dim tmpCNonPayTo As String
Dim tmpC2IC As String
    listrecord = lstCovAdd.ListItems.Count
    ' Insert into Membership Covered Person's table

    tmpPlnCode = cboPLNCode
    tmpINSType = cboINSCode
    If Trim(tmpPlnCode) = "" Then
        tmpPlnCode = frmMBMClinical.cboCisPln
        tmpINSType = Mid(frmMBMClinical.cboCisINS, 1, 1)
    End If

     If listrecord <> 0 Then
         For listloop = 1 To listrecord
            tmpCEffDate = IIf(IsDate(lstCovAdd.ListItems(listloop).SubItems(16)), lstCovAdd.ListItems(listloop).SubItems(16), txtMBMPayorEffDate)
            TmpCExpDate = IIf(IsDate(lstCovAdd.ListItems(listloop).SubItems(17)), lstCovAdd.ListItems(listloop).SubItems(17), txtMBMPayorExpDate)

            If Trim(cboPLNCode) <> "" Then
                SelectAge = GetAge(CDate(tmpCEffDate), lstCovAdd.ListItems(listloop).SubItems(3))
                AgeValue = Mid(GetAgeBand(Trim(CboHLTCode), SelectAge, tmpPlnCode, tmpINSType), 1, 2)
            ElseIf Trim(frmMBMClinical.cboCisPln) <> "" Then
                SelectAge = GetAge(frmMBMClinical.txtCLNEffDate, lstCovAdd.ListItems(listloop).SubItems(3))
                AgeValue = Mid(GetAgeBand(Trim(CboHLTCode), SelectAge, tmpPlnCode, tmpINSType), 1, 2)
            End If
            tmpAdultChild = "A"
            If AgeValue = "01" Then
                tmpAdultChild = "C"
            End If

            CoverID = Format(lstCovAdd.ListItems(listloop).SubItems(7), "00")
            If Trim(CoverID) = "" Then
                CoverID = Format(listloop, "00")
            End If
            CovMBMType = "C"
            SptLmtInd = "A"
            tmpPLNLifeTimeStatus = "N"
            TempMBMCName = Replace(lstCovAdd.ListItems(listloop).Text, "'", "''")
            If Trim(cboPLNCode) <> "" Then
                tmpPlnCode = Trim(cboPLNCode)
                If Len(Trim(lstCovAdd.ListItems(listloop).SubItems(41))) Then
                    tmpPlnCode = Trim(lstCovAdd.ListItems(listloop).SubItems(41))
                End If
                If Trim(CboHLTCode) = "S" Then
                    TmpPlnSql = "Select PLNLifeTimeStatus, PLNRoundUpStatus, PLNProRate from PLN where PLNCode='" & Trim(tmpPlnCode) & _
                                "' AND HLTCode='" & Trim(CboHLTCode) & _
                                "' AND PLNEffDate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by PLNEffDate desc"
                Else
                    TmpPlnSql = "Select PLNLifeTimeStatus, PLNRoundUpStatus, PLNProRate from PLN where PLNCode='" & Trim(tmpPlnCode) & _
                                "' AND HLTCode='" & Trim(CboHLTCode) & "' AND PAYCode='" & Trim(txtPAYCode) & _
                                "' AND PLNEffDate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by PLNEffDate desc"
                End If
                tmpPLNLifeTimeStatus = "N"
                tmpPLNRoundUpStatus = "Y"
                tmpProRate = "Y"
                TmpPlnRec.Open TmpPlnSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                If TmpPlnRec.recordCount Then
                    tmpPLNLifeTimeStatus = IIf(Len(Trim(TmpPlnRec!PLNLifeTimeStatus)), Trim(TmpPlnRec!PLNLifeTimeStatus), "N")
                    tmpPLNRoundUpStatus = IIf(Len(Trim(TmpPlnRec!PLNRoundUpStatus)), Trim(TmpPlnRec!PLNRoundUpStatus), "Y")
                    tmpProRate = IIf(Len(Trim(TmpPlnRec!PLNProRate)), Trim(TmpPlnRec!PLNProRate), "Y")
                End If
                TmpPlnRec.Close
                Set TmpPlnRec = Nothing
                
                SptLmtInd = "A"
                strSptLmt = "Select SuppLimitStatus, SuppLimit from AnnualLimit where PLNCode = '" & Trim(tmpPlnCode) & _
                    "' AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                    "' AND AnnualEffDate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate desc"
                SptPlnLmt.Open strSptLmt, medixmasterconn, adOpenKeyset, adLockOptimistic
                TmpSuppLmt = 0
                If SptPlnLmt.recordCount > 0 Then
                    TmpSuppLmt = IIf(IsNumeric(SptPlnLmt!SuppLimit), SptPlnLmt!SuppLimit, 0)
                    If Trim(SptPlnLmt!SuppLimitStatus) <> "" Then
                        SptLmtInd = Trim(SptPlnLmt!SuppLimitStatus)
                    End If
                End If
                SptPlnLmt.Close
                Set SptPlnLmt = Nothing
                If Trim(SptLmtInd) = "B" And Trim(CoverID) = "01" Or Trim(SptLmtInd) = "C" Then
                    MBMCovAnnLmt = CDbl(TmpSuppLmt)
                    MBMCovAvaLmt = CDbl(TmpSuppLmt)
                End If  ' AnnLmimit Calculation
'------------   Life Time Limit
                If Trim(CboGuaRenewal) <> "" Then
                    If Trim(TempLifeTimeStatus) = "Y" Then
                        tmpKidneyAnnLmt = 0
                        tmpCancerAnnLmt = 0
                        tmpHomeNursAnnLmt = 0
                        Call GetBenefitLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(CboHLTCode), CDate(txtMBMPayorEffDate), tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt, tmpSuppBNFStatus)
                        If (tmpSuppBNFStatus = "B" And CoverID = "01") Or tmpSuppBNFStatus = "C" Then
                            Call SaveMBMBnfAnnLmt(MembershipNo, CoverID, tmpSuppBNFStatus, tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
                        End If
                        If WLifeStatus Then
                            Call SaveMBMLifeTimeLmt(PrevLifeMBMNo, CoverID, SptLmtInd)
                        Else
                            TmpSql2 = "Select top 1 MBMCNumber,MBMCAvailableLimit From MBMCoveredPersons where MBMCNumber='" & Trim(PrevLifeMBMNo) & _
                                     "' AND MBMCCoverID='" & Trim(CoverID) & "'"
                            tmpMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                            If tmpMBM.EOF = False Then
                                If Mid(CboGuaRenewal, 1, 1) <> "3" Then
                                    MBMCovAnnLmt = IIf(IsNumeric(tmpMBM!MBMCAvailableLimit), tmpMBM!MBMCAvailableLimit, 0)
                                    MBMCovAvaLmt = IIf(IsNumeric(tmpMBM!MBMCAvailableLimit), tmpMBM!MBMCAvailableLimit, 0)
                                End If
                            End If
                            tmpMBM.Close
                            Set tmpMBM = Nothing
                        End If
                    End If
                End If
' ----- End of Life Time LImit
                
                SptMCOInd = "A"
                McoCovIDChrg = ""
                strSptMCOSql = "Select SuppMCOStatus, CovIDChrg from MCO where PLNCode = '" & Trim(tmpPlnCode) & _
                       "' AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                       "' AND MCOEffdate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"

                SptPlnMCO.Open strSptMCOSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
                If SptPlnMCO.recordCount > 0 Then
                    If Trim(SptPlnMCO!SuppMCOStatus) <> "" Then
                        SptMCOInd = Trim(SptPlnMCO!SuppMCOStatus)
                    End If
                    If Trim(SptPlnMCO!CovIDChrg) <> "" Then
                        McoCovIDChrg = Trim(SptPlnMCO!CovIDChrg)
                    End If
                Else
                    SptPlnMCO.Close
                    strSptMCOSql = "Select SuppMCOStatus, CovIDChrg from MCO where PLNCode is null " & _
                                   " AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                                   "' AND MCOEffdate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                    SptPlnMCO.Open strSptMCOSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    If Trim(SptPlnMCO!SuppMCOStatus) <> "" Then
                        SptMCOInd = Trim(SptPlnMCO!SuppMCOStatus)
                    End If
                    If Trim(SptPlnMCO!CovIDChrg) <> "" Then
                        McoCovIDChrg = Trim(SptPlnMCO!CovIDChrg)
                    End If
                End If
                SptPlnMCO.Close
                Set SptPlnMCO = Nothing

                If (Trim(SptMCOInd) = "B" And Trim(CoverID) = "01") _
                Or Trim(SptMCOInd) = "C" _
                Or (Trim(SptMCOInd) = "D" And Trim(CoverID) >= Trim(McoCovIDChrg)) Then
                    TmpBasicMCO = Format(GetSuppMCO(Trim(cboINSCode), Trim(CboHLTCode), tmpAdultChild, CDate(txtMBMPayorEffDate), Trim(tmpPlnCode)), "#,##0.00")
                    tmpMCO = TmpBasicMCO
                    If tmpProRate <> "N" Then
                        tmpMCO = InsertMCOFees(CDate(TmpCExpDate), CDate(tmpCEffDate), Format(TmpBasicMCO, "#,##0.00"))
                    End If
                    If tmpPLNRoundUpStatus <> "N" Then
                        tmpMCO = Round(tmpMCO)
                    End If
                End If

                If Trim(CboHLTCode) = "S" Then
                strsqlPrem = "Select SuppPrmStatus, PLNCode, SuppPrmAmt from PRM where " & _
                             " PLNCode = '" & Trim(tmpPlnCode) & "' And AgeCode = '" & Trim(AgeValue) & _
                             "' And INSCode = '" & Mid(cboINSCode, 1, 1) & _
                             "' AND HLTCode='" & Trim(CboHLTCode) & _
                             "' AND PRMEffDate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
                Else
                strsqlPrem = "Select SuppPrmStatus, PLNCode, SuppPrmAmt from PRM where " & _
                             " PLNCode = '" & Trim(tmpPlnCode) & "' And AgeCode = '" & Trim(AgeValue) & _
                             "' And INSCode = '" & Mid(cboINSCode, 1, 1) & _
                             "' AND HLTCode='" & Trim(CboHLTCode) & "' AND PAYCode='" & Trim(txtPAYCode) & _
                             "' AND PRMEffDate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
                End If
                Prem.Open strsqlPrem, medixmasterconn, adOpenKeyset, adLockOptimistic
                TmpSuppPrmStatus = "A"
                TmpSuppPrmAmt = 0
                If Prem.recordCount Then
                    TmpSuppPrmAmt = IIf(IsNumeric(Prem!SuppPrmAmt), Prem!SuppPrmAmt, 0)
                    If Trim(Prem!SuppPrmStatus) <> "" Then
                        TmpSuppPrmStatus = Prem!SuppPrmStatus
                    End If
                End If
                Prem.Close
                Set Prem = Nothing
                
                If (Trim(TmpSuppPrmStatus) = "B" And Trim(CoverID) = "01") Or Trim(TmpSuppPrmStatus) = "C" Then
                    tempSPremium = TmpSuppPrmAmt
                    tempDPremium = Format(GetSPremium(Trim(cboINSCode), Trim(txtPAYCode), AgeValue, Trim(cboPLNCode), Trim(CboHLTCode), CDate(txtMBMPayorEffDate)), "0.00")
                    tempDPremium = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(tempSPremium, "0.00"))
                         
                    If Disc = True Then
                        If lstCovAdd.ListItems(listloop).SubItems(36) <> "" Then
                             tempDPremium = tempDPremium * (100 - CDbl(lstCovAdd.ListItems(listloop).SubItems(36))) / 100
                        End If
                        If lstCovAdd.ListItems(listloop).SubItems(37) <> "" Then
                             tempDPremium = tempDPremium * (100 - CDbl(lstCovAdd.ListItems(listloop).SubItems(37))) / 100
                        End If
                    End If
                End If
            End If  ' IF PLNCode <> ""
        MBMCovAnnLmt = CDbl(IIf(IsNumeric(MBMCovAnnLmt) = True, MBMCovAnnLmt, 0))
        MBMCovAvaLmt = CDbl(IIf(IsNumeric(MBMCovAvaLmt) = True, MBMCovAvaLmt, 0))
        tempSPremium = Format(CDbl(IIf(IsNumeric(tempSPremium) = True, tempSPremium, 0)), "0.00")
        tempDPremium = Format(CDbl(IIf(IsNumeric(tempDPremium) = True, tempDPremium, 0)), "0.00")
        TmpBasicMCO = Format(CDbl(IIf(IsNumeric(TmpBasicMCO) = True, TmpBasicMCO, 0)), "0.00")
        tmpMCO = Format(CDbl(IIf(IsNumeric(tmpMCO) = True, tmpMCO, 0)), "0.00")
        TempMBMCDOB = Format(lstCovAdd.ListItems(listloop).SubItems(3), "YYYY/MM/DD")
        TempMBMCBordxDate = Format(txtMBMBordxDate, "YYYY/MM/DD")
        SuppPlanCode = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(41)), lstCovAdd.ListItems(listloop).SubItems(41), cboPLNCode)
        MembershipNo = MembershipNo
        tmpCIC = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(2)), Trim(lstCovAdd.ListItems(listloop).SubItems(2)), "")
        tmpCSex = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(4)), Trim(lstCovAdd.ListItems(listloop).SubItems(4)), "")
        TmpCAdultChild = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(5)), Trim(lstCovAdd.ListItems(listloop).SubItems(5)), "")
        tmpCRelCode = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(6)), Trim(lstCovAdd.ListItems(listloop).SubItems(6)), "")
        tmpCRegDate = Format(Date, "yyyy/mm/dd")
        tmpCEffDate = IIf(IsDate(tmpCEffDate), Format(tmpCEffDate, "YYYY/MM/DD"), "")
        TmpCExpDate = IIf(IsDate(TmpCExpDate), Format(TmpCExpDate, "YYYY/MM/DD"), "")
        strsqlMBMCov = "EXEC Upload_MBMCovPersonsNew '" & tmpCRelCode & "','" & MembershipNo & "','" & CoverID & "','" & TempMBMCDOB & "'," & _
                "'" & CovMBMType & "','" & TempMBMCName & "','A','" & TempMBMCBordxDate & "','" & txtMBMBatchNo & "', " & _
                "'" & tmpCIC & "','" & tmpCSex & "','" & SuppPlanCode & "', " & _
                "'" & AgeValue & "','" & TmpCAdultChild & "'," & MBMCovAnnLmt & "," & MBMCovAvaLmt & "," & TmpBasicMCO & "," & tmpMCO & ", " & _
                "" & tempSPremium & "," & tempDPremium & ",'" & tmpCRegDate & "','" & tmpCEffDate & "','" & TmpCExpDate & "'"

        medixmasterconn.Execute strsqlMBMCov
        
        tmpUndExcess = lstCovAdd.ListItems(listloop).SubItems(43)
        tmpUndExcess = CDbl(IIf(IsNumeric(tmpUndExcess) = True, tmpUndExcess, 0))

        TmpMBMCRBAmt = CDbl(IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(19)) = True, lstCovAdd.ListItems(listloop).SubItems(19), 0))
        TmpMBMCPayPremGross = CDbl(IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(32)) = True, lstCovAdd.ListItems(listloop).SubItems(32), 0))
        TmpMBMCPayPremNet = CDbl(IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(33)) = True, lstCovAdd.ListItems(listloop).SubItems(33), 0))
        TmpMBMCPayMCOGross = CDbl(IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(34)) = True, lstCovAdd.ListItems(listloop).SubItems(34), 0))
        TmpMBMCPayMCONet = CDbl(IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(35)) = True, lstCovAdd.ListItems(listloop).SubItems(35), 0))
        TmpMBMCFamDiscAmt = CDbl(IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(38)) = True, lstCovAdd.ListItems(listloop).SubItems(38), 0))
        TmpMBMCRenewDiscAmt = CDbl(IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(39)) = True, lstCovAdd.ListItems(listloop).SubItems(39), 0))
        TmpMBMCLoadingPrem = CDbl(IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(40)) = True, lstCovAdd.ListItems(listloop).SubItems(40), 0))
        TempMBMCExlcusion = Replace(lstCovAdd.ListItems(listloop).SubItems(9), "'", "''")
        tempMBMCPolicyDisc = CDbl(IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(36)) = True, lstCovAdd.ListItems(listloop).SubItems(36), 0))
        tempMBMCRenewalDisc = CDbl(IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(37)) = True, lstCovAdd.ListItems(listloop).SubItems(37), 0))

        tmpCNonPayFr = IIf(IsDate(lstCovAdd.ListItems(listloop).SubItems(29)), Format(lstCovAdd.ListItems(listloop).SubItems(29), "yyyy/mm/dd"), "")
        tmpCNonPayTo = IIf(IsDate(lstCovAdd.ListItems(listloop).SubItems(30)), Format(lstCovAdd.ListItems(listloop).SubItems(30), "yyyy/mm/dd"), "")

        tmpCSal = Replace(lstCovAdd.ListItems(listloop).SubItems(1), "'", "''")
        tmpCCovID = Replace(lstCovAdd.ListItems(listloop).SubItems(7), "'", "''")
        tmpCAllergies = Replace(lstCovAdd.ListItems(listloop).SubItems(8), "'", "''")
        tmpCOccup = Replace(lstCovAdd.ListItems(listloop).SubItems(12), "'", "''")
        tmpCPayREm = Replace(lstCovAdd.ListItems(listloop).SubItems(20), "'", "''")
        tmpCPrevINsCom = Replace(lstCovAdd.ListItems(listloop).SubItems(21), "'", "''")
        tmpCWN = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(13)), Trim(lstCovAdd.ListItems(listloop).SubItems(13)), "")
        tmpCSmking = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(14)), Trim(lstCovAdd.ListItems(listloop).SubItems(14)), "")
        tmpCAlcohol = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(15)), Trim(lstCovAdd.ListItems(listloop).SubItems(15)), "")
        tmpCPrevPln = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(18)), Trim(lstCovAdd.ListItems(listloop).SubItems(18)), "")
        tmpCPayMBMNo = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(44)), Trim(lstCovAdd.ListItems(listloop).SubItems(44)), "")
        tmpCPaySeqNo = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(45)), Trim(lstCovAdd.ListItems(listloop).SubItems(45)), "")
        tmpCClientNo = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(46)), Trim(lstCovAdd.ListItems(listloop).SubItems(46)), "")
        tmpCClientID = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(47)), Trim(lstCovAdd.ListItems(listloop).SubItems(47)), "")
        tmpC2IC = Replace(lstCovAdd.ListItems(listloop).SubItems(48), "'", "''")
        strsqlMBMCovTwo = "EXEC Upload_MBMCovPersonsTwo '" & Trim(MembershipNo) & "','" & CoverID & "','" & tmpCSal & "', " & _
            "'" & tmpCOccup & " ','" & tmpCWN & "','" & tmpCCovID & "','" & tmpCAllergies & "','" & tmpCSmking & "', " & _
            "'" & tmpCAlcohol & "','" & tmpCPrevPln & "','" & tmpCPrevINsCom & "','" & tmpCPayREm & "','" & TempMBMCExlcusion & "', " & _
            "'" & tmpCPayMBMNo & "','" & tmpCPaySeqNo & "','" & tmpCClientNo & "','" & tmpCClientID & "'," & TmpMBMCRBAmt & "," & _
            "" & TmpMBMCPayPremNet & ", " & TmpMBMCPayMCONet & ", " & TmpMBMCPayMCOGross & ", " & _
            "'" & tmpC2IC & "', " & TmpMBMCLoadingPrem & ", " & TmpMBMCFamDiscAmt & ", " & TmpMBMCRenewDiscAmt & ", " & TmpMBMCPayPremGross & "," & _
            "" & tmpUndExcess & ",'" & tmpCNonPayFr & "','" & tmpCNonPayTo & "'," & tempMBMCPolicyDisc & "," & tempMBMCRenewalDisc
        medixmasterconn.Execute strsqlMBMCovTwo
        
        'If Trim(frmMBMClinical.cboCisPln) <> "" Then
        '    Call saveCLNCoverPersons(MembershipNo, CoverID, lstCovAdd.ListItems(listloop).SubItems(3), lstCovAdd.ListItems(listloop).SubItems(31))
        'End If
        Next listloop
    End If
End Sub


Private Sub saveCLNCoverPersons(MembershipNo As String, TmpCovID, tmpDOb As String, tmpCisWarranty As String)
Dim FindAge As String
Dim AgeValue As String
Dim CISSuppAnnLmt, CISSuppVisLmt As String
Dim CISSuppBPrem, CISSuppPrem As String
Dim CISSuppBMCO, CISSuppMCO As String
Dim CisDentalLmt As Double
Dim CisSpecLmt As Double
Dim CisMepLmt As Double
Dim CisAnnLmtType As String
Dim CisSuppSql As String
Dim SysDate As String
Dim tmpCisExpDate As String
Dim tmpSuppExpDate As String
'Dim CISPayor As String
Dim CISstrSptLmt As String
Dim CISSptLmtInd As String
Dim CISSptPlnLmt As New ADODB.Recordset
Dim CISstrSptPrem As String
Dim CISSptPremInd As String
Dim CISSptPlnPrem As New ADODB.Recordset
Dim CISstrSptMCOSql As String
Dim CISSptMCOInd As String
Dim CISSptPlnMCO As New ADODB.Recordset
Dim tmpCISPln As String
Dim tmpCisInsType As String
Dim tmpCisEffDate As String
Dim tmpCardType As String
Dim CLNStatus As String

'    CISPayor = IIf(Trim(CISDCStatus) = "Y", "DC", Trim(txtPAYCode))

    tmpCISPln = Mid(frmMBMClinical.cboCisPln, 1, 4)
    tmpCisEffDate = frmMBMClinical.txtCLNEffDate
    tmpCisExpDate = frmMBMClinical.txtClnExpDate
    tmpCisInsType = Mid(frmMBMClinical.cboCisINS, 1, 1)
    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

    ' ******* Clinical Supp Limit
        CISstrSptLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(tmpCISPln) & _
                       "' AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                       "' AND AnnEffDate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by AnnEffDate desc"
        CISSptPlnLmt.Open CISstrSptLmt, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISSptLmtInd = "A"
        If CISSptPlnLmt.recordCount > 0 Then
            If Trim(CISSptPlnLmt!SuppLimitStatus) <> "" Then
                CISSptLmtInd = Trim(CISSptPlnLmt!SuppLimitStatus)
            End If
        End If
        CISSptPlnLmt.Close
        Set CISSptPlnLmt = Nothing
        If Trim(CboHLTCode) = "S" Then
            CISstrSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(tmpCISPln) & _
                             "' AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                             "' AND PRMEffDate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
        Else
            CISstrSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(tmpCISPln) & _
                             "' AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                             "' AND PAYCode='" & Trim(CISDCPayor) & _
                             "' AND PRMEffDate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
        End If
        CISSptPlnPrem.Open CISstrSptPrem, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISSptPremInd = "A"
        If CISSptPlnPrem.recordCount > 0 Then
            If Trim(CISSptPlnPrem!SuppPrmStatus) <> "" Then
                CISSptPremInd = Trim(CISSptPlnPrem!SuppPrmStatus)
            End If
        End If
        CISSptPlnPrem.Close
        Set CISSptPlnPrem = Nothing
        
        CISstrSptMCOSql = "Select SuppMCOStatus from MCO where PLNCode = '" & Trim(tmpCISPln) & _
                         "' AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                         "' AND MCOEffdate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"

        CISSptPlnMCO.Open CISstrSptMCOSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISSptMCOInd = "A"
        If CISSptPlnMCO.recordCount > 0 Then
            If Trim(CISSptPlnMCO!SuppMCOStatus) <> "" Then
                CISSptMCOInd = Trim(CISSptPlnMCO!SuppMCOStatus)
            End If
        Else
            CISSptPlnMCO.Close
            CISstrSptMCOSql = "Select SuppMCOStatus from MCO where PLNCode is null " & _
                              " AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                              "' AND MCOEffdate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"

            CISSptPlnMCO.Open CISstrSptMCOSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If CISSptPlnMCO.recordCount > 0 Then
                If Trim(CISSptPlnMCO!SuppMCOStatus) <> "" Then
                    CISSptMCOInd = Trim(CISSptPlnMCO!SuppMCOStatus)
                End If
            End If
        End If
        CISSptPlnMCO.Close
        Set CISSptPlnMCO = Nothing
'----
        CISSuppBPrem = 0
        CISSuppPrem = 0
        CISSuppBMCO = 0
        CISSuppMCO = 0
        CISSuppAnnLmt = 0
        CISSuppVisLmt = 0
        CisDentalLmt = 0
        CisSpecLmt = 0
        CisMepLmt = 0

        FindAge = GetAge(CDate(tmpCisEffDate), CDate(tmpDOb))
        AgeValue = Mid(GetAgeBand(Trim(CboHLTCode), FindAge, tmpCISPln, tmpCisInsType), 1, 2)
        TmpCovID = Format(TmpCovID, "00")
    
        If Trim(CISSptPremInd) = "B" And Format(TmpCovID, "00") = "01" Or Trim(CISSptPremInd) = "C" Then
            CISSuppPrem = Format(GetCLNSPremium(tmpCisInsType, CISDCPayor, AgeValue, tmpCISPln, Trim(CboHLTCode), CDate(tmpCisEffDate)), "#,##0.00")
            CISSuppBPrem = CISSuppPrem
            CISSuppPrem = InsertPremium(CDate(tmpCisExpDate), CDate(tmpCisEffDate), Format(CISSuppPrem, "#,##0.00"))
            CISSuppPrem = Format(Round(CISSuppPrem, 0), "#,##0.00")
        End If
        If Trim(CISSptMCOInd) = "B" And Format(TmpCovID, "00") = "01" Or Trim(CISSptMCOInd) = "C" Then
            If AgeValue = "01" Then
                CISSuppMCO = Format(GetCLNSMCO(tmpCisInsType, Trim(CboHLTCode), "C", CDate(tmpCisEffDate), tmpCISPln), "#,##0.00")
            Else
                CISSuppMCO = Format(GetCLNSMCO(tmpCisInsType, Trim(CboHLTCode), "A", CDate(tmpCisEffDate), tmpCISPln), "#,##0.00")
            End If
            CISSuppBMCO = CISSuppMCO
            CISSuppMCO = InsertMCOFees(CDate(tmpCisExpDate), CDate(tmpCisEffDate), Format(CISSuppMCO, "#,##0.00"))
            CISSuppMCO = Round(CISSuppMCO)
        End If
        If Trim(CISSptLmtInd) = "B" And Format(TmpCovID, "00") = "01" Or Trim(CISSptLmtInd) = "C" Then
            CISSuppAnnLmt = Format(GetCLNSuppLimit(tmpCisInsType, tmpCISPln, Trim(CboHLTCode), CDate(tmpCisEffDate)), "#,##0.00")
            CISSuppVisLmt = GetCLNSVisitLimit(tmpCisInsType, tmpCISPln, Trim(CboHLTCode), CDate(tmpCisEffDate))
            If Trim(CISSDentStatus) = "S" Then 'stand alone dental limit
                CisAnnLmtType = "SDentalAnnLimit"
                CisDentalLmt = GetCISLimit(tmpCisInsType, tmpCISPln, Trim(CboHLTCode), CDate(tmpCisEffDate), CisAnnLmtType)
            End If
            If Trim(CISSSpecStatus) = "S" Then 'stand alone Spec limit
                CisAnnLmtType = "SSpecAnnLimit"
                CisSpecLmt = GetCISLimit(tmpCisInsType, tmpCISPln, Trim(CboHLTCode), CDate(tmpCisEffDate), CisAnnLmtType)
            End If
            If Trim(CISSMepStatus) = "S" Then 'stand alone MEP limit
                CisAnnLmtType = "SMepAnnLimit"
                CisMepLmt = GetCISLimit(tmpCisInsType, tmpCISPln, Trim(CboHLTCode), CDate(tmpCisEffDate), CisAnnLmtType)
            End If
        End If
       ' MediConnCISDB.Close
       ' Set MediConnCISDB = Nothing
' ****** End of Clinical Supp Limit
        CISSuppBPrem = CDbl(IIf(IsNumeric(CISSuppBPrem) = True, CISSuppBPrem, 0))
        CISSuppPrem = CDbl(IIf(IsNumeric(CISSuppPrem) = True, CISSuppPrem, 0))
        CISSuppBMCO = CDbl(IIf(IsNumeric(CISSuppBMCO) = True, CISSuppBMCO, 0))
        CISSuppMCO = CDbl(IIf(IsNumeric(CISSuppMCO) = True, CISSuppMCO, 0))
        CISSuppAnnLmt = CDbl(IIf(IsNumeric(CISSuppAnnLmt) = True, CISSuppAnnLmt, 0))
        CISSuppVisLmt = CDbl(IIf(IsNumeric(CISSuppVisLmt) = True, CISSuppVisLmt, 0))
        SysDate = Format(Date, "yyyy/mm/dd")
        tmpSuppExpDate = Format(tmpCisExpDate, "yyyy/mm/dd")
        tmpCardType = Replace(frmMBMClinical.CboCardType, "'", "''")
        tmpCisWarranty = Replace(tmpCisWarranty, "'", "''")
        CLNStatus = "A"
        CisSuppSql = "EXEC Upload_MBMCovPersonsCis '" & MembershipNo & "','" & TmpCovID & "', " & _
                    "" & CISSuppMCO & "," & CISSuppBMCO & "," & CISSuppPrem & "," & CISSuppBPrem & ", " & _
                    "" & CISSuppAnnLmt & "," & CISSuppVisLmt & "," & CisDentalLmt & "," & CisSpecLmt & ", " & _
                    "" & CisMepLmt & ",'" & SysDate & "','" & Logon & "'," & _
                    "'" & SysDate & "','" & Logon & "','" & tmpSuppExpDate & "'," & _
                    "'" & tmpCardType & "','" & tmpCisWarranty & "','" & CLNStatus & "','" & Trim(tmpCISPln) & "'"
        medixmasterconn.Execute CisSuppSql
    ' Insert into Membership Clinic Covered Person's table
'    Next listloop
End Sub

Private Sub SaveMBMOthers(MembershipNo As String)
Dim MBMOthers As New ADODB.Recordset
Dim MBMOthers1 As New ADODB.Recordset
Dim strsqlMBMOth As String
Dim FirstJoinDate As String
Dim tmpPremium As String
Dim TmpActMCO As String
Dim TmpBasicMCO As String
Dim TmpBPremium As String
Dim TempPolicyDisc As String
Dim TempRenewalDisc As String
Dim TempGRPCompany As String
Dim TempUNDExcess As String
Dim TmpGuaRenew As String
Dim tmpBranch As String
Dim tmpMBMENDtRefNo As String
Dim tmpBTB As String

     FirstJoinDate = Format(txtMBMDateJoin, "YYYY/MM/DD")
     tmpPremium = Format(CDbl(IIf(IsNumeric(txtMBMPremium) = True, txtMBMPremium, 0)), "0.00")
     TmpActMCO = Format(CDbl(IIf(IsNumeric(txtMCOFees) = True, txtMCOFees, 0)), "0.00")
     TmpBasicMCO = Format(CDbl(IIf(IsNumeric(txtBMco) = True, txtBMco, 0)), "0.00")
     TmpBPremium = Format(CDbl(IIf(IsNumeric(txtBPrem) = True, txtBPrem, 0)), "0.00")
     TempPolicyDisc = CInt(IIf(IsNumeric(frmNewMBMOthers.txtPolicyDisc) = True, frmNewMBMOthers.txtPolicyDisc, 0))
     TempRenewalDisc = CInt(IIf(IsNumeric(frmNewMBMOthers.txtRenewalDisc) = True, frmNewMBMOthers.txtRenewalDisc, 0))
     TempGRPCompany = Replace(txtMBMGroupCompany, "'", "''")
     TempUNDExcess = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtUndExcess) = True, frmNewMBMOthers.txtUndExcess, 0))
     TmpGuaRenew = IIf(Len(CboGuaRenewal), Mid(CboGuaRenewal, 1, 1), "")
     tmpBranch = Trim(cboBRNCode)
     tmpMBMENDtRefNo = Replace(frmNewMBMOthers.txtEndtRefNo, "'", "''")
     tmpBTB = IIf(Len(frmNewMBMOthers.cboPOLCondition), Mid(frmNewMBMOthers.cboPOLCondition, 1, 1), "")
     strsqlMBMOth = "insert into MBMOthers (MBMNumber, MBMGroupCompany, MBMEmployeeNo, " & _
                    "MBMAgentCode, MBMPolSubNo1,  MBMBranch, " & _
                    "MBMPremium, MBMMCOFee, MBMBasicMCO, MBMBasicPrem,  " & _
                    "MBMDepartment, MBMInstallment, RELCODE, MBMUndExcess," & _
                    "MBMFirstMEMNo, MBMFirstPolNo, MBMPrevPolNo, MBMPrevMemNo,MBMFirstJoinedDate, " & _
                    "MBMPolicyDisc, MBMRenewalDisc,MBMPolCondition,MBMGuaRenewal,MBMENDtRefNo)" & _
                    "Values('" & MembershipNo & "', '" & TempGRPCompany & "', '" & txtMBMEmployeeNo & "', " & _
                    "'" & txtMBMAgentCode & "', '" & frmNewMBMOthers.TxtPOLSubNo & "', '" & tmpBranch & "', " & _
                    "" & tmpPremium & ", " & TmpActMCO & ", " & TmpBasicMCO & ", " & TmpBPremium & ", " & _
                    "'" & txtMBMDepartment & "', '" & Mid(CboInsMode, 1, 1) & "', 'P', " & TempUNDExcess & ", " & _
                    "'" & ChkStr(txtFirstMeMNo) & "', '" & txtFirstPolNo & "', '" & txtMBMPrevPolNo & "', '" & txtMBMPrevMEMNo & "', '" & FirstJoinDate & "', " & _
                    "" & TempPolicyDisc & ", " & TempRenewalDisc & ", '" & tmpBTB & "','" & TmpGuaRenew & "','" & tmpMBMENDtRefNo & "')"

    medixmasterconn.Execute strsqlMBMOth
    completeMBMOthers = True
End Sub

Private Sub SaveBatch()
    Dim BAT As New ADODB.Recordset
    Dim strsql As String
    
    strsql = "Select * from BAT Where PAYCode ='" & txtPAYCode & "' " & _
         " And BATBordxDate = '" & Format(txtMBMBordxDate, "MM/DD/YYYY") & "' " & _
         " And BordxType = 'N'"
    BAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                                                    
    If BAT.recordCount = 0 Then
        BAT.AddNew
    End If
    With BAT
        !BordxType = "N"
        !PAYCode = txtPAYCode
        !BATBordxDate = txtMBMBordxDate
        !BATSequence = txtMBMBatchNo
    End With
    BAT.Update
  

    BAT.Close
    Set BAT = Nothing
    completeMBMC = True
End Sub
    
Private Sub UpdateBordxPriting()
Dim BordxPrinting As New ADODB.Recordset

    BordxPrinting.Open "BordxPrinting", medixmasterconn, adOpenKeyset, adLockOptimistic
    BordxPrinting.AddNew
    With BordxPrinting
        !MBMBordxDate = txtMBMBordxDate
        !PAYCode = txtPAYCode
        !MBMBatchNo = txtMBMBatchNo
        !MBMBordxType = "N"
    End With
    BordxPrinting.Update
    BordxPrinting.Close
    Set BordxPrinting = Nothing
End Sub

'  ################# Start Suppplementary ######################


Private Sub cmdAdd_Click()
On Error Resume Next
Dim MemberConveredPerson As ListItem
Dim CountList As String
'Dim RecordSaved

    If txtCovName = "" Then
        MsgBox "Please Enter Covered Person Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtCovName.SetFocus
    ElseIf txtCovIcBcPp = "" Then
        MsgBox "Please Enter Covered Person Identification", vbOKOnly + vbCritical, DialogBoxTitle
        txtCovIcBcPp.SetFocus
    ElseIf txtCovBirthdate = "__/__/____" Then
        MsgBox "Please Enter Birthdate", vbOKOnly + vbCritical, DialogBoxTitle
        txtCovBirthdate.SetFocus
    ElseIf cboCovRelationship = "" Then
        MsgBox "Please Enter Covered Person Relationship", vbOKOnly + vbCritical, DialogBoxTitle
        cboCovRelationship.SetFocus
    Else
        CountList = Format(lstCovAdd.ListItems.Count + 1, "00")
        Set MemberConveredPerson = lstCovAdd.ListItems.Add(, , txtCovName)
        With MemberConveredPerson
            .SubItems(1) = cboCovSalutation
            .SubItems(2) = txtCovIcBcPp
            .SubItems(3) = txtCovBirthdate
            .SubItems(4) = cboCovSex
            .SubItems(5) = cboCovAdultChild
            .SubItems(6) = cboCovRelationship
            '.SubItems(7) = CountList
            .SubItems(8) = txtCovAllergic
            .SubItems(9) = txtCovExclusion
            .SubItems(10) = CountList
            .SubItems(11) = IIf(Trim(txtCovAnnualLimit) = "", 0, txtCovAnnualLimit)
            .SubItems(16) = IIf(IsDate(txtSuppEffDate), txtSuppEffDate, txtMBMPayorEffDate)
            .SubItems(17) = IIf(IsDate(txtSuppExpDate), txtSuppExpDate, txtMBMPayorExpDate)
            .SubItems(18) = txtPrevPLNCode
            .SubItems(19) = txtRBAmt
            .SubItems(20) = txtCOVPayRemarks
            .SubItems(21) = txtPrevINSCom
            .SubItems(28) = Mid(cboSuppStaffDisc, 1, 1)
            .SubItems(29) = TxtSuppClmNonPayFr
            .SubItems(30) = TxtSuppClmNonPayTo
            .SubItems(31) = txtCLNCovExclusion
            .SubItems(32) = txtSuppPayPremGross
            .SubItems(33) = txtSuppPayPremNet
            .SubItems(34) = txtSuppPayMCOGross
            .SubItems(35) = txtSuppPayMCONet
            .SubItems(36) = txtRenewalDisc
            .SubItems(37) = txtPolicyDisc
            .SubItems(38) = txtPolicyDiscAmt
            .SubItems(39) = txtRenewalDiscAmt
            .SubItems(40) = txtLoadingPrem
            .SubItems(41) = Trim(CboSuppPln)
            .SubItems(42) = txtAvailableLimit
            .SubItems(43) = IIf(IsNumeric(txtCovUndExcess), txtCovUndExcess, 0)
            .SubItems(44) = IIf(Len(txtPayMBMNo), Trim(txtPayMBMNo), "")
            .SubItems(45) = IIf(Len(txtPaySeqID), Trim(txtPaySeqID), "")
            .SubItems(46) = IIf(Len(txtClientNo), Trim(txtClientNo), "")
            .SubItems(47) = IIf(Len(txtClientID), Trim(txtClientID), "")
            
        End With
        Call ClearCoveredEntry
        If cboCovSalutation.Enabled = True Then
            cboCovSalutation.SetFocus
        End If
        cboCovRelationship = "S - SON"
    End If

End Sub

Private Sub CmdRemove_Click()
On Error Resume Next
    Dim ForLoop As Integer
    Dim ListCount As Integer
    Dim RecordSaved
    
    RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
      If RecordSaved = vbYes Then
          ListCount = lstCovAdd.ListItems.Count
          For ForLoop = 1 To ListCount
              If lstCovAdd.ListItems(ForLoop).Selected = True Then
                  lstCovAdd.ListItems.Remove (ForLoop)
                  Exit For
              End If
          Next ForLoop
          Call EnableCovered(True)
          Call ClearCoveredEntry
          cboCovSalutation.SetFocus
      End If
    cmdRemove.Enabled = True
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = False
End Sub


Private Sub cmdUpdate_Click()
On Error Resume Next
Dim ForLoop As Integer
    For ForLoop = 1 To lstCovAdd.ListItems.Count
        If lstCovAdd.ListItems(ForLoop - 1).Selected = True Then
            If lstCovAdd.SelectedItem.SubItems(7) = txtcount Then
                lstCovAdd.SelectedItem.Text = txtCovName
                lstCovAdd.SelectedItem.SubItems(1) = cboCovSalutation
                lstCovAdd.SelectedItem.SubItems(2) = txtCovIcBcPp
                lstCovAdd.SelectedItem.SubItems(3) = txtCovBirthdate
                lstCovAdd.SelectedItem.SubItems(4) = cboCovSex
                lstCovAdd.SelectedItem.SubItems(5) = cboCovAdultChild
                lstCovAdd.SelectedItem.SubItems(6) = cboCovRelationship
                lstCovAdd.SelectedItem.SubItems(8) = txtCovAllergic
                lstCovAdd.SelectedItem.SubItems(9) = txtCovExclusion
                lstCovAdd.SelectedItem.SubItems(11) = txtCovAnnualLimit
                lstCovAdd.SelectedItem.SubItems(16) = IIf(IsDate(txtSuppEffDate), txtSuppEffDate, txtMBMPayorEffDate)
                lstCovAdd.SelectedItem.SubItems(17) = IIf(IsDate(txtSuppExpDate), txtSuppExpDate, txtMBMPayorExpDate)
                lstCovAdd.SelectedItem.SubItems(18) = txtPrevPLNCode
                lstCovAdd.SelectedItem.SubItems(19) = txtRBAmt
                lstCovAdd.SelectedItem.SubItems(20) = txtCOVPayRemarks
                lstCovAdd.SelectedItem.SubItems(21) = txtPrevINSCom
                lstCovAdd.SelectedItem.SubItems(28) = Mid(cboSuppStaffDisc, 1, 1)
                lstCovAdd.SelectedItem.SubItems(29) = TxtSuppClmNonPayFr
                lstCovAdd.SelectedItem.SubItems(30) = TxtSuppClmNonPayTo
                lstCovAdd.SelectedItem.SubItems(31) = txtCLNCovExclusion
                lstCovAdd.SelectedItem.SubItems(32) = txtSuppPayPremGross
                lstCovAdd.SelectedItem.SubItems(33) = txtSuppPayPremNet
                lstCovAdd.SelectedItem.SubItems(34) = txtSuppPayMCOGross
                lstCovAdd.SelectedItem.SubItems(35) = txtSuppPayMCONet
                lstCovAdd.SelectedItem.SubItems(36) = txtRenewalDisc
                lstCovAdd.SelectedItem.SubItems(37) = txtPolicyDisc
                lstCovAdd.SelectedItem.SubItems(38) = txtPolicyDiscAmt
                lstCovAdd.SelectedItem.SubItems(39) = txtRenewalDiscAmt
                lstCovAdd.SelectedItem.SubItems(40) = txtLoadingPrem
                lstCovAdd.SelectedItem.SubItems(41) = Trim(CboSuppPln)
                lstCovAdd.SelectedItem.SubItems(42) = txtAvailableLimit
                lstCovAdd.SelectedItem.SubItems(43) = txtCovUndExcess
                lstCovAdd.SelectedItem.SubItems(44) = IIf(Len(txtPayMBMNo), Trim(txtPayMBMNo), "")
                lstCovAdd.SelectedItem.SubItems(45) = IIf(Len(txtPaySeqID), Trim(txtPaySeqID), "")
                lstCovAdd.SelectedItem.SubItems(46) = IIf(Len(txtClientNo), Trim(txtClientNo), "")
                lstCovAdd.SelectedItem.SubItems(47) = IIf(Len(txtClientID), Trim(txtClientID), "")


                Call ClearCoveredEntry
                Exit For
            End If
        End If
    Next ForLoop
    
    cboCovSalutation.SetFocus
    cmdRemove.Enabled = True
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = False
End Sub

Private Sub lstCovAdd_Click()
'On Error Resume Next
    If lstCovAdd.ListItems.Count Then
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovName = lstCovAdd.SelectedItem.Text
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(11), "###,##0.00")
        txtSuppEffDate = IIf(Len(lstCovAdd.SelectedItem.SubItems(16)), lstCovAdd.SelectedItem.SubItems(16), txtMBMPayorEffDate)
        txtSuppExpDate = IIf(Len(lstCovAdd.SelectedItem.SubItems(17)), lstCovAdd.SelectedItem.SubItems(17), txtMBMPayorExpDate)
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(18)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(19)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(20)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(21)
        cboSuppStaffDisc = lstCovAdd.SelectedItem.SubItems(28)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(29)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(30)
        txtCLNCovExclusion = lstCovAdd.SelectedItem.SubItems(31)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(32)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(33)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(34)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(35)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(36)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(37)
        txtPolicyDiscAmt = lstCovAdd.SelectedItem.SubItems(38)
        txtRenewalDiscAmt = lstCovAdd.SelectedItem.SubItems(39)
        txtLoadingPrem = lstCovAdd.SelectedItem.SubItems(40)
        CboSuppPln = lstCovAdd.SelectedItem.SubItems(41)
        txtAvailableLimit = Format(lstCovAdd.SelectedItem.SubItems(42), "###,##0.00")
        If IsNumeric(lstCovAdd.SelectedItem.SubItems(43)) Then
            txtCovUndExcess = lstCovAdd.SelectedItem.SubItems(43)
        End If
        txtPayMBMNo = IIf(Len(lstCovAdd.SelectedItem.SubItems(44)), Trim(lstCovAdd.SelectedItem.SubItems(44)), "")
        txtPaySeqID = IIf(Len(lstCovAdd.SelectedItem.SubItems(45)), Trim(lstCovAdd.SelectedItem.SubItems(45)), "")
        txtClientNo = IIf(Len(lstCovAdd.SelectedItem.SubItems(46)), Trim(lstCovAdd.SelectedItem.SubItems(46)), "")
        txtClientID = IIf(Len(lstCovAdd.SelectedItem.SubItems(47)), Trim(lstCovAdd.SelectedItem.SubItems(47)), "")
        Call EnableCovered(False)
        cmdUpdate.Enabled = True
        cmdRemove.Enabled = True
        cmdAdd.Enabled = False
    End If
    Call EnableCovered(True)
End Sub

Private Sub lstCovAdd_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCovAdd.SortKey = ColumnHeader.Index - 1
    lstCovAdd.Sorted = True
    If lstCovAdd.SortOrder = lvwAscending Then
        lstCovAdd.SortOrder = lvwDescending
    Else
        lstCovAdd.SortOrder = lvwAscending
    End If
End Sub


Private Sub EnableCovered(status As Boolean)
    'txtCovAnnualLimit.Enabled = status
    cboCovSalutation.Enabled = status
    txtCovName.Enabled = status
    txtCovIcBcPp.Enabled = status
    txtCovBirthdate.Enabled = status
    cboCovAdultChild.Enabled = status
    cboCovRelationship.Enabled = status
    txtCovAllergic.Enabled = status
    txtCovExclusion.Enabled = status
End Sub

Private Sub ClearCoveredEntry()
    OverAge = False
    cboCovSalutation.Text = "MR"
    txtCovName = ""
    txtCovIcBcPp = ""
    txtCovBirthdate.mask = ""
    txtCovBirthdate.Text = ""
    txtCovBirthdate.mask = "##/##/####"
    cboCovSex.Text = "M"
    cboCovAdultChild.Text = "A"
    cboCovRelationship.Text = "S - Son"
    txtCovAllergic = ""
    txtCovExclusion = ""
    txtCLNCovExclusion = ""
    txtcount = ""
    txtPrevPLNCode = ""
    txtRBAmt = ""
    txtPrevINSCom = ""
    txtCOVPayRemarks = ""
    txtCovAnnualLimit = ""
    
    TxtSuppClmNonPayFr.mask = ""
    TxtSuppClmNonPayFr.Text = ""
    TxtSuppClmNonPayFr.mask = "##/##/####"
    TxtSuppClmNonPayTo.mask = ""
    TxtSuppClmNonPayTo.Text = ""
    TxtSuppClmNonPayTo.mask = "##/##/####"
    txtSuppPayPremGross = ""
    txtSuppPayPremNet = ""
    txtSuppPayMCOGross = ""
    txtSuppPayMCONet = ""
    txtRenewalDisc = ""
    txtPolicyDisc = ""
    txtPolicyDiscAmt = ""
    txtRenewalDiscAmt = ""
    txtLoadingPrem = ""
    txtAvailableLimit = ""
    txtCovUndExcess = ""
    txtPayMBMNo = ""
    txtPaySeqID = ""
    txtClientNo = ""
    txtClientID = ""
    txtSuppExpDate = txtMBMPayorExpDate
    txtSuppEffDate = txtMBMPayorEffDate
    CboSuppPln.Text = cboPLNCode.Text
End Sub

Private Sub DeleteCoveredID(MembershipNo As String)
Dim DELCOV As New ADODB.Recordset
Dim delstr As String

    delstr = "Delete from MBM Where MBMMemberType = 'C' And MBMNumber = '" & MembershipNo & "'"
    DELCOV.Open delstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    DELCOV.Close
End Sub

Private Sub EnableSupplementary(status As Boolean)
    lstCovAdd.Enabled = status
    cboCovSalutation.Enabled = status
    txtCovName.Enabled = status
    txtCovIcBcPp.Enabled = status
    txtCovBirthdate.Enabled = status
    cboCovAdultChild.Enabled = status
    cboCovRelationship.Enabled = status
    txtCovAllergic.Enabled = status
    txtCovExclusion.Enabled = status
    'txtCovAnnualLimit.Enabled = status
    cmdRemove.Enabled = status
    cmdAdd.Enabled = status
End Sub

'  ################# END Suppplementary ######################


' ################ START Key press effect ###################
Private Sub cboCovAdultChild_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCovAdultChild, KeyAscii)
End Sub

Private Sub cboCovRelationship_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCovRelationship, KeyAscii)
End Sub

Private Sub cboCovSalutation_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCovSalutation, KeyAscii)
End Sub

Private Sub cboCovSex_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCovSex, KeyAscii)
End Sub

Private Sub cboINSCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboINSCode, KeyAscii)
End Sub

Private Sub cboPLNCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboPLNCode, KeyAscii)
End Sub

Private Sub cboRACCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboRACCode, KeyAscii)
End Sub

Private Sub cboMBMRenewal_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMBMRenewal, KeyAscii)
End Sub

Private Sub cboMBMSalutation_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMBMSalutation, KeyAscii)
End Sub

Private Sub cboMBMSex_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMBMSex, KeyAscii)
End Sub


Private Sub cboMBMMaritalStatus_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMBMMaritalStatus, KeyAscii)
End Sub


Private Sub cboNATCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboNATCode, KeyAscii)
End Sub

Private Sub cboMBmtakeover_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMBMTakeover, KeyAscii)
End Sub

Private Sub CboInsMode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboInsMode, KeyAscii)
End Sub

Private Sub cboCTYCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCTYCode, KeyAscii)
End Sub

Private Sub cboCTYCode_LostFocus()
    cboCTYCode = ChkStr(cboCTYCode)
    If Len(cboCTYCode) > 50 Then
        MsgBox "City/Town Exceeded 50 Characters!", vbCritical
        cboCTYCode.SetFocus
    End If
End Sub

Private Sub cboSTACode_LostFocus()
    cboSTACode = ChkStr(cboSTACode)
    If Len(cboSTACode) > 50 Then
        MsgBox "State Exceeded 50 Characters!", vbCritical
        cboSTACode.SetFocus
    End If
End Sub


Private Sub cboSTACode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboSTACode, KeyAscii)
End Sub


Private Sub cboBRNCode_LostFocus()
    If Len(cboBRNCode) > 50 Then
        MsgBox "Branch Exceeded 50 characters!", vbCritical
        cboBRNCode.SetFocus
    End If
    cboBRNCode = UCase(cboBRNCode)
End Sub

Private Sub cboBRNCode_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyCharacter(KeyAscii)
End Sub
' ################ END Key press effect ###################

Private Sub txtRemarks_Click()
    KeyPreview = False
End Sub

Private Sub txtRemarks_LostFocus()
    KeyPreview = True
    txtRemarks = ChkStr(txtRemarks)
End Sub

Private Sub txtSuppEffDate_LostFocus()
Dim TempSuppExpDate As String

    txtSuppEffDate = IIf(IsDate(txtSuppEffDate), txtSuppEffDate, txtMBMPayorEffDate)
    If txtSuppEffDate <> "__/__/____" Then
        If IsDate(txtSuppEffDate) = False Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtSuppEffDate.SetFocus
        Else
            If IsDate(txtMBMPayorEffDate) = False Or IsDate(txtMBMPayorExpDate) = False Then
                MsgBox "Please keyin Principal Effective or Expiry Date ! ", vbCritical
            Else
                If CDate(txtSuppEffDate) < CDate(txtMBMPayorEffDate) _
                    Or CDate(txtSuppEffDate) > CDate(txtMBMPayorExpDate) Then
                    MsgBox "Supllementary Effective Date is out of acceptable range! " & vbNewLine & _
                            "It must be in between " & txtMBMPayorEffDate & " and " & txtMBMPayorExpDate & " !", vbCritical + vbOKOnly
                    txtSuppEffDate.Text = "__/__/____"
                    txtSuppEffDate.mask = "##/##/####"
                    txtSuppEffDate.SetFocus
                Else
                    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                    If Trim(CboSuppPln) <> "" And IsDate(txtCovBirthdate) = True And IsDate(txtSuppEffDate) And IsDate(txtSuppExpDate) Then
                        TempSuppExpDate = GetSuppExpDate(txtCovBirthdate, CboSuppPln, txtSuppEffDate, txtSuppExpDate)
                        txtSuppExpDate = TempSuppExpDate
                    End If
                    'medixmasterconn.Close
                   ' Set medixmasterconn = Nothing
                    If OverAge = True Then
                        MsgBox "Supplementary age is above " & PlnAgeLmt
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Function showMBMOthers() As Integer
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
    
        MBMNumber = txtMBMPrevMEMNo
        With ADOrstMBM
            
            frmNewMBMOthers.Label4 = Trim(MBMNumber)
            If Trim(!MBMCLMNonPayFr) <> "" Then
                frmNewMBMOthers.TxtClmNonPayFr = Format(!MBMCLMNonPayFr, "dd/mm/yyyy")
            End If
            If Trim(!MBMCLMNonPayTo) <> "" Then
                frmNewMBMOthers.TxtClmNonPayTo = Format(!MBMCLMNonPayTo, "dd/mm/yyyy")
            End If
            If Trim(!LabelStatus) = "Y" Then
                frmNewMBMOthers.Check1.Value = 1
            End If
            
            frmNewMBMOthers.TxtPOLSubNo = IIf(Len(Trim(!MBMPolSubNo1)), Trim(!MBMPolSubNo1), "")
            frmNewMBMOthers.txtRenewalDisc = IIf(IsNumeric(!MBMRenewalDisc), !MBMRenewalDisc, 0)
            frmNewMBMOthers.txtPolicyDisc = IIf(IsNumeric(!MBMPolicyDisc), !MBMPolicyDisc, 0)
            frmNewMBMOthers.txtPLNCode = IIf(Len(!MBMNewPlanCode), !MBMNewPlanCode, "")
            
            If Len(!MBMNewPolEffDate) Then
                frmNewMBMOthers.txtNewPolEffDate = Format(!MBMNewPolEffDate, "dd/mm/yyyy")
            End If
            
            txtMBMGroupCompany = IIf(Len(Trim(!MBMGroupCompany)), Trim(!MBMGroupCompany), "")
            txtMBMEmployeeNo = IIf(Len(Trim(!MBMEmployeeNo)), Trim(!MBMEmployeeNo), "")
            txtFirstPolNo = IIf(Len(Trim(!MBMFirstPolNo)), Trim(!MBMFirstPolNo), "")
            txtFirstMeMNo = IIf(Len(Trim(!MBMFirstMemNo)), Trim(!MBMFirstMemNo), "")
            txtMBMAgentCode = IIf(Len(Trim(!MBMAgentCode)), Trim(!MBMAgentCode), "")

            If Trim(!MBMFirstJoinedDate) <> "" Then
                txtMBMDateJoin = !MBMFirstJoinedDate
            End If

            txtMBMDepartment = IIf(Len(Trim(!MBMDepartment)), Trim(!MBMDepartment), "")
            frmNewMBMOthers.cboPOLCondition = IIf(Len(!MBMPolCondition), !MBMPolCondition, "")
            frmNewMBMOthers.txtUndExcess = IIf(IsNumeric(!MBMUndExcess), !MBMUndExcess, 0)
            CboGuaRenewal.Text = IIf(Len(!MBMGuaRenewal), !MBMGuaRenewal, "")
            frmNewMBMOthers.txtEndtRefNo = IIf(Len(!MBMENDtRefNo), !MBMENDtRefNo, "")
        End With
End Function

Private Function showMBMTwo() As Integer
Dim sqlstr As String
Dim MBMTwo As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
    
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    
    Call ClearForm(frmNewMBMOthers)

    If Not ADOrstMBM.EOF Then
        With ADOrstMBM
            If Len(Trim(!MBMSalutation)) Then
               cboMBMSalutation = Trim(!MBMSalutation)
            End If
            
            If Len(Trim(!NATCode)) Then
                cboNATCode = Trim(!NATCode)
            End If
            
            If Trim(!MBMTelnoH) <> "" Then
                frmNewMBMOthers.txtMBMTelNoH = Trim(!MBMTelnoH)
            End If
            
            If Trim(!MBMTelNoM) <> "" Then
               frmNewMBMOthers.txtMBMTelNoM = Trim(!MBMTelNoM)
            End If
            
            If Trim(!MBMTelNoO) <> "" Then
                frmNewMBMOthers.txtMBMTelNoO = Trim(!MBMTelNoO)
            End If

            If Len(!MBMEmail) Then
                frmNewMBMOthers.txtMBMEmail = !MBMEmail
            End If
            
            If Len(!LGNCode) Then
               frmNewMBMOthers.cboLGNCode = !LGNCode
            End If
            
        End With
    End If
End Function

Private Function showMBMOthersTwo() As Integer
Dim sqlstr As String
Dim MBMOthersTwo As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
    
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    
    MBMNumber = txtMBMPrevMEMNo
    If Not ADOrstMBM.EOF Then
        With ADOrstMBM
            frmNewMBMOthers.Label4 = Trim(MBMNumber)
            frmNewMBMOthers.txtMBMPrevPlan = IIf(Len(!MBMPrevPLNCode), !MBMPrevPLNCode, "")
            frmNewMBMOthers.txtMBMRBAmt = IIf(IsNumeric(!MBMRBAmt), !MBMRBAmt, 0)
            frmNewMBMOthers.txtMBMPrevINSCom = IIf(Len(!MBMPrevInsCom), !MBMPrevInsCom, "")
            frmNewMBMOthers.cboCOPayRB = IIf(Len(!MBMCoPayRB), !MBMCoPayRB, "")
            frmNewMBMOthers.TxtGenWP = IIf(Len(!MBMGenWP), !MBMGenWP, "")
            frmNewMBMOthers.txtPREWP = IIf(Len(!MBMPreExWP), !MBMPreExWP, "")
            frmNewMBMOthers.TxtSpeWP = IIf(Len(!MBMSpeIllWP), !MBMSpeIllWP, "")
            frmNewMBMOthers.cboStaffDisc = IIf(Len(!MBMStaffDec), !MBMStaffDec, "")
            If Len(!MBMGenEndWP) Then
                frmNewMBMOthers.txtGENWPEnd = !MBMGenEndWP
            End If
            If Len(!MBMPreExEndWP) Then
                frmNewMBMOthers.txtPREWPEnd = !MBMPreExEndWP
            End If
            If Len(!MBMSpeIllEndWP) Then
                frmNewMBMOthers.txtSPEWPEnd = !MBMSpeIllEndWP
            End If
            frmNewMBMOthers.txtMBMAllergic = IIf(Len(Trim(!MBMAllergic)), Trim(!MBMAllergic), "")
            txtRemarks = IIf(Len(!MBMRemarks), !MBMRemarks, "")
            frmNewMBMOthers.cboMBMOccupation = IIf(Len(!MBMOccupation), !MBMOccupation, "")
            frmNewMBMOthers.txtMBMWorkNature = IIf(Len(!MBMWorkNature), !MBMWorkNature, "")
            frmNewMBMOthers.txtMBMWeight = IIf(Len(!MBMWeight), !MBMWeight, "")
            frmNewMBMOthers.txtMBMHeight = IIf(Len(!MBMHeight), !MBMHeight, "")
            frmNewMBMOthers.cboMBMBlood = IIf(Len(!MBMBlood), !MBMBlood, "")
            frmNewMBMOthers.cboMBMSmoking = IIf(Len(!MBMSmoking), !MBMSmoking, "")
            frmNewMBMOthers.cboMBMAlcohol = IIf(Len(!MBMAlcohol), !MBMAlcohol, "")
            frmNewMBMOthers.txtPayPrem = IIf(IsNumeric(!MBMPayorPremGross), !MBMPayorPremGross, 0)
            frmNewMBMOthers.txtPayPremNett = IIf(IsNumeric(!MBMPayorPremNet), !MBMPayorPremNet, 0)
            frmNewMBMOthers.txtPayMCOGross = IIf(IsNumeric(!MBMPayorMCOGross), !MBMPayorMCOGross, 0)
            frmNewMBMOthers.txtPayMCONett = IIf(IsNumeric(!MBMPayorMCONet), !MBMPayorMCONet, 0)
            frmNewMBMOthers.cboBTBG = IIf(Len(!MBMBTBG), !MBMBTBG, "")
            'frmNewMBMOthers.txtMBMPAYRemarks = IIf(Len(!MBMPAYRemarks), !MBMPAYRemarks, "")
            frmNewMBMOthers.txtPolicyDiscAmt = IIf(IsNumeric(!MBMFamDiscAmt), !MBMFamDiscAmt, 0)
            frmNewMBMOthers.txtRenewalDiscAmt = IIf(IsNumeric(!MBMRenewDiscAmt), !MBMRenewDiscAmt, 0)
            frmNewMBMOthers.txtLoadingPrem = IIf(IsNumeric(!MBMLoadingPrem), !MBMLoadingPrem, 0)
            txtMBMExclusion = IIf(Len(Trim(!MBMExclusion)), Trim(!MBMExclusion), "")
            frmNewMBMOthers.txtBankACNo = IIf(Len(Trim(!MBMBankACNo)), Trim(!MBMBankACNo), "")

        End With
    End If
End Function

Private Sub SaveMBMOthersTwo(MembershipNo As String)
    
    Dim strsqlMBMOthTwo As String
    Dim TmpPayPremium As String
    Dim TmpPayPremiumNett As String
    Dim TmpPayMCOGross As String
    Dim TmpPayMCONett As String
    Dim TmpMBMRBAmt As String
    Dim TmpPolicyDiscAmt As String
    Dim TmpRenewalDiscAmt  As String
    Dim TmpLoadingPrem As String
    Dim TempMBMExclusion As String
    Dim tmpRemarks As String
    Dim tmpPrevInsComp As String
    Dim tmpAllergic As String
    Dim tmpPayRemarks As String
    
    TmpPayPremium = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtPayPrem) = True, frmNewMBMOthers.txtPayPrem, 0))
    TmpPayPremiumNett = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtPayPremNett) = True, frmNewMBMOthers.txtPayPremNett, 0))
    TmpPayMCOGross = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtPayMCOGross) = True, frmNewMBMOthers.txtPayMCOGross, 0))
    TmpPayMCONett = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtPayMCONett) = True, frmNewMBMOthers.txtPayMCONett, 0))
    TmpMBMRBAmt = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtMBMRBAmt) = True, frmNewMBMOthers.txtMBMRBAmt, 0))
    TmpPolicyDiscAmt = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtPolicyDiscAmt) = True, frmNewMBMOthers.txtPolicyDiscAmt, 0))
    TmpRenewalDiscAmt = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtRenewalDiscAmt) = True, frmNewMBMOthers.txtRenewalDiscAmt, 0))
    TmpLoadingPrem = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtLoadingPrem) = True, frmNewMBMOthers.txtLoadingPrem, 0))
    TempMBMExclusion = Replace(txtMBMExclusion, "'", "''")
    tmpRemarks = Replace(txtRemarks, "'", "''")
    tmpPrevInsComp = Replace(frmNewMBMOthers.txtMBMPrevINSCom, "'", "''")
    tmpAllergic = Replace(frmNewMBMOthers.txtMBMAllergic, "'", "''")
    'tmpPayRemarks = Replace(frmNewMBMOthers.txtMBMPAYRemarks, "'", "''")
    ' INsert into Membership Other details table
     
        strsqlMBMOthTwo = "insert into MBMOthersTwo (MBMNumber, MBMWeight, MBMHeight, " & _
                    "MBMOccupation, MBMWorkNature, MBMBlood, MBMSmoking, " & _
                    "MBMAlcohol, MBMPrevPLNCode, MBMPAYRemarks, " & _
                    "MBMPrevInsCom, MBMMaritalStatus, MBMCoPayRB, MBMStaffDec, MBMPayorPremGross, " & _
                    "MBMPayorPremNet,MBMPayorMCOGross,MBMPayorMCONet," & _
                    "MBMRBAmt, MBMFamDiscAmt, MBMRenewDiscAmt,MBMLoadingPrem,MBMExclusion,MBMBankACNo, MBMAllergic,MBMRemarks)" & _
                    "Values('" & MembershipNo & "', '" & frmNewMBMOthers.txtMBMWeight & "', '" & frmNewMBMOthers.txtMBMHeight & "', " & _
                    "'" & Trim(Mid(frmNewMBMOthers.cboMBMOccupation, 1, 2)) & "', '" & Mid(Trim(frmNewMBMOthers.txtMBMWorkNature), 1, 3) & "', '" & frmNewMBMOthers.cboMBMBlood & "', '" & frmNewMBMOthers.cboMBMSmoking & "', " & _
                    "'" & frmNewMBMOthers.cboMBMAlcohol & "' , '" & frmNewMBMOthers.txtMBMPrevPlan & "', '" & tmpPayRemarks & "', " & _
                    "'" & tmpPrevInsComp & "', '" & Mid(cboMBMMaritalStatus, 1, 1) & "', '" & frmNewMBMOthers.cboCOPayRB & "', '" & Mid(frmNewMBMOthers.cboStaffDisc, 1, 1) & "', " & TmpPayPremium & ", " & _
                    "" & TmpPayPremiumNett & ", " & TmpPayMCOGross & ", " & TmpPayMCONett & ", " & _
                    "" & TmpMBMRBAmt & ", " & TmpPolicyDiscAmt & ", " & TmpRenewalDiscAmt & ", " & TmpLoadingPrem & " , '" & ChkStr(TempMBMExclusion) & "', '" & ChkStr(frmNewMBMOthers.txtBankACNo) & "'," & _
                    "'" & tmpAllergic & "','" & tmpRemarks & "') "
        medixmasterconn.Execute strsqlMBMOthTwo
        completeMBMOthersTwo = True
End Sub

Private Sub ClearMBMOthers()
On Error Resume Next

    frmNewMBMOthers.Label4 = ""
    frmNewMBMOthers.TxtClmNonPayFr.mask = ""
    frmNewMBMOthers.TxtClmNonPayFr.Text = ""
    frmNewMBMOthers.TxtClmNonPayFr.mask = "##/##/####"
    frmNewMBMOthers.TxtClmNonPayTo.mask = ""
    frmNewMBMOthers.TxtClmNonPayTo.Text = ""
    frmNewMBMOthers.TxtClmNonPayTo.mask = "##/##/####"
    frmNewMBMOthers.Check1.Value = 0
    frmNewMBMOthers.TxtPOLSubNo = ""
    frmNewMBMOthers.txtMBMAllergic = ""
    'frmNewMBMOthers.txtMBMPAYRemarks = ""
    frmNewMBMOthers.txtPolicyDisc = ""
    frmNewMBMOthers.txtRenewalDisc = ""
    frmNewMBMOthers.txtPLNCode = ""
    frmNewMBMOthers.txtNewPolEffDate.mask = ""
    frmNewMBMOthers.txtNewPolEffDate.Text = ""
    frmNewMBMOthers.txtNewPolEffDate.mask = "##/##/####"
    txtMBMGroupCompany = ""
    txtMBMEmployeeNo = ""
    txtRemarks = ""
    txtMBMExclusion = ""
    txtFirstPolNo = ""
    txtFirstMeMNo = ""
    txtMBMAgentCode = ""
    txtMBMDateJoin.mask = ""
    txtMBMDateJoin.Text = ""
    txtMBMDateJoin.mask = "##/##/####"
    frmNewMBMOthers.txtGENWPEnd.mask = ""
    frmNewMBMOthers.txtGENWPEnd.Text = ""
    frmNewMBMOthers.txtGENWPEnd.mask = "##/##/####"
    frmNewMBMOthers.txtPREWPEnd.mask = ""
    frmNewMBMOthers.txtPREWPEnd.Text = ""
    frmNewMBMOthers.txtPREWPEnd.mask = "##/##/####"
    frmNewMBMOthers.txtSPEWPEnd.mask = ""
    frmNewMBMOthers.txtSPEWPEnd.Text = ""
    frmNewMBMOthers.txtSPEWPEnd.mask = "##/##/####"
    frmNewMBMOthers.TxtClmNonPayFr.Text = ""
    frmNewMBMOthers.TxtClmNonPayFr.mask = "##/##/####"
    frmNewMBMOthers.TxtClmNonPayTo.Text = ""
    frmNewMBMOthers.TxtClmNonPayTo.mask = "##/##/####"
    frmNewMBMOthers.cboBTBG = "N - No"
End Sub

Private Sub SaveMBMCLNOthers(MembershipNo As String)
Dim ClnOthersSql As String
Dim tmpCisPrem, tmpCisBPrem, tmpCISMCO, tmpCisBMCO As Double
Dim tmpPolno, TmpCisWarr, tmpRemark As String
Dim CISPLN As String
Dim CisInsType As String
Dim CLnAnnLmt As New ADODB.Recordset
Dim CLNPln As New ADODB.Recordset
Dim CLNPlnSql, ClnAnnLmtSql As String
Dim CISCovSuppLMtStat, TmpMBMType, TmpCisPayor As String
Dim tmpCisEffDate As String
Dim tmpCisExpDate As String
Dim CisDentalLmt As Double
Dim CisSpecLmt As Double
Dim CisMepLmt As Double
Dim CisAnnLmtType As String
Dim CisVisitLmt As Double
Dim CisBalLmt As Double
Dim tmpCardType As String

    TmpCisPayor = Trim(txtPAYCode)
    CISPLN = Mid(frmMBMClinical.cboCisPln, 1, 4)
    CisInsType = Mid(frmMBMClinical.cboCisINS, 1, 1)
    
    CISCovSuppLMtStat = "A"
   ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

    CLNPlnSql = "SELECT ANNLmtInd, DCStatus, PAYCode from PLN where PLNCode='" & Trim(CISPLN) & _
            "' AND HLTCode='" & Trim(CboHLTCode) & "'"
    CLNPln.Open CLNPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If CLNPln.recordCount Then
        TmpCisPayor = IIf(Trim(CLNPln!DCStatus) = "Y", CLNPln!PAYCode, Trim(txtPAYCode))
        CISDCPayor = Trim(TmpCisPayor)
        If Trim(CLNPln!AnnLmtInd) = "Y" Then
            ClnAnnLmtSql = "Select SuppLimitStatus from AnnualLimit where PLNCode = '" & Trim(CISPLN) & _
                           "' AND HLTCode='" & Trim(CboHLTCode) & "' AND INSCode='" & Trim(CisInsType) & "'"
            CLnAnnLmt.Open ClnAnnLmtSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If CLnAnnLmt.recordCount > 0 Then
                CISCovSuppLMtStat = CLnAnnLmt!SuppLimitStatus
            End If
            CLnAnnLmt.Close
            Set CLnAnnLmt = Nothing
        End If
    End If
    CLNPln.Close
    Set CLNPln = Nothing
    
    CisDentalLmt = 0
    CisSpecLmt = 0
    CisMepLmt = 0
    
    If Trim(CISPDentStatus) = "S" Then 'stand alone dental limit
        CisAnnLmtType = "PDentalAnnLimit"
        CisDentalLmt = GetCISLimit(CisInsType, CISPLN, Trim(CboHLTCode), CDate(frmMBMClinical.txtCLNEffDate), CisAnnLmtType)
    End If
    If Trim(CISPSpecStatus) = "S" Then 'stand alone Spec limit
        CisAnnLmtType = "PSpecAnnLimit"
        CisSpecLmt = GetCISLimit(CisInsType, CISPLN, Trim(CboHLTCode), CDate(frmMBMClinical.txtCLNEffDate), CisAnnLmtType)
    End If
    If Trim(CISPMepStatus) = "S" Then 'stand alone MEP limit
        CisAnnLmtType = "PMepAnnLimit"
        CisMepLmt = GetCISLimit(CisInsType, CISPLN, Trim(CboHLTCode), CDate(frmMBMClinical.txtCLNEffDate), CisAnnLmtType)
    End If
    If Trim(CISCovSuppLMtStat) = "A" Then
        TmpMBMType = "P"
    Else
        TmpMBMType = "Q"
    End If
   ' MediConnCISDB.Close
   ' Set MediConnCISDB = Nothing
    tmpCisPrem = IIf(IsNumeric(frmMBMClinical.lblMBMPremium) = True, frmMBMClinical.lblMBMPremium, 0)
    tmpCisBPrem = IIf(IsNumeric(frmMBMClinical.txtTempPrem) = True, frmMBMClinical.txtTempPrem, 0)
    tmpCISMCO = IIf(IsNumeric(frmMBMClinical.lblMCOFees) = True, frmMBMClinical.lblMCOFees, 0)
    tmpCisBMCO = IIf(IsNumeric(frmMBMClinical.txtTempMCOFees) = True, frmMBMClinical.txtTempMCOFees, 0)
    tmpCisEffDate = Format(frmMBMClinical.txtCLNEffDate, "yyyy/mm/dd")
    tmpCisExpDate = Format(frmMBMClinical.txtClnExpDate, "yyyy/mm/dd")
    tmpPolno = Replace(txtMBMPolicyNo, "'", "''")
    TmpCisWarr = Replace(frmMBMClinical.txtWarranty, "'", "''")
    tmpRemark = Replace(frmMBMClinical.txtRemarks, "'", "''")
    CisBalLmt = IIf(IsNumeric(frmMBMClinical.lblAnnualLimit) = True, frmMBMClinical.lblAnnualLimit, 0)
    CisVisitLmt = IIf(IsNumeric(frmMBMClinical.LblVisitLimit) = True, frmMBMClinical.LblVisitLimit, 0)
    CisDentalLmt = CDbl(IIf(IsNumeric(CisDentalLmt) = True, CisDentalLmt, 0))
    CisSpecLmt = CDbl(IIf(IsNumeric(CisSpecLmt) = True, CisSpecLmt, 0))
    CisMepLmt = CDbl(IIf(IsNumeric(CisMepLmt) = True, CisMepLmt, 0))
    tmpCardType = Replace(frmMBMClinical.CboCardType, "'", "''")

    ClnOthersSql = "Insert Into MBMCLNOthers(MBMNumber, MBMPolNumber, PAYCode, " & _
                 "PLNType, InsType, MBMPolEffDate, MBMPolExpDate, MBMRemarks, MBMCLNWarranty," & _
                 "MBMCLNPremium, MBMCLNBasicPrem, MBMCLNMCOFee, MBMCLNBasicMCO, " & _
                 "MBMCLNMemberType, MBMCLNExpiryStatus, MBMCLNDataStatus, MBMCLNStatus, MBMCardType," & _
                 "MBMCLNVisitLmt, MBMCLNBalLmt, MBMMEPLmt, MBMDentalLmt, MBMSpecLmt) " & _
                 "Values('" & MembershipNo & "','" & Trim(txtMBMPolicyNo) & "', '" & TmpCisPayor & "'," & _
                 "'" & CISPLN & "','" & CisInsType & "','" & tmpCisEffDate & "'," & _
                 "'" & tmpCisExpDate & "','" & Trim(tmpRemark) & "','" & Trim(TmpCisWarr) & "'," & _
                 "" & tmpCisPrem & "," & tmpCisBPrem & "," & tmpCISMCO & "," & tmpCisBMCO & "," & _
                 "'" & TmpMBMType & "','N','A','A'" & ",'" & Trim(tmpCardType) & "'," & _
                 "" & CisVisitLmt & "," & CisBalLmt & "," & CisMepLmt & "," & CisDentalLmt & "," & CisSpecLmt & ")"
                 
    medixmasterconn.Execute ClnOthersSql
    
End Sub

Private Function GetSuppExpDate(DateOfBirth, tmpPlanCode, TmpEffDate, tmpExpDate As String)
Dim CheckYr As Long
Dim PLNRec As New ADODB.Recordset
Dim PLNSql As String
Dim SuppDobExpDate As String
Dim SuppDobEffDate As String

    SuppDobEffDate = ""
    OverAge = False
    PlnAgeLmt = 0
    SuppDobEffDate = Format(DateOfBirth, "dd/mm") & "/" & Year(TmpEffDate)
    SuppDobExpDate = Format(DateOfBirth, "dd/mm") & "/" & Year(tmpExpDate)
    CheckYr = Year(tmpExpDate) - Year(DateOfBirth)
    
    PLNSql = "select PlnAgeLimit from PLN where PLNCode = '" & Trim(tmpPlanCode) & "'"
    PLNRec.Open PLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If PLNRec.recordCount > 0 Then
        If Trim(PLNRec!PlnAgeLimit) <> "" Then
             PlnAgeLmt = IIf(IsNumeric(PLNRec!PlnAgeLimit), PLNRec!PlnAgeLimit, 0)
        End If
    End If
    PLNRec.Close
    Set PLNRec = Nothing
    GetSuppExpDate = tmpExpDate
    If PlnAgeLmt <> 0 Then
        If CInt(CheckYr) = CInt(PlnAgeLmt) Then
            If CDate(SuppDobExpDate) <= CDate(tmpExpDate) Then
                GetSuppExpDate = CDate(SuppDobExpDate)
            End If
            If CDate(SuppDobEffDate) <= CDate(TmpEffDate) Then
                OverAge = True
            End If
        ElseIf CheckYr > PlnAgeLmt Then
            OverAge = True
        End If
    End If
End Function

Private Sub SaveMBMLifeTimeLmt(LMBMNumber, LMBMCovID, LmtType As String)
Dim tmpSql As String
Dim TmpLifeSql As String
Dim TmpLimit As Double
Dim TmpLmtVar As String
Dim cmd As New ADODB.Command
Dim AccumAmtRec As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim TmpSql2 As Integer
Dim tmpAccAmt As Double
    TmpLimit = 0
    tmpAccAmt = 0

    TmpSql2 = 0
    If Trim(LMBMCovID) <> "00" Then
        TmpSql2 = 1
    End If
    tmpSql = " And MBMTopupNumber ='" & Trim(LMBMNumber) & "' AND MBMPatientCovID='" & Trim(LMBMCovID) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchAccAppAmt"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 255, tmpSql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 255, TmpSql2)
    cmd.Parameters.Append Prm2
    AccumAmtRec.CursorLocation = adUseClient
    AccumAmtRec.CursorType = adOpenForwardOnly
    AccumAmtRec.CacheSize = 100
    Set AccumAmtRec = cmd.Execute
    Do While Not AccumAmtRec.EOF
        With AccumAmtRec
            tmpAccAmt = IIf(IsNumeric(!CLMTotalApproved), !CLMTotalApproved, 0)
            AccumAmtRec.MoveNext
        End With
    Loop
    AccumAmtRec.Close
    Set AccumAmtRec = Nothing

    TmpLmtVar = "SuppLifeTimeLimit"
    If LmtType = "A" Then
        TmpLmtVar = "LifeTimeLimit"
    End If
    TmpLimit = GetLifeTimeLimit(TmpLmtVar, Trim(cboINSCode), Trim(cboPLNCode), Trim(CboHLTCode), CDate(txtMBMPayorEffDate))
    TmpLimit = CDbl(TmpLimit) - tmpAccAmt
    TmpLifeSql = "Insert Into MBMLifeTime(MBMNumber, MBMCOVID, MBMLifeTimeLmt) " & _
          "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & TmpLimit & ")"
    medixmasterconn.Execute TmpLifeSql
End Sub

Private Sub SaveMBMBnfAnnLmt(LMBMNumber, LMBMCovID, LmtType, tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
Dim TmpBnfLmtSql As String
    
    TmpBnfLmtSql = "Insert Into MBMBnfLmt(MBMNumber, MBMCoveredID, MBMKidneyAnnLmt, MBMKidneyBalLmt, " & _
                 "MBMCancerAnnLmt, MBMCancerBalLmt, MBMHomeNCAnnLmt, MBMHomeNCBalLmt) " & _
                 "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & tmpKidneyAnnLmt & "," & tmpKidneyAnnLmt & _
                 "," & tmpCancerAnnLmt & "," & tmpCancerAnnLmt & "," & tmpHomeNursAnnLmt & "," & tmpHomeNursAnnLmt & ")"
    medixmasterconn.Execute TmpBnfLmtSql

End Sub

Private Sub showClinicalInfo()
Dim tmpMBMNo As String
Dim MBMClnOth As New ADODB.Recordset
Dim ClnOthSql As String

    tmpMBMNo = txtMBMPrevMEMNo
    ClnOthSql = "SELECT PLNType, INSType, MBMPolEffDate, MBMPolExpDate, MBMCLNWarranty, MBMRemarks " & _
              " From MBMCLNOthers where MBMNumber='" & Trim(tmpMBMNo) & "'"
    MBMClnOth.Open ClnOthSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMClnOth.recordCount Then
        frmMBMClinical.cboCisPln.Text = MBMClnOth!PLNType
        frmMBMClinical.cboCisINS.Text = MBMClnOth!InsType
        frmMBMClinical.txtCLNEffDate = MBMClnOth!MBMPolEffDate
        frmMBMClinical.txtClnExpDate = MBMClnOth!MBMPolExpDate
        frmMBMClinical.txtWarranty = IIf(Len(MBMClnOth!MBMCLNWarranty), MBMClnOth!MBMCLNWarranty, "")
        frmMBMClinical.txtRemarks = IIf(Len(MBMClnOth!MBMRemarks), MBMClnOth!MBMRemarks, "")
    End If
    MBMClnOth.Close
    Set MBMClnOth = Nothing
End Sub

Private Sub PrincipalLifeTimeChecking()
Dim tmpSuppBNFStatus  As String
Dim TmpSql2 As String
Dim TmpLMBM As New ADODB.Recordset
Dim MBMAvailLmt As Double
Dim tmpLmt As Double
    If Trim(TempLifeTimeStatus) = "Y" Then
        TmpSql2 = "Select top 1 MBMNumber, MBMTopupNumber, MBMAvailableLimit From MBM where MBMPolicyNo='" & Trim(txtMBMPolicyNo) & _
                  "' And MBMStatus = 'A' AND PayCode ='" & Trim(txtPAYCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                  "' ORDER BY MBMPayorEffDate , MBMNumber "
        TmpLMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        TmpTopupNo = MembershipNo
        PrevLifeMBMNo = MembershipNo

        If TmpLMBM.EOF = False Then
            TmpTopupNo = IIf(Len(TmpLMBM!MBMTopupNumber), TmpLMBM!MBMTopupNumber, "")
            PrevLifeMBMNo = IIf(Len(TmpLMBM!MBMTopupNumber), TmpLMBM!MBMTopupNumber, "")
            If Mid(CboGuaRenewal, 1, 1) <> "3" Then
                MBMAvailLmt = IIf(IsNumeric(TmpLMBM!MBMAvailableLimit), TmpLMBM!MBMAvailableLimit, 0)
                txtAnnualLimit = MBMAvailLmt
            End If
        End If
        TmpLMBM.Close
        Set TmpLMBM = Nothing
        
        TmpSql2 = "Select MBMNumber, MBMLifeTimeLmt From MBMLifeTime where MBMNumber='" & Trim(PrevLifeMBMNo) & _
                  "' AND MBMCovID='00'"
        TmpLMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        If TmpLMBM.EOF Then
            If Mid(CboGuaRenewal, 1, 1) = "3" Then
                WLifeStatus = True
            End If
        End If
        TmpLMBM.Close
        Set TmpLMBM = Nothing
    End If
End Sub

Private Sub ComboMaintLoad(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, tmpComboBox As ComboBox, NoOfCol As String)
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
    

    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchMaintTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    tmpComboBox.Clear
    Do Until TmpRec.EOF
        With TmpRec
            If NoOfCol = "1" Then
                 tmpComboBox.AddItem !tmpCode
            ElseIf NoOfCol = "2" Then
                tmpComboBox.AddItem !tmpName
            Else
                tmpComboBox.AddItem !tmpCode & " - " & !tmpName
            End If
            TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
End Sub


Private Sub ComboHISDBLoad(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, tmpComboBox As ComboBox, NoOfCol As String)
Dim TmpRec As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
    
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "SearchHISTableRec"
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd1.Parameters.Append Prm2
    Set Prm3 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd1.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd1.Execute
    tmpComboBox.Clear
    Do Until TmpRec.EOF
        With TmpRec
            If NoOfCol = "1" Then
                 tmpComboBox.AddItem !tmpCode
            ElseIf NoOfCol = "2" Then
                tmpComboBox.AddItem !tmpName
            Else
                tmpComboBox.AddItem !tmpCode & " - " & !tmpName
            End If
            TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
End Sub

Private Sub PLNLostFocus()
Dim strsqlPLN, strsql As String
Dim PLN As New ADODB.Recordset
    TempLifeTimeStatus = ""
    tmpPLNProRate = "Y"
    RoundUpStatus = "Y"
    strsqlPLN = "select GRPCompany, PLNRoundUpStatus, PRMInd, MCOInd, AnnLmtIND, PlnAgeLimit, PLNProRate, PLNLifeTimeStatus,PLNPremGenderStatus " & _
                " from PLN where PLNCode = '" & Trim(cboPLNCode) & "' and HLTCode = '" & Trim(CboHLTCode) & "'"
    'If Trim(CboHLTCode) = "N" Then
    '    strsqlPLN = strsqlPLN & " and PAYCode='" & Trim(txtPAYCode) & "'"
    'End If
    PLN.MaxRecords = 1
    PLN.Open strsqlPLN, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    Do While Not PLN.EOF
        TempLifeTimeStatus = IIf(Len(PLN!PLNLifeTimeStatus), PLN!PLNLifeTimeStatus, "")
        tmpPLNProRate = IIf(Len(PLN!PLNProRate), PLN!PLNProRate, "Y")
        RoundUpStatus = IIf(Len(Trim(PLN!PLNRoundUpStatus)), Trim(PLN!PLNRoundUpStatus), "Y")
'        tPLNTopupStatus = IIf(Len(Trim(PLN!PLNTopUpStatus)), Trim(PLN!PLNTopUpStatus), "N")
        tPLNPrmIND = IIf(Len(Trim(PLN!PRMInd)), Trim(PLN!PRMInd), "N")
        tPLNMcoIND = IIf(Len(Trim(PLN!MCOInd)), Trim(PLN!MCOInd), "N")
        tPLNAnnLmtIND = IIf(Len(Trim(PLN!AnnLmtInd)), Trim(PLN!AnnLmtInd), "N")
        tPLNAgeLimit = IIf(Len(Trim(PLN!PlnAgeLimit)), Trim(PLN!PlnAgeLimit), "")
        txtMBMGroupCompany = IIf(Len(PLN!GrpCompany), Trim(PLN!GrpCompany), "")
        tmpPremSexStatus = IIf(Len(PLN!PLNPremGenderStatus), Trim(PLN!PLNPremGenderStatus), "")

        PLN.MoveNext
    Loop
    PLN.Close
    Set PLN = Nothing
    If Trim(txtMBMGroupCompany) = "" Then
        strsql = "Select PLNCode , GRPCompany from PLN where  0=0 "
        strsql = strsql & " AND PLNCode= '" & Trim(cboPLNCode) & "' AND grpcOMPANY IS NOT nuLL"  'need to reverify
        PLN.MaxRecords = 1
        PLN.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        If Not PLN.EOF Then
            txtMBMGroupCompany = IIf(Len(PLN!GrpCompany), Trim(PLN!GrpCompany), "")
        End If
        PLN.Close
        Set PLN = Nothing
    End If
    If Trim(cboPLNCode) <> "" Then
        If txtMBMPayorEffDate <> "__/__/____" Then
            Call CallLimit
            Call MemberCallFees
        End If
    End If
End Sub

Private Sub CheckIC(tmpICNo As String, tmPDupIC As Boolean, tmpMBMNo As String)
Dim strsql As String
Dim TmpRec As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim tmpSelFields As String
Dim tmpSelCriteria As String
Dim tmpTableName As String
    tmpSelFields = "Select TOP 1 MBMNumber "
    tmpTableName = "MBMCrossReference"
    tmpSelCriteria = " AND MBMICBCPP ='" & Trim(tmpICNo) & "'" & _
             " ORDER BY MBMPayorEffDate DESC, MBMNumber DESC"
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    Set cmd1.ActiveConnection = medixmasterconnMaint
    cmd1.CommandText = "SearchMaintTableRec"
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd1.Parameters.Append Prm2
    Set Prm3 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd1.Parameters.Append Prm3
    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Screen.MousePointer = vbHourglass
    Set TmpRec = cmd1.Execute

    Do While Not TmpRec.EOF
        tmPDupIC = True
        tmpMBMNo = IIf(Len(TmpRec!MBMNumber), TmpRec!MBMNumber, "")
        TmpRec.MoveNext
    Loop
    Screen.MousePointer = vbDefault
    TmpRec.Close
    Set TmpRec = Nothing
   ' medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
End Sub

Public Sub MemberCallFees()
Dim SelectAge As String
Dim tmpAge As String
Dim tmpVPlanCode As String
Dim tmpVINSCode As String
Dim tmpMBMDataType As String
Dim tmpVCovID As String

    tmpVPlanCode = cboPLNCode
    tmpVINSCode = cboINSCode
    tmpMBMDataType = "P"
    tmpVCovID = "00"
    If Trim(tmpVPlanCode) = "" Then
        tmpVPlanCode = Mid(frmMBMClinical.cboCisPln, 1, 4)
        tmpVINSCode = Mid(frmMBMClinical.cboCisINS, 1, 1)
    End If
    If IsDate(txtMBMPayorEffDate) And IsDate(txtMBMBirthdate) And Len(Trim(CboHLTCode)) _
        And Trim(tmpVPlanCode) <> "" And Trim(tmpVINSCode) <> "" Then
        '-- Get AgeCode
        SelectAge = GetAge(CDate(txtMBMPayorEffDate), CDate(txtMBMBirthdate))
        txtAGECode = GetAgeBand(Trim(CboHLTCode), SelectAge, tmpVPlanCode, tmpVINSCode)
        tmpAge = Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1))
        If Len(Trim(txtAGECode)) <> 0 Then
            txtMBMAdultChild = "A"
            If Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)) = "01" Then
                txtMBMAdultChild = "C"
            End If
        End If
    
        '-- End of AgeCode
        If tPLNPrmIND = "Y" Then
            txtMBMPremium = Format(GetHISPremium(tmpVINSCode, txtPAYCode, tmpAge, tmpVPlanCode, Trim(CboHLTCode), CDate(txtMBMPayorEffDate), tmpMBMDataType, tmpVCovID, tmpPremSexStatus, Mid(cboMBMSex, 1, 1)), "#,##0.00")
            txtBPrem = txtMBMPremium
            If PolDics > 0 Then
                txtMBMPremium = txtMBMPremium * (100 - PolDics) / 100
            End If
            If RenDisc > 0 Then
                txtMBMPremium = txtMBMPremium * (100 - RenDisc) / 100
            End If
            txtMBMPremium = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMBMPremium, "#,##0.00"))
            txtMBMPremium = Format(txtMBMPremium, "#,##0.00")
        End If
        txtBMco = 0
        txtMCOFees = 0
        If tPLNMcoIND = "Y" Then
            txtMCOFees = Format(GetHISMCO(tmpVINSCode, Trim(CboHLTCode), txtMBMAdultChild, CDate(txtMBMPayorEffDate), tmpVPlanCode, tmpMBMDataType, tmpVCovID), "#,##0.00")
            txtBMco = txtMCOFees
            If tmpPLNProRate <> "N" Then
                txtMCOFees = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtBMco, "#,##0.00"))
            End If
            txtMCOFees = Format(txtMCOFees, "#,##0.00")
            If RoundUpStatus <> "N" Then
                txtMCOFees = Round(txtMCOFees)
            End If
        End If
        txtAnnualLimit = 0
        If tPLNAnnLmtIND = "Y" Then
            txtAnnualLimit = 0
            txtAnnualLimit = GetHISAnnLmt(tmpVINSCode, tmpVPlanCode, Trim(CboHLTCode), CDate(txtMBMPayorEffDate), tmpMBMDataType, tmpVCovID)
            txtAnnualLimit = CDbl(txtAnnualLimit)
        End If
        If tmpMBMDataType = "P" Then
            txtAGECode = txtAGECode
            txtMBMAdultChild = txtMBMAdultChild
            txtMBMPremium = txtMBMPremium
            txtBPrem = txtBPrem
            txtMCOFees = txtMCOFees
            txtBMco = txtBMco
            txtAnnualLimit = txtAnnualLimit
        End If
    End If
    Call MemberCLNCallFees

End Sub

Private Sub EnableFrame()
    FrameMBM1.Enabled = False
    FrameMBM2.Enabled = False
    FrameMBM3.Enabled = False
    cmdShow.Enabled = False
    cmdSearch.Enabled = False
    cmdUpdate.Enabled = False
    cmdRemove.Enabled = False
    If Trim(CboHLTCode) <> "" And Trim(CboPayor) <> "" And IsDate(txtMBMBordxDate) = True And Trim(txtMBMBatchNo) <> "" Then
        FrameMBM1.Enabled = True
        FrameMBM2.Enabled = True
        FrameMBM3.Enabled = True
        cboMBMRenewal.SetFocus
    End If
End Sub

Private Sub ComboKeyPress(tmpComboBox As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(tmpComboBox.Text, tmpComboBox.SelStart) + Chr(KeyAscii) + Mid(tmpComboBox.Text, tmpComboBox.SelStart + tmpComboBox.SelLength + 1))
        ValidCount = 0
        ValidValue = ""
        For i = 0 To tmpComboBox.ListCount - 1
            If NewText = LCase(Left(tmpComboBox.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = tmpComboBox.List(i)
            End If
        Next
        
        If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            tmpComboBox.Text = ValidValue
            tmpComboBox.SelStart = 0
            tmpComboBox.SelLength = Len(ValidValue)
        End If
    End If
    
End Sub


