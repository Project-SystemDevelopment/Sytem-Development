VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdjustment 
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      TabIndex        =   119
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   121
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Member Adjustment"
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
         Left            =   240
         TabIndex        =   120
         Top             =   0
         Width           =   6615
      End
   End
   Begin VB.Frame Frame16 
      Height          =   615
      Left            =   120
      TabIndex        =   157
      Top             =   240
      Width           =   11655
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Go"
         Height          =   285
         Left            =   3840
         TabIndex        =   1
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdSearchMBM 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label76 
         Caption         =   "Mem No."
         Height          =   255
         Left            =   120
         TabIndex        =   158
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame Frame10 
      Height          =   615
      Left            =   120
      TabIndex        =   122
      Top             =   760
      Width           =   11655
      Begin VB.TextBox txtBatchNo 
         Height          =   285
         Left            =   10800
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "1"
         Top             =   255
         Width           =   375
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   3720
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1260
      End
      Begin VB.ComboBox cboEndorseStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmAdjustment.frx":0000
         Left            =   1080
         List            =   "frmAdjustment.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   255
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEndtBordxDate 
         Height          =   285
         Left            =   8760
         TabIndex        =   6
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label59 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   10080
         TabIndex        =   177
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Endt Bordx Date"
         Height          =   255
         Left            =   7560
         TabIndex        =   176
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lblAdj 
         Caption         =   "Endt. Eff. Date"
         Height          =   255
         Left            =   5160
         TabIndex        =   125
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label55 
         Caption         =   "Pol. Status"
         Height          =   255
         Left            =   2880
         TabIndex        =   124
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "Endt. Status"
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   270
         Width           =   1140
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5205
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9181
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pricipal Details"
      TabPicture(0)   =   "frmAdjustment.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Other Details"
      TabPicture(1)   =   "frmAdjustment.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Supplementary"
      TabPicture(2)   =   "frmAdjustment.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "lstCovAdd"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Adjustment History"
      TabPicture(3)   =   "frmAdjustment.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstAdjustmentHis"
      Tab(3).Control(1)=   "Combo3"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame6 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   110
         Top             =   480
         Width           =   10860
         Begin VB.TextBox txtAvailableLimit 
            Height          =   285
            Left            =   6120
            TabIndex        =   68
            Top             =   1200
            Width           =   2655
         End
         Begin VB.ComboBox cboCovRelationship 
            Height          =   315
            Left            =   6120
            TabIndex        =   62
            Text            =   "W -WIFE"
            Top             =   540
            Width           =   1095
         End
         Begin VB.ComboBox cboCovBlood 
            Height          =   315
            Left            =   7920
            TabIndex        =   63
            Text            =   "A"
            Top             =   550
            Width           =   855
         End
         Begin VB.TextBox txtCovAnnualLimit 
            Height          =   285
            Left            =   6120
            TabIndex        =   65
            Top             =   900
            Width           =   2655
         End
         Begin VB.Frame Frame9 
            Height          =   1815
            Left            =   9480
            TabIndex        =   111
            Top             =   240
            Width           =   1185
            Begin VB.CommandButton cmdClear 
               Caption         =   "Clear"
               Height          =   315
               Left            =   135
               TabIndex        =   74
               Top             =   1320
               Width           =   960
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   73
               Top             =   960
               Width           =   960
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   315
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   960
            End
            Begin VB.CommandButton cmdRemove 
               Caption         =   "&Remove"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   72
               Top             =   600
               Width           =   960
            End
         End
         Begin VB.ComboBox cboCovAdultChild 
            Height          =   315
            Left            =   2745
            TabIndex        =   67
            Text            =   "A"
            Top             =   1170
            Width           =   645
         End
         Begin VB.ComboBox cboCovSex 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1215
            TabIndex        =   66
            Text            =   "M"
            Top             =   1200
            Width           =   645
         End
         Begin VB.TextBox txtCovIcBcPp 
            Height          =   285
            Left            =   1215
            MaxLength       =   15
            TabIndex        =   61
            Top             =   540
            Width           =   2145
         End
         Begin VB.TextBox txtCovName 
            Height          =   330
            Left            =   1935
            MaxLength       =   50
            TabIndex        =   60
            Top             =   180
            Width           =   6780
         End
         Begin VB.ComboBox cboCovSalutation 
            Height          =   315
            Left            =   1215
            TabIndex        =   59
            Text            =   "MR"
            Top             =   180
            Width           =   690
         End
         Begin VB.TextBox txtCovAllergic 
            Height          =   690
            Left            =   120
            MaxLength       =   60
            TabIndex        =   69
            Top             =   1800
            Width           =   3975
         End
         Begin VB.TextBox txtCovExclusion 
            Height          =   705
            Left            =   4440
            MaxLength       =   500
            TabIndex        =   70
            Top             =   1800
            Width           =   4335
         End
         Begin MSMask.MaskEdBox txtCovBirthdate 
            Height          =   285
            Left            =   1200
            TabIndex        =   64
            Top             =   855
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   7080
            TabIndex        =   76
            Text            =   "Combo2"
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Available Annual Limit"
            Height          =   255
            Left            =   4440
            TabIndex        =   178
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label47 
            Caption         =   "Warranty Exclusion"
            Height          =   255
            Left            =   4440
            TabIndex        =   175
            Top             =   1560
            Width           =   1500
         End
         Begin VB.Label Label52 
            Caption         =   "Relationship"
            Height          =   255
            Left            =   4440
            TabIndex        =   174
            Top             =   585
            Width           =   1005
         End
         Begin VB.Label Label53 
            Caption         =   "Blood"
            Height          =   255
            Left            =   7320
            TabIndex        =   173
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label Label6 
            Caption         =   "Annual Limit"
            Height          =   255
            Left            =   4440
            TabIndex        =   172
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label51 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   1935
            TabIndex        =   117
            Top             =   1215
            Width           =   825
         End
         Begin VB.Label Label50 
            Caption         =   "Sex (M/F)"
            Height          =   255
            Left            =   90
            TabIndex        =   116
            Top             =   1215
            Width           =   825
         End
         Begin VB.Label Label49 
            Caption         =   "Birthdate"
            Height          =   255
            Left            =   90
            TabIndex        =   115
            Top             =   900
            Width           =   1005
         End
         Begin VB.Label Label65 
            Caption         =   "IC No."
            Height          =   255
            Left            =   90
            TabIndex        =   114
            Top             =   585
            Width           =   1125
         End
         Begin VB.Label Label66 
            Caption         =   "Mem Name"
            Height          =   255
            Left            =   90
            TabIndex        =   113
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label48 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   135
            TabIndex        =   112
            Top             =   1560
            Width           =   600
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4485
         Left            =   -74760
         TabIndex        =   95
         Top             =   480
         Width           =   10500
         Begin VB.TextBox txtMBMEmail 
            Height          =   315
            Left            =   1680
            TabIndex        =   49
            Top             =   1800
            Width           =   2865
         End
         Begin VB.ComboBox txtMBMWorkNature 
            Height          =   315
            Left            =   1680
            TabIndex        =   45
            Top             =   495
            Width           =   2865
         End
         Begin VB.TextBox txtMBMWeight 
            Height          =   315
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   50
            Top             =   2160
            Width           =   2865
         End
         Begin VB.ComboBox cboMBMOccupation 
            Height          =   315
            Left            =   1680
            TabIndex        =   44
            Top             =   195
            Width           =   2865
         End
         Begin VB.TextBox txtMBMTelNoH 
            Height          =   315
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   46
            Top             =   855
            Width           =   2865
         End
         Begin VB.ComboBox cboLGNCode 
            Height          =   315
            Left            =   1680
            TabIndex        =   52
            Top             =   2880
            Width           =   2865
         End
         Begin VB.ComboBox txtMBMClinical 
            Height          =   315
            Left            =   1620
            TabIndex        =   96
            Top             =   4665
            Width           =   1290
         End
         Begin VB.ComboBox cboMBMBlood 
            Height          =   315
            Left            =   1680
            TabIndex        =   53
            Top             =   3230
            Width           =   2865
         End
         Begin VB.ComboBox cboMBMSmoking 
            Height          =   315
            Left            =   1680
            TabIndex        =   54
            Top             =   3550
            Width           =   2865
         End
         Begin VB.ComboBox cboMBMAlcohol 
            Height          =   315
            Left            =   1680
            TabIndex        =   55
            Top             =   3900
            Width           =   2865
         End
         Begin VB.TextBox txtMBMAllergic 
            Height          =   1740
            Left            =   5520
            MaxLength       =   500
            TabIndex        =   57
            Top             =   2520
            Width           =   4500
         End
         Begin VB.TextBox txtMBMExclusion 
            Height          =   1380
            Left            =   5520
            MaxLength       =   500
            TabIndex        =   56
            Top             =   600
            Width           =   4380
         End
         Begin VB.TextBox txtMBMHeight 
            Height          =   315
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   51
            Top             =   2520
            Width           =   2865
         End
         Begin VB.TextBox txtMBMTelNoO 
            Height          =   315
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   48
            Top             =   1485
            Width           =   2865
         End
         Begin VB.TextBox txtMBMTelNoM 
            Height          =   315
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   47
            Top             =   1170
            Width           =   2865
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   6000
            TabIndex        =   58
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Email 
            Caption         =   "Email"
            Height          =   255
            Left            =   240
            TabIndex        =   167
            Top             =   1830
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Occupation"
            Height          =   255
            Left            =   240
            TabIndex        =   132
            Top             =   225
            Width           =   1005
         End
         Begin VB.Label Label37 
            Caption         =   "Blood Type"
            Height          =   255
            Left            =   240
            TabIndex        =   109
            Top             =   3240
            Width           =   1005
         End
         Begin VB.Label Label38 
            Caption         =   "Smoking (Y/N)"
            Height          =   255
            Left            =   240
            TabIndex        =   108
            Top             =   3600
            Width           =   1275
         End
         Begin VB.Label Label39 
            Caption         =   "Alcohol (Y/N)"
            Height          =   255
            Left            =   240
            TabIndex        =   107
            Top             =   3960
            Width           =   1275
         End
         Begin VB.Label Label23 
            Caption         =   "Work Nature"
            Height          =   255
            Left            =   240
            TabIndex        =   106
            Top             =   525
            Width           =   1140
         End
         Begin VB.Label Label46 
            Caption         =   "Clininal (Y/N)"
            Height          =   255
            Left            =   180
            TabIndex        =   105
            Top             =   4710
            Width           =   1455
         End
         Begin VB.Label Label41 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   5520
            TabIndex        =   104
            Top             =   2160
            Width           =   1590
         End
         Begin VB.Label Label40 
            Caption         =   "Warranty Exclusion"
            Height          =   255
            Left            =   5520
            TabIndex        =   103
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label Label36 
            Caption         =   "Pref. Language"
            Height          =   255
            Left            =   240
            TabIndex        =   102
            Top             =   2910
            Width           =   1140
         End
         Begin VB.Label Label35 
            Caption         =   "Height"
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   2550
            Width           =   1005
         End
         Begin VB.Label Label34 
            Caption         =   "Weight"
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   2190
            Width           =   1005
         End
         Begin VB.Label Label33 
            Caption         =   "Tel No. (O)"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   1530
            Width           =   1005
         End
         Begin VB.Label Label32 
            Caption         =   "H/P No."
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   1200
            Width           =   1005
         End
         Begin VB.Label Label28 
            Caption         =   "Tel No. (H)"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   885
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4650
         Left            =   240
         TabIndex        =   78
         Top             =   480
         Width           =   11325
         Begin VB.TextBox txtMBMDateJoin 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6960
            MaxLength       =   25
            TabIndex        =   35
            Top             =   1110
            Width           =   1200
         End
         Begin VB.TextBox txtMBMGroupCompany 
            Height          =   285
            Left            =   6960
            MaxLength       =   150
            TabIndex        =   37
            Top             =   1400
            Width           =   1200
         End
         Begin VB.ComboBox cboSRVCode 
            Height          =   315
            Left            =   6960
            Sorted          =   -1  'True
            TabIndex        =   39
            Text            =   "EA - EUROPE ASSISTANCE"
            Top             =   2040
            Width           =   2220
         End
         Begin VB.TextBox txtMBMAgentCode 
            Height          =   315
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   28
            Top             =   3120
            Width           =   2220
         End
         Begin VB.TextBox txtMBMEmployeeNo 
            Height          =   285
            Left            =   9360
            MaxLength       =   15
            TabIndex        =   38
            Top             =   1400
            Width           =   1275
         End
         Begin VB.TextBox txtMBMAdd3 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   17
            Top             =   795
            Width           =   3705
         End
         Begin VB.TextBox txtMBMAdd2 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   16
            Top             =   495
            Width           =   3705
         End
         Begin VB.TextBox txtMBMAdd1 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   15
            Top             =   160
            Width           =   3705
         End
         Begin VB.ComboBox cboPLNCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9360
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   145
            Width           =   1275
         End
         Begin VB.ComboBox cboINSCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6960
            Sorted          =   -1  'True
            TabIndex        =   29
            Text            =   "I"
            Top             =   145
            Width           =   1200
         End
         Begin VB.ComboBox cboMBMRelCode 
            Height          =   315
            Left            =   6960
            TabIndex        =   40
            Text            =   "P - COMBINED LIMIT"
            Top             =   2400
            Width           =   2220
         End
         Begin VB.ComboBox cboMBMTakeOver 
            Height          =   315
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1700
            Width           =   1455
         End
         Begin VB.ComboBox cboMBMMaritalStatus 
            Height          =   315
            Left            =   3600
            TabIndex        =   24
            Text            =   "M"
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtMBMPolicyYear 
            Enabled         =   0   'False
            Height          =   285
            Left            =   9360
            MaxLength       =   25
            TabIndex        =   36
            Top             =   1080
            Width           =   1275
         End
         Begin VB.ComboBox cboMBMSex 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   21
            Text            =   "M"
            Top             =   1700
            Width           =   885
         End
         Begin VB.ComboBox cboBRNCode 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   2760
            Width           =   2250
         End
         Begin VB.ComboBox cboCTYCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   19
            Top             =   1095
            Width           =   1425
         End
         Begin VB.ComboBox cboSTACode 
            Height          =   315
            Left            =   1320
            TabIndex        =   20
            Top             =   1400
            Width           =   2235
         End
         Begin VB.TextBox txtMBMPostCode 
            Height          =   285
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   18
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox cboRACCode 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   25
            Text            =   "C - CHINESE"
            Top             =   2400
            Width           =   1095
         End
         Begin VB.ComboBox cboNATCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   26
            Text            =   "MY - MALAYSIA"
            Top             =   2400
            Width           =   1455
         End
         Begin MSMask.MaskEdBox txtMBMBirthDate 
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
            Left            =   1320
            TabIndex        =   23
            Top             =   2040
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Frame Frame2 
            Height          =   1250
            Left            =   120
            TabIndex        =   79
            Top             =   3360
            Width           =   10905
            Begin VB.Line Line1 
               X1              =   120
               X2              =   11880
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Line Line6 
               X1              =   120
               X2              =   10920
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label txtMBMCaseCharges 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   9285
               TabIndex        =   155
               Top             =   885
               Width           =   975
            End
            Begin VB.Label Label62 
               Caption         =   "Case Charges"
               Height          =   255
               Left            =   8235
               TabIndex        =   156
               Top             =   900
               Width           =   1095
            End
            Begin VB.Label txtAvailableAnnualLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   2
               EndProperty
               Height          =   255
               Left            =   9285
               TabIndex        =   154
               Top             =   525
               Width           =   960
            End
            Begin VB.Label txtAnnualLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   2
               EndProperty
               Height          =   255
               Left            =   9285
               TabIndex        =   153
               Top             =   195
               Width           =   960
            End
            Begin VB.Label Label4 
               Caption         =   "Avail. Limit"
               Height          =   255
               Left            =   8280
               TabIndex        =   152
               Top             =   570
               Width           =   885
            End
            Begin VB.Label Label5 
               Caption         =   "Annual Limit"
               Height          =   255
               Left            =   8280
               TabIndex        =   151
               Top             =   240
               Width           =   870
            End
            Begin VB.Label txtMBMProviderRedCharges 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   6480
               TabIndex        =   147
               Top             =   555
               Width           =   945
            End
            Begin VB.Label txtMBMProviderEffCharges 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   6480
               TabIndex        =   149
               Top             =   885
               Width           =   945
            End
            Begin VB.Label Label77 
               Caption         =   "Acess Eff Fee"
               Height          =   255
               Left            =   5160
               TabIndex        =   150
               Top             =   885
               Width           =   1380
            End
            Begin VB.Label Label79 
               Caption         =   "Access Red Fee "
               Height          =   255
               Left            =   5115
               TabIndex        =   148
               Top             =   570
               Width           =   1500
            End
            Begin VB.Label Label61 
               Caption         =   "MCO Red"
               Height          =   255
               Left            =   2595
               TabIndex        =   146
               Top             =   525
               Width           =   735
            End
            Begin VB.Label txtMBMMCORed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3360
               TabIndex        =   145
               Top             =   525
               Width           =   945
            End
            Begin VB.Label txtMBMPremiumRed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   255
               Left            =   855
               TabIndex        =   143
               Top             =   525
               Width           =   1035
            End
            Begin VB.Label Label58 
               Caption         =   "Prem Red"
               Height          =   255
               Left            =   120
               TabIndex        =   144
               Top             =   630
               Width           =   900
            End
            Begin VB.Label txtMBMPremium 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   840
               TabIndex        =   86
               Top             =   180
               Width           =   1020
            End
            Begin VB.Label lblPremium 
               Caption         =   "Premium"
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   180
               Width           =   960
            End
            Begin VB.Label txtMBMProviderCharges 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   6480
               TabIndex        =   89
               Top             =   180
               Width           =   945
            End
            Begin VB.Label txtMCOFeesEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3360
               TabIndex        =   88
               Top             =   915
               Width           =   945
            End
            Begin VB.Label txtMBMPremiumEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   255
               Left            =   855
               TabIndex        =   87
               Top             =   915
               Width           =   1035
            End
            Begin VB.Label txtMCOFees 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3360
               TabIndex        =   85
               Top             =   180
               Width           =   945
            End
            Begin VB.Label Label83 
               Caption         =   "Prem Eff"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   930
               Width           =   900
            End
            Begin VB.Label Label85 
               Caption         =   "MCO Eff"
               Height          =   255
               Left            =   2595
               TabIndex        =   83
               Top             =   900
               Width           =   615
            End
            Begin VB.Label lblMCOFee 
               Caption         =   "MCO Fees"
               Height          =   255
               Left            =   2595
               TabIndex        =   81
               Top             =   180
               Width           =   870
            End
            Begin VB.Label Label11 
               Caption         =   "Access Fee"
               Height          =   255
               Left            =   5115
               TabIndex        =   80
               Top             =   240
               Width           =   1140
            End
         End
         Begin MSMask.MaskEdBox txtMBMBordxDate 
            Height          =   285
            Left            =   6960
            TabIndex        =   33
            Top             =   795
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMPayorExpDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   9360
            TabIndex        =   32
            Top             =   495
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
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
            Height          =   285
            Left            =   6960
            TabIndex        =   31
            Top             =   495
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label txtFirstPolNo 
            BackColor       =   &H80000009&
            Height          =   255
            Left            =   9360
            TabIndex        =   189
            Top             =   3105
            Width           =   1215
         End
         Begin VB.Label txtMBMPrevPolNo 
            BackColor       =   &H80000009&
            Height          =   255
            Left            =   6960
            TabIndex        =   188
            Top             =   3105
            Width           =   1215
         End
         Begin VB.Label txtFirstMeMNo 
            BackColor       =   &H80000009&
            Height          =   255
            Left            =   9360
            TabIndex        =   187
            Top             =   2775
            Width           =   1215
         End
         Begin VB.Label txtMBMPrevMEMNo 
            BackColor       =   &H80000009&
            Height          =   255
            Left            =   6960
            TabIndex        =   186
            Top             =   2775
            Width           =   1215
         End
         Begin VB.Label Label68 
            Caption         =   "Prev.Pol.No."
            Height          =   255
            Left            =   5400
            TabIndex        =   185
            Top             =   3105
            Width           =   1095
         End
         Begin VB.Label Label42 
            Caption         =   "Prev.Mem.No."
            Height          =   255
            Left            =   5400
            TabIndex        =   184
            Top             =   2775
            Width           =   1095
         End
         Begin VB.Label Label25 
            Caption         =   "First Mem.No."
            Height          =   255
            Left            =   8280
            TabIndex        =   183
            Top             =   2775
            Width           =   1095
         End
         Begin VB.Label Label26 
            Caption         =   "First.Pol.No."
            Height          =   255
            Left            =   8280
            TabIndex        =   182
            Top             =   3105
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Expiry Date"
            Height          =   255
            Left            =   8400
            TabIndex        =   181
            Top             =   510
            Width           =   915
         End
         Begin VB.Label Label15 
            Caption         =   "Effective Date"
            Height          =   255
            Left            =   5400
            TabIndex        =   180
            Top             =   510
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Batch No"
            Height          =   255
            Left            =   8400
            TabIndex        =   171
            Top             =   810
            Width           =   885
         End
         Begin VB.Label txtMBMBatchNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9360
            TabIndex        =   34
            Top             =   810
            Width           =   1275
         End
         Begin VB.Label Label17 
            Caption         =   "Bordx. Date"
            Height          =   255
            Left            =   5400
            TabIndex        =   170
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label Label22 
            Caption         =   "Serv. Provider"
            Height          =   255
            Left            =   5400
            TabIndex        =   169
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label Label14 
            Caption         =   "Agent Code"
            Height          =   255
            Left            =   240
            TabIndex        =   168
            Top             =   3150
            Width           =   1140
         End
         Begin VB.Label Label18 
            Caption         =   "Grp. Company"
            Height          =   255
            Left            =   5400
            TabIndex        =   166
            Top             =   1430
            Width           =   1545
         End
         Begin VB.Label Label19 
            Caption         =   "E'yee No."
            Height          =   255
            Left            =   8400
            TabIndex        =   165
            Top             =   1415
            Width           =   1545
         End
         Begin VB.Label Label10 
            Caption         =   "Date Join"
            Height          =   255
            Left            =   5400
            TabIndex        =   164
            Top             =   1125
            Width           =   780
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   8400
            TabIndex        =   163
            Top             =   175
            Width           =   825
         End
         Begin VB.Label Label20 
            Caption         =   "Ins Type"
            Height          =   255
            Left            =   5400
            TabIndex        =   162
            Top             =   175
            Width           =   960
         End
         Begin VB.Label Label29 
            Caption         =   "Age Band"
            Height          =   255
            Left            =   5400
            TabIndex        =   161
            Top             =   1710
            Width           =   735
         End
         Begin VB.Label txtAGECode 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6960
            TabIndex        =   160
            Top             =   1740
            Width           =   2220
         End
         Begin VB.Label Label30 
            Caption         =   "Pricipal Type"
            Height          =   255
            Left            =   5400
            TabIndex        =   134
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "Marital Status"
            Height          =   255
            Left            =   2520
            TabIndex        =   133
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label71 
            Caption         =   "Nationality"
            Height          =   255
            Left            =   2520
            TabIndex        =   131
            Top             =   2400
            Width           =   960
         End
         Begin VB.Label Label70 
            Caption         =   "Race"
            Height          =   255
            Left            =   240
            TabIndex        =   130
            Top             =   2400
            Width           =   420
         End
         Begin VB.Label Label12 
            Caption         =   "Branch"
            Height          =   255
            Left            =   240
            TabIndex        =   129
            Top             =   2760
            Width           =   960
         End
         Begin VB.Label Label63 
            Caption         =   "TakeOver"
            Height          =   255
            Left            =   2520
            TabIndex        =   128
            Top             =   1725
            Width           =   825
         End
         Begin VB.Label Label9 
            Caption         =   "Pol. Year"
            Height          =   255
            Left            =   8400
            TabIndex        =   127
            Top             =   1125
            Width           =   690
         End
         Begin VB.Label Label56 
            Caption         =   "Sex "
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   1730
            Width           =   825
         End
         Begin VB.Label Label54 
            Caption         =   "Address"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label57 
            Caption         =   "City/Town"
            Height          =   255
            Left            =   2520
            TabIndex        =   93
            Top             =   1125
            Width           =   900
         End
         Begin VB.Label Label60 
            Caption         =   "State"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   1430
            Width           =   465
         End
         Begin VB.Label Label69 
            Caption         =   "Birthdate"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label Label72 
            Caption         =   "PostCode"
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   1170
            Width           =   1140
         End
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   1665
         Left            =   -74760
         TabIndex        =   75
         Top             =   3240
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   2937
         SortKey         =   10
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Salutation"
            Object.Width           =   0
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
            Text            =   "Blood Type"
            Object.Width           =   1411
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
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Status"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Annual Limit"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstAdjustmentHis 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   77
         Top             =   480
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   8070
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Endt Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Batch no"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Adjustment Date"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Endt Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Endt Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Adjusted By"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -67440
         TabIndex        =   190
         Text            =   "Combo3"
         Top             =   1680
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   120
      TabIndex        =   136
      Top             =   1320
      Width           =   11655
      Begin VB.TextBox txtMBMName 
         Height          =   330
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   9
         Top             =   165
         Width           =   3555
      End
      Begin VB.TextBox txtMBMPolNo 
         Height          =   285
         Left            =   3060
         MaxLength       =   25
         TabIndex        =   12
         Top             =   555
         Width           =   2415
      End
      Begin VB.ComboBox cboPAYCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   11
         Top             =   480
         Width           =   675
      End
      Begin VB.ComboBox cboHLTCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7920
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   560
         Width           =   1425
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   285
         Left            =   7920
         MaxLength       =   15
         TabIndex        =   10
         Top             =   165
         Width           =   2175
      End
      Begin VB.ComboBox cboMBMSalutation 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Text            =   "MR"
         Top             =   165
         Width           =   690
      End
      Begin VB.Label Label74 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   159
         Top             =   510
         Width           =   915
      End
      Begin VB.Label Label75 
         Caption         =   "Health Type"
         Height          =   255
         Left            =   6840
         TabIndex        =   142
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label64 
         Caption         =   "IC. No"
         Height          =   255
         Left            =   6840
         TabIndex        =   141
         Top             =   165
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "Mem Name"
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   195
         Width           =   1140
      End
      Begin VB.Label Label73 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   1920
         TabIndex        =   137
         Top             =   570
         Width           =   780
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Status"
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
         Height          =   255
         Left            =   5640
         TabIndex        =   179
         Top             =   200
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      Height          =   645
      Left            =   120
      TabIndex        =   118
      Top             =   7560
      Width           =   11295
      Begin VB.TextBox txtLastEdtDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   140
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtLastEndStatus 
         Height          =   285
         Left            =   1680
         TabIndex        =   139
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   840
         TabIndex        =   135
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   315
         Left            =   9660
         TabIndex        =   43
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8715
         TabIndex        =   42
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdEditMBM 
         Caption         =   "Edit"
         Height          =   315
         Left            =   7680
         TabIndex        =   41
         Top             =   240
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim startTrans As Boolean
Dim counter As Integer
Dim bExit As Boolean
Dim bTabDone As Boolean



Private Sub Form_Load()
    cboEndorseStatus.Enabled = False
    cboStatus.Enabled = False
    txtDate.Enabled = False
    Frame10.Enabled = False
    Call ComboListing
    Call EnableEntries(True)
    cmdSave.Enabled = False
    'Call TabEnabled(False)
    counter = 0
    cmdRemove.Enabled = False
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    bExit = False
    bTabDone = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim RecordExit
    If bExit = False Then
        RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordExit = vbYes Then
            Unload Me
        Else
            Cancel = True
        End If

    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = SSTab1.Tab
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = 2 Then
            'last tab so we need to wrap to tab 1
            SSTab1.Tab = 3
        ElseIf i = 1 Then
            'last tab so we need to wrap to tab 1
            SSTab1.Tab = 2
        ElseIf i = 0 Then
            'increment the tab
            SSTab1.Tab = 1
        Else
            SSTab1.Tab = 0
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 3 Then
            'last tab so we need to wrap to tab 1
            SSTab1.Tab = 2
        
        ElseIf i = 2 Then
            'last tab so we need to wrap to tab 1
            SSTab1.Tab = 1
        ElseIf i = 1 Then
            'increment the tab
            SSTab1.Tab = 0
        Else
            SSTab1.Tab = 1
        End If
    End If
'    SSTab1.SetFocus
End Sub

Private Sub ComboListing()
    Dim SRV As New ADODB.Recordset
    Dim Work As New ADODB.Recordset
    Dim RAC As New ADODB.Recordset
    Dim NAT As New ADODB.Recordset
    Dim REL As New ADODB.Recordset
    Dim LGN As New ADODB.Recordset
    Dim CTY As New ADODB.Recordset
    Dim STA As New ADODB.Recordset
    Dim BRN As New ADODB.Recordset
    Dim PAY As New ADODB.Recordset
    Dim OCP As New ADODB.Recordset
    Dim HLT As New ADODB.Recordset
    Dim PLN As New ADODB.Recordset
    Dim INS As New ADODB.Recordset
    Dim sqlPAY As String
    
    SRV.Open "select SRVCode, SRVName from  SRV", medixmasterconn, adOpenKeyset, adLockOptimistic
    RAC.Open "select RACCode, RACName from   RAC", medixmasterconn, adOpenKeyset, adLockOptimistic
    NAT.Open "select NATCode, NATName from  NAT", medixmasterconn, adOpenKeyset, adLockOptimistic
    REL.Open "select RELCode, RELDescription from  REL", medixmasterconn, adOpenKeyset, adLockOptimistic
    OCP.Open "select OCPCode, OCPDescription from OCP", medixmasterconn, adOpenKeyset, adLockOptimistic
    LGN.Open "select LGNCode, LGNDescription from LGN", medixmasterconn, adOpenKeyset, adLockOptimistic
    CTY.Open "select ctyCode , CTYDescription from CTY", medixmasterconn, adOpenKeyset, adLockOptimistic
    STA.Open "select STACOde, STADescription from  STA", medixmasterconn, adOpenKeyset, adLockOptimistic
    BRN.Open "Select BRNCOde, BRNNAme from BRN", medixmasterconn, adOpenKeyset, adLockOptimistic
    Work.Open "Select  WorkID, WorkDescription from WorkNature ", medixmasterconn, adOpenKeyset, adLockOptimistic
    sqlPAY = "select PAYCode, PAYCompanyName from PAY "
    PAY.Open sqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    HLT.Open "select HLTCode, HLTDescription from HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    PLN.Open "select PLNCode, PLNDescription from PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
    INS.Open "select INSCode, INSDescription from INS", medixmasterconn, adOpenKeyset, adLockOptimistic

    Do Until PAY.EOF
        cboPAYCode.AddItem ChkStr(PAY!PAYCode) '& " - " & ChkStr(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    Do Until LGN.EOF
        cboLGNCode.AddItem ChkStr(LGN!LGNCode) '& " - " & LGN!LGNDescription
        LGN.MoveNext
    Loop

    Do Until SRV.EOF
        cboSRVCode.AddItem ChkStr(SRV!SRVCode) '& " - " & ChkStr(SRV!SRVname)
        SRV.MoveNext
    Loop

    Do Until RAC.EOF
        cboRACCode.AddItem ChkStr(RAC!RACCode) '& " - " & ChkStr(RAC!RACName)
        RAC.MoveNext
    Loop

    Do Until NAT.EOF
        cboNATCode.AddItem ChkStr(NAT!NATCode) '& " - " & ChkStr(NAT!NATName)
        NAT.MoveNext
    Loop

    Do Until REL.EOF
        cboCovRelationship.AddItem ChkStr(REL!RELCODE) '& " - " & REL!RELdescription
        REL.MoveNext
    Loop

    Do Until OCP.EOF
        cboMBMOccupation.AddItem ChkStr(OCP!OCPCODE) '& " - " & OCP!OCPdescription
        OCP.MoveNext
    Loop

    Do Until Work.EOF
        txtMBMWorkNature.AddItem Work!WorkID '& " - " & Work!WorkDescription
        Work.MoveNext
    Loop
    
    Do Until CTY.EOF
        cboCTYCode.AddItem CTY!CTYDescription
        CTY.MoveNext
    Loop

    Do Until STA.EOF
        cboSTACode.AddItem STA!STADescription
        STA.MoveNext
    Loop

    Do Until BRN.EOF
        cboBRNCode.AddItem ChkStr(BRN!BRNName)
        BRN.MoveNext
    Loop

    Do Until HLT.EOF
        cboHLTCode.AddItem HLT!HLTCode '& " - " & HLT!HLTDescription
        HLT.MoveNext
    Loop

    Do Until PLN.EOF
        cboPLNCode.AddItem PLN!PLNCode '& " - " & PLN!PLNDescription
        PLN.MoveNext
    Loop
    
    Do Until INS.EOF
        cboINSCode.AddItem INS!INSCode '& " - " & INS!INSDescription
        INS.MoveNext
    Loop

    cboMBMRelCode.AddItem "P - Shared Limit"
    cboMBMRelCode.AddItem "Q - Own Limit"


    cboMBMSex.AddItem "M"
    cboMBMSex.AddItem "F"
    
    cboCovSex.AddItem "M"
    cboCovSex.AddItem "F"
    
    cboMBMMaritalStatus.AddItem "S"
    cboMBMMaritalStatus.AddItem "M"
    
    cboMBMBlood.AddItem "A"
    cboMBMBlood.AddItem "B"
    cboMBMBlood.AddItem "AB"
    cboMBMBlood.AddItem "O"
    
    
    cboCovBlood.AddItem "A"
    cboCovBlood.AddItem "B"
    cboCovBlood.AddItem "AB"
    cboCovBlood.AddItem "O"
    
    cboCovAdultChild.AddItem "A"
    cboCovAdultChild.AddItem "C"
    
    cboMBMSalutation.AddItem "MR"
    cboMBMSalutation.AddItem "MS"
    cboMBMSalutation.AddItem "MRS"
    
    cboCovSalutation.AddItem "MR"
    cboCovSalutation.AddItem "MS"
    cboCovSalutation.AddItem "MRS"
    
    cboMBMSmoking.AddItem "Y"
    cboMBMSmoking.AddItem "N"
    
    cboEndorseStatus.AddItem "A" & " - " & "Amendment"
    cboEndorseStatus.AddItem "C" & " - " & "Cancellation"
    cboEndorseStatus.AddItem "S" & " - " & "Add Supplementary"
    cboEndorseStatus.AddItem "T" & " - " & "Delete Supplementary"
    cboEndorseStatus.AddItem "X" & " - " & "Suspend Membership"
    cboEndorseStatus.AddItem "Y" & " - " & "Re-Activate Membership"
    cboEndorseStatus.AddItem "I" & " - " & "Insert New Member"
    cboEndorseStatus.AddItem "D" & " - " & "Delete Policy"
    
    
    cboStatus.AddItem "A" & " - " & "Active"
    cboStatus.AddItem "C" & " - " & "Cancel"
    cboStatus.AddItem "D" & " - " & "Delete"
    cboStatus.AddItem "S" & " - " & "Suspend"
            
            
    cboMBMTakeOver.AddItem "N - None"
    cboMBMTakeOver.AddItem "T - Take Over"
    cboMBMTakeOver.AddItem "C - Conversion"
        
  
    cboMBMAlcohol.AddItem "Y"
    cboMBMAlcohol.AddItem "N"

    PAY.Close
    Set PAY = Nothing
    SRV.Close
    Set SRV = Nothing
    OCP.Close
    Set OCP = Nothing
    RAC.Close
    Set RAC = Nothing
    NAT.Close
    Set NAT = Nothing
    REL.Close
    Set REL = Nothing
    LGN.Close
    Set LGN = Nothing
    CTY.Close
    Set CTY = Nothing
    STA.Close
    Set STA = Nothing
    BRN.Close
    Set BRN = Nothing
    HLT.Close
    Set HLT = Nothing
    INS.Close
    Set INS = Nothing
    PLN.Close
    Set PLN = Nothing
End Sub

'Private Sub BNFListing()
 ''   Dim rst2 As New ADODB.Recordset
  ''  Dim BNF As ListItem
   ' Dim sql As String
'    If lstBenefitDetails.ListItems.Count = 0 And Len(Trim(cboHLTCode)) _
 '    And Len(Trim(cboINSCode)) And Len(Trim(cboPLNCode)) Then
      '  sql = " select BNFCode, " & _
  '        " BNFclaimableAmount, BNFdescription" & _
   '       " from BNF where PLNCOde ='" & Trim(Mid(cboPLNCode, 1, InStr(1, cboPLNCode, "-") - 1)) & "'" & _
    '      " AND HLTCode ='" & Trim(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1)) & "'" & _
     '     " AND INSCODe ='" & Trim(Mid(cboINSCode, 1, InStr(1, cboINSCode, "-") - 1)) & "'"
'        rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
 '       Do Until rst2.EOF
  '          Set BNF = lstBenefitDetails.ListItems.Add(, , rst2!BNFCode)
   '             BNF.SubItems(1) = rst2!BNFDescription
    '            BNF.SubItems(2) = rst2!BNFClaimAbleAmount
 '               rst2.MoveNext
     '   Loop
      '
       ' rst2.Close
'        Set rst2 = Nothing
 '   End If
'End Sub



Private Sub ADJListing()
    Dim ADJ As New ADODB.Recordset
    Dim ADJList As ListItem
    Dim sql As String
    
    If Len(Trim(txtMBMNumber)) Then
         sql = " select MBMAmendID , MBMNumber,  MBMAmendUser , " & _
           " MBMNAme, MBMpolicyNo, MBMICBCPP, MBMAmendEdtType , MBMAmendDate , MBMAENDtEffDate ,   " & _
           " MBMAEndtBordxDate, MBMAEndtBatchNo  " & _
           " from MBMAmendHistory where MBMnumber ='" & Trim(txtMBMNumber) & "'"
         ADJ.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
         Do Until ADJ.EOF
         If Len(ADJ!MBMAEndtBordxDate) Then
             Set ADJList = lstAdjustmentHis.listitems.Add(, , ADJ!MBMAEndtBordxDate)
                 'ADJList.SubItems(1) = ADJ!MBMBatchNo
                 'ADJList.SubItems(2) = ADJ!MBMLastAdjustmentDate
                 'ADJList.SubItems(3) = ADJ!MBMLastEndorseStatus
        Else
        Set ADJList = lstAdjustmentHis.listitems.Add(, , " ")
        End If
                 If Len(Trim(ADJ!MBMAEndtBatchNo)) Then
                    ADJList.SubItems(1) = ADJ!MBMAEndtBatchNo
                 End If
                 If Len(Trim(ADJ!MBMAmendDate)) Then
                    ADJList.SubItems(2) = ADJ!MBMAmendDate
                 End If

                If Len(ADJ!MBMAmendEdtType) Then
                    ADJList.SubItems(3) = ADJ!MBMAmendEdtType
                End If
                
                If Len(ADJ!MBMAEndtEffDate) Then
                    ADJList.SubItems(4) = ADJ!MBMAEndtEffDate
                End If

                 If Len(Trim(ADJ!MBMAmendUser)) Then
                    ADJList.SubItems(5) = ADJ!MBMAmendUser
                 End If
                
                 ADJ.MoveNext
              '   End If
         Loop
         ADJ.Close
         Set ADJ = Nothing
    End If
End Sub

' ######## List out all the cases
'Private Sub ListCases()
 '   Dim CASList As ListItem
  '  Dim strsql As String
   ' Dim CAS As New ADODB.Recordset
'
 '   strsql = "SELECT CASNumber, MBMNumber,  PAYCode, " & _
 '       " MBMPatientCovID, MBMPolicyNo, MBMName, " & _
  '      " INSCode, PLNCode, CASStatus, " & _
   '     " CASRemark, CaseDateOfLoss ,     MBMPayorEffDate, HLTCode, " & _
        " CASHOSCode, CASDateAdmitted, CASDischargeDate , MBMPayorExpDate , " & _
    '    "MBMAvailableLimit, CASReserveAmount"
     '   strsql = strsql & " from VIEW_Cas_MBM_Claims where 0=0"
      '  strsql = strsql & " AND  MBMNumber = '" & RemStr(txtMBMNumber) & "' "
            
'    CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
 '
    'If CAS.recordCount <> 0 Then
  '      Screen.MousePointer = vbHourglass
   '     lstCaseList.ListItems.Clear
    '    Do Until CAS.EOF
     '       Set CASList = lstCaseList.ListItems.Add(, , CAS!CASNumber)
      '          CASList.SubItems(1) = CAS!MBMNumber
       '         CASList.SubItems(2) = CAS!MBMName
        '        CASList.SubItems(3) = CAS!CaseDateOfLoss
'                CASList.SubItems(4) = CAS!CASStatus
 '               CASList.SubItems(5) = CAS!MBMPatientCovID
  '              CASList.SubItems(6) = CAS!MBMPolicyNo
   '             CASList.SubItems(7) = CAS!INSCode
           '     CASList.SubItems(8) = CAS!PLNCode
    '            If Trim(CAS!CASRemark) <> "" Then
     '               CASList.SubItems(9) = CAS!CASRemark
      '          End If
       '         If Len(Trim(CAS!CASDateAdmitted)) Then
        '            CASList.SubItems(10) = CAS!CASDateAdmitted
         '       End If
'          '      If Len(Trim(CAS!CASDischargeDate)) Then
 '                   CASList.SubItems(11) = CAS!CASDischargeDate
  '              End If
   '
             '   CASList.SubItems(12) = CAS!MBMPayorEffDate
    '            CASList.SubItems(13) = CAS!MBMPayorExpDate
     '           If Len(Trim(CAS!CASHOSCode)) Then
      '              CASList.SubItems(14) = CAS!CASHOSCode
       '         End If
        '
         '       CASList.SubItems(15) = CAS!HLTCode
          '      CASList.SubItems(16) = CAS!MBMAvailableLimit
           '     CASList.SubItems(17) = CAS!CASReserveAmount
              '  CAS.MoveNext
            '
        'Loop
'        S'creen.MousePointer = vbDefault
 '   End If
  '  CAS.Close
   ' Set CAS = Nothing
'End Sub


' ####################### Calculation's Function ###############
Private Sub MemberCallAge()
    Dim SelectAge As String
    If Trim(txtMBMPayorEffDate) <> "__/__/____" _
        And Len(Trim(txtMBMBirthDate)) <> 0 _
        And Trim(txtMBMBirthDate) <> "__/__/____" _
        And Len(Trim(cboHLTCode)) <> 0 Then
        If IsDate(txtMBMBirthDate) Then
            SelectAge = GetAge(txtMBMPayorEffDate, txtMBMBirthDate)
            txtAGECode = GetAgeBand(Trim(Mid(cboHLTCode, 1, 1)), SelectAge)
            If Len(Trim(txtAGECode)) = 0 Then
                MsgBox "Please Enter A Valid Date OR Age Out of Acceptable Range", vbOKOnly + vbCritical, DialogBoxTitle
                txtMBMPayorEffDate.SetFocus
            Else
                If Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)) = "01" Then
                    txtMBMAdultChild = "C"
                Else
                    txtMBMAdultChild = "A"
                End If
            End If
        End If
    End If
End Sub

Public Sub MemberCallFees()
    Dim tempPremium, tempMCOFee As Single
    Dim confirm1, confirm2
'        txtMBMPremium = ""
 '       txtMCOFees = ""
    tempPremium = Format(GetPremium(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboPAYCode, 1, 2)), Trim(Mid(txtAGECode, 1, 2)), Trim(Mid(cboPLNCode, 1, 2)), Trim(Mid(cboHLTCode, 1, 1)), txtMBMPayorEffDate), "#,##0.00")
    tempMCOFee = Format(GetMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboHLTCode, 1, 1)), txtMBMAdultChild, txtMBMPayorEffDate), "#,##0.00")
    If Format(txtMBMPremium, "#,##0.00") <> tempPremium Then
        confirm1 = MsgBox("Do you want to change member's premium amount?", vbCritical + vbYesNo)
        If confirm1 = vbYes Then
           txtMBMPremium = tempPremium
        End If
    End If
    If Format(txtMCOFees, "#,##0.00") <> tempMCOFee Then
        confirm2 = MsgBox(" Do you want to change member's MCO Fee?", vbCritical + vbYesNo)
        If confirm2 = vbYes Then
            txtMCOFees = tempMCOFee
        End If
    End If
End Sub

Public Sub ProviderCallFees()
    Dim GetProString As String
    
    If Trim(cboINSCode) <> "" And Trim(cboSRVCode) <> "" And Trim(cboHLTCode) <> "" Then
        txtMBMProviderCharges = ""
        'GetProString = GetProviderCharges(Trim(Mid(cboSRVCode, 1, InStr(1, cboSRVCode, "-") - 1)), Trim(Mid(cboINSCode, 1, InStr(1, cboINSCode, "-") - 1)), Trim(Mid(txtHLTCode, 1, InStr(1, txtHLTCode, "-") - 1)), txtMBMAdultChild)
        If InStr(1, cboSRVCode, "-") <> 0 Then
            GetProString = GetProviderCharges(Trim(Mid(cboSRVCode, 1, InStr(1, cboSRVCode, "-") - 1)), Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild)
        Else
            GetProString = GetProviderCharges(Trim(cboSRVCode), Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild)
        End If
        
        If InStr(1, GetProString, "-") - 1 > 0 Then
            txtMBMProviderCharges = Format(Mid(GetProString, 1, InStr(1, GetProString, "-") - 1), "#,##0.00")
        End If
        'txtMBMCaseCharges = Format(Mid(GetProString, InStr(1, GetProString, "-") + 1), "#,##0.00")
    Else
        txtMBMProviderCharges = 0
     '   txtMBMCaseCharges = 0
    End If
            
End Sub

Public Sub CallLimit()
    txtAnnualLimit = ""
    If Len(Trim(cboPLNCode)) Then
        txtAnnualLimit = Format(GetAnnualLimit(Mid(cboINSCode, 1, 1), Mid(cboPLNCode, 1, 2), Mid(cboHLTCode, 1, 1)), "#,##0.00")
    End If
End Sub


' ######################### END CALCULATIOn###########
' ######################### START CHECKING ###########

Private Sub cboPLNCode_Click()
    Call MemberCallFees
    CallLimit
    Call ProviderCallFees
End Sub

Private Sub cboEndorseStatus_Click()
    If cboEndorseStatus.ListIndex = 0 Then
        If Mid(cboStatus, 1, 1) = "C" Or Mid(cboStatus, 1, 1) = "D" Then
            cboStatus.Enabled = False
        'txtMBMEndorseStatus.Enabled = True
        Else
            'lblAdj.Caption = ""
            'lblAdj.Caption = "Adjustment Date"
            'Call EnableEntries(True)
            Call EnableSupplementary(True)
            'Frame1.Enabled = True
            lstCovAdd.Enabled = True
            txtDate.Mask = ""
            txtDate = Date
            Call TabEnabled(True)
            cmdAdd.Enabled = False
            cmdRemove.Enabled = False
            cmdUpdate.Enabled = True
            cmdClear.Enabled = True
            cboStatus.Enabled = True
        End If
        
    ElseIf cboEndorseStatus.ListIndex = 1 Then
            'lblAdj.Caption = ""
            'lblAdj.Caption = "Cancelation Date"
            'Call EnableEntries(False)
            'Call EnableSupplementary(False)
            cmdExit.Enabled = True
            cmdAdd.Enabled = False
            cmdRemove.Enabled = False
            lstCovAdd.Enabled = False
            txtDate.Mask = ""
            txtDate = Date
            txtDate.Enabled = True
            'txtDate.SetFocus
            'If txtMBMNumber <> "" Then
            '    Frame1.Enabled = False
                'Frame4.Enabled = False
            'End If
            cmdSearchMBM.Enabled = True
            'cmdMBMPolicy.Enabled = True
            'cmdSearchIC.Enabled = True
            'cmdSearchName.Enabled = True
            cmdEditMBM.Enabled = True
    ElseIf cboEndorseStatus.ListIndex = 2 Then
        If Mid(cboStatus, 1, 1) = "C" Or Mid(cboStatus, 1, 1) = "D" Then
            cboStatus.Enabled = False
        ElseIf Mid(cboINSCode, 1, 1) = "I" Or Mid(cboINSCode, 1, 1) = "G" Then
            SSTab1.Tab = 0
        'End If
        Else
            'lblAdj.Caption = ""
            'lblAdj.Caption = "Deletion Date"
            'Call EnableEntries(False)
            Call EnableSupplementary(True)
            cmdExit.Enabled = True
            'cmdEditMBM.Enabled = False
            cmdClear.Enabled = True
            cmdAdd.Enabled = True
            cmdRemove.Enabled = False
            lstCovAdd.Enabled = True
            cboEndorseStatus.Enabled = False
            'Call TabEnabled(True)
            SSTab1.Tab = 2
            txtDate = Date
            txtDate.Enabled = True
            txtDate.SetFocus
            cmdEditMBM.Enabled = True
        End If
    ElseIf cboEndorseStatus.ListIndex = 4 Then
        If Mid(cboStatus, 1, 1) = "C" Or Mid(cboStatus, 1, 1) = "D" Then
            cboStatus.Enabled = False
        Else
            'lblAdj.Caption = ""
            'lblAdj.Caption = "Suspension Date"
            Call EnableEntries(False)
            Call EnableSupplementary(True)
            cmdExit.Enabled = True
            cmdClear.Enabled = True
            'cmdEditMBM.Enabled = True
            cmdAdd.Enabled = False
            cmdRemove.Enabled = True
            lstCovAdd.Enabled = True
            'Call TabEnabled(False)
            'SSTab1.Tab = 2
            txtDate = Date
            txtDate.Enabled = True
            'txtDate.SetFocus
            cmdEditMBM.Enabled = True
        End If
        
    ElseIf cboEndorseStatus.ListIndex = 3 Then
        'Call EnableSupplementary(False)
        'lblAdj.Caption = ""
        'lblAdj.Caption = "Adjustment Date"
        lstCovAdd.Enabled = True
        cmdRemove.Enabled = True
        cmdClear.Enabled = True
        SSTab1.Tab = 2
        
    ElseIf cboEndorseStatus.ListIndex = 6 Then
        'lblAdj.Caption = ""
        'lblAdj.Caption = "Insertion Date"
         Call ClearEntries
         cboINSCode.Enabled = True
         cboPLNCode.Enabled = True
         txtMBMBordxDate.Enabled = True
        lstCovAdd.Enabled = True
        cmdRemove.Enabled = True
        cmdClear.Enabled = True
        Frame1.Enabled = True
        Frame4.Enabled = True
        cboStatus.ListIndex = 0
        txtDate.Enabled = True
        txtDate.SetFocus
    ElseIf Mid(cboEndorseStatus, 1, 1) = "D" Then
        Call DeletePrincipals
    Else ' Reactivate
        'lblAdj.Caption = ""
        'lblAdj.Caption = "Reactivation Date"
        Call TabEnabled(True)
        txtDate = Date
        txtDate.Enabled = True
        'txtDate.SetFocus
        cmdEditMBM.Enabled = True
    End If
    cboEndorseStatus.Enabled = False
    cmdSave.Enabled = True
End Sub

Private Sub cboINSCode_LostFocus()
    'If Len(Trim(cboinsCode.Text)) Then
    '    Call LoadPlanType(cboinsCode.Text)
    'End If
End Sub

Private Sub cboINSCode_Click()
    If Len(Trim(cboINSCode)) Then
        If Trim(cboINSCode) = "F" Or Trim(cboINSCode) = "H" Then
            If Trim(cboINSCode) <> Trim(cboINSCode) Then
                Call EnableSupplementary(True)
            Else
                Call EnableSupplementary(False)
            End If
            cboMBMMaritalStatus = "M"
        ElseIf Trim(cboINSCode) = "I" Or Trim(cboINSCode) = "G" Then
            Call EnableSupplementary(False)
            cboMBMMaritalStatus = "S"
        End If
    End If
End Sub

Private Sub cboMBMSalutation_Click()
    If ChkStr(cboMBMSalutation) = "MR" Then
        cboMBMSex = "M"
    Else
        cboMBMSex = "F"
    End If
End Sub

Private Sub cboMBMSalutation_CHANGE()
    If ChkStr(cboMBMSalutation) = "MR" Then
        cboMBMSex = "M"
    Else
        cboMBMSex = "F"
    End If
End Sub


Private Sub cboCovSalutation_Click()
    If ChkStr(cboCovSalutation) = "MR" Then
        cboCovSex = "M"
    Else
        cboCovSex = "F"
    End If
End Sub

Private Sub cboCovSalutation_Change()
    If ChkStr(cboCovSalutation) = "MR" Then
        cboCovSex = "M"
    Else
        cboCovSex = "F"
    End If
End Sub

Private Sub lstCaseList_DblClick()
    frmCaseInquiry.txtCaseNo = ChkStr(lstCaseList.SelectedItem.text)
    Call frmCaseInquiry.cmdShow_Click
    frmCaseInquiry.show
End Sub

Private Sub lstAdjustmentHis_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstAdjustmentHis.SortKey = ColumnHeader.Index - 1
    lstAdjustmentHis.Sorted = True
    If lstAdjustmentHis.SortOrder = lvwAscending Then
        lstAdjustmentHis.SortOrder = lvwDescending
    Else
        lstAdjustmentHis.SortOrder = lvwAscending
    End If
End Sub


Private Sub txtMBMBirthdate_LostFocus()
    'Call MemberCallFees
End Sub


Private Sub cboSRVCode_LostFocus()
'    Call ProviderCallFees
End Sub

Private Sub txtMBMBirthdate_gotfocus()
    txtMBMBirthDate.SelStart = 0
End Sub


Private Sub txtMBMPayorEffDate_gotfocus()
    txtMBMPayorEffDate.SelStart = 0
    Call MemberCallAge
End Sub

Private Sub txtMBMPayorEffDate_LostFocus()
'    Dim tempdate As Date
   
    'Call MemberCallFees
    'Call InsertMCO
 '   If IsDate(txtMBMPayorEffDate) Then
  '      tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
   '     txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
    '    txtMCOFees = Format(GetMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboHLTCode, 1, 1)), txtMBMAdultChild, txtMBMPayorEffDate), "#,##0.00")
     '   Call MemberCallAge
    'Else
     '   MsgBox "Invalid Date !", vbCritical + vbOKOnly
      '  txtMBMPayorEffDate.SetFocus
    'End If
      Dim MBMValidDate As New ADODB.Recordset
    Dim strsql As String
    
    Dim tempdate As Date
    Dim answer

    If (txtMBMPayorEffDate) <> "__/__/____" Then
        If IsDate(txtMBMPayorEffDate) Then
           strsql = "Select ValidID , MBMEffStart, MBMEffEnd From MBMValidDateRange "
           MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                           
           If MBMValidDate.recordCount Then
                If DateDiff("d", MBMValidDate!MBMEffStart, txtMBMPayorEffDate) >= 0 And DateDiff("d", MBMValidDate!MBMEffEnd, txtMBMPayorEffDate) <= 0 Then
                    tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
                    txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
                Else
                    MsgBox "Policy Effective Date is out of acceptable range! ", vbCritical + vbOKOnly
                    txtMBMPayorEffDate.SetFocus
                End If
           Else
                tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
                txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
           End If
           MBMValidDate.Close
           Set MBMValidDate = Nothing
        Else
            answer = MsgBox("Invalid Date Format!", vbCritical + vbOKCancel)
            If answer = vbOK Then
                txtMBMPayorEffDate.SetFocus
            Else
                txtMBMPayorExpDate.SetFocus
            End If
        End If
    End If
    
    If txtMBMBirthDate <> "__/__/____" Then
        Call MemberCallAge
        Call MemberCallFees
    End If
    

End Sub

' ######################### END CHECKING ###########
' ######################### Start Tab Index ###########
Private Sub cboMBMRelCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
        SSTab1.Tab = 1
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 2 Then
        If Mid(cboINSCode, 1, 1) = "I" Or Mid(cboINSCode, 1, 1) = "G" Then
            MsgBox "This Plan does not have supplementary", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 0
            txtMBMAdd1.SetFocus
        Else
            SSTab1.Tab = 2
            cboCovSalutation.SetFocus
        End If
    ElseIf SSTab1.Tab = 3 Then
        lstAdjustmentHis.listitems.clear
        If lstAdjustmentHis.listitems.Count = 0 Then
            Call ADJListing
        End If
        lstAdjustmentHis.SetFocus
    End If
End Sub


Private Sub SSTab1_LostFocus()
    If SSTab1.Tab = 0 Then
        If txtMBMAdd1.Enabled = True Then
            txtMBMAdd1.SetFocus
        End If
    ElseIf SSTab1.Tab = 1 Then
        If cboMBMOccupation.Enabled = True Then
            cboMBMOccupation.SetFocus
        End If
    ElseIf SSTab1.Tab = 2 Then
        If cboCovSalutation.Enabled = True Then
            cboCovSalutation.SetFocus
        End If
    Else
        If lstAdjustmentHis.Enabled = True Then
            lstAdjustmentHis.SetFocus
        End If
    End If
End Sub

Private Sub txtMBMAllergic_gotfocus()
    bTabDone = False
End Sub


Private Sub Combo1_GotFocus()
    If bTabDone = False Then
        cmdEditMBM.SetFocus
        bTabDone = True
    Else
        txtMBMAllergic.SetFocus
        'bTabDone = False
    End If
End Sub

Private Sub Combo2_GotFocus()
    If bTabDone = False Then
        cmdEditMBM.SetFocus
        bTabDone = True
    Else
       ' lstCovAdd.SetFocus
        bTabDone = False
    End If

End Sub

Private Sub lstAdjustmentHis_GotFocus()
    bTabDone = False
End Sub

Private Sub SSTab1_GotFocus()
    If SSTab1.Tab = 0 Then
        txtMBMAdd1.SetFocus
    End If
End Sub

Private Sub Combo3_GotFocus()
    If bTabDone = False Then
        cmdEditMBM.SetFocus
        bTabDone = True
    Else
        lstAdjustmentHis.SetFocus
    End If
End Sub


Private Sub cmdExit_LostFocus()
    If Me.ActiveControl.TabIndex = 44 Then
        txtMBMNumber.SetFocus
    End If
End Sub
' ######################### START BUTTON CLICKING ###########


Public Sub cmdShow_Click()
    Dim getRecord As Integer
    getRecord = showMBM
    Screen.MousePointer = vbHourglass
    If getRecord = 1 Then
        If txtMBMNumber = "" Then
            MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call showMBMCov
            Call CallLimit
'            Call ProviderCallFees

 '           Call MemberCallFees
           ' Call BNFListing
            'Call ListCases
            'txtMBMPremiumEff = Format(CDbl(txtMBMPremium) + CDbl(MCOPremiumReduction(Format(txtMBMPayorExpDate, "mm/dd/yyyy"), Date, txtMBMPremium)), "######.99")
            'txtMCOFeesEff = Format(CDbl(txtMCOFees) + CDbl(MCOPremiumReduction(Format(txtMBMPayorExpDate, "mm/dd/yyyy"), Date, txtMCOFees)), "######.99")
        End If
    End If
    Screen.MousePointer = Default
End Sub

Private Function showMBM() As Integer
Dim sqlstr, GetProString   As String
Dim MBM As New ADODB.Recordset
    sqlstr = "Select MBMNumber, HLTCode, MBMPolicyNo,  MBMName , MBMIcBcPp,  " & _
            " MBMMemberType ,  MBMStatus, " & _
            " INSCode, PLNCode, " & _
            " MBMAddress1, MBMAddress2, " & _
            " MBMAddress3, MBMPostCode, CTYCode, " & _
            " MBMTelnoH, MBMEmail, MBMSalutation, " & _
            " MBMSex, SRVCodeList, MBMTelNoM, " & _
            " MBMTelNoO, MBMOccupation, " & _
            " MBMHeight, MBMWeight, " & _
            " MBMWorkNature," & _
            " MBMGroupCompany, " & _
            " MBMEmployeeNo, MBMBlood, " & _
            " MBMSmoking, MBMAlcohol, " & _
            " MBMExclusion, MBMAllergic, " & _
            " MBMMaritalStatus, MBMAccessFee, MBMPremium , MBMMCOfee ," & _
            " MBMFirstJoinedDate , BRNCode, " & _
            " MBMDOB, NATCode, RACCode, " & _
            " LGNCode, STACode, MBMFirstPolNo , MBMFirstMEMNo,  " & _
            " BRNCode, MBMTakeOver, MBMAgentCode , MBMPolicyYear"
       sqlstr = sqlstr & ", PAYCode, MBMBatchNo, " & _
                " MBMBordxDate, AGECode ,  MBMPrevMemNo, MBMPrevPolNo ," & _
                " MBMPayorEffDate, MBMPayorExpDate , MBMAdultChild, MBMAvailableLimit, " & _
                " MBMLastEndorseStatus , MBMCancelDate , " & _
                " MBMReinstate , MBMLastAdjustmentDate  "
    sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber ='" & Trim(txtMBMNumber) & "'"

    MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBM.recordCount = 0 Then
       MsgBox "No record found!"
    Else
        showMBM = 1
        With MBM
            If !MBMStatus = "S" Then
                'txtMBMName = Trim(!MBMName) & " (Suspended)"
                txtMBMName = Trim(!MBMName)
                Label7.Caption = "Suspended"
                
            ElseIf !MBMStatus = "C" Then
                'txtMBMName = Trim(!MBMName) & " (Cancelled)"
                txtMBMName = Trim(!MBMName)
                Label7.Caption = "Cancelled"
            ElseIf !MBMStatus = "D" Then
            'txtMBMName = Trim(!MBMName) & " (Deleted)"
                txtMBMName = Trim(!MBMName)
                Label7.Caption = "Deleted"
            Else
            'txtMBMName = Trim(!MBMName) & " (Active)"
                txtMBMName = Trim(!MBMName)
                Label7.Caption = "Active"
            End If
            txtMBMIcBcPP = Trim(!MBMIcBcPp)
            txtMBMPolNo = Trim(!MBMPolicyNo)
            cboINSCode = Trim(!INSCode)
            cboPLNCode = Trim(!PLNCode)
            txtMBMAdd1 = Trim(!MBMAddress1)
            txtMBMAdd2 = Trim(!MBMAddress2)
            txtMBMAdd3 = Trim(!MBMAddress3)
            txtMBMPostCode = Trim(!MBMPostCode)
            cboCTYCode = Trim(!CTYCode)
            If Trim(!MBMTelnoH) <> "" Then
                txtMBMTelNoH = Trim(!MBMTelnoH)
            End If
            If Trim(!MBMEmail) <> "" Then
                txtMBMEmail = Trim(!MBMEmail)
            End If
            If Trim(!MBMSalutation) <> "" Then
                cboMBMSalutation = Trim(!MBMSalutation)
            End If
            If Len(!MBMSex) Then
                cboMBMSex = Trim(!MBMSex)
            End If
            If Trim(!SRVCodeList) <> "" Then
                cboSRVCode = Trim(!SRVCodeList)
            End If
            If Trim(!MBMTelNoM) <> "" Then
                txtMBMTelNoM = Trim(!MBMTelNoM)
            End If
            If Trim(!MBMTelNoO) <> "" Then
                txtMBMTelNoO = Trim(!MBMTelNoO)
            End If
            If Trim(!MBMOccupation) <> "" Then
                cboMBMOccupation = Trim(!MBMOccupation)
            End If
            If Trim(!MBMHeight) <> "" Then
                txtMBMHeight = Trim(!MBMHeight)
            End If
            If Trim(!MBMWeight) <> "" Then
                txtMBMWeight = Trim(!MBMWeight)
            End If
            If Trim(!MBMWorkNature) <> "" Then
                txtMBMWorkNature = Trim(!MBMWorkNature)
            End If
            If Trim(!MBMGroupCompany) <> "" Then
                txtMBMGroupCompany = Trim(!MBMGroupCompany)
            End If
            If Trim(!MBMEmployeeNo) <> "" Then
                txtMBMEmployeeNo = Trim(!MBMEmployeeNo)
            End If
            If Trim(!MBMBlood) <> "" Then
                cboMBMBlood = Trim(!MBMBlood)
            End If
            If Trim(!MBMSmoking) <> "" Then
                cboMBMSmoking = Trim(!MBMSmoking)
            End If
            If Trim(!MBMAlcohol) <> "" Then
                cboMBMAlcohol = Trim(!MBMAlcohol)
            End If
            If Trim(!MBMExclusion) <> "" Then
                txtMBMExclusion = Trim(!MBMExclusion)
            End If
            If Trim(!MBMAllergic) <> "" Then
                txtMBMAllergic = Trim(!MBMAllergic)
            End If
            If Trim(!MBMMaritalStatus) <> "" Then
                cboMBMMaritalStatus = Trim(!MBMMaritalStatus)
            End If
            If Trim(!MBMFirstJoinedDate) <> "" Then
                txtMBMDateJoin = Trim(!MBMFirstJoinedDate)
            End If
            If Trim(!BRNCode) <> "" Then
                cboBRNCode = Trim(!BRNCode)
            End If
            If Trim(!NATCode) <> "" Then
                cboNATCode = Trim(!NATCode)
            End If
            If Trim(!RACCode) <> "" Then
                cboRACCode = Trim(!RACCode)
            End If
            If Trim(!LGNCode) <> "" Then
                cboLGNCode = Trim(!LGNCode)
            End If
            If Trim(!STACode) <> "" Then
                cboSTACode = Trim(!STACode)
            End If
            If Trim(!MBMFirstPolNo) <> "" Then
                txtFirstPolNo = Trim(!MBMFirstPolNo)
            End If
            If Trim(!MBMPrevMemNo) <> "" Then
                txtMBMPrevMEMNo = Trim(!MBMPrevMemNo)
            End If
            If Trim(!MBMPrevPolNo) <> "" Then
                txtMBMPrevPolNo = Trim(!MBMPrevPolNo)
            End If
            If Trim(!MBMDOB) <> "" Then
                txtMBMBirthDate.Mask = ""
                txtMBMBirthDate.text = Trim(!MBMDOB)
            End If
            If Trim(!MBMFirstMEMNo) <> "" Then
                txtFirstMeMNo = Trim(!MBMFirstMEMNo)
            End If
            If Trim(!BRNCode) <> "" Then
                cboBRNCode = Trim(!BRNCode)
            End If
            If Trim(!MBMTakeOver) = "C" Then
                cboMBMTakeOver.ListIndex = 2
            ElseIf Trim(!MBMTakeOver) = "T" Then
                cboMBMTakeOver.ListIndex = 1
            Else
                cboMBMTakeOver.ListIndex = 0
            End If
            If Trim(!MBMAgentCode) = "C" Then
                txtMBMAgentCode = Trim(!MBMAgentCode)
            End If
        '    If Trim(!RELCODE) = "C" Then
         '       cboMBMRelCode = Trim(!RELCODE)
          '  End If
            cboPAYCode = Trim(!PAYCode)
            txtAGECode = Trim(!AgeCode)
            'txtEndtBrodxDate.Mask = ""
            'txtEndtBrodxDate.Text = Trim(!MBMENDtBordxDate)
            If Len(!MBMBordxDate) Then
                txtMBMBordxDate.Mask = ""
                txtMBMBordxDate.text = Trim(!MBMBordxDate)
            End If
            'txtMBMBatchNo = Trim(!MBMBatchNo)
            cboHLTCode = Trim(!HLTCode)
            txtMBMPayorEffDate.Mask = ""
            txtMBMPayorEffDate = Trim(!MBMPayorEffDate)
            txtMBMPayorExpDate.Mask = ""
            txtMBMPayorExpDate = Trim(!MBMPayorEXPDate)
            txtMBMPolicyYear = Trim(!MBMPolicyYear)
            If Len(!MBMAvailableLimit) Then
                txtAvailableAnnualLimit = Trim(!MBMAvailableLimit)
            End If
            txtMBMAdultChild = Trim(!MBMAdultChild)
            txtEndtBordxDate.Mask = ""
            txtEndtBordxDate = Date
            txtEndtBordxDate.Mask = "##/##/####"
            cboMBMRelCode = Trim(!MBMMemberType)
            
            txtMBMPremium = Format(IIf(IsNull(!MBMPremium) = True, Format(GetPremium(Trim(cboINSCode), Trim(cboPAYCode), Trim(txtAGECode), Trim(cboPLNCode), Trim(cboHLTCode), txtMBMPayorEffDate), "#,##0.00"), !MBMPremium), "####0.00")
            txtMCOFees = Format(IIf(IsNull(!MBMMCOFee) = True, Format(GetMCO(Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild, txtMBMPayorEffDate), "#,##0.00"), !MBMMCOFee), "####0.00")
            
            If IsNull(!MBMAccessFee) = True Then
                ProviderCallFees
            Else
                txtMBMProviderCharges = Format(!MBMAccessFee, "####0.00")
                GetProString = GetProviderCharges(cboSRVCode, cboINSCode, cboHLTCode, txtMBMAdultChild)
                txtMBMCaseCharges = Format(Mid(GetProString, InStr(1, GetProString, "-") + 1), "#,##0.00")
            End If
            
            
            If Trim(!MBMStatus) = "A" Then
                cboStatus.ListIndex = 0
            ElseIf Trim(!MBMStatus) = "C" Then
                cboStatus.ListIndex = 1
            ElseIf Trim(!MBMStatus) = "D" Then
                cboStatus.ListIndex = 2
            ElseIf Trim(!MBMStatus) = "S" Then
                cboStatus.ListIndex = 3
            End If
'Trim (!MBMReinstate)
            If Trim(!MBMLastEndorseStatus) = "A" Then
                cboEndorseStatus.ListIndex = 0
                txtDate.Mask = ""
                txtDate = Trim(!MBMLastAdjustmentDate)
                txtLastEndStatus = "A"
                txtLastEdtDate = Trim(!MBMLastAdjustmentDate)
            ElseIf Trim(!MBMLastEndorseStatus) = "C" Then
                cboEndorseStatus.ListIndex = 1
                txtDate.Mask = ""
                txtDate = Trim(!MBMCancelDate)
                txtLastEndStatus = "C"
                txtLastEdtDate = Trim(!MBMLastAdjustmentDate)
            ElseIf Trim(!MBMLastEndorseStatus) = "S" Then
                cboEndorseStatus.ListIndex = 2
                txtDate.Mask = ""
                txtDate = Trim(!MBMLastAdjustmentDate)
                txtLastEndStatus = "S"
                txtLastEdtDate = Trim(!MBMLastAdjustmentDate)
            ElseIf Trim(!MBMLastEndorseStatus) = "T" Then
                cboEndorseStatus.ListIndex = 3
                txtDate.Mask = ""
                txtDate = Trim(!MBMLastAdjustmentDate)
                txtLastEndStatus = "T"
                txtLastEdtDate = Trim(!MBMLastAdjustmentDate)
            ElseIf Trim(!MBMLastEndorseStatus) = "X" Then
                cboEndorseStatus.ListIndex = 4
                txtDate.Mask = ""
                txtDate = Trim(!MBMLastAdjustmentDate)
                txtLastEndStatus = "X"
                txtLastEdtDate = Trim(!MBMLastAdjustmentDate)
            ElseIf Trim(!MBMLastEndorseStatus) = "Y" Then
                cboEndorseStatus.ListIndex = 5
                txtDate.Mask = ""
                'txtDate = lstMBMInquiryList.SelectedItem.SubItems(59)
                txtLastEndStatus = "Y"
                txtLastEdtDate = Trim(!MBMLastAdjustmentDate)
            
            Else
                cboEndorseStatus.ListIndex = 0
                
                'txtDate = Trim(!MBMLastAdjustmentDate)
                
            End If
        End With
    End If
    MBM.Close
    Set MBM = Nothing
End Function

Private Sub showMBMCov()
    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String
                
    sqlstr = "select MBMCNumber, MBMCCoverID, MBMCName, MBMCSalutation," & _
        " MBMCICBCPPNo , MBMCDOB, MBMCSex, MBMCAdultChild, MBMCAgeband, " & _
        " MBMCRELCode, MBMCBloodType, MBMCAllergic,  " & _
        " MBMCExclusion, MBMCStatus , MBMCAnnualLimit" & _
        " from  MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic

    lstCovAdd.listitems.clear
    If MBMCov.recordCount Then
        Dim MemberConveredPerson As ListItem
        Do While Not MBMCov.EOF
            Set MemberConveredPerson = lstCovAdd.listitems.Add(, , MBMCov!MBMCName)
            With MBMCov
                If Len(Trim(!MBMCSalutation)) <> 0 Then
                    MemberConveredPerson.SubItems(1) = ChkStr(!MBMCSalutation)
                End If
                If Len(Trim(!MBMCICBCPPNo)) <> 0 Then
                    MemberConveredPerson.SubItems(2) = !MBMCICBCPPNo
                End If
                If Len(Trim(!MBMCDOB)) Then
                    MemberConveredPerson.SubItems(3) = !MBMCDOB
                End If
                
                If Len(Trim(!MBMCSex)) Then
                    MemberConveredPerson.SubItems(4) = ChkStr(!MBMCSex)
                End If
                
                If Len(Trim(!MBMCAdultChild)) Then
                    MemberConveredPerson.SubItems(5) = ChkStr(!MBMCAdultChild)
                End If
                
                If Len(Trim(!MBMCRELCode)) Then
                    MemberConveredPerson.SubItems(6) = ChkStr(!MBMCRELCode)
                End If
                
                If Len(Trim(!MBMCBloodType)) Then
                    MemberConveredPerson.SubItems(7) = ChkStr(!MBMCBloodType)
                End If
                
                If Len(Trim(!MBMCAllergic)) Then
                    MemberConveredPerson.SubItems(8) = !MBMCAllergic
                End If
                
                If Len(Trim(!MBMCExclusion)) Then
                    MemberConveredPerson.SubItems(9) = !MBMCExclusion
                End If
                MemberConveredPerson.SubItems(10) = !MBMCCoverID
                
                If Len(Trim(!MBMCAnnualLimit)) Then
                    MemberConveredPerson.SubItems(12) = !MBMCAnnualLimit
                End If
            End With
            MBMCov.MoveNext
        Loop
    End If
    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub cmdEditMBM_Click()
    EnableEntries (True)
    cboINSCode.Enabled = False
    cboPLNCode.Enabled = False
    Frame10.Enabled = True
    Frame1.Enabled = False
    SSTab1.Enabled = True
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdRemove.Enabled = False
    cboEndorseStatus.Enabled = True
End Sub

Private Sub CmdExit_Click()
Dim RecordExit
    RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
    End If
End Sub

Private Sub cmdMBMPolicy_Click()
    frmSearchMBM.txtMBMPolicyNo1.value = True
    frmSearchMBM.txtMBMNo1.value = False
    frmSearchMBM.Source = 2
    frmSearchMBM.show
End Sub

Private Sub cmdSearchMBM_Click()
    Call ClearEntries
    frmSearchMBM.Source = 2
    frmSearchMBM.show
End Sub

Private Sub cmdSearchIC_Click()
    frmSearchMBM.txtIcBcPc1.value = True
    frmSearchMBM.txtMBMNo1.value = False
    frmSearchMBM.Source = 2
    frmSearchMBM.show
End Sub

Private Sub cmdSearchName_Click()
    frmSearchMBM.txtMBMName6.value = True
    frmSearchMBM.txtMBMNo1.value = False
    frmSearchMBM.Source = 2
    frmSearchMBM.show
End Sub

Private Sub cmdListBranch_Click()
    frmMBMListBranch.Source = 2
    frmMBMListBranch.show
End Sub

Private Sub cmdListCity_Click()
    frmMBMListCity.Source = 2
    frmMBMListCity.show
End Sub

Private Sub cmdListState_Click()
    frmMBMListState.Source = 2
    frmMBMListState.show
End Sub

' ######## START SUPPLEMENTARY #####################

Private Sub cmdAdd_Click()
    Dim MemberConveredPerson As ListItem
    Dim RecordSaved

    If txtCovName = "" Then
        MsgBox "Please Enter Covered Person Name", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtCovIcBcPp = "" Then
        MsgBox "Please Enter Covered Person Identification", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboCovRelationship = "" Then
        MsgBox "Please Enter Covered Person Relationship", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Set MemberConveredPerson = lstCovAdd.listitems.Add(, , txtCovName)
            MemberConveredPerson.SubItems(1) = cboCovSalutation
            MemberConveredPerson.SubItems(2) = txtCovIcBcPp
            MemberConveredPerson.SubItems(3) = txtCovBirthdate
            MemberConveredPerson.SubItems(4) = cboCovSex
            MemberConveredPerson.SubItems(5) = cboCovAdultChild
            MemberConveredPerson.SubItems(6) = cboCovRelationship
            MemberConveredPerson.SubItems(7) = cboCovBlood
            MemberConveredPerson.SubItems(8) = txtCovAllergic
            MemberConveredPerson.SubItems(9) = txtCovExclusion
            MemberConveredPerson.SubItems(11) = "N"
            cboCovSalutation.SetFocus
        End If
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim ForLoop As Integer
    Dim ListCount As Integer
    Dim strsql As String
    Dim DEL As New ADODB.Recordset
    Dim RecordSaved
    ListCount = 0
    If txtCovName = "" Then
        MsgBox "Please Choose a Record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            ListCount = lstCovAdd.listitems.Count
            For ForLoop = 1 To ListCount
                If lstCovAdd.listitems(ForLoop).Selected = True Then
                    If Trim(lstCovAdd.SelectedItem.SubItems(10)) <> "" Then
                        lstCovAdd.SelectedItem.SubItems(11) = "R"
                    Else
                        lstCovAdd.listitems.Remove (ForLoop)
                    End If
                End If
            Next ForLoop
        End If
    End If
End Sub



Private Sub cmdUpdate_Click()
Dim ForLoop As Integer

Dim RecordSaved
    If txtCovName = "" Then
       MsgBox "Please Enter Covered Person Name", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtCovIcBcPp = "" Then
       MsgBox "Please Enter Covered Person Identification", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboCovRelationship = "" Then
       MsgBox "Please Enter Covered Person Relationship", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            For ForLoop = 1 To lstCovAdd.listitems.Count
                If lstCovAdd.listitems(ForLoop).Selected = True Then
                        lstCovAdd.SelectedItem.text = txtCovName
                        lstCovAdd.SelectedItem.SubItems(1) = Trim(cboCovSalutation)
                        lstCovAdd.SelectedItem.SubItems(2) = txtCovIcBcPp
                        lstCovAdd.SelectedItem.SubItems(3) = txtCovBirthdate
                        lstCovAdd.SelectedItem.SubItems(4) = cboCovSex
                        lstCovAdd.SelectedItem.SubItems(5) = cboCovAdultChild
                        lstCovAdd.SelectedItem.SubItems(6) = cboCovRelationship
                        lstCovAdd.SelectedItem.SubItems(7) = cboCovBlood
                        lstCovAdd.SelectedItem.SubItems(8) = txtCovAllergic
                        lstCovAdd.SelectedItem.SubItems(9) = txtCovExclusion
                        lstCovAdd.SelectedItem.SubItems(11) = "E"
                        lstCovAdd.SelectedItem.SubItems(12) = txtCovAnnualLimit
                        Exit For
                End If
            Next ForLoop
            cboCovSalutation.SetFocus
        End If
       Call ClearCoveredEntry
    End If
End Sub
Private Sub lstCovAdd_Click()
    If lstCovAdd.listitems.Count Then
        txtCovName = lstCovAdd.SelectedItem.text
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.Mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        cboCovBlood = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
       ' txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = lstCovAdd.SelectedItem.SubItems(12)
    End If
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

Private Sub CmdClear_Click()
    Call EnableCovered(True)
    Call ClearCoveredEntry
    
End Sub


Public Sub ClearCoveredEntry()
    cboCovSalutation = "MR"
    txtCovName = ""
    txtCovIcBcPp = ""
    txtCovBirthdate.Mask = ""
    txtCovBirthdate.Mask = "##/##/####"
    cboCovSex = "M"
    cboCovAdultChild = "A"
    cboCovRelationship = "S - SON"
    cboCovBlood = "A"
    txtCovAllergic = ""
    txtCovExclusion = ""
    txtCovAnnualLimit = "0"
End Sub


Private Sub EnableSupplementary(status As Boolean)
    lstCovAdd.Enabled = status
    cboCovSalutation.Enabled = status
    txtCovName.Enabled = status
    txtCovIcBcPp.Enabled = status
    txtCovBirthdate.Enabled = status
    cboCovSex.Enabled = status
    cboCovAdultChild.Enabled = status
    cboCovRelationship.Enabled = status
    cboCovBlood.Enabled = status
    txtCovAllergic.Enabled = status
    txtCovExclusion.Enabled = status
    txtCovAnnualLimit.Enabled = status
    cmdRemove.Enabled = status
    cmdAdd.Enabled = status
End Sub
'  ################# END Suppplementary ######################

' ################################ Enable Tab / button   ######
Public Sub EnableCovered(status As Boolean)
    cboCovSalutation.Enabled = status
    txtCovName.Enabled = status
    txtCovIcBcPp.Enabled = status
    txtCovBirthdate.Enabled = status
    cboCovSex.Enabled = status
    cboCovAdultChild.Enabled = status
    cboCovRelationship.Enabled = status
    cboCovBlood.Enabled = status
    txtCovAllergic.Enabled = status
    txtCovExclusion.Enabled = status
    txtCovAnnualLimit.Enabled = status
End Sub

Private Sub TabEnabled(status As Boolean)
    SSTab1.TabEnabled(0) = status
    SSTab1.TabEnabled(1) = status
    SSTab1.TabEnabled(2) = status
    SSTab1.TabEnabled(3) = status
'    SSTab1.TabEnabled(4) = status

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
         ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Enabled = status
        End If
    Next ctl
    'cmdSave.Enabled = False
    cboEndorseStatus.Enabled = False
    cboStatus.Enabled = False
End Sub

' ################## Start processing

'Private Sub DeleteCoveredID(MembershipNo As String, coverID As Integer)
 '   Dim DELCOV As New ADODB.Recordset
  '  Dim delstr As String
   '
    'delstr = "Delete from MBMCoveredPersons Where MBMCCoverID =" & coverID & " And MBMNumber = '" & MembershipNo & "'"
    'DELCOV.Open delstr, medixmasterconn, adOpenKeyset, adLockOptimistic
'
 '   DELCOV.Close
  '  Set DELCOV = Nothing
'End Sub

Private Sub CancelPrincipalAdjustment()
Dim MBM As New ADODB.Recordset
Dim strsql As String
        
    strsql = "Select * from MBM Where MBMNumber = '" & txtMBMNumber & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                With MBM
                    !MBMLastEndorseStatus = "C"
                    !MBMStatus = "C"
                    !MBMCancelDate = txtDate
                End With
                MBM.Update
    MBM.Close
    'cboStatus.Enabled = True
    cboStatus.ListIndex = 1
    'cboStatus.Enabled = False
End Sub

Private Sub CancelSupplementaryAdjustment()
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
        
    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'MemberHis.Open "MBA", medixmasterconn, adOpenKeyset, adLockOptimistic
    'MBMHIS.Open "MBMHIS", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until MBMCov.EOF
                With MBMCov
                    !MBMCStatus = "C"
                End With
                MBMCov.Update
        MBMCov.MoveNext
    Loop
    MBMCov.Close
End Sub

'  ############### SAVE DATA INTO DATABSE #####################
Private Sub cmdSave_Click()
On Error GoTo ErrorSave
Dim RecordSaved, confirmCancel
Dim strsql As String
Dim CAS As New ADODB.Recordset
    
    If Trim(txtMBMNumber) = "" And Mid(cboEndorseStatus, 1, 1) <> "I" Then
        MsgBox "Please select a record to edit!", vbCritical
        txtMBMNumber.SetFocus
    ElseIf txtEndtBordxDate = "__/__/____" Then
        MsgBox "Please enter ENDtBordx Date!", vbCritical
        txtEndtBordxDate.SetFocus
    ElseIf Trim(txtBatchNo) = "" Then
        MsgBox "Please enter Batch No!", vbCritical
        txtBatchNo.SetFocus
    Else
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                'startTrans = True
                'medixmasterconn.BeginTrans
                Screen.MousePointer = vbHourglass
                If cboEndorseStatus = "" Then
                    MsgBox "Please Choose An Endorsement Type", vbOKOnly + vbCritical, DialogBoxTitle
                    cboEndorseStatus.SetFocus
                ElseIf cboEndorseStatus.ListIndex = 1 Then
                    If txtDate = "__/__/____" Then
                        MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtDate.SetFocus
                    ElseIf txtEndtBordxDate = "__/__/____" Then
                        MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtEndtBordxDate.SetFocus

                    ElseIf Mid(cboStatus, 1, 1) = "D" Then
                        MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                    ElseIf Mid(cboStatus, 1, 1) = "C" Then
                        MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        strsql = "Select CASNumber from CAS Where MBMNumber ='" & txtMBMNumber & "' And CASStatus = 'O'"
                        CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If CAS.recordCount > 0 Then
                            confirmCancel = MsgBox("There is at least an outstanding case under this member!", vbYesNo + vbCritical)
                        
                            'If confirmCancel Then
                            '    Call CancelPrincipalAdjustment
                            '    Call MBMRefund
                            '    Call CancelSupplementaryAdjustment
                            'End If
                        Else
                            Call CancelPrincipalAdjustment
                            Call MBMRefund
                            Call CancelSupplementaryAdjustment
                            MsgBox "This Membership has been cancelled successfully", vbOKOnly + vbInformation, DialogBoxTitle
                        End If
                    End If
                   
                    ElseIf Mid(cboEndorseStatus, 1, 1) = "I" Then
                        If txtDate = "__/__/____" Then
                            MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                            txtDate.SetFocus
                        ElseIf txtEndtBordxDate = "__/__/____" Then
                            MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                            txtEndtBordxDate.SetFocus
    
                        ElseIf Mid(cboStatus, 1, 1) = "D" Then
                            MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                        ElseIf Mid(cboStatus, 1, 1) = "C" Then
                            MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                        Else
                        Call InsertNewRegistrationSub
                        End If
                        
                        
                'ElseIf Mid(cboStatus, 1, 1) = "A" Then
                        ' delete Supplementary
                    ElseIf Mid(cboEndorseStatus, 1, 1) = "T" Then
                        If Mid(cboStatus, 1, 1) = "D" Then
                            MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                        ElseIf Mid(cboStatus, 1, 1) = "C" Then
                            MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                        ElseIf txtEndtBordxDate = "__/__/____" Then
                            MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                            txtEndtBordxDate.SetFocus
                        ElseIf txtDate = "__/__/____" Then
                            MsgBox "Please Enter EndtEffBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                            txtDate.SetFocus
                        Else
                            Call DeleteSupplementary
                        End If
                        ' Add  Supplementary
                 ElseIf Mid(cboEndorseStatus, 1, 1) = "S" Then
                     
                     If Mid(cboStatus, 1, 1) = "D" Then
                         MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf Mid(cboStatus, 1, 1) = "C" Then
                         MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf txtDate = "__/__/____" Then
                         MsgBox "Please Enter EndtEffDate Date", vbOKOnly + vbCritical, DialogBoxTitle
                         txtDate.SetFocus
                    ElseIf txtEndtBordxDate = "__/__/____" Then
                        MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtEndtBordxDate.SetFocus
                     Else
                '     RecordSaved = vbYes
                         Call AddSupplementaryAdjustment
                     End If
                 ' Susspend MEMbership
                 ElseIf Mid(cboEndorseStatus, 1, 1) = "X" Then
                     If Mid(cboStatus, 1, 1) = "D" Then
                         MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf Mid(cboStatus, 1, 1) = "C" Then
                         MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf txtDate = "__/__/____" Then
                         MsgBox "Please Enter EndtEffDate Date", vbOKOnly + vbCritical, DialogBoxTitle
                         txtDate.SetFocus
                    ElseIf txtEndtBordxDate = "__/__/____" Then
                        MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtEndtBordxDate.SetFocus
                     Else
                            Call SuspendPrincipal
                            Call SuspendSupplementary
                            MsgBox "Record is Suspended", vbOKOnly + vbInformation, DialogBoxTitle

                         'RecordSaved = vbYes
                        ' Call SuspendAdjustment
                     End If
                 ElseIf Mid(cboEndorseStatus, 1, 1) = "Y" Then
                    If Mid(cboStatus, 1, 1) = "D" Then
                         MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf Mid(cboStatus, 1, 1) = "C" Then
                         MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf txtDate = "__/__/____" Then
                         MsgBox "Please Enter Suspension Date", vbOKOnly + vbCritical, DialogBoxTitle
                         txtDate.SetFocus
                     ElseIf txtEndtBordxDate = "__/__/____" Then
                        MsgBox "Please Enter EndtBrodx Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtEndtBordxDate.SetFocus
                     Else
                        Call ReactivatePrincipal
                        Call ReactivateSupplementary
                        MsgBox "Record is Activated", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
                 ' Amendment
                 ElseIf Mid(cboEndorseStatus, 1, 1) = "A" Then
                     If txtEndtBordxDate = "__/__/____" Then
                        MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtEndtBordxDate.SetFocus
                     ElseIf txtDate = "__/__/____" Then
                         MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                         txtDate.SetFocus
                     ElseIf Mid(cboStatus, 1, 1) = "C" Then
                         MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf Mid(cboStatus, 1, 1) = "D" Then
                         MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     Else
                         Call SavePrincipalAdjustment
                         Call SaveSupplementaryAdjustment
                         MsgBox "Updated Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
                     End If
              
                'ElseIf cboStatus = "D" Then
                '    strsql = "Select CASNumber from CAS Where MBMNumber ='" & txtMBMNumber & "' And CASStatus = 'O'"
                '    CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                
               '     If CAS.RecordCount > 0 Then
               '         confirmCancel = MsgBox("There is at least an outstanding case under this member!", vbYesNo + vbCritical)
               '         If confirmCancel Then
               '             Call DeleteSupplementary
               '         End If
               '     Else
                        'RecordSaved = vbYes
                        
               '     End If
                ElseIf Mid(cboEndorseStatus, 1, 1) = "D" Then
                     If txtEndtBordxDate = "__/__/____" Then
                            MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                         txtEndtBordxDate.SetFocus
                     ElseIf txtDate = "__/__/____" Then
                         MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                         txtDate.SetFocus
                     ElseIf Mid(cboStatus, 1, 1) = "C" Then
                         MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf Mid(cboStatus, 1, 1) = "D" Then
                         MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     Else
                        Call DeletePrincipals
                     End If
                Else
                    If cboStatus = "D" Then
                            MsgBox "This Membership is Deleted", vbOKOnly + vbInformation, DialogBoxTitle
                    ElseIf cboStatus = "C" Then
                            MsgBox "This Membership is Cancelled", vbOKOnly + vbInformation, DialogBoxTitle
                    ElseIf cboStatus = "S" And Mid(cboEndorseStatus, 1, 1) <> "C" Then
                            MsgBox "This Membership is Suspended", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
                End If
                   End If
            Screen.MousePointer = vbDefault
            'If startTrans = True Then
            '    medixmasterconn.CommitTrans
            'End If
    'End If

    End If
    cmdEditMBM.Enabled = True
    cboEndorseStatus.Enabled = False
    cmdSave.Enabled = False
    Call ADJListing
    
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        medixmasterconn.RollbackTrans
    End If
    MsgBox ERR.Description
End Sub

Public Sub MBMRefund()
    Dim ServiceChargeRefundValue As Double
    Dim MBMRefund As New ADODB.Recordset
    Dim strsql As String
    Dim MCORefundValue As String
    Dim MBMMCOFees As Double
    Dim PremiumRefundValue As String
    Dim MCOEffective As String
    Dim PremiumEffective As String
    Dim FindAge As String
    Dim MBMPremium As Double
    Dim PlanCode As String
    Dim InsuredCode As String
    Dim HLTCode As String
    Dim Payor As String
    Dim AgeValue As String
    Dim PayEffDate As String
    Dim PayExpDate As String
    Dim MBMDOB As String
    Dim ProviderRefValue As Double
    Dim ProviderEff As String
    
    HLTCode = Mid(cboHLTCode, 1, 1)
    If InStr(1, cboPLNCode, "-") > 1 Then
        PlanCode = Mid(cboPLNCode, 1, InStr(1, cboPLNCode, "-") - 1)
    Else
        PlanCode = Trim(Mid(cboPLNCode, 1, 4))
    End If
    InsuredCode = Mid(cboINSCode, 1, 1)
    Payor = Mid(cboPAYCode, 1, 2)
    MBMDOB = txtMBMBirthDate
    PayEffDate = txtMBMPayorEffDate
    PayExpDate = txtMBMPayorExpDate
        
    strsql = "Select * from MBMRefund Where MBMNumber = '" & txtMBMNumber & "'"
    MBMRefund.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

    FindAge = GetAge(CDate(PayEffDate), CDate(MBMDOB))
    AgeValue = Mid(GetAgeBand(HLTCode, FindAge), 1, 2)
    
    MBMPremium = GetPremium(InsuredCode, Payor, AgeValue, PlanCode, HLTCode)

    If AgeValue = "01" Then
        MBMMCOFees = GetMCO(InsuredCode, HLTCode, "C")
    Else
        MBMMCOFees = GetMCO(InsuredCode, HLTCode, "A")
    End If
    
    MCORefundValue = MCORefund(CDate(PayExpDate), CDate(txtDate), MBMMCOFees)
    PremiumRefundValue = MCOPremiumReduction(CDate(PayExpDate), CDate(txtDate), MBMPremium)
    ProviderRefValue = EAFeesReduction(CDate(PayExpDate), CDate(txtDate), txtMBMProviderCharges)
    
    PremiumEffective = CDbl(MBMPremium) - CDbl(PremiumRefundValue)
    MCOEffective = CDbl(MBMMCOFees) - CDbl(MCORefundValue)
    ProviderEff = CDbl(txtMBMProviderCharges) - CDbl(ProviderRefValue)
        MBMRefund.AddNew
        With MBMRefund
            !MBMNumber = txtMBMNumber
            !MBMPREMEffective = PremiumEffective
            !MBMPremReduction = PremiumRefundValue
            !MBMMCOEffective = MCOEffective
            !MBMMCORefund = MCORefundValue
            !MBMProviderRef = ProviderRefValue
            !MBMProviderEff = ProviderEff
        End With
        MBMRefund.Update
    txtMBMPremiumEff = Format(PremiumEffective, "#,##0.00")
    txtMCOFeesEff = Format(MCOEffective, "#,##0.00")
    txtMBMPremiumRed = Format(PremiumRefundValue, "#,##0.00")
    txtMBMMCORed = Format(MCORefundValue, "#,##0.00")
    txtMBMProviderRedCharges = Format(ProviderRefValue, "#,##0.00")
    txtMBMProviderEffCharges = Format(ProviderEff, "#,##0.00")
End Sub

Private Sub SavePrincipalAdjustment()

    Dim MemberAmendHis As New ADODB.Recordset
    Dim MBM As New ADODB.Recordset
    Dim MBMOthers As New ADODB.Recordset
    Dim strsql As String
    Dim strsqlOthers As String

    
    strsql = "Select * from MBM Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MemberAmendHis.AddNew
    With MemberAmendHis
        
        !MBMNumber = MBM!MBMNumber
        !MBMPolicyNo = MBM!MBMPolicyNo
        !MBMIcBcPp = MBM!MBMIcBcPp
        !MBMName = MBM!MBMName
        !MBMExpiryDate = MBM!MBMPayorEXPDate
        !MBMPlaNCode = MBM!PLNCode
        !MBMAmendEdtType = Mid(cboEndorseStatus, 1, 1)
        If IsDate(txtDate) Then
            !MBMAmendDate = txtDate
        End If
        !MBMAmendUser = Logon
        !MBMAEndtBatchNo = txtBatchNo
        !MBMAEndtEffDate = txtDate
        MemberAmendHis.Update
    End With
    
    With MBM
        !MBMSalutation = cboMBMSalutation
        !MBMName = txtMBMName
        !MBMPolicyNo = Trim(txtMBMPolNo)
        !MBMSex = cboMBMSex
        !MBMName = txtMBMName
        !MBMAddress1 = Trim(txtMBMAdd1)
        !MBMAddress2 = Trim(txtMBMAdd2)
        !MBMAddress3 = Trim(txtMBMAdd3)
        !CTYCode = Trim(cboCTYCode)
        !STACode = Trim(cboSTACode)
        !MBMPostCode = txtMBMPostCode
        !MBMEmail = txtMBMEmail
        !MBMIcBcPp = txtMBMIcBcPP
        !MBMDOB = txtMBMBirthDate
        !RACCode = Trim(Mid(cboRACCode, 1, 2))
        !LGNCode = Trim(Mid(cboLGNCode, 1, 2))
        If InStr(1, cboNATCode, "-") > 0 Or Len(cboNATCode) = 2 Then
            If InStr(1, cboNATCode, "-") > 0 Then
                !NATCode = Trim(Mid(cboNATCode, 1, InStr(1, cboNATCode, "-") - 1))
            Else
                !NATCode = Trim(Mid(cboNATCode, 1, 2))
            End If
        End If
        
        If IsDate(txtMBMPayorEffDate) Then
            !MBMPayorEffDate = txtMBMPayorEffDate
        End If
        If IsDate(txtMBMPayorExpDate) Then
            !MBMPayorEXPDate = txtMBMPayorExpDate
        End If
        !BRNCode = cboBRNCode
        !MBMTelnoH = txtMBMTelNoH
        !MBMTelNoM = txtMBMTelNoM
        !MBMTelNoO = txtMBMTelNoO
        !MBMLastAdjustmentDate = Date
        !MBMLastEndorseStatus = Trim(Mid(cboEndorseStatus, 1, InStr(1, cboEndorseStatus, "-") - 1))
        !MBMBordxDate = txtMBMBordxDate
        MBM.Update
    End With
    
    ' INsert into Membership Other details table
    strsqlOthers = "select * from MBMOthers where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    MBMOthers.Open strsqlOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
        With MBMOthers
             !MBMWeight = txtMBMWeight
             !MBMHeight = txtMBMHeight
             If InStr(1, txtMBMWorkNature, "-") > 0 Then
                !MBMWorkNature = Trim(Mid(txtMBMWorkNature, 1, InStr(1, txtMBMWorkNature, "-") - 1))
             Else
               !MBMWorkNature = Trim(txtMBMWorkNature)
             End If
             !MBMExclusion = txtMBMExclusion
             !MBMAllergic = txtMBMAllergic
             !MBMGroupCompany = txtMBMGroupCompany
             !MBMEmployeeNo = txtMBMEmployeeNo
             !MBMBlood = Trim(Mid(cboMBMBlood, 1, 2))
             !MBMSmoking = Trim(Mid(cboMBMSmoking, 1, 2))
             !MBMAlcohol = Mid(cboMBMAlcohol, 1, 2)
             !MBMAgentCode = txtMBMAgentCode
             !RELCODE = Trim(Mid(cboMBMRelCode, 1, 1))
             If Len(Trim(cboMBMMaritalStatus)) Then
                !MBMMaritalStatus = Mid(cboMBMMaritalStatus, 1, 1)
             End If
            !MBMFirstMEMNo = Trim(txtFirstMeMNo)
            !MBMFirstPolNo = txtFirstPolNo
            !MBMPrevPolNo = txtMBMPrevPolNo
            !MBMPrevMemNo = txtMBMPrevMEMNo
            !MBMFirstJoinedDate = txtMBMDateJoin
            If InStr(1, cboMBMOccupation, "-") > 0 Then
                !MBMOccupation = Trim(Mid(cboMBMOccupation, 1, InStr(1, cboMBMOccupation, "-") - 1))
            Else
                !MBMOccupation = Trim(cboMBMOccupation)
            End If
     End With
     MBMOthers.Update
        '!MBMUsualDoctor = txtMBMUsualDoctor
        '!MBMClinicName = txtMBMClinicName
        '!MBMClinicTelNo = txtMBMClinicTelNo
        '!MBMClininal = txtMBMClinical
        '!MBMClinicAdd = txtMBMClinicAdd
    
    MemberAmendHis.Close
    Set MemberAmendHis = Nothing
    MBM.Close
    Set MBM = Nothing
    MBMOthers.Close
    Set MBMOthers = Nothing
End Sub

Private Sub SaveSupplementaryAdjustment()
'On Error Resume Next
Dim arrayCnt, arrayCurrRow As Integer
Dim strsql As String
Dim ListCount As Integer
Dim strsql2 As String
Dim MBMCov As New ADODB.Recordset
Dim LastCoveredID As Integer
Dim listloop As Integer
    
    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    LastCoveredID = GetLastCoveredID(txtMBMNumber)
        ListCount = lstCovAdd.listitems.Count
        For listloop = 1 To ListCount
        If lstCovAdd.listitems(listloop).SubItems(11) = "E" Then
            With MBMCov
                MBMCov.Filter = ("MBMCCoverID = '" & lstCovAdd.listitems(listloop).SubItems(10) & "'")
                  !MBMCName = lstCovAdd.listitems(listloop).text
                  !MBMCSalutation = lstCovAdd.listitems(listloop).SubItems(1)
                  !MBMCICBCPPNo = lstCovAdd.listitems(listloop).SubItems(2)
                  !MBMCDOB = lstCovAdd.listitems(listloop).SubItems(3)
                  !MBMCSex = lstCovAdd.listitems(listloop).SubItems(4)
                  !MBMCAdultChild = lstCovAdd.listitems(listloop).SubItems(5)
                  !MBMCRELCode = Mid(lstCovAdd.listitems(listloop).SubItems(6), 1, 1)
                  !MBMCBloodType = lstCovAdd.listitems(listloop).SubItems(7)
                  !MBMCAllergic = lstCovAdd.listitems(listloop).SubItems(8)
                  !MBMCExclusion = lstCovAdd.listitems(listloop).SubItems(9)
                  If Len(lstCovAdd.listitems(listloop).SubItems(12)) Then
                    !MBMCAnnualLimit = lstCovAdd.listitems(listloop).SubItems(12)
                  End If
                  lstCovAdd.listitems(listloop).SubItems(11) = ""
                MBMCov.Update
            End With
            MBMCov.MoveNext
           'For arrayCurrRow = 1 To 10
           '     If lstString(arrayCurrRow) = 0 Then
           '         lstString(arrayCurrRow) = listloop
           '         Exit For
           '     End If
           ' Next

      End If
    Next listloop
    MBMCov.Close
    Set MBMCov = Nothing
    
End Sub

Private Sub DeleteSupplementary()
Dim MBM As New ADODB.Recordset
Dim strsql As String
Dim listrecord As Integer
Dim listloop As Integer
Dim ListCount As Integer
Dim DEL As New ADODB.Recordset
Dim LastCoveredID As Integer
Dim SelectAge As String
Dim strdelcov As String
Dim MBMCov As New ADODB.Recordset
Dim ForLoop As Integer
Dim planstring
Dim MBMCovPersons As New ADODB.Recordset


strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    LastCoveredID = GetLastCoveredID(txtMBMNumber)
        ListCount = lstCovAdd.listitems.Count
        For listloop = 1 To ListCount
        If lstCovAdd.listitems(listloop).SubItems(11) = "R" Then
            strdelcov = "Delete from MBMCoveredPersons Where MBMCNumber ='" & txtMBMNumber & "' And MBMCCoverID ='" & lstCovAdd.listitems(listloop).SubItems(10) & "'"
            MBMCovPersons.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic
            'lstCovAdd.ListItems(listloop).SubItems(11) = "W"
            'MBMCovPersons.MoveNext
      End If
    Next listloop
    
    Call DelSuppLoop
End Sub

Private Sub AddSupplementaryAdjustment()
    Dim listloop As Integer
    Dim ListCount As Integer
    Dim strsql As String
    Dim DEL As New ADODB.Recordset
    Dim LastCoveredID As Integer
    Dim SelectAge As String
    Dim strdelcov As String
    Dim MBMCov As New ADODB.Recordset
    Dim ForLoop As Integer
    Dim planstring
    
    strsql = "Select * from MBMCoveredPersons "
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        LastCoveredID = GetLastCoveredID(txtMBMNumber)
        ListCount = lstCovAdd.listitems.Count
        For listloop = 1 To ListCount
        If lstCovAdd.listitems(listloop).SubItems(11) = "N" Then
              MBMCov.AddNew
              With MBMCov
                  SelectAge = GetAge(txtMBMPayorEffDate, lstCovAdd.listitems(listloop).SubItems(3))
                  !MBMCAGEBand = Mid(GetAgeBand("S", SelectAge), 1, 2)
                  !MBMCStatus = Mid(cboStatus, 1, 1)
                  !MBMCBordxDate = txtMBMBordxDate
                  !MBMCNumber = txtMBMNumber
                  !MBMCSalutation = lstCovAdd.listitems(listloop).SubItems(1)
                  !MBMCName = lstCovAdd.listitems(listloop).text
                  !MBMCICBCPPNo = lstCovAdd.listitems(listloop).SubItems(2)
                  !MBMCDOB = lstCovAdd.listitems(listloop).SubItems(3)
                  !MBMCSex = lstCovAdd.listitems(listloop).SubItems(4)
                  !MBMCAdultChild = lstCovAdd.listitems(listloop).SubItems(5)
                  !MBMCRELCode = Mid(lstCovAdd.listitems(listloop).SubItems(6), 1, 1)
                  !MBMCBloodType = lstCovAdd.listitems(listloop).SubItems(7)
                  !MBMCAllergic = lstCovAdd.listitems(listloop).SubItems(8)
                  !MBMCExclusion = lstCovAdd.listitems(listloop).SubItems(9)
                  !MBMCMemberType = "C"
                  !MBMCCoverID = LastCoveredID
                  lstCovAdd.listitems(listloop).SubItems(11) = ""
                  LastCoveredID = LastCoveredID + 1
                  
             MBMCov.Update
             End With
      End If
    Next listloop
    
    MBMCov.Close
    Set MBMCov = Nothing
End Sub


' ################ START Key press effect ###################

Private Sub cboLGNCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboLGNCode.text, cboLGNCode.SelStart) + Chr(KeyAscii) + Mid(cboLGNCode.text, cboLGNCode.SelStart + cboLGNCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboLGNCode.ListCount - 1
            
            If NewText = LCase(Left(cboLGNCode.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboLGNCode.list(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboLGNCode.text = ValidValue
            cboLGNCode.SelStart = 0
            cboLGNCode.SelLength = Len(ValidValue)
        End If
    End If
End Sub


Private Sub cboMBMAlcohol_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMAlcohol.text, cboMBMAlcohol.SelStart) + Chr(KeyAscii) + Mid(cboMBMAlcohol.text, cboMBMAlcohol.SelStart + cboMBMAlcohol.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMAlcohol.ListCount - 1
            
            If NewText = LCase(Left(cboMBMAlcohol.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMAlcohol.list(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboMBMAlcohol.text = ValidValue
            cboMBMAlcohol.SelStart = 0
            cboMBMAlcohol.SelLength = Len(ValidValue)
        End If
    End If
End Sub


Private Sub cboMBMSmoking_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMSmoking.text, cboMBMSmoking.SelStart) + Chr(KeyAscii) + Mid(cboMBMSmoking.text, cboMBMSmoking.SelStart + cboMBMSmoking.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMSmoking.ListCount - 1
            
            If NewText = LCase(Left(cboMBMSmoking.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMSmoking.list(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboMBMSmoking.text = ValidValue
            cboMBMSmoking.SelStart = 0
            cboMBMSmoking.SelLength = Len(ValidValue)
        End If
    End If
End Sub



Private Sub cboMBMBlood_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMBlood.text, cboMBMBlood.SelStart) + Chr(KeyAscii) + Mid(cboMBMBlood.text, cboMBMBlood.SelStart + cboMBMBlood.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMBlood.ListCount - 1
            
            If NewText = LCase(Left(cboMBMBlood.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMBlood.list(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboMBMBlood.text = ValidValue
            cboMBMBlood.SelStart = 0
            cboMBMBlood.SelLength = Len(ValidValue)
        End If
    End If
End Sub



Private Sub cboHLTCode_KeyPress(KeyAscii As Integer)

On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboHLTCode.text, cboHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode.text, cboHLTCode.SelStart + cboHLTCode.SelLength + 1))

        ValidCount = 0
        ValidValue = ""

        For i = 0 To cboHLTCode.ListCount - 1

            If NewText = LCase(Left(cboHLTCode.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboHLTCode.list(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboHLTCode.text = ValidValue
            cboHLTCode.SelStart = 0
            cboHLTCode.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub cboPAYCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboPAYCode.text, cboPAYCode.SelStart) + Chr(KeyAscii) + Mid(cboPAYCode.text, cboPAYCode.SelStart + cboPAYCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPAYCode.ListCount - 1
            
            If NewText = LCase(Left(cboPAYCode.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPAYCode.list(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboPAYCode.text = ValidValue
            cboPAYCode.SelStart = 0
            cboPAYCode.SelLength = Len(ValidValue)
        End If
    End If
End Sub


Private Sub cboMBMRelCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMRelCode.text, cboMBMRelCode.SelStart) + Chr(KeyAscii) + Mid(cboMBMRelCode.text, cboMBMRelCode.SelStart + cboMBMRelCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMRelCode.ListCount - 1
            
            If NewText = LCase(Left(cboMBMRelCode.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMRelCode.list(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboMBMRelCode.text = ValidValue
            cboMBMRelCode.SelStart = 0
            cboMBMRelCode.SelLength = Len(ValidValue)
        End If
    End If
 
End Sub

Private Sub cboBRNCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboBRNCode.text, cboBRNCode.SelStart) + Chr(KeyAscii) + Mid(cboBRNCode.text, cboBRNCode.SelStart + cboBRNCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboBRNCode.ListCount - 1
            
            If NewText = LCase(Left(cboBRNCode.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboBRNCode.list(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboBRNCode.text = ValidValue
            cboBRNCode.SelStart = 0
            cboBRNCode.SelLength = Len(ValidValue)
        End If
    End If
End Sub



Private Sub cboCovRelationship_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboCovRelationship.text, cboCovRelationship.SelStart) + Chr(KeyAscii) + Mid(cboCovRelationship.text, cboCovRelationship.SelStart + cboCovRelationship.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboCovRelationship.ListCount - 1
            
            If NewText = LCase(Left(cboCovRelationship.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboCovRelationship.list(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboCovRelationship.text = ValidValue
            cboCovRelationship.SelStart = 0
            cboCovRelationship.SelLength = Len(ValidValue)
        End If
    End If
            
End Sub

Private Sub cboCovSalutation_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboCovSalutation.text, cboCovSalutation.SelStart) + Chr(KeyAscii) + Mid(cboCovSalutation.text, cboCovSalutation.SelStart + cboCovSalutation.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboCovSalutation.ListCount - 1
            
            If NewText = LCase(Left(cboCovSalutation.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboCovSalutation.list(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboCovSalutation.text = ValidValue
            cboCovSalutation.SelStart = 0
            cboCovSalutation.SelLength = Len(ValidValue)
        End If
    End If
            
End Sub

Private Sub cboCovSex_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboCovSex.text, cboCovSex.SelStart) + Chr(KeyAscii) + Mid(cboCovSex.text, cboCovSex.SelStart + cboCovSex.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboCovSex.ListCount - 1
        
        If NewText = LCase(Left(cboCovSex.list(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCovSex.list(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboCovSex.text = ValidValue
               cboCovSex.SelStart = 0
               cboCovSex.SelLength = Len(ValidValue)
             End If
        End If

End Sub


Private Sub cboINSCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINSCode.text, cboINSCode.SelStart) + Chr(KeyAscii) + Mid(cboINSCode.text, cboINSCode.SelStart + cboINSCode.SelLength + 1))
  
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINSCode.ListCount - 1
        
        If NewText = LCase(Left(cboINSCode.list(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINSCode.list(i)
        
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             
             If ValidCount = 1 Then
               cboINSCode.text = ValidValue
               cboINSCode.SelStart = 0
               cboINSCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNCode.text, cboPLNCode.SelStart) + Chr(KeyAscii) + Mid(cboPLNCode.text, cboPLNCode.SelStart + cboPLNCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNCode.ListCount - 1

        If NewText = LCase(Left(cboPLNCode.list(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNCode.list(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

           If ValidCount = 1 Then
             cboPLNCode.text = ValidValue
             cboPLNCode.SelStart = 0
             cboPLNCode.SelLength = Len(ValidValue)
            End If
 
        End If
End Sub

Private Sub cboRACCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboRACCode.text, cboRACCode.SelStart) + Chr(KeyAscii) + Mid(cboRACCode.text, cboRACCode.SelStart + cboRACCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboRACCode.ListCount - 1
            
            If NewText = LCase(Left(cboRACCode.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboRACCode.list(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboRACCode.text = ValidValue
                   cboRACCode.SelStart = 0
                   cboRACCode.SelLength = Len(ValidValue)
                 End If
            
            End If

End Sub

Private Sub cboSRVCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboSRVCode.text, cboSRVCode.SelStart) + Chr(KeyAscii) + Mid(cboSRVCode.text, cboSRVCode.SelStart + cboSRVCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboSRVCode.ListCount - 1
            
            If NewText = LCase(Left(cboSRVCode.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboSRVCode.list(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                If ValidCount = 1 Then
                  cboSRVCode.text = ValidValue
                  cboSRVCode.SelStart = 0
                  cboSRVCode.SelLength = Len(ValidValue)
                End If
            
            End If
End Sub


Private Sub cboMBMSalutation_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMSalutation.text, cboMBMSalutation.SelStart) + Chr(KeyAscii) + Mid(cboMBMSalutation.text, cboMBMSalutation.SelStart + cboMBMSalutation.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMSalutation.ListCount - 1
            
            If NewText = LCase(Left(cboMBMSalutation.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMSalutation.list(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboMBMSalutation.text = ValidValue
                   cboMBMSalutation.SelStart = 0
                   cboMBMSalutation.SelLength = Len(ValidValue)
                 End If
            
            End If
            
End Sub


Private Sub cboMBMSex_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMSex.text, cboMBMSex.SelStart) + Chr(KeyAscii) + Mid(cboMBMSex.text, cboMBMSex.SelStart + cboMBMSex.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMSex.ListCount - 1
            
            If NewText = LCase(Left(cboMBMSex.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMSex.list(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboMBMSex.text = ValidValue
                   cboMBMSex.SelStart = 0
                   cboMBMSex.SelLength = Len(ValidValue)
                 End If
            End If
End Sub

Private Sub cboNATCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboNATCode.text, cboNATCode.SelStart) + Chr(KeyAscii) + Mid(cboNATCode.text, cboNATCode.SelStart + cboNATCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboNATCode.ListCount - 1
            
            If NewText = LCase(Left(cboNATCode.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboNATCode.list(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboNATCode.text = ValidValue
                   cboNATCode.SelStart = 0
                   cboNATCode.SelLength = Len(ValidValue)
                 End If
            End If
End Sub

Private Sub cboMBMMaritalStatus_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMMaritalStatus.text, cboMBMMaritalStatus.SelStart) + Chr(KeyAscii) + Mid(cboMBMMaritalStatus.text, cboMBMMaritalStatus.SelStart + cboMBMMaritalStatus.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMMaritalStatus.ListCount - 1
            
            If NewText = LCase(Left(cboMBMMaritalStatus.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMMaritalStatus.list(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboMBMMaritalStatus.text = ValidValue
                   cboMBMMaritalStatus.SelStart = 0
                   cboMBMMaritalStatus.SelLength = Len(ValidValue)
                 End If
            End If
End Sub

Private Sub cboMBMOccupation_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMOccupation.text, cboMBMOccupation.SelStart) + Chr(KeyAscii) + Mid(cboMBMOccupation.text, cboMBMOccupation.SelStart + cboMBMOccupation.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMOccupation.ListCount - 1
            
            If NewText = LCase(Left(cboMBMOccupation.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMOccupation.list(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboMBMOccupation.text = ValidValue
                   cboMBMOccupation.SelStart = 0
                   cboMBMOccupation.SelLength = Len(ValidValue)
                 End If
            End If
End Sub



Private Sub txtMBMPayorExpDate_KeyPress(KeyAscii As Integer)
Dim MBMMCOFees As String
Dim MBMPremium As String
Dim AccessFees As String
    
    MBMMCOFees = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), txtMCOFees, "#,##0.00")
    MBMPremium = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), txtMBMPremium, "#,##0.00")
    AccessFees = InsertProvider(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), txtMBMProviderCharges, "#,##0.00")

    txtMCOFees = MBMMCOFees
    txtMBMPremium = MBMPremium
    txtMBMProviderCharges = AccessFees
End Sub

Private Sub txtMBMWorkNature_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(txtMBMWorkNature.text, txtMBMWorkNature.SelStart) + Chr(KeyAscii) + Mid(txtMBMWorkNature.text, txtMBMWorkNature.SelStart + txtMBMWorkNature.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To txtMBMWorkNature.ListCount - 1
            
            If NewText = LCase(Left(txtMBMWorkNature.list(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = txtMBMWorkNature.list(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   txtMBMWorkNature.text = ValidValue
                   txtMBMWorkNature.SelStart = 0
                   txtMBMWorkNature.SelLength = Len(ValidValue)
                 End If
            End If
End Sub

Private Sub SuspendPrincipal()
    Dim MemberAmendHis As New ADODB.Recordset
    Dim MBM As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
    Dim strsqldel As String
    Dim DEL As New ADODB.Recordset
    Dim SelectAge As String
    
    strsql = "Select * from MBM Where MBMNumber = '" & txtMBMNumber & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MemberAmendHis.AddNew
    With MemberAmendHis
        !MBMNumber = MBM!MBMNumber
        !MBMPolicyNo = MBM!MBMPolicyNo
        !MBMIcBcPp = MBM!MBMIcBcPp
        !MBMName = MBM!MBMName
        !MBMExpiryDate = MBM!MBMPayorEXPDate
        !MBMPlaNCode = MBM!PLNCode
        !MBMAmendEdtType = Mid(cboEndorseStatus, 1, 1)
        If IsDate(txtDate) Then
            !MBMAmendDate = txtDate
        End If
        !MBMAmendDate = txtDate
        '!MBMValueDate = Date
        MemberAmendHis.Update
    End With

    
    'MemberHis.Open "MBA", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'Do Until MBM.EOF
                With MBM
                    !MBMLastEndorseStatus = Mid(cboEndorseStatus, 1, 1)
                    !MBMStatus = "S"
                    '!MBMSuspendDate = txtDate
                    !MBMLastAdjustmentDate = txtDate
                    MBM.Update
                End With
        'MBM.MoveNext
    'Loop
    MBM.Close
    cboStatus.ListIndex = 3

    'Call EntriesEnable(False)
    'txtMBMEndorseStatus.Enabled = False
    'cmdSave.Enabled = False
    'txtMBACancelDate.Enabled = False
End Sub

Private Sub SuspendSupplementary()
    Dim arrayCnt, arrayCurrRow As Integer
    Dim MemberAmendHis As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
    Dim ListCount As Integer
    Dim listloop As Integer
    Dim LastCoveredID As Integer
    Dim lstString()
    Dim item
    lstString = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    LastCoveredID = GetLastCoveredID(txtMBMNumber)

    ListCount = lstCovAdd.listitems.Count
    For listloop = 1 To ListCount

                With MBMCov
                    MBMCov.Filter = ("MBMCCoverID = '" & lstCovAdd.listitems(listloop).SubItems(10) & "'")
                    !MBMCStatus = "S"
                    !MBMCLastUpdateDate = txtDate
                    MBMCov.Update
                End With
            
            For arrayCurrRow = 1 To 10
                If lstString(arrayCurrRow) = 0 Then
                    lstString(arrayCurrRow) = listloop
                    Exit For
                End If
            Next
        Next listloop
    MBMCov.Close
    Set MBMCov = Nothing
    

    'Call EntriesEnable(False)
    'txtMBMEndorseStatus.Enabled = False
    'cmdSave.Enabled = False
    'txtMBACancelDate.Enabled = False
End Sub

Private Sub ReactivateSupplementary()
    Dim arrayCnt, arrayCurrRow As Integer
    Dim MemberHis As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
    Dim ListCount As Integer
    Dim listloop As Integer
    Dim LastCoveredID As Integer
    Dim lstString()
    Dim item
    lstString = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'MemberHis.Open "MBA", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    LastCoveredID = GetLastCoveredID(txtMBMNumber)

    ListCount = lstCovAdd.listitems.Count
    For listloop = 1 To ListCount

                With MBMCov
                    MBMCov.Filter = ("MBMCCoverID = '" & lstCovAdd.listitems(listloop).SubItems(10) & "'")
                    !MBMCStatus = "A"
                    !MBMCLastUpdateDate = txtDate
                    MBMCov.Update
                End With
            
            For arrayCurrRow = 1 To 10
                If lstString(arrayCurrRow) = 0 Then
                    lstString(arrayCurrRow) = listloop
                    Exit For
                End If
            Next
        Next listloop
    MBMCov.Close
    Set MBMCov = Nothing
    

    'Call EntriesEnable(False)
    'txtMBMEndorseStatus.Enabled = False
    'cmdSave.Enabled = False
    'txtMBACancelDate.Enabled = False
End Sub

Private Sub ReactivatePrincipal()
    Dim MemberAmendHis As New ADODB.Recordset
    Dim MBM As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
    Dim strsqldel As String
    Dim DEL As New ADODB.Recordset
    Dim SelectAge As String
    
    strsql = "Select * from MBM Where MBMNumber = '" & txtMBMNumber & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic

    MemberAmendHis.AddNew
    With MemberAmendHis
        !MBMNumber = MBM!MBMNumber
        !MBMPolicyNo = MBM!MBMPolicyNo
        !MBMIcBcPp = MBM!MBMIcBcPp
        !MBMName = MBM!MBMName
        !MBMExpiryDate = MBM!MBMPayorEXPDate
        !MBMPlaNCode = MBM!PLNCode
        !MBMAmendEdtType = Mid(cboEndorseStatus, 1, 1)
        If IsDate(txtDate) Then
            !MBMAmendDate = txtDate
        End If
        '!MBMValueDate = Date
        MemberAmendHis.Update
    End With

    
    'MemberHis.Open "MBA", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'Do Until MBM.EOF
                With MBM
                    !MBMLastEndorseStatus = "Y"
                    !MBMStatus = "A"
                    '!MBMSuspendDate = txtDate
                    !MBMLastAdjustmentDate = txtDate
                    MBM.Update
                End With
        'MBM.MoveNext
    'Loop
    MBM.Close
    cboStatus.ListIndex = 0
    'Call EntriesEnable(False)
    'txtMBMEndorseStatus.Enabled = False
    'cmdSave.Enabled = False
    'txtMBACancelDate.Enabled = False
End Sub

Private Sub InsertNewRegistrationSub()
    Dim MBM As New ADODB.Recordset
    Dim MBMOthers As New ADODB.Recordset
    Dim MBMCoveredPersons As New ADODB.Recordset
    Dim planstring
    Dim listloop As Integer
    Dim listrecord As Integer
    Dim MembershipNo As String
    Dim SelectAge As String
    Dim RecordSaved
    Dim completeMBM, completeMBMOthers, completeMBMC As Boolean
    
    planstring = cboPLNCode
    ' INsert into Membership table
    MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    ' Begin transaction
    medixmasterconn.beginTrans
    startTrans = True
    MembershipNo = Trim(Mid(cboPAYCode, 1, 2)) & Trim(Mid(cboPLNCode, 1, 2)) & Trim(Mid(cboINSCode, 1, 1)) & Mid(BusinessYear, 3, 2) & GenerateMembershipNo(Trim(Mid(cboPAYCode, 1, 2)), Trim(Mid(cboPLNCode, 1, 2)), Trim(Mid(cboINSCode, 1, 1)))
    
    MBM.AddNew
         With MBM
             !MBMNumber = MembershipNo
             !MBMPolicyNo = txtMBMPolNo
             !MBMIcBcPp = txtMBMIcBcPP
             !MBMName = txtMBMName
             !MBMSalutation = cboMBMSalutation
             !HLTCode = Trim(Mid(Trim(cboHLTCode), 1, 2))
             !PAYCode = Trim(Mid(cboPAYCode, 1, 2))
             !MBMAddress1 = txtMBMAdd1
             !MBMAddress2 = txtMBMAdd2
             !MBMAddress3 = txtMBMAdd3
             !CTYCode = cboCTYCode
             !MBMPostCode = txtMBMPostCode
             !MBMTelnoH = txtMBMTelNoH
             !MBMTelNoM = txtMBMTelNoM
             !MBMTelNoO = txtMBMTelNoO
             !MBMEmail = txtMBMEmail
             '!MBMReceiptDate = CDate(txtMBMBordxDate)
             If Len(Trim(!BRNCode)) Then
                !BRNCode = Trim(Mid(cboBRNCode, 1, 4))
             End If
             txtMBMBirthDate.Mask = ""
             !MBMDOB = txtMBMBirthDate
             !NATCode = Trim(Mid(cboNATCode, 1, 2))
             !RACCode = Trim(Mid(cboRACCode, 1, 2))
             !MBMSex = cboMBMSex
             !MBMPayorEffDate = CDate(txtMBMPayorEffDate)
             '!MBMDateJoin = CDate(txtMBMDateJoin)
             !MBMPayorEXPDate = CDate(txtMBMPayorExpDate)
             !AgeCode = Trim(Mid(txtAGECode, 1, 2))
             !INSCode = Trim(Mid(cboINSCode, 1, 2))
             !PLNCode = Trim(Mid(cboPLNCode, 1, 2))
             !MBMAdultChild = txtMBMAdultChild
             'txtDate.Mask =
             '!MBMInsertionDate = txtDate
             'Call CallLimit
             '!MBMAvailableLimit = txtAvailableAnnualLimit
             !SRVCodeList = Trim(Mid(cboSRVCode, 1, InStr(1, cboSRVCode, "-") - 1))
             !STACode = Trim(cboSTACode)
             !LGNCode = Trim(Mid(cboLGNCode, 1, 2))
             'txtMBMBordxDate.Mask = ""
             '!MBMBordxDate = CDate(txtMBMBordxDate)
             !MBMTakeOver = Mid(ChkStr(cboMBMTakeOver), 1, 1)
             '!MBMRenewFlag = cboMBMRenewal
             !BRNCode = cboBRNCode
             !MBMStatus = "A"
             !MBMValueDate = Date
             !MBMPolicyYear = 1
             !MBMTakeOver = Trim(Mid(cboMBMTakeOver, 1, 1))
             !MBMMemberType = Trim(Mid(Trim(cboMBMRelCode), 1, 1))
             
            '!MBMBatchNo = txtMBMBatchNo
             Call MemberCallAge
             Call MemberCallFees
             Call ProviderCallFees
             Call InsertMCO
             Call InsertPrem
             Call InsertProviderFees
             .Update
     End With
    completeMBM = True
    ' INsert into Membership Other details table
     MBMOthers.Open "MBMOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
     MBMOthers.AddNew
     With MBMOthers
             !MBMNumber = MembershipNo
             !MBMWeight = txtMBMWeight
             !MBMHeight = txtMBMHeight
             !MBMOccupation = Trim(Mid(cboMBMOccupation, 1, 2))
             !MBMWorkNature = txtMBMWorkNature
             !MBMExclusion = txtMBMExclusion
             !MBMAllergic = txtMBMAllergic
             !MBMGroupCompany = txtMBMGroupCompany
             !MBMEmployeeNo = txtMBMEmployeeNo
             !MBMBlood = cboMBMBlood
             !MBMSmoking = cboMBMSmoking
             !MBMAlcohol = cboMBMAlcohol
             !MBMAgentCode = txtMBMAgentCode
             
             If Len(Trim(cboMBMMaritalStatus)) Then
                !MBMMaritalStatus = Mid(cboMBMMaritalStatus, 1, 1)
             End If
             !MBMAccessFee = txtMBMProviderCharges
             !MBMPremium = txtMBMPremium
             !MBMMCOFee = txtMCOFees
             'If ChkStr(cboMBMRenewal) = "Y" Then
             '   !MBMFirstMEMNo = txtFirstMeMNo
             '   !MBMFirstPolNo = txtFirstPolNo
             '   !MBMPrevPolNo = txtMBMPrevPolNo
             '   !MBMPrevMemNo = txtMBMPrevMEMNo
             '   !MBMFirstJoinedDate = txtMBMDateJoin
             'Else
             '   !MBMFirstMEMNo = MembershipNo
             '   !MBMFirstPolNo = txtMBMPolicyNo
             '   !MBMFirstJoinedDate = txtMBMDateJoin
             'End If
             
     End With
     MBMOthers.Update
     completeMBMOthers = True
     
    If Trim(MBM!INSCode) = "I" Or Trim(MBM!INSCode) = "G" Then
        completeMBMC = False
     ElseIf Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
     listrecord = lstCovAdd.listitems.Count
    ' Insert into Membership Covered Person's table
     If listrecord <> 0 Then
         MBMCoveredPersons.Open "MBMCOveredPersons", medixmasterconn, adOpenKeyset, adLockOptimistic
         For listloop = 1 To listrecord
             MBMCoveredPersons.AddNew
             With MBMCoveredPersons
                 !MBMCNumber = MembershipNo
                 !MBMCSalutation = lstCovAdd.listitems(listloop).SubItems(1)
                 !MBMCName = lstCovAdd.listitems(listloop).text
                 !MBMCICBCPPNo = lstCovAdd.listitems(listloop).SubItems(2)
                 !MBMCDOB = lstCovAdd.listitems(listloop).SubItems(3)
                 SelectAge = GetAge(txtMBMPayorEffDate, lstCovAdd.listitems(listloop).SubItems(3))
                 !MBMCAGEBand = Mid(GetAgeBand(Trim(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1)), SelectAge), 1, 2)
                 !MBMCSex = lstCovAdd.listitems(listloop).SubItems(4)
                 !MBMCAdultChild = lstCovAdd.listitems(listloop).SubItems(5)
                 !MBMCRELCode = Mid(lstCovAdd.listitems(listloop).SubItems(6), 1, 1)
                 !MBMCBloodType = lstCovAdd.listitems(listloop).SubItems(7)
                 !MBMCAllergic = lstCovAdd.listitems(listloop).SubItems(8)
                 !MBMCExclusion = lstCovAdd.listitems(listloop).SubItems(9)
                 If Len(Trim(lstCovAdd.listitems(listloop).SubItems(12))) Then
                    !MBMCAnnualLimit = lstCovAdd.listitems(listloop).SubItems(12)
                 End If
                 !MBMCCoverID = listloop
                 .Update
            End With
         Next listloop
         MBMCoveredPersons.Close
         Set MBMCoveredPersons = Nothing
    End If
   
    completeMBMC = True
     End If
          
    'If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
       If completeMBM = True And completeMBMC = True And completeMBMOthers = True Then
            ' commit changes
            medixmasterconn.CommitTrans
        'End If
    'ElseIf Trim(MBM!INSCode) = "I" Or Trim(MBM!INSCode) = "G" Then
       ElseIf completeMBM = True And completeMBMOthers = True And completeMBMC = False Then
            medixmasterconn.CommitTrans
        'End If
        Else
            'Roll back
            medixmasterconn.RollbackTrans
        End If
    startTrans = False
    MBM.Close
    Set MBM = Nothing
   
    MBMOthers.Close
    Set MBMOthers = Nothing
     
     
     'cboMBMRenewal.Enabled = True
     MsgBox "Record Saved, Your Membership No is : " & MembershipNo, vbOKOnly + vbInformation, DialogBoxTitle
     ' Clering previous record
     'txtMBMPayorExpDate.Enabled = True
     'Call ClearEntries
     'txtMBMPayorExpDate.Enabled = False
    
     SSTab1.Tab = 0
     'cboMBMRenewal.SetFocus
End Sub

Private Sub InsertMCO()
Dim MBMInsertMCOFees As Double
Dim MBMMCOFess As Double
Dim PayExpDate As String

MBMMCOFess = txtMCOFees
PayExpDate = txtMBMPayorExpDate

    MBMInsertMCOFees = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMCOFees, "#,##0.00"))
    txtMCOFees = ""
    txtMCOFees = Format(MBMInsertMCOFees, "#,##0.00")
End Sub

Private Sub InsertPrem()
Dim MBMInsertPrem As Double
Dim MBMPrem As Double
Dim PayExpDate As String

MBMPrem = txtMCOFees
PayExpDate = txtMBMPayorExpDate

    MBMInsertPrem = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMBMPremium, "#,##0.00"))
    txtMBMPremium = ""
    txtMBMPremium = Format(MBMInsertPrem, "#,##0.00")
End Sub

Private Sub ClearEntries()
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.text = ""
        ElseIf TypeOf ctl Is ComboBox Then
            cboSRVCode.text = "EA - EUROPE ASSISTANCE"
            cboHLTCode.text = ""
            'cboPatientList.Clear
        ElseIf TypeOf ctl Is ListView Then
            ctl.listitems.clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Mask = ""
            ctl.text = ""
            ctl.Mask = "##/##/####"
        ElseIf TypeOf ctl Is Label Then
            txtMBMPremium = ""
            txtMBMPremiumRed = ""
            txtMBMPremiumEff = ""
            txtMCOFees = ""
            txtMBMMCORed = ""
            txtMCOFeesEff = ""
            txtMBMProviderCharges = ""
            txtMBMProviderRedCharges = ""
            txtMBMProviderEffCharges = ""
            txtAnnualLimit = ""
            txtAvailableAnnualLimit = ""
            txtMBMCaseCharges = ""
        End If
    Next ctl
    '  Call ComboListing
      'cmdSaveKIV.Enabled = True
      'cmdGenerate.Enabled = True
      
End Sub

Private Sub DeletePrincipals()
Dim strsqlMBM As String
Dim MBM As New ADODB.Recordset
Dim MBMArchieve As New ADODB.Recordset
Dim strdelPrincipal As String
Dim DelPrincipal As New ADODB.Recordset
Dim DelOthers As String
Dim MBMOthers As New ADODB.Recordset

    
    strsqlMBM = "SELECT * From MBM where MBMNumber = '" & txtMBMNumber & "'"
    MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
             
             MBM!MBMStatus = "D"
             MBM!MBMValueDate = Date
             'txtDate.Mask = "__/__/____"
             MBM!MBMENDtEffDate = txtDate
             
        MBM.Update
            cboStatus.ListIndex = 2
        'txtMBMNumber = MBM!MBMNumber
        'txtMBMType = MBM!MBMMemberType
        
        If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
            Call DeleteSupplementaryPol
        End If

    MBM.Close
    Set MBM = Nothing
End Sub

Private Sub DeleteSupplementaryPol()
Dim strsqlMBMCov As String
Dim MBMCov As New ADODB.Recordset
        

      strsqlMBMCov = "SELECT * From MBMCoveredPersons  Where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
      MBMCov.Open strsqlMBMCov, medixmasterconn, adOpenKeyset, adLockOptimistic
      
        Do Until MBMCov.EOF
                With MBMCov
                    !MBMCStatus = "D"
                End With
                MBMCov.Update
        MBMCov.MoveNext
        Loop

    
    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub DelSuppLoop()
Dim listrecord As Integer
Dim listloop As Integer
Dim ListCount As Integer
Dim ForLoop As Integer
Dim LastCoveredID As Integer

        ListCount = lstCovAdd.listitems.Count
        For listloop = 1 To ListCount
        If listloop > ListCount Then
            Exit Sub
       ElseIf listloop <= ListCount Then
       If lstCovAdd.listitems(listloop).SubItems(11) = "R" Then
            lstCovAdd.listitems.Remove (listloop)
            'listloop = listloop - 1
            ListCount = ListCount - 1
          End If
          End If
        'If listloop <> ListCount Then
            'Exit Sub
        'Else
    Next listloop
      
    'Else
    'Exit Sub
        'End If
        
End Sub

Public Sub InsertProviderFees()
Dim MBMInsertProviderFees As Double
'Dim MBMProviderFees As Double
'Dim PayExpDate As String

'MBMPrem = txtMCOFees
'PayExpDate = txtMBMPayorExpDate

    MBMInsertProviderFees = InsertProvider(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMCOFees, "#,##0.00"))
    txtMBMProviderCharges = ""
    txtMBMProviderCharges = Format(MBMInsertProviderFees, "#,##0.00")

End Sub

