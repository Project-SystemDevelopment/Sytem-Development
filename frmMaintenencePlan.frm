VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMaintenancePlan 
   Caption         =   "Plan Maintenance"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   12255
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      Height          =   705
      Left            =   360
      TabIndex        =   81
      Top             =   7080
      Width           =   10935
      Begin VB.TextBox temptxtEffDate 
         Height          =   285
         Left            =   360
         TabIndex        =   130
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5040
         TabIndex        =   77
         Top             =   285
         Width           =   840
      End
      Begin VB.TextBox txtPrmCode 
         Height          =   285
         Left            =   120
         TabIndex        =   114
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   840
         TabIndex        =   72
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   7320
         TabIndex        =   78
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   4200
         TabIndex        =   76
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   73
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   74
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   8160
         TabIndex        =   79
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9000
         TabIndex        =   80
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3360
         TabIndex        =   75
         Top             =   285
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7725
      Left            =   120
      TabIndex        =   83
      Top             =   240
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   13626
      _Version        =   393216
      Tabs            =   11
      Tab             =   6
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "HealthType"
      TabPicture(0)   =   "frmMaintenencePlan.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(2)"
      Tab(0).Control(1)=   "lstHLTList"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Insured Type"
      TabPicture(1)   =   "frmMaintenencePlan.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(3)"
      Tab(1).Control(1)=   "lstINSList"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Plan Code"
      TabPicture(2)   =   "frmMaintenencePlan.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(1)"
      Tab(2).Control(1)=   "lstPlanList"
      Tab(2).Control(2)=   "Frame2"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Premium"
      TabPicture(3)   =   "frmMaintenencePlan.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Frame1(0)"
      Tab(3).Control(2)=   "lstPmPRMList"
      Tab(3).Control(3)=   "FrameAGEBand"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Annual Limit"
      TabPicture(4)   =   "frmMaintenencePlan.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1(6)"
      Tab(4).Control(1)=   "lstAnnualList"
      Tab(4).Control(2)=   "Frame4"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "MCO Fee"
      TabPicture(5)   =   "frmMaintenencePlan.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame5"
      Tab(5).Control(1)=   "lstMCOList"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Special Note"
      TabPicture(6)   =   "frmMaintenencePlan.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "lstSpecialNoteList"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame1(4)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Annual Limit (Benefit)"
      TabPicture(7)   =   "frmMaintenencePlan.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame1(5)"
      Tab(7).Control(1)=   "lstALB"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Product Category"
      TabPicture(8)   =   "frmMaintenencePlan.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lstProductCat"
      Tab(8).Control(1)=   "Frame3"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Supp. MCO Status"
      TabPicture(9)   =   "frmMaintenencePlan.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame11"
      Tab(9).Control(1)=   "lstMCOStatus"
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "Case Mgmt. Fee"
      TabPicture(10)  =   "frmMaintenencePlan.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame1(7)"
      Tab(10).Control(1)=   "lstMgmtFeeList"
      Tab(10).ControlCount=   2
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Index           =   1
         Left            =   -74760
         TabIndex        =   220
         Top             =   1080
         Width           =   10935
         Begin VB.TextBox TxtGrpCom 
            Height          =   285
            Left            =   1750
            TabIndex        =   285
            Top             =   800
            Width           =   6615
         End
         Begin VB.CommandButton CmdPolicyWording 
            Caption         =   " >"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            TabIndex        =   287
            Top             =   2160
            Width           =   330
         End
         Begin VB.TextBox TxtPolicyWording 
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
            Left            =   1750
            TabIndex        =   286
            Text            =   " "
            Top             =   2160
            Width           =   2220
         End
         Begin VB.Frame Frame9 
            Height          =   1730
            Left            =   8520
            TabIndex        =   248
            Top             =   120
            Width           =   2295
            Begin VB.ComboBox cboDisIndicator 
               Height          =   315
               Left            =   1560
               TabIndex        =   253
               Top             =   1320
               Width           =   615
            End
            Begin VB.ComboBox CboLifeTimeStatus 
               Height          =   315
               Left            =   1560
               Sorted          =   -1  'True
               TabIndex        =   252
               Top             =   1040
               Width           =   615
            End
            Begin VB.ComboBox cboPLNMCO 
               Height          =   315
               Left            =   1560
               Sorted          =   -1  'True
               TabIndex        =   251
               Top             =   750
               Width           =   615
            End
            Begin VB.ComboBox cboPLNPremium 
               Height          =   315
               Left            =   1560
               Sorted          =   -1  'True
               TabIndex        =   250
               Top             =   480
               Width           =   615
            End
            Begin VB.ComboBox cboPLNAnnLimit 
               Height          =   315
               ItemData        =   "frmMaintenencePlan.frx":0134
               Left            =   1560
               List            =   "frmMaintenencePlan.frx":0136
               Sorted          =   -1  'True
               TabIndex        =   249
               Top             =   195
               Width           =   615
            End
            Begin VB.Label Label59 
               Caption         =   "Annual Limit *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   180
               TabIndex        =   258
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label50 
               Caption         =   "Premium  *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   180
               TabIndex        =   257
               Top             =   530
               Width           =   975
            End
            Begin VB.Label Label51 
               Caption         =   "M C O  *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   180
               TabIndex        =   256
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label62 
               Caption         =   "Life Time Status *"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   180
               TabIndex        =   255
               Top             =   1100
               Width           =   1335
            End
            Begin VB.Label Label86 
               Caption         =   "Per Dis. Indicator *"
               Height          =   255
               Left            =   170
               TabIndex        =   254
               Top             =   1400
               Width           =   1455
            End
         End
         Begin VB.ComboBox cboClientPLN 
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
            ItemData        =   "frmMaintenencePlan.frx":0138
            Left            =   5250
            List            =   "frmMaintenencePlan.frx":013A
            Sorted          =   -1  'True
            TabIndex        =   247
            Top             =   2160
            Width           =   1365
         End
         Begin VB.ComboBox cboPolicyWording 
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
            Left            =   1750
            Sorted          =   -1  'True
            TabIndex        =   246
            Top             =   2160
            Width           =   2565
         End
         Begin VB.TextBox txtEndAge 
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
            Left            =   3700
            MaxLength       =   3
            TabIndex        =   245
            Top             =   1830
            Width           =   615
         End
         Begin VB.TextBox txtStartAge 
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
            Left            =   3700
            MaxLength       =   2
            TabIndex        =   244
            Top             =   1560
            Width           =   615
         End
         Begin VB.ComboBox CboTax 
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
            Left            =   7335
            Sorted          =   -1  'True
            TabIndex        =   243
            Top             =   1560
            Width           =   1095
         End
         Begin VB.ComboBox CboPLNMgmtStatus 
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
            Left            =   1750
            Sorted          =   -1  'True
            TabIndex        =   242
            Top             =   1880
            Width           =   615
         End
         Begin VB.ComboBox CboMeal 
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
            Left            =   5250
            Sorted          =   -1  'True
            TabIndex        =   241
            Top             =   1880
            Width           =   1365
         End
         Begin VB.ComboBox CboNursing 
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
            Left            =   1750
            Sorted          =   -1  'True
            TabIndex        =   240
            Top             =   1600
            Width           =   615
         End
         Begin VB.ComboBox CboMRI 
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
            Left            =   7335
            Sorted          =   -1  'True
            TabIndex        =   239
            Top             =   1320
            Width           =   1095
         End
         Begin VB.ComboBox CboSOF 
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
            Left            =   9735
            Sorted          =   -1  'True
            TabIndex        =   238
            Top             =   2130
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox cboProCat 
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
            Left            =   5250
            Sorted          =   -1  'True
            TabIndex        =   237
            Top             =   1610
            Width           =   1365
         End
         Begin VB.ComboBox CboSMPlan 
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
            Left            =   5250
            Sorted          =   -1  'True
            TabIndex        =   236
            Top             =   1320
            Width           =   1365
         End
         Begin VB.TextBox TxtSpecGPeriod 
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
            Left            =   3700
            MaxLength       =   3
            TabIndex        =   235
            Top             =   1320
            Width           =   615
         End
         Begin VB.ComboBox CboSpecGPeriod 
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
            Left            =   1750
            Sorted          =   -1  'True
            TabIndex        =   234
            Top             =   1320
            Width           =   615
         End
         Begin VB.ComboBox CboCoPayment 
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
            Left            =   3700
            Sorted          =   -1  'True
            TabIndex        =   232
            Top             =   1050
            Width           =   615
         End
         Begin VB.ComboBox CboTopupStatus 
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
            Left            =   1750
            Sorted          =   -1  'True
            TabIndex        =   231
            Top             =   1050
            Width           =   615
         End
         Begin VB.ComboBox cboPLNPayor 
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
            Left            =   1750
            Sorted          =   -1  'True
            TabIndex        =   230
            Top             =   520
            Width           =   6540
         End
         Begin VB.ComboBox cboPLNHLTCode 
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
            Left            =   1750
            Sorted          =   -1  'True
            TabIndex        =   229
            Top             =   240
            Width           =   3375
         End
         Begin VB.ComboBox CboNoofPlan 
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
            ItemData        =   "frmMaintenencePlan.frx":013C
            Left            =   6730
            List            =   "frmMaintenencePlan.frx":013E
            Sorted          =   -1  'True
            TabIndex        =   228
            Top             =   240
            Width           =   1680
         End
         Begin VB.CheckBox chkPLNPayor 
            Caption         =   "Check2"
            Height          =   255
            Left            =   8280
            TabIndex        =   227
            Top             =   550
            Width           =   255
         End
         Begin VB.TextBox txtClientPolNo 
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
            Left            =   6720
            TabIndex        =   226
            Top             =   2160
            Width           =   1875
         End
         Begin VB.TextBox txtCopayPerc 
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
            Left            =   5250
            TabIndex        =   225
            Top             =   1080
            Width           =   1365
         End
         Begin VB.ComboBox cmbVIPInd 
            Height          =   315
            ItemData        =   "frmMaintenencePlan.frx":0140
            Left            =   4560
            List            =   "frmMaintenencePlan.frx":014A
            Sorted          =   -1  'True
            TabIndex        =   224
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox txtLGPrivate 
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
            Left            =   1750
            TabIndex        =   223
            Text            =   " "
            Top             =   2520
            Width           =   1875
         End
         Begin VB.TextBox txtFileName 
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
            Left            =   6720
            TabIndex        =   222
            Text            =   " "
            Top             =   2640
            Width           =   3435
         End
         Begin VB.CommandButton cmdFileLocation 
            Caption         =   " >"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10320
            TabIndex        =   221
            Top             =   2640
            Width           =   320
         End
         Begin MSMask.MaskEdBox txtCoPayEffdate 
            Height          =   315
            Left            =   7305
            TabIndex        =   233
            Top             =   1050
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   10560
            Top             =   2400
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   8520
            Top             =   1920
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label30 
            Caption         =   "Health Type  *"
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
            Left            =   75
            TabIndex        =   284
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "Payor Code"
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
            Left            =   75
            TabIndex        =   283
            Top             =   525
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Group Company *"
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
            Left            =   75
            TabIndex        =   282
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label37 
            Caption         =   "Top Up Status *"
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
            Left            =   75
            TabIndex        =   281
            Top             =   1125
            Width           =   1575
         End
         Begin VB.Label Label52 
            Caption         =   "Co-Payment *"
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
            Left            =   2420
            TabIndex        =   280
            Top             =   1115
            Width           =   1215
         End
         Begin VB.Label Label53 
            Caption         =   "Eff Date  *"
            Height          =   255
            Left            =   6720
            TabIndex        =   279
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label54 
            Caption         =   "Mgmt Fee Status *"
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
            TabIndex        =   278
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Special GPeriod *"
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
            Left            =   75
            TabIndex        =   277
            Top             =   1395
            Width           =   1575
         End
         Begin VB.Label Label55 
            Caption         =   "GPeriod Day *"
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
            Left            =   2415
            TabIndex        =   276
            Top             =   1365
            Width           =   1215
         End
         Begin VB.Label Label56 
            Caption         =   "New Plan"
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
            Left            =   4395
            TabIndex        =   275
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label64 
            Caption         =   "No of Plan Code  "
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
            Left            =   5400
            TabIndex        =   274
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label65 
            Caption         =   "Fee Sch"
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
            TabIndex        =   273
            Top             =   2190
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label78 
            Caption         =   "Prod Cat *"
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
            Left            =   4440
            TabIndex        =   272
            Top             =   1650
            Width           =   1935
         End
         Begin VB.Label Label79 
            Caption         =   "Meals *"
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
            Left            =   4440
            TabIndex        =   271
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label80 
            Caption         =   "Nursing Chrg *"
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
            TabIndex        =   270
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label81 
            Caption         =   "Tax *"
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
            TabIndex        =   269
            Top             =   1600
            Width           =   615
         End
         Begin VB.Label Label82 
            Caption         =   "MRI *"
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
            TabIndex        =   268
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "End Age (Year)*"
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
            Index           =   81
            Left            =   2420
            TabIndex        =   267
            Top             =   1880
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Start Age (Day)*"
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
            Index           =   80
            Left            =   2420
            TabIndex        =   266
            Top             =   1600
            Width           =   1455
         End
         Begin VB.Label Label79 
            Caption         =   "Policy Wording *"
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
            Left            =   120
            TabIndex        =   265
            Top             =   2205
            Width           =   1335
         End
         Begin VB.Label Label56 
            Caption         =   "Client Plan"
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
            Left            =   4440
            TabIndex        =   264
            Top             =   2220
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Client Pol. No"
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
            Left            =   6720
            TabIndex        =   263
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Label17 
            Caption         =   "Copay Perc"
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
            Left            =   4395
            TabIndex        =   262
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblVIP 
            Caption         =   "VIP Ind *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   261
            Top             =   2565
            Width           =   1095
         End
         Begin VB.Label lblLGInd 
            Caption         =   "LG Number *"
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
            Index           =   3
            Left            =   120
            TabIndex        =   260
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label lblLGInd 
            Caption         =   "E-Card *"
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
            TabIndex        =   259
            Top             =   2640
            Width           =   1575
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   202
         Top             =   960
         Width           =   10800
         Begin VB.TextBox txtMCOCode 
            Height          =   285
            Left            =   2300
            MaxLength       =   1
            TabIndex        =   204
            Top             =   360
            Width           =   800
         End
         Begin VB.TextBox txtMCOStatus 
            Height          =   285
            Left            =   2280
            TabIndex        =   203
            Top             =   800
            Width           =   3500
         End
         Begin VB.Label Label84 
            Caption         =   "MCO Code"
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
            Left            =   480
            TabIndex        =   206
            Top             =   360
            Width           =   1700
         End
         Begin VB.Label Label83 
            Caption         =   "MCO Status          *"
            Height          =   255
            Left            =   480
            TabIndex        =   205
            Top             =   800
            Width           =   1700
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Index           =   7
         Left            =   -74760
         TabIndex        =   192
         Top             =   960
         Width           =   10875
         Begin VB.CheckBox chkMgmtPLN 
            Caption         =   "Check3"
            Height          =   255
            Left            =   5160
            TabIndex        =   219
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox TxtMgmtFee 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6840
            MaxLength       =   8
            TabIndex        =   196
            Top             =   720
            Width           =   3225
         End
         Begin VB.ComboBox CboMgmtPlan 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   195
            Top             =   720
            Width           =   3255
         End
         Begin VB.ComboBox CboMgmtHltCode 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   194
            Top             =   360
            Width           =   3255
         End
         Begin VB.ComboBox CboMgmtIns 
            Height          =   315
            ItemData        =   "frmMaintenencePlan.frx":015B
            Left            =   6840
            List            =   "frmMaintenencePlan.frx":015D
            Sorted          =   -1  'True
            TabIndex        =   193
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label90 
            Caption         =   "Mgmt. Fee          *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   200
            Top             =   800
            Width           =   1335
         End
         Begin VB.Label Label89 
            Caption         =   "Plan Code        *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   199
            Top             =   800
            Width           =   1335
         End
         Begin VB.Label Label88 
            Caption         =   "Health Type    *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   198
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label Label87 
            Caption         =   "Insured Type  *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   197
            Top             =   450
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView lstProductCat 
         Height          =   3855
         Left            =   -74760
         TabIndex        =   187
         Top             =   2400
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   6800
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   -74760
         TabIndex        =   186
         Top             =   960
         Width           =   10800
         Begin VB.TextBox txtProName 
            Height          =   285
            Left            =   2280
            TabIndex        =   191
            Top             =   800
            Width           =   3500
         End
         Begin VB.TextBox txtProCode 
            Height          =   285
            Left            =   2300
            MaxLength       =   2
            TabIndex        =   190
            Top             =   360
            Width           =   800
         End
         Begin VB.Label Label75 
            Caption         =   "Product Name             *"
            Height          =   255
            Left            =   480
            TabIndex        =   189
            Top             =   800
            Width           =   1700
         End
         Begin VB.Label Label73 
            Caption         =   "Product Code       *"
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
            Left            =   480
            TabIndex        =   188
            Top             =   360
            Width           =   1700
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Index           =   5
         Left            =   -74760
         TabIndex        =   167
         Top             =   960
         Width           =   10875
         Begin VB.CheckBox chkALBPLN 
            Caption         =   "Check2"
            Height          =   255
            Left            =   5100
            TabIndex        =   218
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox txtALBHomeNursLmt 
            Height          =   315
            Left            =   7400
            TabIndex        =   184
            Top             =   1440
            Width           =   2000
         End
         Begin VB.TextBox txtALBCancerLmt 
            Height          =   315
            Left            =   7400
            TabIndex        =   183
            Top             =   1080
            Width           =   2000
         End
         Begin VB.TextBox txtALBKidDiaLmt 
            Height          =   315
            Left            =   7400
            TabIndex        =   182
            Top             =   720
            Width           =   2000
         End
         Begin VB.ComboBox cboALBPlnCode 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   171
            Top             =   720
            Width           =   3255
         End
         Begin VB.ComboBox cboALBHltCode 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   170
            Top             =   360
            Width           =   3255
         End
         Begin VB.ComboBox cboALBInsType 
            Height          =   315
            ItemData        =   "frmMaintenencePlan.frx":015F
            Left            =   1800
            List            =   "frmMaintenencePlan.frx":0161
            Sorted          =   -1  'True
            TabIndex        =   169
            Top             =   1080
            Width           =   3255
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   7395
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   168
            Top             =   360
            Width           =   3180
         End
         Begin MSMask.MaskEdBox txtALBEffDate 
            Height          =   315
            Left            =   1800
            TabIndex        =   172
            Top             =   1440
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label69 
            Caption         =   "Home Nursing Care Limit"
            Height          =   375
            Left            =   5400
            TabIndex        =   181
            Top             =   1440
            Width           =   1900
         End
         Begin VB.Label Label67 
            Caption         =   "O/P Cancer Limit"
            Height          =   375
            Left            =   5400
            TabIndex        =   180
            Top             =   1080
            Width           =   1900
         End
         Begin VB.Label Label66 
            Caption         =   "O/P Kidney Dialysis Limit"
            Height          =   255
            Left            =   5400
            TabIndex        =   179
            Top             =   720
            Width           =   1900
         End
         Begin VB.Label Label74 
            Caption         =   "Eff. Date            *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   177
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label72 
            Caption         =   "Plan Code        *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   176
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label71 
            Caption         =   "Health Type    *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   175
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label70 
            Caption         =   "Insured Type  *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   174
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label68 
            Caption         =   "Supp. Ann Lmt Status   "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5400
            TabIndex        =   173
            Top             =   360
            Width           =   1905
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Index           =   4
         Left            =   240
         TabIndex        =   160
         Top             =   960
         Width           =   10860
         Begin VB.CheckBox chkSpecialGrpCom 
            Caption         =   "Check2"
            Height          =   255
            Left            =   10560
            TabIndex        =   217
            Top             =   380
            Width           =   255
         End
         Begin VB.CheckBox chkSpecialPLN 
            Caption         =   "Check2"
            Height          =   255
            Left            =   3720
            TabIndex        =   216
            Top             =   370
            Width           =   255
         End
         Begin VB.ComboBox CboSpecialGrpCom 
            Height          =   315
            Left            =   5700
            Sorted          =   -1  'True
            TabIndex        =   70
            Top             =   360
            Width           =   4800
         End
         Begin VB.ComboBox CboSpecialPlan 
            Height          =   315
            Left            =   2040
            Sorted          =   -1  'True
            TabIndex        =   69
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtSpecialNote 
            Height          =   1575
            Left            =   2040
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   71
            Top             =   720
            Width           =   8450
         End
         Begin VB.Label Label61 
            Caption         =   "Group Company  *"
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
            TabIndex        =   164
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label58 
            Caption         =   "Plan Code        *"
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
            TabIndex        =   162
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label57 
            Caption         =   "Special Note          *"
            Height          =   255
            Left            =   240
            TabIndex        =   161
            Top             =   720
            Width           =   1530
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "AgeBand"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -69120
         TabIndex        =   133
         Top             =   2640
         Width           =   5400
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4560
            MaxLength       =   8
            TabIndex        =   39
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3840
            MaxLength       =   8
            TabIndex        =   38
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3120
            MaxLength       =   8
            TabIndex        =   37
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   36
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1680
            MaxLength       =   8
            TabIndex        =   35
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   960
            MaxLength       =   8
            TabIndex        =   34
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            MaxLength       =   8
            TabIndex        =   33
            Top             =   480
            Width           =   690
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Std Supplementary Premium"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label11 
            Caption         =   "07"
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
            Left            =   4920
            TabIndex        =   143
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label47 
            Caption         =   "06"
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
            Left            =   4200
            TabIndex        =   139
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label46 
            Caption         =   "05"
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
            Left            =   3480
            TabIndex        =   138
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label41 
            Caption         =   "04"
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
            TabIndex        =   137
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label40 
            Caption         =   "03"
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
            Left            =   2040
            TabIndex        =   136
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label39 
            Caption         =   "02"
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
            Left            =   1320
            TabIndex        =   135
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label38 
            Caption         =   "01"
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
            Left            =   600
            TabIndex        =   134
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   108
         Top             =   960
         Width           =   10935
         Begin VB.Frame Frame8 
            Caption         =   "Supp. Insured Type *"
            Height          =   1335
            Left            =   5520
            TabIndex        =   154
            Top             =   960
            Width           =   5415
            Begin VB.TextBox TxtSFCAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   4080
               MaxLength       =   8
               TabIndex        =   65
               Top             =   480
               Width           =   945
            End
            Begin VB.TextBox TxtSFAAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1560
               MaxLength       =   8
               TabIndex        =   64
               Top             =   480
               Width           =   945
            End
            Begin VB.TextBox TxtSHAAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1560
               MaxLength       =   8
               TabIndex        =   66
               Top             =   720
               Width           =   945
            End
            Begin VB.TextBox TxtSHCAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   4080
               MaxLength       =   8
               TabIndex        =   67
               Top             =   720
               Width           =   945
            End
            Begin VB.TextBox TxtCoverID 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1560
               MaxLength       =   2
               TabIndex        =   68
               Top             =   960
               Width           =   945
            End
            Begin VB.Label Label14 
               Caption         =   "Covered ID Chg."
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   185
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label14 
               Caption         =   "Family (Adult)"
               Height          =   255
               Index           =   15
               Left            =   240
               TabIndex        =   158
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Family (Child)"
               Height          =   255
               Index           =   14
               Left            =   2640
               TabIndex        =   157
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Family (Adult)"
               Height          =   255
               Index           =   11
               Left            =   240
               TabIndex        =   156
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Family (Child)"
               Height          =   255
               Index           =   3
               Left            =   2640
               TabIndex        =   155
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Principal Insured Type *"
            Height          =   1335
            Left            =   0
            TabIndex        =   145
            Top             =   960
            Width           =   5535
            Begin VB.TextBox TxtICAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   4440
               MaxLength       =   8
               TabIndex        =   57
               Top             =   240
               Width           =   945
            End
            Begin VB.TextBox TxtIAAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1800
               MaxLength       =   8
               TabIndex        =   56
               Top             =   240
               Width           =   945
            End
            Begin VB.TextBox TxtFCAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   4440
               MaxLength       =   8
               TabIndex        =   59
               Top             =   480
               Width           =   945
            End
            Begin VB.TextBox TxtGCAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   4440
               MaxLength       =   8
               TabIndex        =   61
               Top             =   720
               Width           =   945
            End
            Begin VB.TextBox TxtHCAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   4440
               MaxLength       =   8
               TabIndex        =   63
               Top             =   960
               Width           =   945
            End
            Begin VB.TextBox TxtFAAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1800
               MaxLength       =   8
               TabIndex        =   58
               Top             =   480
               Width           =   945
            End
            Begin VB.TextBox TxtGAAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1800
               MaxLength       =   8
               TabIndex        =   60
               Top             =   720
               Width           =   945
            End
            Begin VB.TextBox TxtHAAmt 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1800
               MaxLength       =   8
               TabIndex        =   62
               Top             =   960
               Width           =   945
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Family (Child)"
               Height          =   255
               Index           =   10
               Left            =   2880
               TabIndex        =   153
               Top             =   990
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Family (Adult)"
               Height          =   255
               Index           =   9
               Left            =   240
               TabIndex        =   152
               Top             =   990
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Ind. (Child)"
               Height          =   255
               Index           =   8
               Left            =   2880
               TabIndex        =   151
               Top             =   735
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Grp Ind. (Adult)"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   150
               Top             =   735
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Family (Child)"
               Height          =   255
               Index           =   6
               Left            =   2880
               TabIndex        =   149
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label14 
               Caption         =   "Family (Adult)"
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   148
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label14 
               Caption         =   "Individual (Child)"
               Height          =   255
               Index           =   4
               Left            =   2880
               TabIndex        =   147
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label LblIA 
               Caption         =   "Individual (Adult)"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   146
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.ComboBox CboMCOStatus 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   54
            Top             =   550
            Width           =   3690
         End
         Begin VB.TextBox TxtMCOPlan 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7680
            MaxLength       =   4
            TabIndex        =   53
            Top             =   240
            Width           =   1185
         End
         Begin VB.ComboBox CboMCOHlt 
            Height          =   315
            ItemData        =   "frmMaintenencePlan.frx":0163
            Left            =   1800
            List            =   "frmMaintenencePlan.frx":0165
            Sorted          =   -1  'True
            TabIndex        =   52
            Top             =   200
            Width           =   3705
         End
         Begin MSMask.MaskEdBox txtEffDate 
            Height          =   315
            Left            =   7680
            TabIndex        =   55
            Top             =   600
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label22 
            Caption         =   "Supp. MCO Status   *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "Plan Code *"
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
            Index           =   2
            Left            =   6120
            TabIndex        =   141
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Health Type     *"
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
            Left            =   120
            TabIndex        =   110
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label34 
            Caption         =   "MCO  Eff. Date *"
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
            TabIndex        =   109
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Index           =   6
         Left            =   -74760
         TabIndex        =   100
         Top             =   960
         Width           =   10875
         Begin VB.CheckBox chkPLN 
            Caption         =   "Check2"
            Height          =   255
            Left            =   5100
            TabIndex        =   210
            Top             =   750
            Width           =   255
         End
         Begin VB.TextBox txtDisabilityLmt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   209
            Top             =   1800
            Width           =   3300
         End
         Begin VB.TextBox TxtSuppLifeTimeLimit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7300
            MaxLength       =   8
            TabIndex        =   47
            Top             =   1440
            Width           =   3180
         End
         Begin VB.TextBox TxtLifeTimeLimit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   43
            Top             =   1440
            Width           =   3300
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   7300
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   720
            Width           =   3180
         End
         Begin VB.TextBox txtANNVersion 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10050
            MaxLength       =   8
            TabIndex        =   49
            Top             =   1800
            Width           =   420
         End
         Begin VB.ComboBox CboInsType 
            Height          =   315
            ItemData        =   "frmMaintenencePlan.frx":0167
            Left            =   1800
            List            =   "frmMaintenencePlan.frx":0169
            Sorted          =   -1  'True
            TabIndex        =   42
            Top             =   1080
            Width           =   3255
         End
         Begin VB.ComboBox CboHltCode 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   40
            Top             =   360
            Width           =   3255
         End
         Begin VB.ComboBox CboPlnCode 
            Height          =   315
            ItemData        =   "frmMaintenencePlan.frx":016B
            Left            =   1800
            List            =   "frmMaintenencePlan.frx":016D
            Sorted          =   -1  'True
            TabIndex        =   41
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox TxtAnnualLimit 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7300
            MaxLength       =   8
            TabIndex        =   44
            Top             =   360
            Width           =   3180
         End
         Begin VB.TextBox TxtSupplyAL 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7300
            MaxLength       =   8
            TabIndex        =   46
            Top             =   1080
            Width           =   3180
         End
         Begin MSMask.MaskEdBox txtALEffDate 
            Height          =   315
            Left            =   7300
            TabIndex        =   48
            Top             =   1800
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label91 
            Caption         =   "Per Disability Limit"
            Height          =   255
            Left            =   400
            TabIndex        =   208
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label36 
            Caption         =   "Supp. Life Time Lmt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   166
            Top             =   1470
            Width           =   1695
         End
         Begin VB.Label Label63 
            Caption         =   "Life Time Limit          "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   400
            TabIndex        =   165
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label28 
            Caption         =   "Supp. Ann Lmt Status   "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   132
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label23 
            Caption         =   "Annual Limit Version *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8450
            TabIndex        =   131
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label45 
            Caption         =   "Insured Type  *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   400
            TabIndex        =   125
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label35 
            Caption         =   "Health Type    *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   400
            TabIndex        =   105
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label49 
            Caption         =   "Plan Code        *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   400
            TabIndex        =   104
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label44 
            Caption         =   "Annual Limit                 *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   103
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label43 
            Caption         =   "Annual Limit Eff. Date *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   102
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label42 
            Caption         =   "Supp. Ann Lmt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   101
            Top             =   1110
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Index           =   0
         Left            =   -74760
         TabIndex        =   93
         Top             =   960
         Width           =   10980
         Begin VB.CheckBox chkPmPLN 
            Caption         =   "Check4"
            Height          =   255
            Left            =   10680
            TabIndex        =   215
            Top             =   250
            Width           =   255
         End
         Begin VB.CheckBox chkPmAge 
            Caption         =   "Check3"
            Height          =   255
            Left            =   10680
            TabIndex        =   214
            Top             =   650
            Width           =   255
         End
         Begin VB.CheckBox chkPmPayor 
            Caption         =   "Check2"
            Height          =   255
            Left            =   5550
            TabIndex        =   213
            Top             =   960
            Width           =   255
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1320
            Width           =   3690
         End
         Begin VB.TextBox txtPREVersion 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   9720
            MaxLength       =   8
            TabIndex        =   23
            Top             =   945
            Width           =   900
         End
         Begin VB.ComboBox cboPmAgeCode 
            Height          =   315
            Left            =   7200
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   600
            Width           =   3400
         End
         Begin VB.ComboBox cboPmPAYCode 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   960
            Width           =   3690
         End
         Begin VB.ComboBox cboPmPLNCode 
            Height          =   315
            Left            =   7200
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   3400
         End
         Begin VB.ComboBox cboPmINSCode 
            Height          =   315
            ItemData        =   "frmMaintenencePlan.frx":016F
            Left            =   1800
            List            =   "frmMaintenencePlan.frx":0171
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   600
            Width           =   3690
         End
         Begin VB.ComboBox cboPmHLTCode 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   3690
         End
         Begin MSMask.MaskEdBox txtPRMEffdate 
            Height          =   315
            Left            =   7200
            TabIndex        =   22
            Top             =   960
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label21 
            Caption         =   "Prem Version  *"
            Height          =   255
            Left            =   8400
            TabIndex        =   129
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Age Type    *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5880
            TabIndex        =   113
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Insured Type   *"
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
            Left            =   360
            TabIndex        =   97
            Top             =   630
            Width           =   1380
         End
         Begin VB.Label Label7 
            Caption         =   "Payor Code   "
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
            Left            =   360
            TabIndex        =   96
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Label10 
            Caption         =   "Plan Code   *"
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
            TabIndex        =   95
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Health Type    *"
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
            Left            =   360
            TabIndex        =   94
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label48 
            Caption         =   "Supp. Prem Status   "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   140
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label32 
            Caption         =   "Eff Date From  *"
            Height          =   255
            Left            =   5880
            TabIndex        =   98
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   2
         Left            =   -74760
         TabIndex        =   87
         Top             =   960
         Width           =   10860
         Begin VB.TextBox txtHLTdescription 
            Height          =   1095
            Left            =   2040
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   720
            Width           =   8535
         End
         Begin VB.TextBox txtHLTcode 
            Height          =   285
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   0
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Health Description      *"
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   720
            Width           =   1770
         End
         Begin VB.Label Label12 
            Caption         =   "Health Type         *"
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
            TabIndex        =   88
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Index           =   3
         Left            =   -74760
         TabIndex        =   84
         Top             =   1080
         Width           =   10770
         Begin VB.TextBox txtINSdescription 
            Height          =   735
            Left            =   2160
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   720
            Width           =   8295
         End
         Begin VB.TextBox txtINScode 
            Height          =   285
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "Insured Type           *"
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
            Left            =   210
            TabIndex        =   86
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label Label15 
            Caption         =   "Insured Type Description *  "
            Height          =   255
            Left            =   210
            TabIndex        =   85
            Top             =   840
            Width           =   2205
         End
      End
      Begin MSComctlLib.ListView lstPlanList 
         Height          =   1920
         Left            =   -74760
         TabIndex        =   90
         Top             =   4920
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   3387
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   31
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Plan Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "New Plan Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Plan Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Payor Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Group Company"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Top Up Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Annual Limit Ind"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Premium Ind"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "MCO Ind"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Co-Payment"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Eff Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Text            =   "Special GPeriod Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "Special GPeriod Days"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   15
            Text            =   "New SM Plan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   16
            Text            =   "Life Time Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   17
            Text            =   "Schedule of Fee"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Product Category"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   19
            Text            =   "Meals"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   20
            Text            =   "Nursing Charges"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   21
            Text            =   "Tax"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   22
            Text            =   "MRI"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   23
            Text            =   "Mgmt. Fee"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   24
            Text            =   "Dis. Indicator"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "Policy Wording"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Client Plan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "Client Pol No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "LG Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "VIP Indicator"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "E-Card Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstHLTList 
         Height          =   3045
         Left            =   -74760
         TabIndex        =   91
         Top             =   3120
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   5371
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComctlLib.ListView lstINSList 
         Height          =   3330
         Left            =   -74760
         TabIndex        =   92
         Top             =   2880
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   5874
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
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Insured Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstPmPRMList 
         Height          =   2280
         Left            =   -74760
         TabIndex        =   99
         Top             =   3960
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   4022
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
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Premium Code"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Insured Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Payor Code"
            Object.Width           =   2535
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Health Code"
            Object.Width           =   2535
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Age Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Plan Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Premium Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Supp Premium Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Eff.Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Prem Version"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Status"
            Object.Width           =   353
         EndProperty
      End
      Begin MSComctlLib.ListView lstAnnualList 
         Height          =   2520
         Left            =   -74760
         TabIndex        =   106
         Top             =   3720
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   4445
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Plan Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Insured Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Supply. Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Annual Limit Eff. Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Payor Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Annual Version"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Annualindex"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Life Time Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "Supp Life Time Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Disability Lmt"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame FrameAGEBand 
         Caption         =   "AgeBand"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   115
         Top             =   2640
         Width           =   5640
         Begin VB.TextBox txt07 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4680
            MaxLength       =   8
            TabIndex        =   31
            Top             =   480
            Width           =   690
         End
         Begin VB.CheckBox chkSTDPremium 
            Caption         =   "Std Premium"
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txt01 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   360
            MaxLength       =   8
            TabIndex        =   25
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt02 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            MaxLength       =   8
            TabIndex        =   26
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt03 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   27
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt04 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2520
            MaxLength       =   8
            TabIndex        =   28
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt05 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3240
            MaxLength       =   8
            TabIndex        =   29
            Top             =   480
            Width           =   690
         End
         Begin VB.TextBox txt06 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   30
            Top             =   480
            Width           =   690
         End
         Begin VB.Label Label2 
            Caption         =   "07"
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
            TabIndex        =   142
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl01 
            Caption         =   "01"
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
            Left            =   720
            TabIndex        =   121
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl02 
            Caption         =   "02"
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
            Left            =   1440
            TabIndex        =   120
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl03 
            Caption         =   "03"
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
            Left            =   2160
            TabIndex        =   119
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl04 
            Caption         =   "04"
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
            Left            =   2880
            TabIndex        =   118
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl05 
            Caption         =   "05"
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
            TabIndex        =   117
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbl06 
            Caption         =   "06"
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
            Left            =   4320
            TabIndex        =   116
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   -74760
         TabIndex        =   126
         Top             =   3050
         Width           =   10875
         Begin VB.CheckBox chkGrpCom 
            Caption         =   "Check4"
            Height          =   255
            Left            =   10550
            TabIndex        =   212
            Top             =   250
            Width           =   255
         End
         Begin VB.CheckBox chkPayor 
            Caption         =   "Check3"
            Height          =   255
            Left            =   5100
            TabIndex        =   211
            Top             =   250
            Width           =   255
         End
         Begin VB.ComboBox CboGrpCom 
            Height          =   315
            ItemData        =   "frmMaintenencePlan.frx":0173
            Left            =   7300
            List            =   "frmMaintenencePlan.frx":0175
            Sorted          =   -1  'True
            TabIndex        =   51
            Top             =   240
            Width           =   3180
         End
         Begin VB.ComboBox CboPayorCode 
            Height          =   315
            ItemData        =   "frmMaintenencePlan.frx":0177
            Left            =   1800
            List            =   "frmMaintenencePlan.frx":0179
            Sorted          =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label20 
            Caption         =   "Group Company  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   128
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label5 
            Caption         =   "Payor Code    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   127
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView lstMCOList 
         Height          =   2835
         Left            =   -74760
         TabIndex        =   107
         Top             =   3360
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   5001
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
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Insured Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Plan Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Supp MCO Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Prin. MCO Fees (Adult)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Prin. MCO Fees (Child)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Supp. MCO Fees (Adult)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Supp. MCO Fees (Child)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Eff.Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Covered ID"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstSpecialNoteList 
         Height          =   2685
         Left            =   240
         TabIndex        =   163
         Top             =   3480
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   4736
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Plan Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Group Company"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Special Note"
            Object.Width           =   17639
         EndProperty
      End
      Begin MSComctlLib.ListView lstALB 
         Height          =   3120
         Left            =   -74760
         TabIndex        =   178
         Top             =   3120
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   5503
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Plan Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Insured Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Annual Limit Eff. Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Kidney Dialysis Lmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cancer Lmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Home Nursing Care Lmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ALBCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMgmtFeeList 
         Height          =   3600
         Left            =   -74760
         TabIndex        =   201
         Top             =   2640
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   6350
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Plan Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Insured Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Mgmt. Fee"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMCOStatus 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   207
         Top             =   2400
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   6800
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   -74760
         TabIndex        =   122
         Top             =   4080
         Width           =   10935
         Begin VB.TextBox txtPLN1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Index           =   5
            Left            =   9045
            MaxLength       =   4
            TabIndex        =   14
            Top             =   120
            Width           =   1410
         End
         Begin VB.TextBox txtPLN1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   7485
            MaxLength       =   4
            TabIndex        =   12
            Top             =   120
            Width           =   1410
         End
         Begin VB.TextBox txtPLN1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   5880
            MaxLength       =   4
            TabIndex        =   10
            Top             =   120
            Width           =   1410
         End
         Begin VB.TextBox txtPLN1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   4365
            MaxLength       =   4
            TabIndex        =   8
            Top             =   120
            Width           =   1410
         End
         Begin VB.TextBox txtPLN1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2925
            MaxLength       =   4
            TabIndex        =   6
            Top             =   120
            Width           =   1290
         End
         Begin VB.TextBox txtPLN1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1485
            MaxLength       =   4
            TabIndex        =   4
            Top             =   120
            Width           =   1290
         End
         Begin VB.TextBox txtDesc1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   1485
            MaxLength       =   200
            TabIndex        =   5
            Top             =   360
            Width           =   1290
         End
         Begin VB.TextBox txtDesc1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   2925
            MaxLength       =   200
            TabIndex        =   7
            Top             =   360
            Width           =   1290
         End
         Begin VB.TextBox txtDesc1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   4365
            MaxLength       =   200
            TabIndex        =   9
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox txtDesc1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   5880
            MaxLength       =   200
            TabIndex        =   11
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox txtDesc1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   7485
            MaxLength       =   200
            TabIndex        =   13
            Top             =   360
            Width           =   1410
         End
         Begin VB.TextBox txtDesc1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   9045
            MaxLength       =   200
            TabIndex        =   15
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label Label19 
            Caption         =   "Plan Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   124
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label Label6 
            Caption         =   "Plan Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   123
            Top             =   120
            Width           =   1350
         End
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Plan Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   159
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label18 
      Caption         =   "TOTAL RECORDS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   112
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   111
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label31 
      Caption         =   "Bold - Searchable        * - Mandatory Fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   82
      Top             =   8160
      Width           =   4695
   End
End
Attribute VB_Name = "frmMaintenancePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim bExit As Boolean
Public PRMCriteria As String
Public PLNCriteria As String
Public HLTCriteria As String
Public INSCriteria As String
Public AnnualCriteria As String
Public MCOCriteria As String
Public SpecialCriteria As String
Public ALBCriteria As String
Public ProCatCriteria As String
Public MCOStatCriteria As String
Public s As Boolean
Public ListClick As Boolean
Public Save As Boolean
Private m_blnDirection(1 To 25) As Boolean

Private Sub cboPLNPayor_change()
    cboClientPLN.Clear
    If Mid(cboPLNPayor, 1, 2) = "AX" Then
        cboClientPLN.AddItem "001"
        cboClientPLN.AddItem "002"
        cboClientPLN.AddItem "003"
        cboClientPLN.AddItem "004"
        cboClientPLN.AddItem "005"
        cboClientPLN.AddItem "006"
        cboClientPLN.AddItem "007"
        cboClientPLN.AddItem "008"
        cboClientPLN.AddItem "009"
        cboClientPLN.AddItem "010"
        cboClientPLN.AddItem "011"
        cboClientPLN.AddItem "012"
        
    ElseIf Mid(cboPLNPayor, 1, 2) = "WM" Or Mid(cboPLNPayor, 1, 2) = "WK" Or Mid(cboPLNPayor, 1, 2) = "ET" Or Mid(cboPLNPayor, 1, 2) = "EC" Then
        cboClientPLN.AddItem "PLAN1"
        cboClientPLN.AddItem "PLAN2"
        cboClientPLN.AddItem "PLAN3"
        cboClientPLN.AddItem "PLAN4"
        cboClientPLN.AddItem "PLAN5"
        cboClientPLN.AddItem "PLAN6"
        cboClientPLN.AddItem "PLAN7"
        cboClientPLN.AddItem "PLAN8"
        cboClientPLN.AddItem "PLAN9"
        cboClientPLN.AddItem "PLAN10"
        cboClientPLN.AddItem "PLANA"
        cboClientPLN.AddItem "PLANB"
        cboClientPLN.AddItem "PLANC"
        cboClientPLN.AddItem "PLAND"
        cboClientPLN.AddItem "PLANE"
        cboClientPLN.AddItem "PLANF"
        cboClientPLN.AddItem "PLANG"
        cboClientPLN.AddItem "PLANH"
    Else
        cboClientPLN.Clear
    End If

End Sub

Private Sub cboPLNPayor_Click()
    cboClientPLN.Clear
    If Mid(cboPLNPayor, 1, 2) = "AX" Then
        cboClientPLN.AddItem "001"
        cboClientPLN.AddItem "002"
        cboClientPLN.AddItem "003"
        cboClientPLN.AddItem "004"
        cboClientPLN.AddItem "005"
        cboClientPLN.AddItem "006"
        cboClientPLN.AddItem "007"
        cboClientPLN.AddItem "008"
        cboClientPLN.AddItem "009"
        cboClientPLN.AddItem "010"
        cboClientPLN.AddItem "011"
        cboClientPLN.AddItem "012"
        
    ElseIf Mid(cboPLNPayor, 1, 2) = "WM" Or Mid(cboPLNPayor, 1, 2) = "WK" Or Mid(cboPLNPayor, 1, 2) = "ET" Or Mid(cboPLNPayor, 1, 2) = "EC" Then
        cboClientPLN.AddItem "PLAN1"
        cboClientPLN.AddItem "PLAN2"
        cboClientPLN.AddItem "PLAN3"
        cboClientPLN.AddItem "PLAN4"
        cboClientPLN.AddItem "PLAN5"
        cboClientPLN.AddItem "PLAN6"
        cboClientPLN.AddItem "PLAN7"
        cboClientPLN.AddItem "PLAN8"
        cboClientPLN.AddItem "PLAN9"
        cboClientPLN.AddItem "PLAN10"
        cboClientPLN.AddItem "PLANA"
        cboClientPLN.AddItem "PLANB"
        cboClientPLN.AddItem "PLANC"
        cboClientPLN.AddItem "PLAND"
        cboClientPLN.AddItem "PLANE"
        cboClientPLN.AddItem "PLANF"
        cboClientPLN.AddItem "PLANG"
        cboClientPLN.AddItem "PLANH"
    Else
        cboClientPLN.Clear
    End If
    
    
End Sub

Private Sub cboPLNPayor_LostFocus()
    cboClientPLN.Clear
    If Mid(cboPLNPayor, 1, 2) = "AX" Then
        cboClientPLN.AddItem "001"
        cboClientPLN.AddItem "002"
        cboClientPLN.AddItem "003"
        cboClientPLN.AddItem "004"
        cboClientPLN.AddItem "005"
        cboClientPLN.AddItem "006"
        cboClientPLN.AddItem "007"
        cboClientPLN.AddItem "008"
        cboClientPLN.AddItem "009"
        cboClientPLN.AddItem "010"
        cboClientPLN.AddItem "011"
        cboClientPLN.AddItem "012"
        
    ElseIf Mid(cboPLNPayor, 1, 2) = "WM" Or Mid(cboPLNPayor, 1, 2) = "WK" Or Mid(cboPLNPayor, 1, 2) = "ET" Or Mid(cboPLNPayor, 1, 2) = "EC" Then
        cboClientPLN.AddItem "PLAN1"
        cboClientPLN.AddItem "PLAN2"
        cboClientPLN.AddItem "PLAN3"
        cboClientPLN.AddItem "PLAN4"
        cboClientPLN.AddItem "PLAN5"
        cboClientPLN.AddItem "PLAN6"
        cboClientPLN.AddItem "PLAN7"
        cboClientPLN.AddItem "PLAN8"
        cboClientPLN.AddItem "PLAN9"
        cboClientPLN.AddItem "PLAN10"
        cboClientPLN.AddItem "PLANA"
        cboClientPLN.AddItem "PLANB"
        cboClientPLN.AddItem "PLANC"
        cboClientPLN.AddItem "PLAND"
        cboClientPLN.AddItem "PLANE"
        cboClientPLN.AddItem "PLANF"
        cboClientPLN.AddItem "PLANG"
        cboClientPLN.AddItem "PLANH"
    Else
    
    
        cboClientPLN.Clear
    End If

End Sub

Private Sub cboPolicyWording_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPolicyWording.Text, cboPolicyWording.SelStart) & _
                Chr(KeyAscii) + Mid(cboPolicyWording.Text, cboPolicyWording.SelStart + cboPolicyWording.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPolicyWording.ListCount - 1
            
            If NewText = ChkStr(Left(cboPolicyWording.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPolicyWording.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboPolicyWording.Text = ValidValue
              cboPolicyWording.SelStart = 0
              cboPolicyWording.SelLength = Len(ValidValue)
            End If

End Sub

Private Sub chkALBPLN_Click()
Dim PLN As New ADODB.Recordset

cboALBPlnCode.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkALBPLN.Value = 1 Then
    PLN.Open "select PLNCode, PLNDescription  from PLN", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until PLN.EOF
        cboALBPlnCode.AddItem PLN!PLNCode & " - " & PLN!PLNDescription
        PLN.MoveNext
    Loop
    PLN.Close
    Set PLN = Nothing
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
End If
End Sub

Private Sub ChkGrpCom_Click()
Dim GrpCom As New ADODB.Recordset

CboGrpCom.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkGrpCom.Value = 1 Then
    GrpCom.Open "select Distinct GRPCompany from PLN ", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until GrpCom.EOF
        If Len(GrpCom!Grpcompany) Then
           CboGrpCom.AddItem GrpCom!Grpcompany
        End If
        GrpCom.MoveNext
    Loop
    GrpCom.Close
    Set GrpCom = Nothing
End If
'medixmasterconn.Close
'set medixmasterconn = Nothing
End If
End Sub

Private Sub chkMgmtPLN_Click()
Dim PLN As New ADODB.Recordset

CboMgmtPlan.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkMgmtPLN.Value = 1 Then
    PLN.Open "select PLNCode, PLNDescription  from PLN", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until PLN.EOF
        CboMgmtPlan.AddItem PLN!PLNCode & " - " & PLN!PLNDescription
        PLN.MoveNext
    Loop
    PLN.Close
    Set PLN = Nothing
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
End If
End Sub

Private Sub chkPLN_Click()
Dim PLN As New ADODB.Recordset

CboPlnCode.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkPLN.Value = 1 Then
    PLN.Open "select PLNCode, PLNDescription  from PLN", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until PLN.EOF
        CboPlnCode.AddItem PLN!PLNCode & " - " & PLN!PLNDescription
        PLN.MoveNext
    Loop
    PLN.Close
    Set PLN = Nothing
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
End If
End Sub

Private Sub chkPayor_Click()
Dim PAY As New ADODB.Recordset

CboPayorCode.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkPayor.Value = 1 Then
    PAY.Open "select PAYCOde, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until PAY.EOF
        CboPayorCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
End If
End Sub

Private Sub chkPLNPayor_Click()
Dim PAY As New ADODB.Recordset

cboPLNPayor.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkPLNPayor.Value = 1 Then
    PAY.Open "select PAYCOde, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until PAY.EOF
        cboPLNPayor.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
End If
End Sub

Private Sub chkPmAge_Click()
Dim AGE As New ADODB.Recordset
Dim strAge As String

cboPmAgeCode.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkPmAge.Value = 1 Then
    
    strAge = "Select Distinct AgeCode, AGEDescription From Age"
    AGE.Open strAge, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until AGE.EOF
       cboPmAgeCode.AddItem AGE!AGECode & " - " & AGE!AGEDescription
       AGE.MoveNext
    Loop
    AGE.Close
    Set AGE = Nothing
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
End If
End Sub

Private Sub chkPmPayor_Click()
Dim PAY As New ADODB.Recordset

cboPmPAYCode.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkPmPayor.Value = 1 Then
    PAY.Open "select PAYCOde, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until PAY.EOF
        cboPmPAYCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
End If
End Sub

Private Sub chkPmPLN_Click()
Dim PLN As New ADODB.Recordset

cboPmPLNCode.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkPmPLN.Value = 1 Then
    PLN.Open "select PLNCode, PLNDescription  from PLN", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until PLN.EOF
        cboPmPLNCode.AddItem PLN!PLNCode & " - " & PLN!PLNDescription
        PLN.MoveNext
    Loop
    PLN.Close
    Set PLN = Nothing
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
End If
End Sub

Private Sub chkSpecialGrpCom_Click()
Dim GrpCom As New ADODB.Recordset

CboSpecialGrpCom.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkSpecialGrpCom.Value = 1 Then
    GrpCom.Open "select Distinct GRPCompany from PLN ", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until GrpCom.EOF
        If Len(GrpCom!Grpcompany) Then
           CboSpecialGrpCom.AddItem GrpCom!Grpcompany
        End If
        GrpCom.MoveNext
    Loop
    GrpCom.Close
    Set GrpCom = Nothing
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
End If
End Sub

Private Sub chkSpecialPLN_Click()
Dim PLN As New ADODB.Recordset

CboSpecialPlan.Clear
If Save = False Then
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
If chkSpecialPLN.Value = 1 Then
    PLN.Open "select PLNCode, PLNDescription  from PLN", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Do Until PLN.EOF
        CboSpecialPlan.AddItem PLN!PLNCode
        PLN.MoveNext
    Loop
    PLN.Close
    Set PLN = Nothing
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
End If
End Sub


Private Sub cmdFileLocation_Click()
     Dim sFname As String
     sFname = "\\97.1.1.9\Ecards\"
     CommonDialog1.CancelError = True
      On Error GoTo ErrHandler
      CommonDialog1.InitDir = "\\97.1.1.9\Ecards"
      CommonDialog1.Flags = cdlOFNHideReadOnly
      CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
      CommonDialog1.ShowOpen
      sFname = Replace(CommonDialog1.Filename, sFname, "")
      
      txtFileName = sFname
      Exit Sub
ErrHandler:
      Exit Sub
End Sub

Private Sub CmdPolicyWording_Click()
    'To Change the Policy Wording Combobox to textbox - By CHEng 4 Dec 2016
     Dim sFname1 As String
     Dim IntCharNum As Integer
     
     sFname1 = "\\svr-doc\PolicyWording\"
     CommonDialog2.CancelError = True
      On Error GoTo ErrHandler
      CommonDialog2.InitDir = "\\svr-doc\PolicyWording"
      CommonDialog2.Flags = cdlOFNHideReadOnly
      CommonDialog2.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
      CommonDialog2.ShowOpen
      sFname1 = Replace(CommonDialog2.Filename, sFname1, "")
      IntCharNum = InStr(1, sFname1, "\", vbTextCompare)
      sFname1 = Mid(sFname1, 1, IntCharNum - 1) 'Len(sFname1) - IntCharNum)
      
      TxtPolicyWording.Text = sFname1
      Exit Sub
ErrHandler:
      Exit Sub
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

Private Sub cboDisIndicator_change()
    If InStr(cboDisIndicator.Text, "null") Then
       cboDisIndicator.Enabled = False
    End If
End Sub

Private Sub cboDisIndicator_KeyPress(KeyAscii As Integer)

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDisIndicator.Text, cboDisIndicator.SelStart) + Chr(KeyAscii) + Mid(cboDisIndicator.Text, cboDisIndicator.SelStart + cboDisIndicator.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboDisIndicator.ListCount - 1
 
        If NewText = LCase(Left(cboDisIndicator.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboDisIndicator.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboDisIndicator.Text = ValidValue
               cboDisIndicator.SelStart = 0
               cboDisIndicator.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboProCat_Change()
        If InStr(cboProCat.Text, "null") Then
       cboProCat.Enabled = False
    End If
End Sub

Private Sub CboProCat_KeyPress(KeyAscii As Integer)

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboProCat.Text, cboProCat.SelStart) + Chr(KeyAscii) + Mid(cboProCat.Text, cboProCat.SelStart + cboProCat.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboProCat.ListCount - 1
 
        If NewText = LCase(Left(cboProCat.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboProCat.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboProCat.Text = ValidValue
               cboProCat.SelStart = 0
               cboProCat.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboGrpCom_Change()
        If InStr(CboGrpCom.Text, "null") Then
       CboGrpCom.Enabled = False
    End If
End Sub

Private Sub CboGrpCom_KeyPress(KeyAscii As Integer)

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGrpCom.Text, CboGrpCom.SelStart) + Chr(KeyAscii) + Mid(CboGrpCom.Text, CboGrpCom.SelStart + CboGrpCom.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboGrpCom.ListCount - 1
 
        If NewText = LCase(Left(CboGrpCom.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGrpCom.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboGrpCom.Text = ValidValue
               CboGrpCom.SelStart = 0
               CboGrpCom.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboHLTCode_Change()
    
    If InStr(CboHltCode.Text, "null") Then
       CboMCOHlt.Enabled = False
    End If
   
End Sub

Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboHltCode.Text, CboHltCode.SelStart) + Chr(KeyAscii) + Mid(CboHltCode.Text, CboHltCode.SelStart + CboHltCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboHltCode.ListCount - 1
 
        If NewText = LCase(Left(CboHltCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboHltCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboHltCode.Text = ValidValue
               CboHltCode.SelStart = 0
               CboHltCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboInsType_Change()
    If InStr(CboInsType.Text, "null") Then
       CboInsType.Enabled = False
    End If
End Sub

Private Sub CboInsType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboInsType.Text, CboInsType.SelStart) + Chr(KeyAscii) + Mid(CboInsType.Text, CboInsType.SelStart + CboInsType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboInsType.ListCount - 1
 
        If NewText = LCase(Left(CboInsType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboInsType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboInsType.Text = ValidValue
               CboInsType.SelStart = 0
               CboInsType.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboMCOHlt_Change()
    If InStr(CboMCOHlt.Text, "null") Then
       CboMCOHlt.Enabled = False
    End If
End Sub

Private Sub CboMCOHlt_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboMCOHlt.Text, CboMCOHlt.SelStart) + Chr(KeyAscii) + Mid(CboMCOHlt.Text, CboMCOHlt.SelStart + CboMCOHlt.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboMCOHlt.ListCount - 1
 
        If NewText = LCase(Left(CboMCOHlt.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboMCOHlt.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboMCOHlt.Text = ValidValue
               CboMCOHlt.SelStart = 0
               CboMCOHlt.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNCode_Change()
    If InStr(CboPlnCode.Text, "null") Then
       CboPlnCode.Enabled = False
    End If
End Sub

Private Sub cboPLNCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPlnCode.Text, CboPlnCode.SelStart) + Chr(KeyAscii) + Mid(CboPlnCode.Text, CboPlnCode.SelStart + CboPlnCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPlnCode.ListCount - 1
 
        If NewText = LCase(Left(CboPlnCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPlnCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPlnCode.Text = ValidValue
               CboPlnCode.SelStart = 0
               CboPlnCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNHLTCode_Change()
Dim AA As Integer
Dim EnStatus As Boolean
    
    If InStr(cboPLNHLTCode.Text, "null") Then
       cboPLNHLTCode.Enabled = False
    End If
    
    If Mid(Trim(cboPLNHLTCode), 1, 1) = "N" Then
        EnStatus = False
    Else
        EnStatus = True
    End If
    For AA = 0 To txtPLN1.UBound
        txtPLN1(AA).Enabled = EnStatus
    Next AA
End Sub

Private Sub cboPayorCode_Change()
    If InStr(CboPayorCode.Text, "null") Then
       CboPayorCode.Enabled = False
    End If
End Sub

Private Sub cboPayorCode_KeyPress(KeyAscii As Integer)

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayorCode.Text, CboPayorCode.SelStart) + Chr(KeyAscii) + Mid(CboPayorCode.Text, CboPayorCode.SelStart + CboPayorCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPayorCode.ListCount - 1
 
        If NewText = LCase(Left(CboPayorCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayorCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPayorCode.Text = ValidValue
               CboPayorCode.SelStart = 0
               CboPayorCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub
Private Sub cboPLNHLTCode_Click()
Dim AA As Integer
Dim EnStatus As Boolean

    If InStr(cboPLNHLTCode.Text, "null") Then
       cboPLNHLTCode.Enabled = False
    End If
    
    If Mid(Trim(cboPLNHLTCode), 1, 1) = "N" Then
        EnStatus = False
    Else
        EnStatus = True
    End If
    For AA = 0 To txtPLN1.UBound
        txtPLN1(AA).Enabled = EnStatus
    Next AA
End Sub

Private Sub cboPmHLTCode_Change()
    If InStr(cboPmHLTCode.Text, "null") Then
       CboMCOHlt.Enabled = False
    End If
End Sub

Private Sub cboPmHLTCode_Click()
Dim PAY As New ADODB.Recordset
Dim strsql As String
    
    'cboPmPAYCode.Clear
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    'StrSql = "Select  PAYCode, PAYCompanyName From PAY "
    'StrSql = StrSql + " Where HLTCode ='" & Mid(Trim(cboPmHLTCode), 1, 1) & "'"
    
    'PAY.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'cboPmPAYCode.Clear
    'Do Until PAY.EOF
       'cboPmPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       'PAY.MoveNext
    'Loop
    'PAY.Close
    'Set PAY = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing

End Sub
Private Sub cboALBHltCode_KeyPress(KeyAscii As Integer)

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboALBHltCode.Text, cboALBHltCode.SelStart) + Chr(KeyAscii) + Mid(cboALBHltCode.Text, cboALBHltCode.SelStart + cboALBHltCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboALBHltCode.ListCount - 1
 
        If NewText = LCase(Left(cboALBHltCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboALBHltCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboALBHltCode.Text = ValidValue
               cboALBHltCode.SelStart = 0
               cboALBHltCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboALBPlnCode_KeyPress(KeyAscii As Integer)

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboALBPlnCode.Text, cboALBPlnCode.SelStart) + Chr(KeyAscii) + Mid(cboALBPlnCode.Text, cboALBPlnCode.SelStart + cboALBPlnCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboALBHltCode.ListCount - 1
 
        If NewText = LCase(Left(cboALBPlnCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboALBPlnCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboALBPlnCode.Text = ValidValue
               cboALBPlnCode.SelStart = 0
               cboALBPlnCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboALBInsType_KeyPress(KeyAscii As Integer)

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboALBInsType.Text, cboALBInsType.SelStart) + Chr(KeyAscii) + Mid(cboALBInsType.Text, cboALBInsType.SelStart + cboALBInsType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboALBInsType.ListCount - 1
 
        If NewText = LCase(Left(cboALBInsType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboALBInsType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboALBInsType.Text = ValidValue
               cboALBInsType.SelStart = 0
               cboALBInsType.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub Check1_Click()
Dim AgeBand As String
    If Check1.Value = 1 Then
        If Text2 <> "" Then
            AgeBand = Text2
            Text3 = AgeBand
            Text4 = AgeBand
            Text5 = AgeBand
            Text6 = AgeBand
            Text7 = AgeBand
            Text1 = AgeBand
        End If
    End If
End Sub

Private Sub chkSTDPremium_Click()
Dim AgeBand As String
    If chkSTDPremium.Value = 1 Then
        If txt01 <> "" Then
            AgeBand = txt01
            txt02 = AgeBand
            txt03 = AgeBand
            txt04 = AgeBand
            txt05 = AgeBand
            txt06 = AgeBand
            txt07 = AgeBand
        End If
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    Call EntriesEnable(True)
    'Select Case SSTab1.Tab
        'Case 0
            
        'Case 1
            
        'Case 2
            
        'Case 3
            
        'Case 4
            
        'Case 5
            'CmdAdd.Enabled = True
        'Case 6
        
        'Case 7
        
        'Case 8
        
    'End Select
    lstPlanList.ListItems.Clear
    lstHLTList.ListItems.Clear
    lstINSList.ListItems.Clear
    lstPmPRMList.ListItems.Clear
    lstAnnualList.ListItems.Clear
    lstMCOList.ListItems.Clear
    lstSpecialNoteList.ListItems.Clear
    lstALB.ListItems.Clear
    lstProductCat.ListItems.Clear
    lstMgmtFeeList.ListItems.Clear
    lstMCOStatus.ListItems.Clear
    txtTotalRecord = 0
    CmdAdd.Enabled = True
    CmdSave.Enabled = False
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    'Select Case SSTab1.Tab
        'Case 0
            'Call PrintHLT
        'Case 1
            'Call PrintINS
        'Case 2
            'Call PrintPLN
        'Case 3
            'Call PrintPRM
        'Case 4
            'Call PrintAnnual
        'Case 5
            'Call PrintMCO
        'Case 6
            'Call PrintSpecial
    'End Select
    Select Case SSTab1.Tab
        Case 0
            If lstHLTList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstHLTList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 1
            If lstINSList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstINSList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 2
            If lstPlanList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstPlanList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 3
            If lstPmPRMList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstPmPRMList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 4
            If lstAnnualList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstAnnualList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 5
            If lstMCOList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstMCOList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 6
            If lstSpecialNoteList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstSpecialNoteList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 7
            If lstALB.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstALB)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 8
            If lstProductCat.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstProductCat)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 9
            If lstMCOStatus.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstMCOStatus)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 10
            If lstMgmtFeeList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstMgmtFeeList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        End Select
    
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub Combo1_Change()
    If Mid(Combo1, 1, 1) = "B" Or Mid(Combo1, 1, 1) = "C" Then
        TxtSupplyAL.Enabled = True
        TxtSuppLifeTimeLimit.Enabled = True
    Else
        TxtSupplyAL.Enabled = False
        TxtSuppLifeTimeLimit.Enabled = False
    End If
End Sub

Private Sub Combo1_Click()
    If Mid(Combo1, 1, 1) = "B" Or Mid(Combo1, 1, 1) = "C" Then
        TxtSupplyAL.Enabled = True
        TxtSuppLifeTimeLimit.Enabled = True
    Else
        TxtSupplyAL.Enabled = False
        TxtSuppLifeTimeLimit.Enabled = False
    End If
End Sub

Private Sub CboMCOStatus_Change()
    If Mid(CboMCOStatus, 1, 1) = "B" Or Mid(CboMCOStatus, 1, 1) = "C" Then
        Frame8.Enabled = True
        TxtCoverID.Enabled = False
    ElseIf Mid(CboMCOStatus, 1, 1) = "D" Then
        Frame8.Enabled = True
        TxtCoverID.Enabled = True
    Else
        Frame8.Enabled = False
    End If
End Sub

Private Sub CboMCOStatus_Click()
    If Mid(CboMCOStatus, 1, 1) = "B" Or Mid(CboMCOStatus, 1, 1) = "C" Then
        Frame8.Enabled = True
        TxtCoverID.Enabled = False
    ElseIf Mid(CboMCOStatus, 1, 1) = "D" Then
        Frame8.Enabled = True
        TxtCoverID.Enabled = True
    Else
        Frame8.Enabled = False
    End If
    'If Mid(CboMCOStatus, 1, 1) = "B" Or Mid(CboMCOStatus, 1, 1) = "C" Or Mid(CboMCOStatus, 1, 1) = "D" Then
        'Frame8.Enabled = True
    'Else
        'Frame8.Enabled = False
    'End If
End Sub

Private Sub Combo3_Change()
    If Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
        Frame6.Enabled = True
    Else
        Frame6.Enabled = False
    End If
End Sub

Private Sub Combo3_Click()
    If Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
        Frame6.Enabled = True
    Else
        Frame6.Enabled = False
    End If
End Sub




Private Sub lstAnnualList_KeyDown(KeyCode As Integer, Shift As Integer)

    If lstAnnualList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboHltCode = lstAnnualList.SelectedItem.Text
        CboPlnCode = lstAnnualList.SelectedItem.SubItems(1)
        CboInsType = lstAnnualList.SelectedItem.SubItems(2)
        TxtAnnualLimit = lstAnnualList.SelectedItem.SubItems(3)
        TxtSupplyAL = lstAnnualList.SelectedItem.SubItems(4)
        txtALEffDate = lstAnnualList.SelectedItem.SubItems(5)
        CboPayorCode = lstAnnualList.SelectedItem.SubItems(6)
        CboGrpCom = lstAnnualList.SelectedItem.SubItems(7)
        txtANNVersion = lstAnnualList.SelectedItem.SubItems(8)
        TxtLifeTimeLimit = lstAnnualList.SelectedItem.SubItems(11)
        TxtSuppLifeTimeLimit = lstAnnualList.SelectedItem.SubItems(12)
        
        If lstAnnualList.SelectedItem.SubItems(9) = "A" Then
            Combo1.ListIndex = 0
        ElseIf lstAnnualList.SelectedItem.SubItems(9) = "B" Then
            Combo1.ListIndex = 1
        ElseIf lstAnnualList.SelectedItem.SubItems(9) = "C" Then
            Combo1.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If

End Sub

Private Sub lstAnnualList_KeyUp(KeyCode As Integer, Shift As Integer)

    If lstAnnualList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboHltCode = lstAnnualList.SelectedItem.Text
        CboPlnCode = lstAnnualList.SelectedItem.SubItems(1)
        CboInsType = lstAnnualList.SelectedItem.SubItems(2)
        TxtAnnualLimit = lstAnnualList.SelectedItem.SubItems(3)
        TxtSupplyAL = lstAnnualList.SelectedItem.SubItems(4)
        txtALEffDate = lstAnnualList.SelectedItem.SubItems(5)
        CboPayorCode = lstAnnualList.SelectedItem.SubItems(6)
        CboGrpCom = lstAnnualList.SelectedItem.SubItems(7)
        txtANNVersion = lstAnnualList.SelectedItem.SubItems(8)
        TxtLifeTimeLimit = lstAnnualList.SelectedItem.SubItems(11)
        TxtSuppLifeTimeLimit = lstAnnualList.SelectedItem.SubItems(12)
        
        If lstAnnualList.SelectedItem.SubItems(9) = "A" Then
            Combo1.ListIndex = 0
        ElseIf lstAnnualList.SelectedItem.SubItems(9) = "B" Then
            Combo1.ListIndex = 1
        ElseIf lstAnnualList.SelectedItem.SubItems(9) = "C" Then
            Combo1.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If

End Sub

Private Sub lstHLTList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstHLTList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtHLTcode = lstHLTList.SelectedItem.Text
        txtHLTcode.Enabled = False
        txtHLTdescription = lstHLTList.SelectedItem.SubItems(1)
        txtHLTdescription.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstHLTList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstHLTList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtHLTcode = lstHLTList.SelectedItem.Text
        txtHLTcode.Enabled = False
        txtHLTdescription = lstHLTList.SelectedItem.SubItems(1)
        txtHLTdescription.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstINSList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstINSList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        txtINScode = lstINSList.SelectedItem.Text
        txtINScode.Enabled = False
        txtINSdescription = lstINSList.SelectedItem.SubItems(1)
        txtINSdescription.Enabled = False
        'Call EntriesEnable(False)
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstINSList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstINSList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        txtINScode = lstINSList.SelectedItem.Text
        txtINScode.Enabled = False
        txtINSdescription = lstINSList.SelectedItem.SubItems(1)
        txtINSdescription.Enabled = False
        'Call EntriesEnable(False)
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstMCOList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim MCO As New ADODB.Recordset
Dim strsql As String
Dim Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Amt7, Amt8 As String
Dim Amt9, Amt10, Amt11, Amt12 As String
   
   Amt1 = 0
   Amt2 = 0
   Amt3 = 0
   Amt4 = 0
   Amt5 = 0
   Amt6 = 0
   Amt7 = 0
   Amt8 = 0
   Amt9 = 0
   Amt10 = 0
   Amt11 = 0
   Amt12 = 0
   
  ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
   
   If lstMCOList.ListItems.Count <> 0 Then
        Call ClearForm(Me)
        
        strsql = " Select * From MCO where PLNCode = '" & lstMCOList.SelectedItem.SubItems(2) & "' AND " & _
                 " MCOEffDate = '" & Format(lstMCOList.SelectedItem.SubItems(8), "mm/dd/yyyy") & "'"
        
        MCO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        Do Until MCO.EOF
            If MCO!SuppMCOStatus = "A" Then
                If MCO!INSCode = "I" Then 'And MCO!AdultChild = "A" Then
                    Amt1 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "I" And MCO!AdultChild = "C" Then
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then 'And MCO!AdultChild = "A" Then
                    Amt3 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "F" And MCO!AdultChild = "C" Then
                    Amt4 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "G" Then 'And MCO!AdultChild = "A" Then
                    Amt5 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "G" And MCO!AdultChild = "C" Then
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then 'And MCO!AdultChild = "A" Then
                    Amt7 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "H" And MCO!AdultChild = "C" Then
                    Amt8 = MCO!MCOAmountChild
                End If
            ElseIf MCO!SuppMCOStatus = "B" Then
                If MCO!INSCode = "I" Then 'And MCO!AdultChild = "A" Then
                    Amt1 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "I" And MCO!AdultChild = "C" Then
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then 'And MCO!AdultChild = "A" Then
                    Amt3 = MCO!MCOAmountAdult
                    Amt9 = MCO!MCOSuppAdultAmt
                'ElseIf MCO!INSCode = "F" And MCO!AdultChild = "C" Then
                    Amt4 = MCO!MCOAmountChild
                    Amt10 = MCO!MCOSuppChildAmt
                ElseIf MCO!INSCode = "G" Then 'And MCO!AdultChild = "A" Then
                    Amt5 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "G" And MCO!AdultChild = "C" Then
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then 'And MCO!AdultChild = "A" Then
                    Amt7 = MCO!MCOAmountAdult
                    Amt11 = MCO!MCOSuppAdultAmt
                'ElseIf MCO!INSCode = "H" And MCO!AdultChild = "C" Then
                    Amt8 = MCO!MCOAmountChild
                    Amt12 = MCO!MCOSuppChildAmt
                End If
            ElseIf MCO!SuppMCOStatus = "C" Then
                If MCO!INSCode = "I" Then 'And MCO!AdultChild = "A" Then
                    Amt1 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "I" And MCO!AdultChild = "C" Then
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then 'And MCO!AdultChild = "A" Then
                    Amt3 = MCO!MCOAmountAdult
                    Amt9 = MCO!MCOSuppAdultAmt
                'ElseIf MCO!INSCode = "F" And MCO!AdultChild = "C" Then
                    Amt4 = MCO!MCOAmountChild
                    Amt10 = MCO!MCOSuppChildAmt
                ElseIf MCO!INSCode = "G" Then 'And MCO!AdultChild = "A" Then
                    Amt5 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "G" And MCO!AdultChild = "C" Then
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then 'And MCO!AdultChild = "A" Then
                    Amt7 = MCO!MCOAmountAdult
                    Amt11 = MCO!MCOSuppAdultAmt
                'ElseIf MCO!INSCode = "H" And MCO!AdultChild = "C" Then
                    Amt8 = MCO!MCOAmountChild
                    Amt12 = MCO!MCOSuppChildAmt
                End If
            ElseIf MCO!SuppMCOStatus = "D" Then
                If MCO!INSCode = "I" Then
                    Amt1 = MCO!MCOAmountAdult
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then
                    Amt3 = MCO!MCOAmountAdult
                    Amt9 = MCO!MCOSuppAdultAmt
                    Amt4 = MCO!MCOAmountChild
                    Amt10 = MCO!MCOSuppChildAmt
                ElseIf MCO!INSCode = "G" Then
                    Amt5 = MCO!MCOAmountAdult
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then
                    Amt7 = MCO!MCOAmountAdult
                    Amt11 = MCO!MCOSuppAdultAmt
                    Amt8 = MCO!MCOAmountChild
                    Amt12 = MCO!MCOSuppChildAmt
                End If
            End If
        MCO.MoveNext
        Loop
        MCO.Close
        Set MCO = Nothing
    
                
        CboMCOHlt = lstMCOList.SelectedItem.Text
        TxtMCOPlan = lstMCOList.SelectedItem.SubItems(2)
        
        If lstMCOList.SelectedItem.SubItems(3) = "A" Then
            CboMCOStatus.ListIndex = 0
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "B" Then
            CboMCOStatus.ListIndex = 1
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "C" Then
            CboMCOStatus.ListIndex = 2
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "D" Then
            CboMCOStatus.ListIndex = 3
        End If
        
        TxtIAAmt = Format(Amt1, "#,##0.00")
        TxtICAmt = Format(Amt2, "#,##0.00")
    
        TxtFAAmt = Format(Amt3, "#,##0.00")
        TxtFCAmt = Format(Amt4, "#,##0.00")
        TxtSFAAmt = Format(Amt9, "#,##0.00")
        TxtSFCAmt = Format(Amt10, "#,##0.00")
    
        TxtGAAmt = Format(Amt5, "#,##0.00")
        TxtGCAmt = Format(Amt6, "#,##0.00")
    
        TxtHAAmt = Format(Amt7, "#,##0.00")
        TxtHCAmt = Format(Amt8, "#,##0.00")
        TxtSHAAmt = Format(Amt11, "#,##0.00")
        TxtSHCAmt = Format(Amt12, "#,##0.00")
        
        txtEffDate = lstMCOList.SelectedItem.SubItems(8)
        TxtCoverID = lstMCOList.SelectedItem.SubItems(9)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
'    medixmasterconn.Close
 '   Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
'    Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub lstMCOList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim MCO As New ADODB.Recordset
Dim strsql As String
Dim Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Amt7, Amt8 As String
Dim Amt9, Amt10, Amt11, Amt12 As String
   
   Amt1 = 0
   Amt2 = 0
   Amt3 = 0
   Amt4 = 0
   Amt5 = 0
   Amt6 = 0
   Amt7 = 0
   Amt8 = 0
   Amt9 = 0
   Amt10 = 0
   Amt11 = 0
   Amt12 = 0
   
  ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
   
   If lstMCOList.ListItems.Count <> 0 Then
        Call ClearForm(Me)
        
        strsql = " Select * From MCO where PLNCode = '" & lstMCOList.SelectedItem.SubItems(2) & "' AND " & _
                 " MCOEffDate = '" & Format(lstMCOList.SelectedItem.SubItems(8), "mm/dd/yyyy") & "'"
                 
        MCO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MCO.recordCount Then
        End If
        Do Until MCO.EOF
            If MCO!SuppMCOStatus = "A" Then
                If MCO!INSCode = "I" Then 'And MCO!AdultChild = "A" Then
                    Amt1 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "I" And MCO!AdultChild = "C" Then
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then 'And MCO!AdultChild = "A" Then
                    Amt3 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "F" And MCO!AdultChild = "C" Then
                    Amt4 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "G" Then 'And MCO!AdultChild = "A" Then
                    Amt5 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "G" And MCO!AdultChild = "C" Then
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then 'And MCO!AdultChild = "A" Then
                    Amt7 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "H" And MCO!AdultChild = "C" Then
                    Amt8 = MCO!MCOAmountChild
                End If
            ElseIf MCO!SuppMCOStatus = "B" Then
                If MCO!INSCode = "I" Then 'And MCO!AdultChild = "A" Then
                    Amt1 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "I" And MCO!AdultChild = "C" Then
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then 'And MCO!AdultChild = "A" Then
                    Amt3 = MCO!MCOAmountAdult
                    Amt9 = MCO!MCOSuppAdultAmt
                'ElseIf MCO!INSCode = "F" And MCO!AdultChild = "C" Then
                    Amt4 = MCO!MCOAmountChild
                    Amt10 = MCO!MCOSuppChildAmt
                ElseIf MCO!INSCode = "G" Then 'And MCO!AdultChild = "A" Then
                    Amt5 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "G" And MCO!AdultChild = "C" Then
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then 'And MCO!AdultChild = "A" Then
                    Amt7 = MCO!MCOAmountAdult
                    Amt11 = MCO!MCOSuppAdultAmt
                'ElseIf MCO!INSCode = "H" And MCO!AdultChild = "C" Then
                    Amt8 = MCO!MCOAmountChild
                    Amt12 = MCO!MCOSuppChildAmt
                End If
            ElseIf MCO!SuppMCOStatus = "C" Then
                If MCO!INSCode = "I" Then 'And MCO!AdultChild = "A" Then
                    Amt1 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "I" And MCO!AdultChild = "C" Then
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then 'And MCO!AdultChild = "A" Then
                    Amt3 = MCO!MCOAmountAdult
                    Amt9 = MCO!MCOSuppAdultAmt
                'ElseIf MCO!INSCode = "F" And MCO!AdultChild = "C" Then
                    Amt4 = MCO!MCOAmountChild
                    Amt10 = MCO!MCOSuppChildAmt
                ElseIf MCO!INSCode = "G" Then 'And MCO!AdultChild = "A" Then
                    Amt5 = MCO!MCOAmountAdult
                'ElseIf MCO!INSCode = "G" And MCO!AdultChild = "C" Then
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then 'And MCO!AdultChild = "A" Then
                    Amt7 = MCO!MCOAmountAdult
                    Amt11 = MCO!MCOSuppAdultAmt
                'ElseIf MCO!INSCode = "H" And MCO!AdultChild = "C" Then
                    Amt8 = MCO!MCOAmountChild
                    Amt12 = MCO!MCOSuppChildAmt
                End If
            ElseIf MCO!SuppMCOStatus = "D" Then
                If MCO!INSCode = "I" Then
                    Amt1 = MCO!MCOAmountAdult
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then
                    Amt3 = MCO!MCOAmountAdult
                    Amt9 = MCO!MCOSuppAdultAmt
                    Amt4 = MCO!MCOAmountChild
                    Amt10 = MCO!MCOSuppChildAmt
                ElseIf MCO!INSCode = "G" Then
                    Amt5 = MCO!MCOAmountAdult
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then
                    Amt7 = MCO!MCOAmountAdult
                    Amt11 = MCO!MCOSuppAdultAmt
                    Amt8 = MCO!MCOAmountChild
                    Amt12 = MCO!MCOSuppChildAmt
                End If
            End If
        MCO.MoveNext
        Loop
        MCO.Close
        Set MCO = Nothing
    
                
        CboMCOHlt = lstMCOList.SelectedItem.Text
        TxtMCOPlan = lstMCOList.SelectedItem.SubItems(2)
        
        If lstMCOList.SelectedItem.SubItems(3) = "A" Then
            CboMCOStatus.ListIndex = 0
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "B" Then
            CboMCOStatus.ListIndex = 1
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "C" Then
            CboMCOStatus.ListIndex = 2
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "D" Then
            CboMCOStatus.ListIndex = 3
        End If
                
        TxtIAAmt = Format(Amt1, "#,##0.00")
        TxtICAmt = Format(Amt2, "#,##0.00")
    
        TxtFAAmt = Format(Amt3, "#,##0.00")
        TxtFCAmt = Format(Amt4, "#,##0.00")
        TxtSFAAmt = Format(Amt9, "#,##0.00")
        TxtSFCAmt = Format(Amt10, "#,##0.00")
    
        TxtGAAmt = Format(Amt5, "#,##0.00")
        TxtGCAmt = Format(Amt6, "#,##0.00")
    
        TxtHAAmt = Format(Amt7, "#,##0.00")
        TxtHCAmt = Format(Amt8, "#,##0.00")
        TxtSHAAmt = Format(Amt11, "#,##0.00")
        TxtSHCAmt = Format(Amt12, "#,##0.00")
        
        txtEffDate = lstMCOList.SelectedItem.SubItems(8)
        TxtCoverID = lstMCOList.SelectedItem.SubItems(9)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
'    Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub lstPlanList_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstPlanList.ListItems.Count <> 0 Then
        Call lstPlanListShow
    End If
End Sub

Private Sub lstPlanList_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstPlanList.ListItems.Count <> 0 Then
        Call lstPlanListShow
    End If
End Sub

Private Sub lstPmPRMList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim prm As New ADODB.Recordset
Dim strsql As String
Dim AGE01, AGE02, AGE03, AGE04, AGE05, AGE06, AGE07 As String
ListClick = True
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If lstPmPRMList.ListItems.Count <> 0 Then
    
    Call setCmdButton(Me, False)
    'Call ClearForm(Me)
       
    txtPRMEffdate.Enabled = False
    FrameAGEBand.Enabled = False
    Combo3.Enabled = False
    txtPREVersion.Enabled = False
    
    strsql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
    prm.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until prm.EOF
        If prm!AGECode = "01" Then
            AGE01 = prm!PRMAmount
        ElseIf prm!AGECode = "02" Then
            AGE02 = prm!PRMAmount
        ElseIf prm!AGECode = "03" Then
            AGE03 = prm!PRMAmount
        ElseIf prm!AGECode = "04" Then
            AGE04 = prm!PRMAmount
        ElseIf prm!AGECode = "05" Then
            AGE05 = prm!PRMAmount
        ElseIf prm!AGECode = "06" Then
            AGE06 = prm!PRMAmount
        ElseIf prm!AGECode = "07" Then
            AGE07 = prm!PRMAmount
   
   End If
   prm.MoveNext
    Loop
   prm.Close
   Set prm = Nothing
   
    
    txtPrmCode = lstPmPRMList.SelectedItem.Text
    cboPmINSCode.Text = lstPmPRMList.SelectedItem.SubItems(1)
    cboPmINSCode.Enabled = False
    cboPmHLTCode.Text = lstPmPRMList.SelectedItem.SubItems(3)
    cboPmHLTCode.Enabled = False
    cboPmPAYCode.Text = lstPmPRMList.SelectedItem.SubItems(2)
    cboPmPAYCode.Enabled = False
    cboPmAgeCode.Text = lstPmPRMList.SelectedItem.SubItems(4)
    cboPmAgeCode.Enabled = False
    cboPmPLNCode.Text = lstPmPRMList.SelectedItem.SubItems(5)
    cboPmPLNCode.Enabled = False
    txt01 = Format(AGE01, "#,###0.00")
    txt02 = Format(AGE02, "#,###0.00")
    txt03 = Format(AGE03, "#,###0.00")
    txt04 = Format(AGE04, "#,###0.00")
    txt05 = Format(AGE05, "#,###0.00")
    txt06 = Format(AGE06, "#,###0.00")
    txt07 = Format(AGE07, "#,###0.00")
    FrameAGEBand.Enabled = False
    AGE01 = 0
    AGE02 = 0
    AGE03 = 0
    AGE04 = 0
    AGE05 = 0
    AGE06 = 0
    AGE07 = 0
    
    strsql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
    prm.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until prm.EOF
        If prm!AGECode = "01" Then
            If Len(prm!SuppPrmAmt) Then
                AGE01 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "02" Then
            If Len(prm!SuppPrmAmt) Then
                AGE02 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "03" Then
            If Len(prm!SuppPrmAmt) Then
                AGE03 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "04" Then
            If Len(prm!SuppPrmAmt) Then
                AGE04 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "05" Then
            If Len(prm!SuppPrmAmt) Then
                AGE05 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "06" Then
            If Len(prm!SuppPrmAmt) Then
                AGE06 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "07" Then
            If Len(prm!SuppPrmAmt) Then
                AGE07 = prm!SuppPrmAmt
            End If
   
   End If
   prm.MoveNext
    Loop
   prm.Close
   Set prm = Nothing
    
    Text2 = Format(AGE01, "#,###0.00")
    Text3 = Format(AGE02, "#,###0.00")
    Text4 = Format(AGE03, "#,###0.00")
    Text5 = Format(AGE04, "#,###0.00")
    Text6 = Format(AGE05, "#,###0.00")
    Text7 = Format(AGE06, "#,###0.00")
    Text1 = Format(AGE07, "#,###0.00")
    Frame6.Enabled = False
    
    txtPRMEffdate.mask = ""
    txtPRMEffdate = lstPmPRMList.SelectedItem.SubItems(8)
    txtPRMEffdate.mask = "##/##/####"
    txtPREVersion = lstPmPRMList.SelectedItem.SubItems(9)
    
    If lstPmPRMList.SelectedItem.SubItems(10) = "A" Then
        Combo3.ListIndex = 0
    ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "B" Then
        Combo3.ListIndex = 1
    ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "C" Then
        Combo3.ListIndex = 2
    End If
    
    CmdAdd.Enabled = True
  
    End If
    
 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub lstPmPRMList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim prm As New ADODB.Recordset
Dim strsql As String
Dim AGE01, AGE02, AGE03, AGE04, AGE05, AGE06, AGE07 As String
ListClick = True
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If lstPmPRMList.ListItems.Count <> 0 Then
    
    Call setCmdButton(Me, False)
    'Call ClearForm(Me)
       
    txtPRMEffdate.Enabled = False
    FrameAGEBand.Enabled = False
    Combo3.Enabled = False
    txtPREVersion.Enabled = False
    
    strsql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
    prm.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until prm.EOF
        If prm!AGECode = "01" Then
            AGE01 = prm!PRMAmount
        ElseIf prm!AGECode = "02" Then
            AGE02 = prm!PRMAmount
        ElseIf prm!AGECode = "03" Then
            AGE03 = prm!PRMAmount
        ElseIf prm!AGECode = "04" Then
            AGE04 = prm!PRMAmount
        ElseIf prm!AGECode = "05" Then
            AGE05 = prm!PRMAmount
        ElseIf prm!AGECode = "06" Then
            AGE06 = prm!PRMAmount
        ElseIf prm!AGECode = "07" Then
            AGE07 = prm!PRMAmount
   
   End If
   prm.MoveNext
    Loop
   prm.Close
   Set prm = Nothing
   
    txtPrmCode = lstPmPRMList.SelectedItem.Text
    cboPmINSCode.Text = lstPmPRMList.SelectedItem.SubItems(1)
    cboPmINSCode.Enabled = False
    cboPmHLTCode.Text = lstPmPRMList.SelectedItem.SubItems(3)
    cboPmHLTCode.Enabled = False
    cboPmPAYCode.Text = lstPmPRMList.SelectedItem.SubItems(2)
    cboPmPAYCode.Enabled = False
    cboPmAgeCode.Text = lstPmPRMList.SelectedItem.SubItems(4)
    cboPmAgeCode.Enabled = False
    cboPmPLNCode.Text = lstPmPRMList.SelectedItem.SubItems(5)
    cboPmPLNCode.Enabled = False
    txt01 = Format(AGE01, "#,###0.00")
    txt02 = Format(AGE02, "#,###0.00")
    txt03 = Format(AGE03, "#,###0.00")
    txt04 = Format(AGE04, "#,###0.00")
    txt05 = Format(AGE05, "#,###0.00")
    txt06 = Format(AGE06, "#,###0.00")
    txt07 = Format(AGE07, "#,###0.00")
    FrameAGEBand.Enabled = False
        
    AGE01 = 0
    AGE02 = 0
    AGE03 = 0
    AGE04 = 0
    AGE05 = 0
    AGE06 = 0
    AGE07 = 0
    
    strsql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
    prm.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until prm.EOF
        If prm!AGECode = "01" Then
            If Len(prm!SuppPrmAmt) Then
                AGE01 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "02" Then
            If Len(prm!SuppPrmAmt) Then
                AGE02 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "03" Then
            If Len(prm!SuppPrmAmt) Then
                AGE03 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "04" Then
            If Len(prm!SuppPrmAmt) Then
                AGE04 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "05" Then
            If Len(prm!SuppPrmAmt) Then
                AGE05 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "06" Then
            If Len(prm!SuppPrmAmt) Then
                AGE06 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "07" Then
            If Len(prm!SuppPrmAmt) Then
                AGE07 = prm!SuppPrmAmt
            End If
   
   End If
   prm.MoveNext
    Loop
   prm.Close
   Set prm = Nothing
    
    Text2 = Format(AGE01, "#,###0.00")
    Text3 = Format(AGE02, "#,###0.00")
    Text4 = Format(AGE03, "#,###0.00")
    Text5 = Format(AGE04, "#,###0.00")
    Text6 = Format(AGE05, "#,###0.00")
    Text7 = Format(AGE06, "#,###0.00")
    Text1 = Format(AGE07, "#,###0.00")
    Frame6.Enabled = False
    
    txtPRMEffdate.mask = ""
    txtPRMEffdate = lstPmPRMList.SelectedItem.SubItems(8)
    txtPRMEffdate.mask = "##/##/####"
    txtPREVersion = lstPmPRMList.SelectedItem.SubItems(9)
    
    If lstPmPRMList.SelectedItem.SubItems(10) = "A" Then
        Combo3.ListIndex = 0
    ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "B" Then
        Combo3.ListIndex = 1
    ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "C" Then
        Combo3.ListIndex = 2
    End If
    
    CmdAdd.Enabled = True
  
    End If
    
'    medixmasterconn.Close
 '   Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    
End Sub


Private Sub lstSpecialNoteList_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If lstSpecialNoteList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboSpecialPlan = lstSpecialNoteList.SelectedItem.Text
        CboSpecialGrpCom = lstSpecialNoteList.SelectedItem.SubItems(1)
        TxtSpecialNote = lstSpecialNoteList.SelectedItem.SubItems(2)
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstSpecialNoteList_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If lstSpecialNoteList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboSpecialPlan = lstSpecialNoteList.SelectedItem.Text
        CboSpecialGrpCom = lstSpecialNoteList.SelectedItem.SubItems(1)
        TxtSpecialNote = lstSpecialNoteList.SelectedItem.SubItems(2)
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstALB_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If lstALB.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        cboALBHltCode = lstALB.SelectedItem.Text
        cboALBPlnCode = lstALB.SelectedItem.SubItems(1)
        cboALBInsType = lstALB.SelectedItem.SubItems(2)
        txtALBEffDate = lstALB.SelectedItem.SubItems(3)
        txtALBKidDiaLmt = lstALB.SelectedItem.SubItems(4)
        txtALBCancerLmt = lstALB.SelectedItem.SubItems(5)
        txtALBHomeNursLmt = lstALB.SelectedItem.SubItems(6)
        
        If lstALB.SelectedItem.SubItems(7) = "A" Then
            Combo4.ListIndex = 0
        ElseIf lstALB.SelectedItem.SubItems(7) = "B" Then
            Combo4.ListIndex = 1
        ElseIf lstALB.SelectedItem.SubItems(7) = "C" Then
            Combo4.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstALB_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If lstALB.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        cboALBHltCode = lstALB.SelectedItem.Text
        cboALBPlnCode = lstALB.SelectedItem.SubItems(1)
        cboALBInsType = lstALB.SelectedItem.SubItems(2)
        txtALBEffDate = lstALB.SelectedItem.SubItems(3)
        txtALBKidDiaLmt = lstALB.SelectedItem.SubItems(4)
        txtALBCancerLmt = lstALB.SelectedItem.SubItems(5)
        txtALBHomeNursLmt = lstALB.SelectedItem.SubItems(6)
        
        If lstALB.SelectedItem.SubItems(7) = "A" Then
            Combo4.ListIndex = 0
        ElseIf lstALB.SelectedItem.SubItems(7) = "B" Then
            Combo4.ListIndex = 1
        ElseIf lstALB.SelectedItem.SubItems(7) = "C" Then
            Combo4.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstProductCat_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstProductCat.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtProCode = lstProductCat.SelectedItem.Text
        txtProCode.Enabled = False
        txtProName = lstProductCat.SelectedItem.SubItems(1)
        txtProName.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstProductCat_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strsql As String
    
    If lstProductCat.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtProCode = lstProductCat.SelectedItem.Text
        txtProCode.Enabled = False
        txtProName = lstProductCat.SelectedItem.SubItems(1)
        txtProName.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
End Sub

Private Sub lstALB_Click()
    
    If lstALB.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        cboALBHltCode = lstALB.SelectedItem.Text
        cboALBPlnCode = lstALB.SelectedItem.SubItems(1)
        cboALBInsType = lstALB.SelectedItem.SubItems(2)
        txtALBEffDate = lstALB.SelectedItem.SubItems(3)
        txtALBKidDiaLmt = lstALB.SelectedItem.SubItems(4)
        txtALBCancerLmt = lstALB.SelectedItem.SubItems(5)
        txtALBHomeNursLmt = lstALB.SelectedItem.SubItems(6)
        
        If lstALB.SelectedItem.SubItems(7) = "A" Then
            Combo4.ListIndex = 0
        ElseIf lstALB.SelectedItem.SubItems(7) = "B" Then
            Combo4.ListIndex = 1
        ElseIf lstALB.SelectedItem.SubItems(7) = "C" Then
            Combo4.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
        
    Save = False
    Screen.MousePointer = vbHourglass
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
'    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    Select Case SSTab1.Tab
    Case 0
        Call ClearForm(Me)
        lstHLTList.ListItems.Clear
        CmdRefresh.Enabled = False
        CmdEdit.Enabled = False
        CmdSave.Enabled = True
        CmdDelete.Enabled = False
        txtTotalRecord = 0
    Case 1
        Call ClearForm(Me)
        lstINSList.ListItems.Clear
        CmdRefresh.Enabled = False
        CmdEdit.Enabled = False
        CmdSave.Enabled = True
        CmdDelete.Enabled = False
        txtTotalRecord = 0
    Case 2
        Call ClearForm(Me)
        lstPlanList.ListItems.Clear
        If cboPLNPayor.ListCount = 0 Then
            Call PLNComboListing
        End If
        CmdRefresh.Enabled = True
        CmdEdit.Enabled = False
        CmdSave.Enabled = True
        CmdDelete.Enabled = False
        txtTotalRecord = 0
    Case 3
        Call ClearForm(Me)
        CmdEdit.Enabled = False
        CmdDelete.Enabled = False
        CmdRefresh.Enabled = True
        CmdSave.Enabled = True
        Frame4.Enabled = True
        If cboPmHLTCode.ListCount = 0 Then
            Call PRMComboListing
        End If
        lstPmPRMList.ListItems.Clear
        txtTotalRecord = 0
        
        If Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
            Frame6.Enabled = True
        Else
            Frame6.Enabled = False
        End If
    Case 4
        Call ClearForm(Me)
        CmdEdit.Enabled = False
        CmdDelete.Enabled = False
        CmdRefresh.Enabled = True
        CmdSave.Enabled = True
        If CboPlnCode.ListCount = 0 Then
            Call AnnualComboListing
        End If
        lstAnnualList.ListItems.Clear
        txtTotalRecord = 0
        
        If Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
            TxtSupplyAL.Enabled = True
            TxtSuppLifeTimeLimit.Enabled = True
        Else
            TxtSupplyAL.Enabled = False
            TxtSuppLifeTimeLimit.Enabled = False
        End If
        
    Case 5
        Call ClearForm(Me)
        CmdEdit.Enabled = False
        CmdDelete.Enabled = False
        CmdRefresh.Enabled = True
        CmdSave.Enabled = True
        Call MCOComboListing
        lstMCOList.ListItems.Clear
        txtTotalRecord = 0
        
    Case 6
        Call ClearForm(Me)
        CmdEdit.Enabled = False
        CmdDelete.Enabled = False
        CmdRefresh.Enabled = True
        CmdSave.Enabled = True
        'Call SpecialNoteComboListing
        lstSpecialNoteList.ListItems.Clear
        txtTotalRecord = 0
        
    Case 7
        Call ClearForm(Me)
        CmdEdit.Enabled = False
        CmdDelete.Enabled = False
        CmdRefresh.Enabled = True
        CmdSave.Enabled = True
        Call ALBComboListing
        lstALB.ListItems.Clear
        txtTotalRecord = 0
        
    Case 8
        
        
    Case 9
        
        
    Case 10
        Call ClearForm(Me)
        CmdEdit.Enabled = False
        CmdDelete.Enabled = False
        CmdRefresh.Enabled = True
        CmdSave.Enabled = True
        Call MgmtFeeCombolisting
        lstMgmtFeeList.ListItems.Clear
        txtTotalRecord = 0
    
        
    End Select
'    Set medixmasterconn = Nothing
 '   medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    Screen.MousePointer = Default
End Sub


Private Sub CmdRefresh_Click()
Screen.MousePointer = vbHourglass
'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
Save = True
Call ClearForm(Me)
Call EntriesEnable(True)
Label5.Visible = True

Select Case SSTab1.Tab
    Case 0
        'Call HLTListing
    Case 1
        'Call INSListing
    Case 2
        'Call PLNListing
    Case 3
        Call PRMComboListing
    Case 4
        Call AnnualComboListing
    Case 5
        Call MCOComboListing
    Case 6
        'Call SpecialNoteComboListing
    Case 7
        Call ALBComboListing
    Case 8
        'Call ProCatCombolisting
    Case 9
        'Call MCOStatComboListing
    Case 10
        Call MgmtFeeCombolisting
        
End Select
Save = False
txtTotalRecord = 0
'medixmasterconn.Close
'Set medixmasterconn = Nothing
'medixmasterconnMaint.Close
'Set medixmasterconnMaint = Nothing
Screen.MousePointer = Default
End Sub

'**************************Start Form******************************
Private Sub Form_Load()
    
    CmdRefresh.Enabled = False
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    Frame4.Enabled = True
    Me.Width = SSTab1.Width + 500

    PRMCriteria = ""
    PLNCriteria = ""
    HLTCriteria = ""
    INSCriteria = ""
    AnnualCriteria = ""
    MCOCriteria = ""
    ALBCriteria = ""
    ProCatCriteria = ""
    MCOStatCriteria = ""
    frmMaintenancePlan.SSTab1.Tab = 2
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub
'**************************End Form******************************

'**************************Start Add, Edit, Delete, Search, Save, Exit******************************
Private Sub cmdAdd_Click()
Dim rst As New ADODB.Recordset
Dim strsql As String
Dim tmpCode As String
Dim AscCode
Dim strCount
    
    AscCode = 0
    tmpCode = ""
    EditFlag = False
    txtTotalRecord = 0
    Call ClearForm(Me)
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    cboPmAgeCode.Enabled = False
    
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Select Case SSTab1.Tab
        Case 0
            txtHLTcode.SetFocus
        Case 1
            txtINScode.SetFocus
        Case 2
            cboPLNHLTCode.SetFocus
            Frame2.Enabled = True
        Case 3
            FrameAGEBand.Enabled = True
            'Frame6.Enabled = True
            If Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
                Frame6.Enabled = True
            Else
                Frame6.Enabled = False
            End If
            'cboPmAgeCode.Enabled = True
            cboPmHLTCode.SetFocus
        Case 4
            CboHltCode.SetFocus
            Frame2.Enabled = True
            Frame4.Enabled = True
        Case 5
            CboMCOHlt.SetFocus
            Call setCmdButton(Me, False)
            CmdSave.Enabled = True
            CmdDelete.Enabled = False
        Case 6
            CboSpecialPlan.SetFocus
        Case 7
            cboALBHltCode.SetFocus
        Case 8
            txtProCode.SetFocus
            CmdEdit.Enabled = False
            CmdDelete.Enabled = False
            CmdRefresh.Enabled = False
        Case 9
        
            strsql = "Select MCOCode From SuppMCOStatus"
            rst.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
            txtMCOCode.Enabled = False
            txtMCOStatus.SetFocus
            CmdEdit.Enabled = False
            CmdDelete.Enabled = False
            CmdRefresh.Enabled = False
            If rst.recordCount = 0 Then
                tmpCode = Chr(65)
                txtMCOCode = tmpCode
            Else
                For strCount = 65 To 90
                    Do Until rst.EOF
                        If strCount = Asc(rst!MCOCode) Then
                            AscCode = Asc(rst!MCOCode) + 1
                            rst.MoveNext
                            Exit Do
                        Else
                            AscCode = strCount
                            Exit For
                        End If
                    Loop
                Next
                tmpCode = Chr(AscCode)
                txtMCOCode = tmpCode
                AscCode = 0
            End If
            rst.Close
            Set rst = Nothing
        Case 10
            CboMgmtHltCode.SetFocus
            CmdRefresh.Enabled = True
    End Select
   ' medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
        
End Sub

Private Sub CmdEdit_Click()
    EditFlag = True
     Select Case SSTab1.Tab
        Case 2
        txtLGPrivate.Enabled = True
        Case Default
             Call EntriesEnable(True)
    End Select
    CmdSave.Enabled = True
    CmdAdd.Enabled = True
    CmdDelete.Enabled = False
    Frame4.Enabled = True
              
    If txtHLTcode <> "" Then
       txtHLTcode.Enabled = False
       txtHLTdescription.SetFocus
    End If
    
    If txtINScode <> "" Then
       txtINScode.Enabled = False
       txtINSdescription.Enabled = True
       txtINSdescription.SetFocus
    End If
    
    If cboPLNHLTCode.Text <> "" Then
       cboPLNHLTCode.Enabled = True
       cboPLNPayor.Enabled = True
       Frame2.Enabled = True
       Call PLNDisEnable
    End If
    If cboPmINSCode <> "" Then
       cboPmPAYCode.Enabled = False
       cboPmHLTCode.Enabled = False
       cboPmAgeCode.Enabled = False
       cboPmPLNCode.Enabled = False
       cboPmINSCode.Enabled = False
       txtPRMEffdate.Enabled = False
    End If
            
    If Mid(Combo1, 1, 1) = "A" Then
      TxtSupplyAL.Enabled = False
      TxtSuppLifeTimeLimit.Enabled = False
    Else
      TxtSupplyAL.Enabled = True
      TxtSuppLifeTimeLimit.Enabled = True
    End If
    
    If CboSpecialPlan <> "" Then
       CboSpecialPlan.Enabled = False
       CboSpecialGrpCom.Enabled = False
    End If
    
    If txtProCode <> "" Then
        txtProCode.Enabled = False
        txtProName.Enabled = True
    End If
    
    If txtMCOStatus <> "" Then
        txtMCOCode.Enabled = False
        txtMCOStatus.Enabled = True
        CmdSave.Enabled = True
        txtMCOStatus.SetFocus
    End If
    
    If CboMgmtHltCode <> "" Then
        CboMgmtHltCode.Enabled = False
        CboMgmtPlan.Enabled = False
        CboMgmtIns.Enabled = False
        TxtMgmtFee.Enabled = True
    End If

End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    If SSTab1.Tab = 0 Then
       Call DeleteHLT
       CmdAdd.Enabled = True
    ElseIf SSTab1.Tab = 1 Then
      Call DeleteINS
      CmdAdd.Enabled = True
    ElseIf SSTab1.Tab = 2 Then
       Call DeletePLN
       CmdAdd.Enabled = True
    ElseIf SSTab1.Tab = 3 Then
      Call DeletePRM
      CmdAdd.Enabled = True
    ElseIf SSTab1.Tab = 4 Then
      Call DeleteAnnual
      CmdAdd.Enabled = True
    ElseIf SSTab1.Tab = 5 Then
        'Call DeleteMCOFees
    ElseIf SSTab1.Tab = 6 Then
      Call DeleteSpecialNote
      CmdAdd.Enabled = True
    ElseIf SSTab1.Tab = 7 Then
      Call DeleteALB
      CmdAdd.Enabled = True
    ElseIf SSTab1.Tab = 8 Then
      Call DeleteProCat
      CmdAdd.Enabled = True
    ElseIf SSTab1.Tab = 9 Then
      Call DeleteMCOStatus
      CmdAdd.Enabled = True
    ElseIf SSTab1.Tab = 10 Then
      Call DeleteCaseMgmtFee
      CmdAdd.Enabled = True
    End If
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing

Screen.MousePointer = Default
End Sub

Private Sub CmdSearch_Click()
Screen.MousePointer = vbHourglass
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
'    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    CmdSearch.Enabled = False
    Select Case SSTab1.Tab
        Case 0
            Call SearchHLT
            Call ClearForm(Me)
            txtHLTcode.Enabled = True
            txtHLTdescription.Enabled = True
            txtHLTcode.SetFocus
        Case 1
            Call SearchINS
            txtINScode.Enabled = True
            txtINSdescription.Enabled = True
            txtINScode.SetFocus
        Case 2
            Call SearchPLN
            cboPLNHLTCode.Enabled = True
            TxtGrpCom.Enabled = True
            
            cboPLNHLTCode.SetFocus
        Case 3
            Call SearchPRM
            cboPmINSCode.Enabled = True
            cboPmAgeCode.Enabled = True
            cboPmPAYCode.Enabled = True
            cboPmPLNCode.Enabled = True
            cboPmHLTCode.Enabled = True
            cboPmHLTCode.SetFocus
        Case 4
            Call SearchAnnual
            CboPlnCode.Enabled = True
            CboHltCode.Enabled = True
            CboInsType.Enabled = True
            Frame4.Enabled = True
            TxtSupplyAL.Enabled = False
            TxtSuppLifeTimeLimit.Enabled = False
            txtALEffDate.Enabled = False
            CboPlnCode.SetFocus
               
        Case 5
            Call SearchMCO
            CboMCOHlt.Enabled = True
            txtEffDate.Enabled = True
            CboMCOHlt.SetFocus
            
        Case 6
            Call SearchSpecialNote
            CboSpecialPlan.Enabled = True
            CboSpecialGrpCom.Enabled = True
            CboSpecialPlan.SetFocus
        
        Case 7
            Call SearchALB
            cboALBHltCode.SetFocus
        
        Case 8
            Call SearchProCat
            txtProCode.SetFocus
        
        Case 9
            Call SearchMCOStat
            txtMCOStatus.SetFocus
            
        Case 10
            Call SearchMgmtFee
            'CboMgmtHltCode.SetFocus
            
       End Select
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
      '  medixmasterconnMaint.Close
      '  Set medixmasterconnMaint = Nothing
        CmdSearch.Enabled = True
        CmdAdd.Enabled = True
Screen.MousePointer = Default
End Sub

Private Sub CmdSave_Click()
    Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SaveHLT
        Case 1
            Call SaveINS
        Case 2
        If EditFlag = True Then
        Call UpdatePlan
        Else
        Call SavePLN
            Frame2.Enabled = True
        End If
            
        Case 3
            Call SavePRM
        Case 4
            Call SaveAnnualLimit
        Case 5
            Call SaveMCOFees
        Case 6
            Call SaveSpecialNote
        Case 7
            Call SaveALB
        Case 8
            Call SaveProCat
        Case 9
            Call SaveMCOStatus
        Case 10
            Call SaveMgmtFee
    End Select
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub CmdExit_Click()
'Dim PLNSeq As String
 'PLNSeq = GeneratePLNSeq(Mid(TxtGrpCom, 1, 1))
    Unload Me
End Sub
'**************************End Add, Edit, Delete, Search, Save, Exit******************************

'**************************Start EntriesEnable******************************
Private Sub EntriesEnable(status As Boolean)
    Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Enabled = status
        End If
    Next ctl
End Sub
'**************************Start EntriesEnable******************************

' ############################## PRM ##############################

'List down all the premium list into the comboBox
Private Sub PRMComboListing()
Dim INS As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim HLT As New ADODB.Recordset
Dim AGE As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
    
  INS.Open "INS", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
  PAY.Open "PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
  HLT.Open "HLT", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
  AGE.Open "Age", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
  PLN.Open "PLN", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
  cboPmINSCode.Clear
  cboPmHLTCode.Clear
  cboPmPAYCode.Clear
  cboPmPLNCode.Clear
  cboPmAgeCode.Clear
  Combo3.Clear
  
  Do Until INS.EOF
     cboPmINSCode.AddItem INS!INSCode & " - " & INS!INSDescription
     INS.MoveNext
  Loop
    
  'Do Until PAY.EOF
  '   cboPmPAYCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
  '   PAY.MoveNext
  'Loop
  
    
  Do Until HLT.EOF
     cboPmHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
     HLT.MoveNext
  Loop
    
  'Do Until AGE.EOF
  '   cboPmAgeCode.AddItem AGE!AgeCode & " - " & AGE!AGEDescription
  '   AGE.MoveNext
  'Loop
    
  'Do Until PLN.EOF
  '   cboPmPLNCode.AddItem PLN!PLNCode & " - " & PLN!PLNDescription
  '   PLN.MoveNext
  'Loop
  
  Combo3.AddItem "A" & " - " & "Share Limit for Family"
  Combo3.AddItem "B" & " - " & "Share Limit for Supplementary"
  Combo3.AddItem "C" & " - " & "Separate Limit for Supplementary"
  Combo3.Text = "A" & " - " & "Share Limit for Family"
    
    INS.Close
    Set INS = Nothing
    PAY.Close
    Set PAY = Nothing
    HLT.Close
    Set HLT = Nothing
    AGE.Close
    Set AGE = Nothing
    PLN.Close
    Set PLN = Nothing
End Sub

'List down all the premium list into the listbox
Private Sub PRMListing(Optional strAppend As String)
    On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim prm As ListItem
    Dim sql As String
    Dim TotalPRMCount As Integer
    
    TotalPRMCount = 0
    lstPmPRMList.ListItems.Clear
        
    sql = "SELECT PRMCode, INSCode, " & _
      " PAYCode, HLTCode, PRMVersion, SuppPrmAmt, " & _
      " AGECode, PLNCode, PRMAmount , PRMEffDate, SuppPrmStatus " & _
      " From PRM where 0=0 "
      
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    
    Do Until rst2.EOF
            Set prm = lstPmPRMList.ListItems.Add(, , rst2!PRMCode)
                prm.SubItems(1) = rst2!INSCode
                prm.SubItems(2) = rst2!PAYCode
                prm.SubItems(3) = rst2!HLTCode
                prm.SubItems(4) = rst2!AGECode
                prm.SubItems(5) = rst2!PLNCode
                prm.SubItems(6) = Format(rst2!PRMAmount, "#,##0.00")
                prm.SubItems(7) = Format(rst2!SuppPrmAmt, "#,##0.00")
                prm.SubItems(8) = rst2!PRMEffDAte
                prm.SubItems(9) = rst2!PRMVersion
                prm.SubItems(10) = rst2!SuppPrmStatus

                rst2.MoveNext
           TotalPRMCount = TotalPRMCount + 1
           txtTotalRecord = TotalPRMCount
           Loop

    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

' To do the sorting
Private Sub lstPmPRMList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstPmPRMList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstPmPRMList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstPmPRMList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstPmPRMList, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstPmPRMList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    End Select
End Sub

Private Sub Text2_Change()
Dim AgeBand As String
    If Check1.Value = 1 Then
        If Text2 <> "" Then
            AgeBand = Text2
            Text3 = AgeBand
            Text4 = AgeBand
            Text5 = AgeBand
            Text6 = AgeBand
            Text7 = AgeBand
            Text1 = AgeBand
        ElseIf Text2 = "" Then
            Text3 = ""
            Text4 = ""
            Text5 = ""
            Text6 = ""
            Text7 = ""
            Text1 = ""
        End If
    End If
End Sub

Private Sub txt01_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt02_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt03_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt04_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt05_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt06_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txt07_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub



Private Sub txt01_Change()
Dim AgeBand As String
    If chkSTDPremium.Value = 1 Then
        If txt01 <> "" Then
            AgeBand = txt01
            txt02 = AgeBand
            txt03 = AgeBand
            txt04 = AgeBand
            txt05 = AgeBand
            txt06 = AgeBand
            txt07 = AgeBand
        ElseIf txt01 = "" Then
            txt01 = ""
            txt02 = ""
            txt03 = ""
            txt04 = ""
            txt05 = ""
            txt06 = ""
            txt07 = ""
        End If
    End If
End Sub


Private Sub txtALBCancerLmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtALBHomeNursLmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtALBKidDiaLmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtALEffDate_LostFocus()
    If IsDate(txtALEffDate) Then
    Else
        If txtALEffDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtALEffDate.SetFocus
        End If
    End If
End Sub

Private Sub txtCoPayEffdate_LostFocus()
    If IsDate(txtCoPayEffdate) Then
    Else
        If txtCoPayEffdate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtCoPayEffdate.SetFocus
        End If
    End If
End Sub

Private Sub txtEndAge_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtLifeTimeLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPRMEffdate_LostFocus()
    If IsDate(txtPRMEffdate) Then
    Else
        If txtPRMEffdate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPRMEffdate.SetFocus
        End If
    End If
End Sub

Private Sub TxtSpecGPeriod_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub txtStartAge_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSuppLifeTimeLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSupplyAL_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtAnnualLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub lstPMPRMList_Click()
Dim prm As New ADODB.Recordset
Dim strsql As String
Dim AGE01, AGE02, AGE03, AGE04, AGE05, AGE06, AGE07 As String
ListClick = True
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    If lstPmPRMList.ListItems.Count <> 0 Then
    
    Call setCmdButton(Me, False)
    'Call ClearForm(Me)
       
    txtPRMEffdate.Enabled = False
    FrameAGEBand.Enabled = False
    Combo3.Enabled = False
    txtPREVersion.Enabled = False
    
    strsql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
    prm.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until prm.EOF
        If prm!AGECode = "01" Then
            AGE01 = prm!PRMAmount
        ElseIf prm!AGECode = "02" Then
            AGE02 = prm!PRMAmount
        ElseIf prm!AGECode = "03" Then
            AGE03 = prm!PRMAmount
        ElseIf prm!AGECode = "04" Then
            AGE04 = prm!PRMAmount
        ElseIf prm!AGECode = "05" Then
            AGE05 = prm!PRMAmount
        ElseIf prm!AGECode = "06" Then
            AGE06 = prm!PRMAmount
        ElseIf prm!AGECode = "07" Then
            AGE07 = prm!PRMAmount
   
   End If
   prm.MoveNext
    Loop
   prm.Close
   Set prm = Nothing
     
                
    'if lstPmPRMList.SelectedItem.SubItems(1) =
    txtPrmCode = lstPmPRMList.SelectedItem.Text
    cboPmINSCode.Text = lstPmPRMList.SelectedItem.SubItems(1)
    cboPmINSCode.Enabled = False
    cboPmHLTCode.Text = lstPmPRMList.SelectedItem.SubItems(3)
    cboPmHLTCode.Enabled = False
    cboPmPAYCode.Text = lstPmPRMList.SelectedItem.SubItems(2)
    cboPmPAYCode.Enabled = False
    cboPmAgeCode.Text = lstPmPRMList.SelectedItem.SubItems(4)
    cboPmAgeCode.Enabled = False
    cboPmPLNCode.Text = lstPmPRMList.SelectedItem.SubItems(5)
    cboPmPLNCode.Enabled = False
    txt01 = Format(AGE01, "#,###0.00")
    txt02 = Format(AGE02, "#,###0.00")
    txt03 = Format(AGE03, "#,###0.00")
    txt04 = Format(AGE04, "#,###0.00")
    txt05 = Format(AGE05, "#,###0.00")
    txt06 = Format(AGE06, "#,###0.00")
    txt07 = Format(AGE07, "#,###0.00")
    FrameAGEBand.Enabled = False
    
    AGE01 = 0
    AGE02 = 0
    AGE03 = 0
    AGE04 = 0
    AGE05 = 0
    AGE06 = 0
    AGE07 = 0
    'txtPmPRMAmount.Text = Format(lstPmPRMList.SelectedItem.SubItems(6), "#,###0.00")
    'txtPmPRMAmount.Enabled = False
    
    strsql = "Select * from PRM where INSCode = '" & lstPmPRMList.SelectedItem.SubItems(1) & "' And PLNCode = '" & lstPmPRMList.SelectedItem.SubItems(5) & "'"
    prm.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until prm.EOF
        If prm!AGECode = "01" Then
            If Len(prm!SuppPrmAmt) Then
                AGE01 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "02" Then
            If Len(prm!SuppPrmAmt) Then
                AGE02 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "03" Then
            If Len(prm!SuppPrmAmt) Then
                AGE03 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "04" Then
            If Len(prm!SuppPrmAmt) Then
                AGE04 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "05" Then
            If Len(prm!SuppPrmAmt) Then
                AGE05 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "06" Then
            If Len(prm!SuppPrmAmt) Then
                AGE06 = prm!SuppPrmAmt
            End If
        ElseIf prm!AGECode = "07" Then
            If Len(prm!SuppPrmAmt) Then
                AGE07 = prm!SuppPrmAmt
            End If
   
   End If
   prm.MoveNext
    Loop
   prm.Close
   Set prm = Nothing
    
    Text2 = Format(AGE01, "#,###0.00")
    Text3 = Format(AGE02, "#,###0.00")
    Text4 = Format(AGE03, "#,###0.00")
    Text5 = Format(AGE04, "#,###0.00")
    Text6 = Format(AGE05, "#,###0.00")
    Text7 = Format(AGE06, "#,###0.00")
    Text1 = Format(AGE07, "#,###0.00")
    Frame6.Enabled = False
    
    txtPRMEffdate.mask = ""
    txtPRMEffdate = lstPmPRMList.SelectedItem.SubItems(8)
    txtPRMEffdate.mask = "##/##/####"
    txtPREVersion = lstPmPRMList.SelectedItem.SubItems(9)
    
    If lstPmPRMList.SelectedItem.SubItems(10) = "A" Then
        Combo3.ListIndex = 0
    ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "B" Then
        Combo3.ListIndex = 1
    ElseIf lstPmPRMList.SelectedItem.SubItems(10) = "C" Then
        Combo3.ListIndex = 2
    End If
    CmdAdd.Enabled = True
    End If
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
 '   medixmasterconnMaint.Close
 '   Set medixmasterconnMaint = Nothing
End Sub

    ' '''''''''''''''''' Save, edit , Add, delete PRM..... ''''''''''''''

Private Sub SavePRM()
Dim prm As New ADODB.Recordset
Dim SavePRM As New ADODB.Recordset
Dim RecordSaved
Dim strsql As String
Dim listloop As Integer
Dim ListCount As Integer

    If Len(Trim(cboPmINSCode)) = 0 Then
        MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPmINSCode.SetFocus
    ElseIf Len(Trim(cboPmHLTCode)) = 0 Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPmHLTCode.SetFocus
    'ElseIf Len(Trim(cboPmAgeCode)) = 0 Then
    '    MsgBox "Please Enter Age Code", vbOKOnly + vbCritical, DialogBoxTitle
    '    cboPmAgeCode.SetFocus
    ElseIf Len(Trim(cboPmPLNCode)) = 0 Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPmPLNCode.SetFocus
    ElseIf txtPRMEffdate = "__/__/____" Or Not IsDate(txtPRMEffdate) Then
        MsgBox "Please validate your Premium Effective date", vbOKOnly + vbCritical, DialogBoxTitle
        txtPRMEffdate.SetFocus
    ElseIf Trim(txtPREVersion) = "" Then
        MsgBox "Please Enter Premium Version", vbOKOnly + vbCritical, DialogBoxTitle
        txtPREVersion.SetFocus
    'ElseIf Mid(Combo3, 1, 1) = "B" Or Mid(Combo3, 1, 1) = "C" Then
    '    If Text2 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text2.SetFocus
    '    ElseIf Text3 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text3.SetFocus
    '    ElseIf Text4 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text4.SetFocus
    '    ElseIf Text5 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text5.SetFocus
    '    ElseIf Text6 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text6.SetFocus
    '    ElseIf Text7 = "" Then
    '        MsgBox "Please Enter Std Supp Premium ", vbOKOnly + vbCritical, DialogBoxTitle
    '        Text7.SetFocus
    '    'End If
    '
    'ElseIf Len(Trim(txtPmPRMAmount)) = 0 Then
    '    MsgBox "Please Enter Premium Amount", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtPmPRMAmount.SetFocus
    Else
        Dim locationPLN As Integer
        Dim locationPAY As Integer
        strsql = "Select * From PRM "
        locationPLN = InStr(1, cboPmPLNCode, "-") - 1
        If locationPLN = -1 Then
            locationPLN = 4
        End If
        
        locationPAY = InStr(1, cboPmPAYCode, "-") - 1
        If locationPAY = -1 Then
            locationPAY = 2
        End If
        strsql = strsql & " Where " & _
                " INSCode = '" & Trim(Mid(cboPmINSCode, 1, 2)) & "' " & _
                " AND HLTCode = '" & Trim(Mid(cboPmHLTCode, 1, 2)) & "' " & _
                " AND PLNCode = '" & Trim(Mid(cboPmPLNCode, 1, locationPLN)) & "' " & _
                " and AGECode = '" & Trim(Mid(cboPmAgeCode, 1, 2)) & "' " & _
                " and PRMeffdate = '" & Format(txt01, "mm/dd/yyyy") & "'"
                'and PRMeffdate = '" & Format(txtPRMEffdate, "mm/dd/yyyy") & "'"
'" PRMCode =" & txtPrmCode
        If ChkStr(Mid(cboPmHLTCode, 1, 2)) <> "S" Then
               strsql = strsql & " and (PAYCode = '" & Trim(Mid(cboPmPAYCode, 1, locationPAY)) & "') "
        End If
        
        prm.MaxRecords = 1
        prm.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            If prm.recordCount > 0 And EditFlag = False Then
               MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
               prm.Close
               cboPmINSCode.SetFocus
               Set prm = Nothing
               Exit Sub
            
            Else
            RecordSaved = MsgBox("Do you want to save this record?", vbYesNo + vbExclamation)
            Save = True
            If RecordSaved = vbYes Then
               
               If EditFlag = False Then
                  SavePRM.Open "PRM", medixmasterconn, adOpenKeyset, adLockOptimistic
                  ListCount = 7
                  For listloop = 1 To ListCount
                  SavePRM.AddNew
                    
                    
                    With SavePRM
                        !INSCode = ChkStr(Mid(cboPmINSCode, 1, 2))
                        !PAYCode = ChkStr(Mid(cboPmPAYCode, 1, 2))
                        !HLTCode = ChkStr(Mid(cboPmHLTCode, 1, 2))
                        !AGECode = ChkStr(Mid(cboPmAgeCode, 1, 2))
                        
                        If Mid(cboPmHLTCode, 1, 1) = "S" Then
                            !PLNCode = ChkStr(Mid(cboPmPLNCode, 1, 2))
                        Else
                            !PLNCode = ChkStr(Mid(cboPmPLNCode, 1, 4))
                        End If
                        
                        If Mid(Combo3, 1, 1) = "A" Then
                            !SuppPrmStatus = "A"
                            
                            If listloop = 1 Then
                                If txt01 <> "" Then
                                    !PRMAmount = txt01
                                    !AGECode = "01"
                                End If
                            ElseIf listloop = 2 Then
                                If txt02 <> "" Then
                                    !PRMAmount = txt02
                                    !AGECode = "02"
                                End If
                            ElseIf listloop = 3 Then
                                If txt03 <> "" Then
                                    !PRMAmount = txt03
                                    !AGECode = "03"
                                End If
                            ElseIf listloop = 4 Then
                                If txt04 <> "" Then
                                    !PRMAmount = txt04
                                    !AGECode = "04"
                                End If
                            ElseIf listloop = 5 Then
                                If txt05 <> "" Then
                                    !PRMAmount = txt05
                                    !AGECode = "05"
                                End If
                            ElseIf listloop = 6 Then
                                If txt06 <> "" Then
                                    !PRMAmount = txt06
                                    !AGECode = "06"
                                End If
                            ElseIf listloop = 7 Then
                                If txt07 <> "" Then
                                    !PRMAmount = txt07
                                    !AGECode = "07"
                                End If
                            End If
                            
                        ElseIf Mid(Combo3, 1, 1) = "B" Then
                            !SuppPrmStatus = "B"
                            
                            If listloop = 1 Then
                                If txt01 <> "" Then
                                    !PRMAmount = txt01
                                    If Text2 = "" Then
                                        !SuppPrmAmt = Text2
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "01"
                                End If
                            ElseIf listloop = 2 Then
                                If txt02 <> "" Then
                                    !PRMAmount = txt02
                                    If Text3 = "" Then
                                        !SuppPrmAmt = Text3
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "02"
                                End If
                            ElseIf listloop = 3 Then
                                If txt03 <> "" Then
                                    !PRMAmount = txt03
                                    If Text4 = "" Then
                                        !SuppPrmAmt = Text4
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "03"
                                End If
                            ElseIf listloop = 4 Then
                                If txt04 <> "" Then
                                    !PRMAmount = txt04
                                    If Text5 = "" Then
                                        !SuppPrmAmt = Text5
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "04"
                                End If
                            ElseIf listloop = 5 Then
                                If txt05 <> "" Then
                                    !PRMAmount = txt05
                                    If Text6 = "" Then
                                        !SuppPrmAmt = Text6
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "05"
                                End If
                            ElseIf listloop = 6 Then
                                If txt06 <> "" Then
                                    !PRMAmount = txt06
                                    If Text7 = "" Then
                                        !SuppPrmAmt = Text7
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "06"
                                End If
                            ElseIf listloop = 7 Then
                                If txt07 <> "" Then
                                    !PRMAmount = txt07
                                    If Text1 = "" Then
                                        !SuppPrmAmt = Text1
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "07"
                                End If
                            End If
                        
                        ElseIf Mid(Combo3, 1, 1) = "C" Then
                            !SuppPrmStatus = "C"
                            
                            If listloop = 1 Then
                                If txt01 <> "" Then
                                    !PRMAmount = txt01
                                    If Text2 = "" Then
                                        !SuppPrmAmt = Text2
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "01"
                                End If
                            ElseIf listloop = 2 Then
                                If txt02 <> "" Then
                                    !PRMAmount = txt02
                                    If Text3 = "" Then
                                        !SuppPrmAmt = Text3
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "02"
                                End If
                            ElseIf listloop = 3 Then
                                If txt03 <> "" Then
                                    !PRMAmount = txt03
                                    If Text4 = "" Then
                                        !SuppPrmAmt = Text4
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "03"
                                End If
                            ElseIf listloop = 4 Then
                                If txt04 <> "" Then
                                    !PRMAmount = txt04
                                    If Text5 = "" Then
                                        !SuppPrmAmt = Text5
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "04"
                                End If
                            ElseIf listloop = 5 Then
                                If txt05 <> "" Then
                                    !PRMAmount = txt05
                                    If Text6 = "" Then
                                        !SuppPrmAmt = Text6
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "05"
                                End If
                            ElseIf listloop = 6 Then
                                If txt06 <> "" Then
                                    !PRMAmount = txt06
                                    If Text7 = "" Then
                                        !SuppPrmAmt = Text7
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "06"
                                End If
                            ElseIf listloop = 7 Then
                                If txt07 <> "" Then
                                    !PRMAmount = txt07
                                    If Text1 = "" Then
                                        !SuppPrmAmt = Text1
                                    Else
                                        !SuppPrmAmt = "0"
                                    End If
                                    !AGECode = "07"
                                End If
                            End If
                        End If
                        
                        
                        !PRMEffDAte = CDate(txtPRMEffdate)
                        !PRMVersion = txtPREVersion
                        
                        SavePRM!PRMLastUpdateUser = Logon
                        SavePRM!PRMLastUpdateDate = Date
                    End With
                    SavePRM.Update
                 'End If
                 Next listloop

               Else
                  SavePRM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                  SavePRM.Update
               End If
                     Call ClearForm(Me)
                
                'Call PRMListing(PRMCriteria)
                
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                SavePRM.Close
                Set SavePRM = Nothing
            End If
            Save = False
            EditFlag = False
        End If
        End If
        'End If
        cboPmAgeCode.Enabled = True

  'End If
End Sub

Private Sub DeletePRM()
'On Error GoTo DeleleErr
Dim rsDeletePRM As New ADODB.Recordset
Dim rsCheckPRM As New ADODB.Recordset
Dim strsql As String
Dim sqlPRM As String
Dim DeleteRecord

EditFlag = False
    
 If cboPmINSCode = "" Then
    MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
    'cboPmINSCode.SetFocus
    ElseIf lstPmPRMList.ListItems.Count = 0 Then
       MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
   sqlPRM = "Select MBMNumber From MBM"
   
   sqlPRM = sqlPRM & " Where (INSCode = '" & Trim(cboPmINSCode) & "') AND (HLTCode = '" & Trim(cboPmHLTCode) & "')" & _
   " AND (PLNCode = '" & Trim(cboPmPLNCode) & "') and (AGECode = '" & Trim(cboPmAgeCode) & "') "
   If ChkStr(cboPmHLTCode) <> "S" Then
        sqlPRM = sqlPRM & " and (PAYCode = '" & Trim(cboPmPAYCode) & "') "
    End If
    
    
    rsCheckPRM.MaxRecords = 1
    rsCheckPRM.Open sqlPRM, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rsCheckPRM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Membership table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    
        If DeleteRecord = vbYes And rsCheckPRM.recordCount = 0 Then
             strsql = "Delete From PRM Where PRMCode = '" & ChkStr(lstPmPRMList.SelectedItem.Text) & "'"
             rsDeletePRM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             Call ClearForm(Me)
             Call PRMListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             CmdAdd.Enabled = True
             CmdSave.Enabled = True
        Else
             CmdDelete.Enabled = False
             CmdEdit.Enabled = False
        End If
    End If
    Exit Sub
  End If
End Sub


Private Sub SearchPRM()
Dim strAppend As String
Dim location As Integer
    
    Screen.MousePointer = vbHourglass
    
    strAppend = ""
    
    If Len(Trim(cboPmINSCode)) Then
        strAppend = strAppend + " And INSCode = '" & Trim(Mid(cboPmINSCode, 1, IIf(InStr(1, cboPmINSCode, "-") = 0, 4, InStr(1, cboPmINSCode, "-") - 1))) & "'"
        'strAppend = strAppend + " And INSCode = '" & Trim(Mid(cboPmINSCode, 1, 2)) & "'"
    End If
    If Len(Trim(cboPmPAYCode)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(cboPmPAYCode, 1, IIf(InStr(1, cboPmPAYCode, "-") = 0, 4, InStr(1, cboPmPAYCode, "-") - 1))) & "'"
        'strAppend = strAppend + " And PAYCode = '" & Trim(Mid(cboPmPAYCode, 1, 2)) & "'"
    End If
    If Len(Trim(cboPmHLTCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Trim(Mid(cboPmHLTCode, 1, IIf(InStr(1, cboPmHLTCode, "-") = 0, 4, InStr(1, cboPmHLTCode, "-") - 1))) & "'"
        'strAppend = strAppend + " And HLTCode = '" & Trim(Mid(cboPmHLTCode, 1, 2)) & "'"
    End If
    If Len(Trim(cboPmAgeCode)) Then
        strAppend = strAppend + " And AgeCode = '" & Trim(Mid(cboPmAgeCode, 1, IIf(InStr(1, cboPmAgeCode, "-") = 0, 4, InStr(1, cboPmAgeCode, "-") - 1))) & "'"
        'strAppend = strAppend + " And AgeCode = '" & Trim(Mid(cboPmAgeCode, 1, 2)) & "'"
    End If
    If Len(Trim(cboPmPLNCode)) Then
        location = InStr(1, cboPmPLNCode, "-") - 1
        If location = -1 Then
            location = 2
        End If
        strAppend = strAppend + " And PLNCOde = '" & Trim(Mid(cboPmPLNCode, 1, location)) & "'"
    End If
   
    PRMCriteria = strAppend
    Call PRMListing(strAppend)
    Screen.MousePointer = Default
End Sub

'Private Sub PrintPRM()
    'RptPrmList.show
'End Sub

    ' ''''''''''''''''  END Save, edit , Add, delete ..... ''''''''''''''

' '''''''''''''''''  to configure the drop down list

Private Sub cboPmINSCode_KeyPress(KeyAscii As Integer)
  ' If secondeditflag = False Then
     
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPmINSCode.Text, cboPmINSCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboPmINSCode.Text, cboPmINSCode.SelStart + cboPmINSCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPmINSCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboPmINSCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPmINSCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              cboPmINSCode.Text = ValidValue
              cboPmINSCode.SelStart = 0
              cboPmINSCode.SelLength = Len(ValidValue)
            End If
   End If
'End If
End Sub

'Private Sub cboPmHLTCode_Click()
'Dim PAY As New ADODB.Recordset
'Dim INS As New ADODB.Recordset
'Dim strsql As String
'Dim strsql2 As String

'cboPmPAYCode.Clear
'cboPmINSCode.Clear



    'If Mid(cboPmHLTCode, 1, 1) = "N" Then
       'strsql = "Select  Distinct INSCode, INSDescription From View_INS "
       'strsql = strsql + " Where HLTCode ='" & Mid(cboPmHLTCode, 1, 1) & "'"
       'strsql2 = " Select * from PAY Where PAYCode Like 'W%'"
    'Else
       'strsql = "Select INSCode, INSDescription From INS "
       'strsql2 = "Select  PAYCode, PAYCompanyName From PAY "
       'strsql2 = strsql2 + " Where PAYCode Not Like 'W%'"
    'End If
    
    'INS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'cboPmINSCode.Clear
    'Do Until INS.EOF
       'cboPmINSCode.AddItem ChkStr(INS!INSCode) & " - " & ChkStr(INS!INSDescription)
       'INS.MoveNext
    'Loop
    'INS.Close
    'Set INS = Nothing
    
    'PAY.Open strsql2, medixmasterconn, adOpenKeyset, adLockOptimistic
    'cboPmPAYCode.Clear
    'Do Until PAY.EOF
       'cboPmPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       'PAY.MoveNext
    'Loop
    
    'PAY.Close
    'Set PAY = Nothing
  
'End Sub

Private Sub cboPmHLTCode_KeyPress(KeyAscii As Integer)
  ' If secondeditflag = False Then
     
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPmHLTCode.Text, cboPmHLTCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboPmHLTCode.Text, cboPmHLTCode.SelStart + cboPmHLTCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPmHLTCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboPmHLTCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPmHLTCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              cboPmHLTCode.Text = ValidValue
              cboPmHLTCode.SelStart = 0
              cboPmHLTCode.SelLength = Len(ValidValue)
            End If
   End If
'End If
End Sub

Private Sub cboPmPLNCode_KeyPress(KeyAscii As Integer)
   ' If secondeditflag = False Then
    
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPmPLNCode.Text, cboPmPLNCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboPmPLNCode.Text, cboPmPLNCode.SelStart + cboPmPLNCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPmPLNCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboPmPLNCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPmPLNCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              cboPmPLNCode.Text = ValidValue
              cboPmPLNCode.SelStart = 0
              cboPmPLNCode.SelLength = Len(ValidValue)
            End If
   End If
'End If
End Sub


Private Sub cboPmPAYCode_KeyPress(KeyAscii As Integer)
  '  If secondeditflag = False Then
    
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPmPAYCode.Text, cboPmPAYCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboPmPAYCode.Text, cboPmPAYCode.SelStart + cboPmPAYCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPmPAYCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboPmPAYCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPmPAYCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              cboPmPAYCode.Text = ValidValue
              cboPmPAYCode.SelStart = 0
              cboPmPAYCode.SelLength = Len(ValidValue)
            End If
   End If
'End If
End Sub

Private Sub cboPmAgeCode_KeyPress(KeyAscii As Integer)
   ' If secondeditflag = False Then
    
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPmAgeCode.Text, cboPmAgeCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboPmAgeCode.Text, cboPmAgeCode.SelStart + cboPmAgeCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPmAgeCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboPmAgeCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPmAgeCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
              cboPmAgeCode.Text = ValidValue
              cboPmAgeCode.SelStart = 0
              cboPmAgeCode.SelLength = Len(ValidValue)
            End If
   End If
'End If
End Sub

' ''''''' End ''''  to configure the drop down list
' ############################## END PRM ##############################



' ############################## Start PLN (Plan)  ##############################

Private Sub PLNComboListing()
    Dim PAY As New ADODB.Recordset
    Dim HLT As New ADODB.Recordset
    Dim INS As New ADODB.Recordset
    Dim ProCat As New ADODB.Recordset
    Dim PolCat As New ADODB.Recordset
    
    PolCat.Open "PolicyCategory", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    PAY.Open "select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    HLT.Open "select  HLTCOde  , HLTDescription from HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    ProCat.Open "Select ProductCode, ProductName From ProductCategory", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    INS.Open "INS", medixmasterconn, adOpenKeyset, adLockOptimistic
    CboInsType.Clear
    cboPLNHLTCode.Clear
    cboPLNPayor.Clear
    cboProCat.Clear
    
    Do Until HLT.EOF
       cboPLNHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
       HLT.MoveNext
    Loop
    
    Do Until INS.EOF
       CboInsType.AddItem INS!INSCode & " - " & INS!INSDescription
       INS.MoveNext
    Loop
  
    'Do Until PAY.EOF
    '   cboPLNPayor.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
    '   PAY.MoveNext
    'Loop
    
    Do Until ProCat.EOF
        cboProCat.AddItem ProCat!ProductCode & " - " & ProCat!ProductName
        ProCat.MoveNext
    Loop
    
    Do Until PolCat.EOF
        cboPolicyWording.AddItem PolCat!PolicyWording
        PolCat.MoveNext
    Loop

    CboTopupStatus.AddItem "N - No"
    CboTopupStatus.AddItem "Y - Yes"
    
    'CboLGInd.AddItem "C - Co-Payment"
    'CboLGInd.AddItem "S - Single Beded"
    'CboLGInd.AddItem "D - Double Beded"
    
    CboCoPayment.AddItem "N - No"
    CboCoPayment.AddItem "Y - Yes"
    
    cboPLNAnnLimit.AddItem "Yes"
    cboPLNAnnLimit.AddItem "No"
    
    cboPLNPremium.AddItem "Yes"
    cboPLNPremium.AddItem "No"

    cboPLNMCO.AddItem "Yes"
    cboPLNMCO.AddItem "No"
    
    CboSpecGPeriod.AddItem "Y - Yes"
    CboSpecGPeriod.AddItem "N - No"
    
    CboSMPlan.AddItem "Y - Yes"
    CboSMPlan.AddItem "N - No"
    
    CboLifeTimeStatus.AddItem "Y - Yes"
    CboLifeTimeStatus.AddItem "N - No"
    
    CboSOF.AddItem "N - No"
    CboSOF.AddItem "Y - Yes"
    
    CboMeal.AddItem "N - No"
    CboMeal.AddItem "Y - Yes"
    
    CboNursing.AddItem "N - No"
    CboNursing.AddItem "Y - Yes"
    
    CboTax.AddItem "N - No"
    CboTax.AddItem "Y - Yes"
    
    CboMRI.AddItem "N - No"
    CboMRI.AddItem "Y - Yes"
    
    CboNoofPlan.AddItem "1"
    CboNoofPlan.AddItem "2"
    CboNoofPlan.AddItem "3"
    CboNoofPlan.AddItem "4"
    CboNoofPlan.AddItem "5"
    CboNoofPlan.AddItem "6"
    
    CboPLNMgmtStatus.AddItem "N - No"
    CboPLNMgmtStatus.AddItem "Y - Yes"
    
    cboDisIndicator.AddItem "N - No"
    cboDisIndicator.AddItem "Y - Yes"
        
    INS.Close
    Set INS = Nothing
    HLT.Close
    Set HLT = Nothing
    PAY.Close
    Set PAY = Nothing
    PolCat.Close
    Set PolCat = Nothing
End Sub

Private Sub PLNListing(Optional strAppend As String)
    Dim rsListPlan As New ADODB.Recordset
    Dim TotalPLNCount As Integer
    Dim PLNMaster As ListItem
    Dim SqlPln As String
    
    TotalPLNCount = 0
    lstPlanList.ListItems.Clear
    SqlPln = "select * From PLN where 0=0 "
    
    If Len(Trim(strAppend)) Then
        SqlPln = SqlPln + strAppend
    End If
    
    rsListPlan.Open SqlPln, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If rsListPlan.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
    Do Until rsListPlan.EOF
         Set PLNMaster = lstPlanList.ListItems.Add(, , rsListPlan!HLTCode)
             PLNMaster.SubItems(1) = rsListPlan!PLNCode
             
             If rsListPlan!NewPLNCode <> "" Then
                PLNMaster.SubItems(2) = rsListPlan!NewPLNCode
             End If
             
             If rsListPlan!PLNDescription <> "" Then
                PLNMaster.SubItems(3) = rsListPlan!PLNDescription
             End If
             
             PLNMaster.SubItems(4) = IIf(Trim(rsListPlan!PAYCode) <> "", Trim(rsListPlan!PAYCode), "")
             PLNMaster.SubItems(5) = IIf(Trim(rsListPlan!Grpcompany) <> "", rsListPlan!Grpcompany, "")
             PLNMaster.SubItems(6) = IIf(Trim(rsListPlan!PLNTopUpStatus) <> "", rsListPlan!PLNTopUpStatus, "")
             'PLNMaster.SubItems(7) = IIf(Trim(rsListPlan!PLNIndex) <> "", rsListPlan!PLNIndex, "")
             PLNMaster.SubItems(8) = IIf(Trim(rsListPlan!AnnLmtInd) <> "", rsListPlan!AnnLmtInd, "")
             PLNMaster.SubItems(9) = IIf(Trim(rsListPlan!PRMInd) <> "", rsListPlan!PRMInd, "")
             PLNMaster.SubItems(10) = IIf(Trim(rsListPlan!MCOInd) <> "", rsListPlan!MCOInd, "")
             PLNMaster.SubItems(11) = IIf(Trim(rsListPlan!CoPaymentInd) <> "", rsListPlan!CoPaymentInd, "")
             PLNMaster.SubItems(12) = IIf(Trim(rsListPlan!PLNEffDate) <> "", rsListPlan!PLNEffDate, "")
             PLNMaster.SubItems(13) = IIf(Trim(rsListPlan!SpeGPeriodStatus) <> "", rsListPlan!SpeGPeriodStatus, "")
             PLNMaster.SubItems(14) = IIf(Trim(rsListPlan!SpeGPeriodDays) <> "", rsListPlan!SpeGPeriodDays, "")
             PLNMaster.SubItems(15) = IIf(Trim(rsListPlan!NewSMPlan) <> "", rsListPlan!NewSMPlan, "")
             PLNMaster.SubItems(16) = IIf(Trim(rsListPlan!PLNLifeTimeStatus) <> "", rsListPlan!PLNLifeTimeStatus, "")
             PLNMaster.SubItems(17) = IIf(Trim(rsListPlan!PLNSOF) <> "", rsListPlan!PLNSOF, "")
             PLNMaster.SubItems(18) = IIf(Trim(rsListPlan!PLNProCat) <> "", rsListPlan!PLNProCat, "")
             PLNMaster.SubItems(19) = IIf(Trim(rsListPlan!PLNMeal) <> "", rsListPlan!PLNMeal, "")
             PLNMaster.SubItems(20) = IIf(Trim(rsListPlan!PLNNursing) <> "", rsListPlan!PLNNursing, "")
             PLNMaster.SubItems(21) = IIf(Trim(rsListPlan!PLNTax) <> "", rsListPlan!PLNTax, "")
             PLNMaster.SubItems(22) = IIf(Trim(rsListPlan!PLNMRI) <> "", rsListPlan!PLNMRI, "")
             PLNMaster.SubItems(23) = IIf(Trim(rsListPlan!PLNMgmtFee) <> "", rsListPlan!PLNMgmtFee, "")
             PLNMaster.SubItems(24) = IIf(Trim(rsListPlan!PLNDisIndicator) <> "", rsListPlan!PLNDisIndicator, "")
             PLNMaster.SubItems(25) = IIf(Trim(rsListPlan!PLNPolicyWording) <> "", rsListPlan!PLNPolicyWording, "")
             PLNMaster.SubItems(26) = IIf(Trim(rsListPlan!PLNPayorPlan) <> "", rsListPlan!PLNPayorPlan, "")
             PLNMaster.SubItems(27) = IIf(Trim(rsListPlan!PLNPolNo) <> "", rsListPlan!PLNPolNo, "")
             PLNMaster.SubItems(28) = IIf(IsNumeric(rsListPlan!PLNLGPrivate), rsListPlan!PLNLGPrivate, "")
             If rsListPlan!VIPInd = "Y" Then
                PLNMaster.SubItems(29) = "Y-Yes"
             ElseIf rsListPlan!VIPInd = "N" Then
                PLNMaster.SubItems(29) = "N-No"
             End If
             PLNMaster.SubItems(30) = IIf(IsNumeric(rsListPlan!PLNCardName), rsListPlan!PLNCardName, "")
             rsListPlan.MoveNext
         TotalPLNCount = TotalPLNCount + 1
         txtTotalRecord = TotalPLNCount
    Loop
    rsListPlan.Close
    Set rsListPlan = Nothing
    End If
End Sub

Private Sub lstPlanList_Click()
    If lstPlanList.ListItems.Count <> 0 Then
        Call lstPlanListShow
    End If
End Sub

Private Sub lstPlanList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstPlanList.SortKey = ColumnHeader.Index - 1
    lstPlanList.Sorted = True
    If lstPlanList.SortOrder = lvwAscending Then
       lstPlanList.SortOrder = lvwDescending
    Else
       lstPlanList.SortOrder = lvwAscending
    End If
 
End Sub

    ' '''''''''''''''''' Save, edit , Add, delete PLN..... ''''''''''''''
Private Sub DeletePLN()
    'On Error GoTo DeleleErr
    
    Dim rsDeletePLN As New ADODB.Recordset
    Dim rsCheckMBM As New ADODB.Recordset
    Dim strsql As String
    Dim sqlMBM As String
    Dim DeleteRecord
    EditFlag = False
    
    
    'If TxtPLNCode = "" Then
    '   MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
      ' TxtPLNCode.SetFocus
    If lstPlanList.ListItems.Count = 0 Then
       MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    sqlMBM = "Select MBMNumber From MBM where PLNCode = '" & RemStr(lstPlanList.SelectedItem.SubItems(2)) & "'"
    
    rsCheckMBM.MaxRecords = 1
    rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Membership table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    
        If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
             strsql = "Delete From PLN Where PLNCode = '" & ChkStr(lstPlanList.SelectedItem.SubItems(1)) & "'"
             rsDeletePLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             Call ClearForm(Me)
             Call PLNListing(PLNCriteria)
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             CmdAdd.Enabled = True
             CmdSave.Enabled = True
        Else
             CmdDelete.Enabled = False
             CmdEdit.Enabled = False
        End If
    End If
    Exit Sub
  End If
End Sub
    
Private Sub SearchPLN()
    Dim strAppend As String
        
    Screen.MousePointer = vbHourglass
    
    strAppend = ""
        
    If Len(Trim(cboPLNHLTCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Mid(cboPLNHLTCode, 1, 1) & "'"
    End If
    
    If Len(Trim(cboPLNPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Mid(cboPLNPayor, 1, 2) & "'"
    End If
    
    If Len(Trim(TxtGrpCom)) Then
        strAppend = strAppend + " And GRPCompany like '%" & ChkStr(TxtGrpCom) & "%'"
    End If
    
    If Len(Trim(CboTopupStatus)) Then
        strAppend = strAppend + " And PLNTopUPStatus = '" & Trim(Mid(CboTopupStatus, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboCoPayment)) Then
        strAppend = strAppend + " And CoPaymentInd = '" & Trim(Mid(CboCoPayment, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboSOF)) Then
        strAppend = strAppend + " And PLNSOF = '" & Trim(Mid(CboSOF, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboSpecGPeriod)) Then
        strAppend = strAppend + " And SpeGPeriodStatus = '" & Trim(Mid(CboSpecGPeriod, 1, 1)) & "'"
    End If
    
    If Len(Trim(cboProCat)) Then
        strAppend = strAppend + " AND PLNProCat = '" & Trim(Mid(cboProCat, 1, 2)) & "'"
    End If
    
    If Len(Trim(txtPLN1(0))) Then
        strAppend = strAppend + " And PLNCode like '%" & ChkStr(txtPLN1(0)) & "%'"
    End If
        
    If Len(Trim(CboMeal)) Then
        strAppend = strAppend + " And PLNMeal = '" & Trim(Mid(CboMeal, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboNursing)) Then
        strAppend = strAppend + " And PLNNursing = '" & Trim(Mid(CboNursing, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboTax)) Then
        strAppend = strAppend + " And PLNTax = '" & Trim(Mid(CboTax, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboMRI)) Then
        strAppend = strAppend + " And PLNMRI = '" & Trim(Mid(CboMRI, 1, 1)) & "'"
    End If
    
    'If Len(Trim(CboLGInd)) Then
        'strAppend = strAppend + " And LGInd = '" & Trim(Mid(CboLGInd, 1, 1)) & "'"
    'End If
    
    If Len(Trim(cboDisIndicator)) Then
        strAppend = strAppend + " And PLNDisIndicator = '" & Trim(Mid(cboDisIndicator, 1, 1)) & "'"
    End If
    
    PLNCriteria = strAppend
    Call PLNListing(strAppend)
    Screen.MousePointer = Default
    
End Sub

Private Sub SavePLN()
Dim RecordSaved
Dim SavePLN As New ADODB.Recordset
Dim ListCount As Integer
Dim listloop As Integer
Dim PLNSeq As String

    If Trim(cboPLNHLTCode) = "" Then
        MsgBox "Please Enter Health Type ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf IsNull(txtLGPrivate) Then
        MsgBox "Please Enter LG Number ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(cboPLNHLTCode, 1, 1)) = "N" And Trim(cboPLNPayor) = "" Then
        MsgBox "Please Enter Payor Code ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf IsNull(cmbVIPInd) Then
        MsgBox "Please Enter VIP Indicator ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(cboPLNHLTCode, 1, 1)) = "N" And Trim(CboNoofPlan) = "" Then
        MsgBox "Please Enter No of Plan Code ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Mid(cboPLNAnnLimit, 1, 1) = "" Then
        MsgBox "Please Enter Annual Limit Ind.", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Mid(CboLifeTimeStatus, 1, 1) = "" Then
        MsgBox "Please Enter Life Limit Ind.", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(TxtGrpCom) = "" Then
        MsgBox "Please Enter Group Company", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Mid(cboPLNPremium, 1, 1) = "" Then
        MsgBox "Please Enter Premium Ind.", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Mid(cboPLNMCO, 1, 1) = "" Then
        MsgBox "Please Enter MCO Ind.", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(CboTopupStatus, 1, 1)) = "" Then
        MsgBox "Please Enter Topup Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(CboSpecGPeriod, 1, 1)) = "" Then
        MsgBox "Please Enter Special GPeriod Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(CboSpecGPeriod, 1, 1)) = "Y" And TxtSpecGPeriod = "" Then
        MsgBox "Please Enter Special GPeriod Days", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(CboCoPayment, 1, 1)) = "" Then
        MsgBox "Please Enter Co-Payment Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtCoPayEffdate = "__/__/____" Then
        MsgBox "Please Enter Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(cboPLNHLTCode, 1, 1)) = "S" And Trim(txtPLN1(0)) = "" Then
        MsgBox "Please Enter Plan Code ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboProCat = "" Then
        MsgBox "Please Select Product Category", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(CboMeal, 1, 1)) = "" Then
        MsgBox "Please Enter Meals Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(CboNursing, 1, 1)) = "" Then
        MsgBox "Please Enter Nursing Charges Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(CboTax, 1, 1)) = "" Then
        MsgBox "Please Enter Tax Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(CboMRI, 1, 1)) = "" Then
        MsgBox "Please Enter MRI Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(CboPLNMgmtStatus, 1, 1)) = "" Then
        MsgBox "Please Enter Mgmt. Fee Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Mid(cboDisIndicator, 1, 1)) = "" Then
        MsgBox "Please Select a Dis. Indicator", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(txtStartAge) = "" Then
        MsgBox "Please Enter Policy Start Age", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(txtEndAge) = "" Then
        MsgBox "Please Enter Policy End Age", vbOKOnly + vbCritical, DialogBoxTitle
    'To Change the Policy Wording Combobox to textbox - By CHEng 4 Dec 2016
    'ElseIf Trim(cboPolicyWording) = "" Then
    ElseIf Trim(TxtPolicyWording.Text) = "" Then
        MsgBox "Please Enter Policy Category", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Mid(cboPLNPayor, 1, 2) = "AX" And cboClientPLN = "" Then
        MsgBox "Please Select Client Plan", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(txtFileName) = "" Then
        MsgBox "Please Select E-Card type", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Mid(cboPLNPayor, 1, 2) = "AX" And txtClientPolNo = "" Then
        MsgBox "Please Enter Client Policy Number", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Save = True
            If Trim(Mid(cboPLNHLTCode, 1, 1)) = "N" Then
                ListCount = Trim(CboNoofPlan) - 1
            Else
                ListCount = txtPLN1.UBound
            End If
            For listloop = 0 To ListCount
                If (Trim(Mid(cboPLNHLTCode, 1, 1)) = "S" And Trim(txtPLN1(listloop)) <> "") Or Trim(Mid(cboPLNHLTCode, 1, 1)) = "N" Then
                    SavePLN.Open "PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
                    SavePLN.AddNew
                    If Trim(Mid(cboPLNHLTCode, 1, 1)) = "N" Then
                        PLNSeq = GeneratePLNSeq(Mid(TxtGrpCom, 1, 1))
                    Else
                        PLNSeq = txtPLN1(listloop)
                    End If
                    With SavePLN
                        !PLNCode = ChkStr(PLNSeq)
                        !PLNDescription = ChkStr(txtDesc1(listloop))
                        !HLTCode = Trim(Mid(cboPLNHLTCode, 1, IIf(InStr(1, cboPLNHLTCode, "-") = 0, 1, InStr(1, cboPLNHLTCode, "-") - 1)))
                        !PAYCode = Trim(Mid(cboPLNPayor, 1, IIf(InStr(1, cboPLNPayor, "-") = 0, 4, InStr(1, cboPLNPayor, "-") - 1)))
                        !Grpcompany = ChkStr(TxtGrpCom)
                        !PLNTopUpStatus = Mid(CboTopupStatus, 1, 1)
                        !AnnLmtInd = Trim(Mid(cboPLNAnnLimit, 1, 1))
                        !PRMInd = Trim(Mid(cboPLNPremium, 1, 1))
                        !MCOInd = Trim(Mid(cboPLNMCO, 1, 1))
                        !CoPaymentInd = Mid(CboCoPayment, 1, 1)
                        !PLNMeal = Trim(Mid(CboMeal, 1, 1))
                        !PLNNursing = Trim(Mid(CboNursing, 1, 1))
                        !PLNTax = Trim(Mid(CboTax, 1, 1))
                        !PLNMRI = Trim(Mid(CboMRI, 1, 1))
                        !PLNSOF = Trim(Mid(CboSOF, 1, 1))
                        !NewSMPlan = Trim(Mid(CboSMPlan, 1, 1))
                        !SpeGPeriodStatus = Trim(Mid(CboSpecGPeriod, 1, 1))
                        !PLNLifeTimeStatus = Trim(Mid(CboLifeTimeStatus, 1, 1))
                        !PLNProCat = Trim(Mid(cboProCat, 1, 2))
                        If Len(TxtSpecGPeriod) Then
                            !SpeGPeriodDays = TxtSpecGPeriod
                        End If
                        If txtCoPayEffdate <> "__/__/____" Then
                            !PLNEffDate = CDate(txtCoPayEffdate)
                        End If
                        !PLNMgmtFee = Mid(CboPLNMgmtStatus, 1, 1)
                        !PLNLastUpdateUser = Logon
                        !PLNLastUpdateDate = Date
                        !PLNDisIndicator = Mid(cboDisIndicator, 1, 1)
                        !PLNStartAge = txtStartAge
                        !PLNEndAge = txtEndAge
                        '!PLNPolicyWording = cboPolicyWording
                        !PLNPolicyWording = TxtPolicyWording.Text
                        'Change PolicyWording combobox to textbox - By CHEng 2 Dec 2016
                        If Len(cboClientPLN) Then
                            !PLNPayorPlan = cboClientPLN
                        End If
                        
                        If Len(txtClientPolNo) Then
                            !PLNPolNo = txtClientPolNo
                        End If
                        
                        If Len(txtCopayPerc) Then
                            !PLNCoPayPerc = txtCopayPerc
                        End If
                        
                        If Mid(cboPLNPayor, 1, 2) = "ET" Then
                        
                            !PLNLGGov = "5"
                            !PLNLGPrivate = "5"
                            !PLNExcessPrivate = "2"
                            !PLNCoPayPerc = "20"
                            !CoPaymentInd = "Y"
                        End If
                        
                        If Mid(cboPLNPayor, 1, 2) = "EC" Then
                            !PLNLGGov = "5"
                            !PLNLGPrivate = "5"
                            !PLNExcessPrivate = "2"
                            !PLNCoPayPerc = "20"
                            !CoPaymentInd = "Y"
                        End If
                        If Mid(cboPLNPayor, 1, 2) = "TK" Then
                            !PLNLGPrivate = "9"
                        Else
                            !PLNLGPrivate = IIf(IsNumeric(txtLGPrivate), txtLGPrivate, 0)
                        End If
                        If Len(cmbVIPInd) Then
                        !VIPInd = Mid(cmbVIPInd, 1, 1)
                            If Mid(cmbVIPInd, 1, 1) = "Y" Then
                                If Mid(cboPLNPayor, 1, 2) = "ST" Then
                                    !PLNSpecialNote = "ST-VVIP Member. Full Cover including exclusion"
                                Else
                                    !PLNSpecialNote = "VIP Member. Full Cover including exclusion"
                                End If
                            End If
                        End If
                        !PLNCardName = IIf(Len(txtFileName), txtFileName, "")
                        SavePLN.Update
                    End With
                    SavePLN.Close
                    Set SavePLN = Nothing
                End If
            Next listloop
        End If
        Call ClearForm(Me)
        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
    End If
    Save = False
End Sub

Private Sub cboPLNHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNHLTCode.Text, cboPLNHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboPLNHLTCode.Text, cboPLNHLTCode.SelStart + cboPLNHLTCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNHLTCode.ListCount - 1
 
        If NewText = LCase(Left(cboPLNHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNHLTCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPLNHLTCode.Text = ValidValue
               cboPLNHLTCode.SelStart = 0
               cboPLNHLTCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboPLNPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboPLNPayor.Text, cboPLNPayor.SelStart) & _
                Chr(KeyAscii) + Mid(cboPLNPayor.Text, cboPLNPayor.SelStart + cboPLNPayor.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPLNPayor.ListCount - 1
            
            If NewText = ChkStr(Left(cboPLNPayor.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPLNPayor.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboPLNPayor.Text = ValidValue
              cboPLNPayor.SelStart = 0
              cboPLNPayor.SelLength = Len(ValidValue)
              End If
            
'End If

End Sub

Private Sub CboTopupStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(CboTopupStatus.Text, CboTopupStatus.SelStart) & _
                Chr(KeyAscii) + Mid(CboTopupStatus.Text, CboTopupStatus.SelStart + CboTopupStatus.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboTopupStatus.ListCount - 1
            
            If NewText = ChkStr(Left(CboTopupStatus.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboTopupStatus.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              CboTopupStatus.Text = ValidValue
              CboTopupStatus.SelStart = 0
              CboTopupStatus.SelLength = Len(ValidValue)
            End If
            
'End If

End Sub

Private Sub CboSOF_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSOF.Text, CboSOF.SelStart) + Chr(KeyAscii) + Mid(CboSOF.Text, CboSOF.SelStart + CboSOF.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSOF.ListCount - 1
 
        If NewText = LCase(Left(CboSOF.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSOF.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSOF.Text = ValidValue
               CboSOF.SelStart = 0
               CboSOF.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboCoPayment_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(CboCoPayment.Text, CboCoPayment.SelStart) & _
                Chr(KeyAscii) + Mid(CboCoPayment.Text, CboCoPayment.SelStart + CboCoPayment.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboCoPayment.ListCount - 1
            
            If NewText = ChkStr(Left(CboCoPayment.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboCoPayment.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              CboCoPayment.Text = ValidValue
              CboCoPayment.SelStart = 0
              CboCoPayment.SelLength = Len(ValidValue)
            End If

End Sub

Private Sub CboSMPlan_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(CboSMPlan.Text, CboSMPlan.SelStart) & _
                Chr(KeyAscii) + Mid(CboSMPlan.Text, CboSMPlan.SelStart + CboSMPlan.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboSMPlan.ListCount - 1
            
            If NewText = ChkStr(Left(CboSMPlan.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboSMPlan.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              CboSMPlan.Text = ValidValue
              CboSMPlan.SelStart = 0
              CboSMPlan.SelLength = Len(ValidValue)
            End If

End Sub

'Private Sub CboLGInd_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
'
'    If KeyAscii > 32 And KeyAscii < 127 Then
'        NewText = ChkStr(Left(CboLGInd.Text, CboLGInd.SelStart) & _
'                Chr(KeyAscii) + Mid(CboLGInd.Text, CboLGInd.SelStart + CboLGInd.SelLength + 1))
'
'        ValidCount = 0
'        ValidValue = ""
'
'        For i = 0 To CboLGInd.ListCount - 1
'
'            If NewText = ChkStr(Left(CboLGInd.List(i), Len(NewText))) Then
'                    ValidCount = ValidCount + 1
'                    ValidValue = CboLGInd.List(i)
'            End If
'        Next
'            If ValidCount <= 1 Then KeyAscii = 0
'            End If
'            If ValidCount = 1 Then
'              CboLGInd.Text = ValidValue
'              CboLGInd.SelStart = 0
'              CboLGInd.SelLength = Len(ValidValue)
'            End If'
'End Sub

Private Sub CboMeal_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(CboMeal.Text, CboMeal.SelStart) & _
                Chr(KeyAscii) + Mid(CboMeal.Text, CboMeal.SelStart + CboMeal.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboMeal.ListCount - 1
            
            If NewText = ChkStr(Left(CboMeal.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboMeal.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              CboMeal.Text = ValidValue
              CboMeal.SelStart = 0
              CboMeal.SelLength = Len(ValidValue)
            End If

End Sub

Private Sub CboNursing_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(CboNursing.Text, CboNursing.SelStart) & _
                Chr(KeyAscii) + Mid(CboNursing.Text, CboNursing.SelStart + CboNursing.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboNursing.ListCount - 1
            
            If NewText = ChkStr(Left(CboNursing.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboNursing.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              CboNursing.Text = ValidValue
              CboNursing.SelStart = 0
              CboNursing.SelLength = Len(ValidValue)
            End If

End Sub

Private Sub CboTax_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(CboTax.Text, CboTax.SelStart) & _
                Chr(KeyAscii) + Mid(CboTax.Text, CboTax.SelStart + CboTax.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboTax.ListCount - 1
            
            If NewText = ChkStr(Left(CboTax.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboTax.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              CboTax.Text = ValidValue
              CboTax.SelStart = 0
              CboTax.SelLength = Len(ValidValue)
            End If

End Sub

Private Sub CboMRI_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(CboMRI.Text, CboMRI.SelStart) & _
                Chr(KeyAscii) + Mid(CboMRI.Text, CboMRI.SelStart + CboMRI.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboMRI.ListCount - 1
            
            If NewText = ChkStr(Left(CboMRI.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboMRI.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              CboMRI.Text = ValidValue
              CboMRI.SelStart = 0
              CboMRI.SelLength = Len(ValidValue)
            End If

End Sub


' ############################## End PLN (Plan)  ##############################

' ############################## Start HLT (Health)  ##############################
'Private Sub HLTcombolisting()
'Dim HLT As New ADODB.Recordset
'Dim sqlHLT As String

'cboHLTCode.Clear
' sqlHLT = "select HLTcode,HLTdescription from HLT"
 
'HLT.Open sqlHLT, medixmasterconn, adOpenKeyset, adLockOptimistic

'Do Until HLT.EOF
'With HLT
 '   cboHLTCode.AddItem !HLTCode & "-" & !HLTDescription
'HLT.MoveNext
'End With
'Loop

'HLT.Close

'End Sub

Private Sub HLTListing(Optional strAppend As String)
    Dim rslistHLT As New ADODB.Recordset
    Dim rsListHLT2 As New ADODB.Recordset
    Dim HLTMaster As ListItem
    Dim sqlHLT As String
    Dim TotalHLTCount As Integer
         
    
    TotalHLTCount = 0
    lstHLTList.ListItems.Clear
    sqlHLT = "select HLTCode, HLTDescription From HLT where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sqlHLT = sqlHLT + strAppend
    End If
    
    rslistHLT.Open sqlHLT, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rslistHLT.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    
    Do Until rslistHLT.EOF
        With rslistHLT
            Set HLTMaster = lstHLTList.ListItems.Add(, , !HLTCode)
                HLTMaster.SubItems(1) = !HLTDescription
            .MoveNext
        End With
        TotalHLTCount = TotalHLTCount + 1
        txtTotalRecord = TotalHLTCount
    Loop
    End If
    rslistHLT.Close
    Set rslistHLT = Nothing
        
End Sub


Private Sub lstHLTList_Click()
Dim strsql As String
    
    If lstHLTList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtHLTcode = lstHLTList.SelectedItem.Text
        txtHLTcode.Enabled = False
        txtHLTdescription = lstHLTList.SelectedItem.SubItems(1)
        txtHLTdescription.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstHLTList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstHLTList.SortKey = ColumnHeader.Index - 1
    lstHLTList.Sorted = True
    If lstHLTList.SortOrder = lvwAscending Then
       lstHLTList.SortOrder = lvwDescending
    Else
       lstHLTList.SortOrder = lvwAscending
    End If
    
    txtHLTcode.Enabled = False
End Sub

Private Sub DeleteHLT()
    Dim rsDeleteHLT As New ADODB.Recordset
    Dim rsCheckHLT As New ADODB.Recordset
    Dim rsCheckAGE As New ADODB.Recordset
    Dim rsCheckBNF As New ADODB.Recordset
    Dim rsCheckMCO As New ADODB.Recordset
    Dim rsCheckPLN As New ADODB.Recordset
    Dim rsCheckPRM As New ADODB.Recordset
    Dim rsCheckPRV As New ADODB.Recordset
    Dim strsql, sqlHLT, sqlAGE, sqlBNF, sqlMBM, sqlMCO, sqlPRM, sqlPRV, SqlPln As String
    Dim DeleteRecord
    EditFlag = False
        
    If txtHLTcode = "" Then
       MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
       'txtHLTcode.SetFocus
    ElseIf lstHLTList.ListItems.Count = 0 Then
       MsgBox "Please search th record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    sqlMBM = "Select MBMNumber From MBM where HLTCode = '" & RemStr(txtHLTcode) & "'" '(ChkStr(lstHLTList.SelectedItem.Text) & "'"
    SqlPln = "Select PLNCode From PLN where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
    sqlAGE = "Select AgeCode From AGE where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
    sqlBNF = "Select BNFCode From BNF where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
    sqlMCO = "Select MCOCode From MCO where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
    sqlPRM = "Select PRMCode From PRM where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
    sqlPRV = "Select PRVIndex From PRV where HLTCode = '" & RemStr(txtHLTcode) & "'" 'ChkStr(lstHLTList.SelectedItem.Text) & "'"
    
    
    rsCheckHLT.MaxRecords = 1
    rsCheckHLT.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckPLN.MaxRecords = 1
    rsCheckPLN.Open SqlPln, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckAGE.MaxRecords = 1
    rsCheckAGE.Open sqlAGE, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckBNF.MaxRecords = 1
    rsCheckBNF.Open sqlBNF, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckMCO.MaxRecords = 1
    rsCheckMCO.Open sqlMCO, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckPRM.MaxRecords = 1
    rsCheckPRM.Open sqlPRM, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckPRV.MaxRecords = 1
    rsCheckPRV.Open sqlPRV, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
    If rsCheckHLT.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckPLN.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Plan table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckAGE.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Age table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckBNF.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Benefit table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckMCO.recordCount Then
        MsgBox "This record cannot be deleted because it is used in MCO table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckPRM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Premium table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckPRV.recordCount Then
        MsgBox "This record cannot be deleted because it is used in PRV table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes And rsCheckHLT.recordCount = 0 Then
             strsql = "Delete From HLT Where HLTCode = '" & ChkStr(lstHLTList.SelectedItem.Text) & "'"
             rsDeleteHLT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call HLTListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             txtHLTcode.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
    End If
    Exit Sub
End Sub
    
Private Sub SearchHLT()
Dim strAppend As String
    
  Screen.MousePointer = vbHourglass
  strAppend = ""
  
  If Len(Trim(txtHLTcode)) Then
  strAppend = strAppend + " And HLTCode = '" & Trim(Mid(txtHLTcode, 1, 2)) & "'"
  'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  HLTCriteria = strAppend
  Call HLTListing(strAppend)
  Screen.MousePointer = Default
  
End Sub

'Private Sub PrintHLT()
    'RptHltList.show
'End Sub

Private Sub SaveHLT()
   On Error Resume Next
   Dim RecordSaved
   Dim HLT As New ADODB.Recordset
   Dim SaveHLT As New ADODB.Recordset
   Dim strsql As String
   
    If txtHLTcode = "" Then
        MsgBox "Please Enter Health Type", vbOKOnly + vbCritical, DialogBoxTitle
        txtHLTcode.SetFocus
    ElseIf txtHLTdescription = "" Then
        MsgBox "Please Enter Health Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtHLTdescription.SetFocus
    Else
        strsql = "Select * From HLT where HLTCode =  '" & Trim(txtHLTcode) & "'"
        HLT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If HLT.recordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           HLT.Close
           txtHLTcode.SetFocus
           txtHLTcode.SelStart = 0
           txtHLTcode.SelLength = Len(txtHLTcode)
           Set HLT = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveHLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
                 SaveHLT.AddNew
              Else
                 strsql = "Select * From HLT Where HLTCode = '" & ChkStr(txtHLTcode) & "'"
                 SaveHLT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveHLT
                !HLTCode = ChkStr(txtHLTcode)
                !HLTDescription = ChkStr(txtHLTdescription)
                !HLTLastUpdateUser = Logon
                !HLTLastUpdateDate = Date
                
              End With
              SaveHLT.Update
            
             Call ClearForm(Me)
             Call HLTListing(HLTCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             HLT.Close
             Set SaveHLT = Nothing
        End If
        EditFlag = False
     End If
   End If
End Sub

Private Sub INSListing(Optional strAppend As String)
Dim rsListINS As New ADODB.Recordset
Dim INSMaster As ListItem
Dim sqlINS As String
Dim TotalINSCount As Integer
        
  
    TotalINSCount = 0
    lstINSList.ListItems.Clear
    sqlINS = "select INSCode, INSDescription From INS where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sqlINS = sqlINS + strAppend
    End If
    
    rsListINS.Open sqlINS, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If rsListINS.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
    Do Until rsListINS.EOF
        With rsListINS
            Set INSMaster = lstINSList.ListItems.Add(, , !INSCode)
                INSMaster.SubItems(1) = !INSDescription
            .MoveNext
        End With
        TotalINSCount = TotalINSCount + 1
        txtTotalRecord = TotalINSCount
    Loop
    End If
    rsListINS.Close
    Set rsListINS = Nothing
    
End Sub

Private Sub lstINSList_Click()
Dim strsql As String
    
    If lstINSList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        txtINScode = lstINSList.SelectedItem.Text
        txtINScode.Enabled = False
        txtINSdescription = lstINSList.SelectedItem.SubItems(1)
        txtINSdescription.Enabled = False
        'Call EntriesEnable(False)
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstINSList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstINSList.SortKey = ColumnHeader.Index - 1
    lstINSList.Sorted = True
    If lstINSList.SortOrder = lvwAscending Then
       lstINSList.SortOrder = lvwDescending
    Else
       lstINSList.SortOrder = lvwAscending
    End If
    
    txtINScode.Enabled = False
End Sub

' '''''''''''''''''' Save, edit , Add, delete INS..... ''''''''''''''
Private Sub DeleteINS()
    
    Dim rsDeleteINS As New ADODB.Recordset
    Dim rsCheckINS As New ADODB.Recordset
    Dim rsCheckBNF As New ADODB.Recordset
    'Dim rsCheckEAF As New ADODB.Recordset
    Dim rsCheckPRM As New ADODB.Recordset
    Dim rsCheckMBMSEQ As New ADODB.Recordset
    Dim rsCheckPRV As New ADODB.Recordset
    Dim rsCheckMCO As New ADODB.Recordset
    Dim strsql, sqlINS, sqlBNF, sqlPRM, sqlMBMSEQ, sqlPRV, sqlMCO As String
    Dim DeleteRecord
    EditFlag = False
        
    If txtINScode = "" Then
       MsgBox "Please Enter Insured Type", vbOKOnly + vbCritical, DialogBoxTitle
       'txtINScode.Enabled = True
       'txtINScode.SetFocus
    ElseIf lstINSList.ListItems.Count = 0 Then
       MsgBox "Please search th record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    sqlINS = "Select MBMNumber From MBM where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    sqlBNF = "Select BNFCode From BNF where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    'sqlEAF = "Select EAFCode From EAF where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    sqlPRM = "Select PRMCode From PRM where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    sqlMBMSEQ = "Select SEQCode From MBMSEQ where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    sqlPRV = "Select PRVIndex From PRV where INSCode = '" & RemStr(txtINScode) & "'" 'ChkStr(lstINSList.SelectedItem.Text) & "'"
    sqlMCO = "Select MCOCode From MCO where INSCode = '" & ChkStr(lstINSList.SelectedItem.Text) & "'"
    
    rsCheckBNF.MaxRecords = 1
    rsCheckBNF.Open sqlBNF, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckMCO.MaxRecords = 1
    rsCheckMCO.Open sqlMCO, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckINS.MaxRecords = 1
    rsCheckINS.Open sqlINS, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'rsCheckEAF.MaxRecords = 1
    'rsCheckEAF.Open sqlEAF, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckPRM.MaxRecords = 1
    rsCheckPRM.Open sqlPRM, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckMBMSEQ.MaxRecords = 1
    rsCheckMBMSEQ.Open sqlMBMSEQ, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    rsCheckPRV.MaxRecords = 1
    rsCheckPRV.Open sqlPRV, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rsCheckBNF.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Benefit table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckMCO.recordCount Then
        MsgBox "This record cannot be deleted because it is used in MCO table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckINS.recordCount Then
        MsgBox "This record cannot be deleted because it is used in MBM table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    'ElseIf rsCheckEAF.recordCount Then
        'MsgBox "This record cannot be deleted because it is used in EAF table!", vbCritical, "Error in deleting record..."
        'Call ClearForm(Me)
        'CmdDelete.Enabled = False
        'CmdEdit.Enabled = False
    ElseIf rsCheckMBMSEQ.recordCount Then
        MsgBox "This record cannot be deleted because it is used in MBMSEQ table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckPRM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in PRM table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckPRV.recordCount Then
        MsgBox "This record cannot be deleted because it is used in PRV table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    ElseIf rsCheckINS.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes And rsCheckINS.recordCount = 0 Then
             strsql = "Delete From INS Where INSCode = '" & ChkStr(lstINSList.SelectedItem.Text) & "'"
             rsDeleteINS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call INSListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             txtINScode.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
 End If
    Exit Sub

End Sub
    
Private Sub SearchINS()
    Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(txtINScode)) Then
        strAppend = strAppend + " And INSCode = '" & ChkStr(txtINScode) & "'"
    End If
    
    INSCriteria = strAppend
    Call INSListing(strAppend)
    Screen.MousePointer = Default
    
End Sub

Private Sub SaveINS()
Dim RecordSaved
Dim INS As New ADODB.Recordset
Dim SaveINS As New ADODB.Recordset
Dim strsql As String
    
    If txtINScode = "" Then
        MsgBox "Please Enter Insured Type", vbOKOnly + vbCritical, DialogBoxTitle
        txtINScode.SetFocus
    ElseIf txtINSdescription = "" Then
        MsgBox "Please Enter Insured Type Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtINSdescription.SetFocus
    Else
        strsql = "Select * From INS where INSCode = '" & Trim(txtINScode) & "'"
        INS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If INS.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            INS.Close
            txtINScode.SetFocus
            txtINScode.SelStart = 0
            txtINScode.SelLength = Len(txtINScode)
            Set INS = Nothing
            Set SaveINS = Nothing
            Exit Sub
            
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                   SaveINS.Open "INS", medixmasterconn, adOpenKeyset, adLockOptimistic
                   SaveINS.AddNew
                Else
                   strsql = "Select * From INS Where INSCode = '" & ChkStr(txtINScode) & "'"
                   SaveINS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                End If
                With SaveINS
                     !INSCode = ChkStr(txtINScode)
                     !INSDescription = ChkStr(txtINSdescription)
                     !INSLastUpdateUser = Logon
                     !INSLastUpdateDate = Date
                     
                     
                End With
                SaveINS.Update
                Call ClearForm(Me)
                Call INSListing(INSCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        
                SaveINS.Close
            End If
            EditFlag = False
      End If
  End If
End Sub

'Private Sub cboINSCode_KeyPress(KeyAscii As Integer)
'If secondeditflag = False Then


'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer

'If KeyAscii > 32 And KeyAscii < 127 Then
'NewText = ChkStr(Left(cboINSCode.Text, cboINSCode.SelStart) & _
'            Chr(KeyAscii) + Mid(cboINSCode.Text, cboINSCode.SelStart + cboINSCode.SelLength + 1))
'        ValidCount = 0
'        ValidValue = ""
'        For i = 0 To cboINSCode.ListCount - 1
'          If NewText = ChkStr(Left(cboINSCode.list(i), Len(NewText))) Then
'              ValidCount = ValidCount + 1
'              ValidValue = cboINSCode.list(i)
'          End If
'          Next
'        If ValidCount <= 1 Then KeyAscii = 0
'        End If
'
'        If ValidCount = 1 Then
'        cboINSCode.Text = ValidValue
'        cboINSCode.SelStart = 0
'        cboINSCode.SelLength = Len(ValidValue)
'        End If
       
        
     
'End If

'End Sub

'Private Sub PrintINS()
    'RptInsList.show
'End Sub
' '''''''''''''''''' END ==> Save, edit , Add, delete INS .... ''''''''''''''


'**********************************Start MCOFees**************************************
Private Sub SaveMCOFees()
Dim RecordSaved
Dim PLN As New ADODB.Recordset
Dim MCO1 As New ADODB.Recordset
Dim MCO2 As New ADODB.Recordset
Dim MCO3 As New ADODB.Recordset
Dim MCO4 As New ADODB.Recordset
Dim MCO5 As New ADODB.Recordset
Dim MCO6 As New ADODB.Recordset
Dim MCO7 As New ADODB.Recordset
Dim MCO8 As New ADODB.Recordset
Dim SaveMCO As New ADODB.Recordset
Dim strsql As String
Dim StrSql1 As String
Dim strsql2 As String
Dim strsql3 As String
Dim strsql4 As String
Dim strsql5 As String
Dim strsql6 As String
Dim strsql7 As String
Dim strsql8 As String
Dim ListCount As Integer
Dim listloop As Integer

    If CboMCOHlt = "" Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical
        CboMCOHlt.SetFocus
    ElseIf TxtMCOPlan = "" Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical
        TxtMCOPlan.SetFocus
    ElseIf Mid(CboMCOStatus, 1, 1) = "D" And TxtCoverID = "" Then
        MsgBox "Please Enter Covered ID Charge", vbOKOnly + vbCritical
        TxtCoverID.SetFocus
    ElseIf txtEffDate = "__/__/____" Or Not IsDate(txtEffDate) Then
        MsgBox "Please validate your MCO Effective date", vbOKOnly + vbCritical
        txtEffDate.SetFocus
    ElseIf TxtIAAmt = "" And TxtICAmt = "" And TxtFAAmt = "" And TxtFCAmt = "" _
           And TxtGAAmt = "" And TxtGCAmt = "" And TxtHAAmt = "" And TxtHCAmt = "" Then
        MsgBox "Please Enter MCO Fees ", vbOKOnly + vbCritical
        TxtIAAmt.SetFocus
    Else
    
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        
            strsql = " Select PLNCode, HLTCode From PLN where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                     " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'"
            PLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            If PLN.recordCount = 0 Then
                MsgBox "Please validate your Plan Code", vbOKOnly + vbCritical
                TxtMCOPlan.SetFocus
                TxtMCOPlan.SelStart = 0
                TxtMCOPlan.SelLength = Len(TxtMCOPlan)
                Set PLN = Nothing
                Exit Sub
            Else
            
            If RecordSaved = vbYes Then
                ListCount = 4
                For listloop = 1 To ListCount
                    
                If Mid(CboMCOStatus, 1, 1) = "A" Then
                    
                    If listloop = 1 Then
                        If Trim(TxtIAAmt) <> "" Then
                            StrSql1 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                      " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                      " AND INSCode = 'I' AND MCOAmountAdult = " & Trim(TxtIAAmt) & "" & _
                                      " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                      " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                            MCO1.Open StrSql1, medixmasterconn, adOpenKeyset, adLockOptimistic
                            If MCO1.recordCount > 0 Then
                               MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                               MCO1.Close
                               TxtIAAmt.SetFocus
                               TxtIAAmt.SelStart = 0
                               TxtIAAmt.SelLength = Len(TxtIAAmt)
                               Set MCO1 = Nothing
                               Exit Sub
                            End If
                                SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                                SaveMCO.AddNew
                                SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                                SaveMCO!MCOAmountAdult = Format(TxtIAAmt, "#,##0.00")
                                SaveMCO!MCOAmountChild = Format(TxtICAmt, "#,##0.00")
                                SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                                SaveMCO!INSCode = "I"
                                'SaveMCO!AdultChild = "A"
                                SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                                SaveMCO!SuppMCOStatus = "A"
                                SaveMCO!MCOLastUpdateUser = Logon
                                SaveMCO!MCOLastUpdateDate = Date
                                
                                SaveMCO.Update
                                SaveMCO.Close
                                Set SaveMCO = Nothing
                        'Else
                            'Exit For
                        End If
                    
                ElseIf listloop = 2 Then
                    If Trim(TxtFAAmt) <> "" Then
                        strsql3 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'F' AND MCOAmountAdult = " & Trim(TxtFAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO3.Open strsql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO3.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO3.Close
                           TxtFAAmt.SetFocus
                           TxtFAAmt.SelStart = 0
                           TxtFAAmt.SelLength = Len(TxtFAAmt)
                           Set MCO3 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtFAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtFCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "F"
                            'SaveMCO!AdultChild = "A"
                            SaveMCO!SuppMCOStatus = "A"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    'Else
                        'Exit For
                    End If
                    
                                    
                ElseIf listloop = 3 Then
                    If Trim(TxtGAAmt) <> "" Then
                        strsql5 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'G' AND MCOAmountAdult = " & Trim(TxtGAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO5.Open strsql5, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO5.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO5.Close
                           TxtGAAmt.SetFocus
                           TxtGAAmt.SelStart = 0
                           TxtGAAmt.SelLength = Len(TxtGAAmt)
                           Set MCO5 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtGAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtGCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "G"
                            SaveMCO!SuppMCOStatus = "A"
                            'SaveMCO!AdultChild = "A"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    'Else
                        'Exit For
                    End If
                    
                
                ElseIf listloop = 4 Then
                    If Trim(TxtHAAmt) <> "" Then
                        strsql7 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'H' AND MCOAmountAdult = " & Trim(TxtHAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO7.Open strsql7, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO7.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO7.Close
                           TxtHAAmt.SetFocus
                           TxtHAAmt.SelStart = 0
                           TxtHAAmt.SelLength = Len(TxtHAAmt)
                           Set MCO7 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtHAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtHCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "H"
                            'SaveMCO!AdultChild = "A"
                            SaveMCO!SuppMCOStatus = "A"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    'Else
                        'Exit For
                    End If
                
                Else
                    Exit For
                End If
            
                
            ElseIf Mid(CboMCOStatus, 1, 1) = "B" Then
            
                If listloop = 1 Then
                    If Trim(TxtIAAmt) <> "" Then
                        StrSql1 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'I' AND MCOAmountAdult = " & Trim(TxtIAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO1.Open StrSql1, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO1.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO1.Close
                           TxtIAAmt.SetFocus
                           TxtIAAmt.SelStart = 0
                           TxtIAAmt.SelLength = Len(TxtIAAmt)
                           Set MCO1 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtIAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtICAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "I"
                            SaveMCO!SuppMCOStatus = "B"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    End If
                    
                ElseIf listloop = 2 Then
                    If Trim(TxtFAAmt) <> "" Then
                        strsql3 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'F' AND MCOAmountAdult = " & Trim(TxtFAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO3.Open strsql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO3.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO3.Close
                           TxtFAAmt.SetFocus
                           TxtFAAmt.SelStart = 0
                           TxtFAAmt.SelLength = Len(TxtFAAmt)
                           Set MCO3 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtFAAmt, "#,##0.00")
                            SaveMCO!MCOSuppAdultAmt = Format(TxtSFAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtFCAmt, "#,##0.00")
                            SaveMCO!MCOSuppChildAmt = Format(TxtSFCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "F"
                            SaveMCO!SuppMCOStatus = "B"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    End If
                
                    
                ElseIf listloop = 3 Then
                    If Trim(TxtGAAmt) <> "" Then
                        strsql5 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'G' AND MCOAmountAdult = " & Trim(TxtGAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO5.Open strsql5, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO5.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO5.Close
                           TxtGAAmt.SetFocus
                           TxtGAAmt.SelStart = 0
                           TxtGAAmt.SelLength = Len(TxtGAAmt)
                           Set MCO5 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtGAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtGCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "G"
                            SaveMCO!SuppMCOStatus = "B"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    End If
                    
                ElseIf listloop = 4 Then
                    If Trim(TxtHAAmt) <> "" Then
                        strsql7 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'H' AND MCOAmountAdult = " & Trim(TxtHAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO7.Open strsql7, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO7.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO7.Close
                           TxtHAAmt.SetFocus
                           TxtHAAmt.SelStart = 0
                           TxtHAAmt.SelLength = Len(TxtHAAmt)
                           Set MCO7 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtHAAmt, "#,##0.00")
                            SaveMCO!MCOSuppAdultAmt = Format(TxtSHAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtHCAmt, "#,##0.00")
                            SaveMCO!MCOSuppChildAmt = Format(TxtSHCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "H"
                            SaveMCO!SuppMCOStatus = "B"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    End If
                        
                Else
                    Exit For
                End If
            
            
            ElseIf Mid(CboMCOStatus, 1, 1) = "C" Then
                       
                If listloop = 1 Then
                    If Trim(TxtIAAmt) <> "" Then
                        StrSql1 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'I' AND MCOAmountAdult = " & Trim(TxtIAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO1.Open StrSql1, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO1.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO1.Close
                           TxtIAAmt.SetFocus
                           TxtIAAmt.SelStart = 0
                           TxtIAAmt.SelLength = Len(TxtIAAmt)
                           Set MCO1 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtIAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtICAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "I"
                            SaveMCO!SuppMCOStatus = "C"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    End If
                
                    
                ElseIf listloop = 2 Then
                    If Trim(TxtFAAmt) <> "" Then
                        strsql3 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'F' AND MCOAmountAdult = " & Trim(TxtFAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO3.Open strsql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO3.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO3.Close
                           TxtFAAmt.SetFocus
                           TxtFAAmt.SelStart = 0
                           TxtFAAmt.SelLength = Len(TxtFAAmt)
                           Set MCO3 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtFAAmt, "#,##0.00")
                            SaveMCO!MCOSuppAdultAmt = Format(TxtSFAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtFCAmt, "#,##0.00")
                            SaveMCO!MCOSuppChildAmt = Format(TxtSFCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "F"
                            SaveMCO!SuppMCOStatus = "C"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    End If
                
                    
                ElseIf listloop = 3 Then
                    If Trim(TxtGAAmt) <> "" Then
                        strsql5 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'G' AND MCOAmountAdult = " & Trim(TxtGAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO5.Open strsql5, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO5.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO5.Close
                           TxtGAAmt.SetFocus
                           TxtGAAmt.SelStart = 0
                           TxtGAAmt.SelLength = Len(TxtGAAmt)
                           Set MCO5 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtGAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtGCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "G"
                            SaveMCO!SuppMCOStatus = "C"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    
                    End If
                    
                                    
                ElseIf listloop = 4 Then
                    If Trim(TxtHAAmt) <> "" Then
                        strsql7 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'H' AND MCOAmountAdult = " & Trim(TxtHAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO7.Open strsql7, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO7.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO7.Close
                           TxtHAAmt.SetFocus
                           TxtHAAmt.SelStart = 0
                           TxtHAAmt.SelLength = Len(TxtHAAmt)
                           Set MCO7 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtHAAmt, "#,##0.00")
                            SaveMCO!MCOSuppAdultAmt = Format(TxtSHAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtHCAmt, "#,##0.00")
                            SaveMCO!MCOSuppChildAmt = Format(TxtSHCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "H"
                            SaveMCO!SuppMCOStatus = "C"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    
                    End If
                        
                Else
                    Exit For
                End If
            
                ElseIf Mid(CboMCOStatus, 1, 1) = "D" Then
                       
                If listloop = 1 Then
                    If Trim(TxtIAAmt) <> "" Then
                        StrSql1 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'I' AND MCOAmountAdult = " & Trim(TxtIAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO1.Open StrSql1, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO1.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO1.Close
                           TxtIAAmt.SetFocus
                           TxtIAAmt.SelStart = 0
                           TxtIAAmt.SelLength = Len(TxtIAAmt)
                           Set MCO1 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtIAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtICAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "I"
                            SaveMCO!SuppMCOStatus = "D"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    End If
                
                    
                ElseIf listloop = 2 Then
                    If Trim(TxtFAAmt) <> "" Then
                        strsql3 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'F' AND MCOAmountAdult = " & Trim(TxtFAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO3.Open strsql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO3.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO3.Close
                           TxtFAAmt.SetFocus
                           TxtFAAmt.SelStart = 0
                           TxtFAAmt.SelLength = Len(TxtFAAmt)
                           Set MCO3 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtFAAmt, "#,##0.00")
                            SaveMCO!MCOSuppAdultAmt = Format(TxtSFAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtFCAmt, "#,##0.00")
                            SaveMCO!MCOSuppChildAmt = Format(TxtSFCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "F"
                            SaveMCO!CovIDChrg = TxtCoverID
                            SaveMCO!SuppMCOStatus = "D"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    
                    End If
                
                    
                ElseIf listloop = 3 Then
                    If Trim(TxtGAAmt) <> "" Then
                        strsql5 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'G' AND MCOAmountAdult = " & Trim(TxtGAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO5.Open strsql5, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO5.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO5.Close
                           TxtGAAmt.SetFocus
                           TxtGAAmt.SelStart = 0
                           TxtGAAmt.SelLength = Len(TxtGAAmt)
                           Set MCO5 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtGAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtGCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "G"
                            SaveMCO!SuppMCOStatus = "D"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    
                    End If
                    
                                    
                ElseIf listloop = 4 Then
                    If Trim(TxtHAAmt) <> "" Then
                        strsql7 = " Select * From MCO where PLNCode = '" & Trim(TxtMCOPlan) & "'" & _
                                  " AND HLTCode = '" & Trim(Mid(CboMCOHlt, 1, 1)) & "'" & _
                                  " AND INSCode = 'H' AND MCOAmountAdult = " & Trim(TxtHAAmt) & "" & _
                                  " AND MCOEffDate = '" & Format(txtEffDate, "mm/dd/yyyy") & "'" & _
                                  " AND SuppMCOStatus = '" & Trim(Mid(CboMCOStatus, 1, 1)) & "'"
                        MCO7.Open strsql7, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MCO7.recordCount > 0 Then
                           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                           MCO7.Close
                           TxtHAAmt.SetFocus
                           TxtHAAmt.SelStart = 0
                           TxtHAAmt.SelLength = Len(TxtHAAmt)
                           Set MCO7 = Nothing
                           Exit Sub
                        End If
                            SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                            SaveMCO.AddNew
                            SaveMCO!PLNCode = ChkStr(TxtMCOPlan)
                            SaveMCO!MCOAmountAdult = Format(TxtHAAmt, "#,##0.00")
                            SaveMCO!MCOSuppAdultAmt = Format(TxtSHAAmt, "#,##0.00")
                            SaveMCO!MCOAmountChild = Format(TxtHCAmt, "#,##0.00")
                            SaveMCO!MCOSuppChildAmt = Format(TxtSHCAmt, "#,##0.00")
                            SaveMCO!HLTCode = Trim(Mid(CboMCOHlt, 1, 1))
                            SaveMCO!INSCode = "H"
                            SaveMCO!CovIDChrg = TxtCoverID
                            SaveMCO!SuppMCOStatus = "D"
                            SaveMCO!MCOEffDate = Format(txtEffDate, "dd/mm/yyyy")
                            SaveMCO!MCOLastUpdateUser = Logon
                            SaveMCO!MCOLastUpdateDate = Date
                            
                            SaveMCO.Update
                            SaveMCO.Close
                            Set SaveMCO = Nothing
                    
                    End If
                        
                Else
                    Exit For
                End If
                
            End If
            
            Next listloop
            
                'Call ClearForm(Me)
                Call MCOListing(MCOCriteria)
                    MsgBox "Record Saved...", vbOKOnly + vbInformation
            End If
        End If
    End If
End Sub

Private Sub DeleteMCOFees()

Dim DeleteMCO As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If CboMCOHlt = "" Then
   MsgBox "Please Enter Health Type", vbOKOnly + vbCritical, DialogBoxTitle
   'CboMCOHlt.SetFocus
ElseIf lstMCOList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where INSCode = '" & ChkStr(lstMCOList.SelectedItem.Text) & "'" & _
            " and MBMPayorEffdate >='" & Format(lstMCOList.SelectedItem.SubItems(6), "mm/dd/yyyy") & "'"
   
   rsCheckMBM.MaxRecords = 1
    rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
     
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From MCO Where MCOcode = '" & ChkStr(lstMCOList.SelectedItem.Text) & "'"
           DeleteMCO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call MCOListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtINScode.SetFocus
         End If
     End If
End If
Exit Sub
End Sub

Private Sub SearchMCO()
    Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(CboMCOHlt)) Then
        strAppend = strAppend + " AND HLTCode = '" & Mid(CboMCOHlt, 1, 1) & "'"
    End If
    
    If Len(Trim(TxtMCOPlan)) Then
        strAppend = strAppend + " AND PLNCode = '" & Trim(TxtMCOPlan) & "'"
    End If
    

    If txtEffDate <> "__/__/____" Then
        strAppend = strAppend + " AND MCOEffDate = '" & Format(CDate(txtEffDate), "MM/DD/YYYY") & "'"
    End If
        
    MCOCriteria = strAppend
    Call MCOListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub MCOListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
    Dim MCOMaster As ListItem
    Dim sql As String
    Dim TotalMCOCount As Integer
    
    TotalMCOCount = 0
    lstMCOList.ListItems.Clear
    sql = "Select * from MCO where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
    sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation
        
    Else
        Do Until rst2.EOF
            Set MCOMaster = lstMCOList.ListItems.Add(, , rst2!HLTCode)
                MCOMaster.SubItems(1) = Trim(rst2!INSCode)
                If Len(rst2!PLNCode) Then
                    MCOMaster.SubItems(2) = Trim(rst2!PLNCode)
                End If
                If Len(rst2!SuppMCOStatus) Then
                    MCOMaster.SubItems(3) = Trim(rst2!SuppMCOStatus)
                End If
                MCOMaster.SubItems(4) = Format(rst2!MCOAmountAdult, "#,##0.00")
                MCOMaster.SubItems(5) = Format(rst2!MCOAmountChild, "#,##0.00")
                MCOMaster.SubItems(6) = Format(rst2!MCOSuppAdultAmt, "#,##0.00")
                MCOMaster.SubItems(7) = Format(rst2!MCOSuppChildAmt, "#,##0.00")
                MCOMaster.SubItems(8) = IIf(Len(Trim(rst2!MCOEffDate)) >= 1, rst2!MCOEffDate, "")
                If Len(rst2!CovIDChrg) Then
                    MCOMaster.SubItems(9) = Trim(rst2!CovIDChrg)
                End If
                
                rst2.MoveNext
        TotalMCOCount = TotalMCOCount + 1
        txtTotalRecord = TotalMCOCount
        Loop
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub MCOComboListing()
Dim HLT As New ADODB.Recordset
Dim MCO As New ADODB.Recordset
    
    HLT.Open "HLT", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    MCO.Open "SuppMCOStatus", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    CboMCOHlt.Clear
    CboMCOStatus.Clear
    
    Do Until HLT.EOF
        CboMCOHlt.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
        HLT.MoveNext
    Loop
    HLT.Close
    Set HLT = Nothing
        
    Do Until MCO.EOF
        CboMCOStatus.AddItem MCO!MCOCode & " - " & MCO!MCOStatus
        MCO.MoveNext
    Loop
    MCO.Close
    Set MCO = Nothing
    
End Sub

Private Sub lstMCOList_Click()
Dim MCO As New ADODB.Recordset
Dim strsql As String
Dim Amt1, Amt2, Amt3, Amt4, Amt5, Amt6, Amt7, Amt8 As String
Dim Amt9, Amt10, Amt11, Amt12 As String
    
'    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

   Amt1 = 0
   Amt2 = 0
   Amt3 = 0
   Amt4 = 0
   Amt5 = 0
   Amt6 = 0
   Amt7 = 0
   Amt8 = 0
   Amt9 = 0
   Amt10 = 0
   Amt11 = 0
   Amt12 = 0
   
   If lstMCOList.ListItems.Count <> 0 Then
        Call ClearForm(Me)
        
        strsql = " Select * From MCO where PLNCode = '" & lstMCOList.SelectedItem.SubItems(2) & "'  AND " & _
                 " MCOEffDate = '" & Format(lstMCOList.SelectedItem.SubItems(8), "mm/dd/yyyy") & "'"
                 
        MCO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        Do Until MCO.EOF
            If MCO!SuppMCOStatus = "A" Then
                If MCO!INSCode = "I" Then
                    Amt1 = MCO!MCOAmountAdult
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then
                    Amt3 = MCO!MCOAmountAdult
                    Amt4 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "G" Then
                    Amt5 = MCO!MCOAmountAdult
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then
                    Amt7 = MCO!MCOAmountAdult
                    Amt8 = MCO!MCOAmountChild
                End If
            ElseIf MCO!SuppMCOStatus = "B" Then
                If MCO!INSCode = "I" Then
                    Amt1 = MCO!MCOAmountAdult
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then
                    Amt3 = MCO!MCOAmountAdult
                    Amt9 = MCO!MCOSuppAdultAmt
                    Amt4 = MCO!MCOAmountChild
                    Amt10 = MCO!MCOSuppChildAmt
                ElseIf MCO!INSCode = "G" Then
                    Amt5 = MCO!MCOAmountAdult
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then
                    Amt7 = MCO!MCOAmountAdult
                    Amt11 = MCO!MCOSuppAdultAmt
                    Amt8 = MCO!MCOAmountChild
                    Amt12 = MCO!MCOSuppChildAmt
                End If
            ElseIf MCO!SuppMCOStatus = "C" Then
                If MCO!INSCode = "I" Then
                    Amt1 = MCO!MCOAmountAdult
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then
                    Amt3 = MCO!MCOAmountAdult
                    Amt9 = MCO!MCOSuppAdultAmt
                    Amt4 = MCO!MCOAmountChild
                    Amt10 = MCO!MCOSuppChildAmt
                ElseIf MCO!INSCode = "G" Then
                    Amt5 = MCO!MCOAmountAdult
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then
                    Amt7 = MCO!MCOAmountAdult
                    Amt11 = MCO!MCOSuppAdultAmt
                    Amt8 = MCO!MCOAmountChild
                    Amt12 = MCO!MCOSuppChildAmt
                End If
            ElseIf MCO!SuppMCOStatus = "D" Then
                If MCO!INSCode = "I" Then
                    Amt1 = MCO!MCOAmountAdult
                    Amt2 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "F" Then
                    Amt3 = MCO!MCOAmountAdult
                    Amt9 = MCO!MCOSuppAdultAmt
                    Amt4 = MCO!MCOAmountChild
                    Amt10 = MCO!MCOSuppChildAmt
                ElseIf MCO!INSCode = "G" Then
                    Amt5 = MCO!MCOAmountAdult
                    Amt6 = MCO!MCOAmountChild
                ElseIf MCO!INSCode = "H" Then
                    Amt7 = MCO!MCOAmountAdult
                    Amt11 = MCO!MCOSuppAdultAmt
                    Amt8 = MCO!MCOAmountChild
                    Amt12 = MCO!MCOSuppChildAmt
                End If
            End If
        MCO.MoveNext
        Loop
        MCO.Close
        Set MCO = Nothing
    
                
        CboMCOHlt = lstMCOList.SelectedItem.Text
        TxtMCOPlan = lstMCOList.SelectedItem.SubItems(2)
        
        If lstMCOList.SelectedItem.SubItems(3) = "A" Then
            CboMCOStatus.ListIndex = 0
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "B" Then
            CboMCOStatus.ListIndex = 1
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "C" Then
            CboMCOStatus.ListIndex = 2
        ElseIf lstMCOList.SelectedItem.SubItems(3) = "D" Then
            CboMCOStatus.ListIndex = 3
        End If
        
        TxtIAAmt = Format(Amt1, "#,##0.00")
        TxtICAmt = Format(Amt2, "#,##0.00")
    
        TxtFAAmt = Format(Amt3, "#,##0.00")
        TxtFCAmt = Format(Amt4, "#,##0.00")
        TxtSFAAmt = Format(Amt9, "#,##0.00")
        TxtSFCAmt = Format(Amt10, "#,##0.00")
    
        TxtGAAmt = Format(Amt5, "#,##0.00")
        TxtGCAmt = Format(Amt6, "#,##0.00")
    
        TxtHAAmt = Format(Amt7, "#,##0.00")
        TxtHCAmt = Format(Amt8, "#,##0.00")
        TxtSHAAmt = Format(Amt11, "#,##0.00")
        TxtSHCAmt = Format(Amt12, "#,##0.00")
        
        txtEffDate = lstMCOList.SelectedItem.SubItems(8)
        TxtCoverID = lstMCOList.SelectedItem.SubItems(9)
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing

End Sub

Private Sub lstMCOList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstMCOList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstMCOList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstMCOList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstMCOList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstMCOList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstMCOList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstMCOList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstMCOList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstMCOList, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstMCOList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    End Select
End Sub
    
   
Private Sub SaveAnnualLimit()
Dim Annual As New ADODB.Recordset
Dim SaveAnnual As New ADODB.Recordset
Dim RecordSaved
Dim strsql As String
Dim LifeTime As Boolean
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset

         
    If InStr(1, CboPlnCode, "-") > 0 Then
        strsqlPLN = " Select PLNCode, PLNLifeTimeStatus From PLN Where PLNCode = '" & Trim(Mid(CboPlnCode, 1, InStr(1, CboPlnCode, "-") - 1)) & "'" & _
                    " AND PLNLifeTimeStatus = 'Y'"
    Else
        strsqlPLN = "Select PLNCode, PLNLifeTimeStatus From PLN Where PLNCode = '" & Trim(Mid(CboPlnCode, 1, 4)) & "' AND PLNLifeTimeStatus = 'Y'"
    End If
    
    PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If PLN.recordCount > 0 Then
        LifeTime = True
    Else
        LifeTime = False
    End If
    
    PLN.Close
    Set PLN = Nothing
    
    If Len(Trim(CboPlnCode)) = 0 Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboPlnCode.SetFocus
    ElseIf Len(Trim(CboHltCode)) = 0 Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboHltCode.SetFocus
    ElseIf Len(Trim(CboInsType)) = 0 Then
        MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboInsType.SetFocus
    ElseIf Len(Trim(TxtAnnualLimit)) = 0 Then
        MsgBox "Please Enter Annual Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtAnnualLimit.SetFocus
    ElseIf LifeTime = True And TxtLifeTimeLimit = "" Then
        MsgBox "Please Enter Life Time Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtLifeTimeLimit.SetFocus
    ElseIf txtALEffDate = "__/__/____" Then
        MsgBox "Please Enter Annual Limit Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtALEffDate.SetFocus
    ElseIf txtANNVersion = "" Then
        MsgBox "Please Enter Annual Limit Version", vbOKOnly + vbCritical, DialogBoxTitle
        txtANNVersion.SetFocus
    ElseIf Mid(Combo1, 1, 1) = "B" And TxtSupplyAL = "" Then
        MsgBox "Please Enter Supp Premium Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSupplyAL.SetFocus
    ElseIf Mid(Combo1, 1, 1) = "C" And TxtSupplyAL = "" Then
        MsgBox "Please Enter Supp Premium Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSupplyAL.SetFocus
    ElseIf Mid(Combo1, 1, 1) = "B" And TxtSuppLifeTimeLimit = "" Then
        MsgBox "Please Enter Supp Life Time Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSuppLifeTimeLimit.SetFocus
    ElseIf Mid(Combo1, 1, 1) = "C" And TxtSuppLifeTimeLimit = "" Then
        MsgBox "Please Enter Supp Life Time Limit", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSuppLifeTimeLimit.SetFocus
    Else
        
        strsql = "Select * From AnnualLimit "
        strsql = strsql & " Where " & _
                " INSCode = '" & Trim(Mid(CboInsType, 1, 2)) & "' " & _
                " AND HLTCode = '" & Trim(Mid(CboHltCode, 1, 2)) & "' " & _
                " AND PLNCode = '" & Trim(Mid(CboPlnCode, 1, 4)) & "' "
        Annual.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
      If Annual.recordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! " & vbNewLine & _
                " Please Change The Corresponding Field", vbCritical, "Error"
        Annual.Close
        CboInsType.SetFocus
        CboInsType.SelStart = 0
        CboInsType.SelLength = Len(CboInsType)
        Set Annual = Nothing
        Exit Sub
      Else
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                Save = True
                If EditFlag = False Then
                    SaveAnnual.Open "AnnualLimit", medixmasterconn, adOpenKeyset, adLockOptimistic
                    SaveAnnual.AddNew
                Else
                    strsql = "Select * From AnnualLimit Where AnnualIndex = '" & ChkStr(lstAnnualList.SelectedItem.SubItems(10)) & "'"
                    'If Trim(temptxtAnnEffDate) <> "" Then
                        'strsql = strsql & " and AnnualEffDate = '" & Format(temptxtAnnEffDate, "mm/dd/yyyy") & "' "
                    'Else
                        'strsql = strsql & " and (AnnualEffDate) IS  NULL "
                    'End If
                    SaveAnnual.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                End If
            
                With SaveAnnual
                    If InStr(1, CboPlnCode, "-") > 0 Then
                        !PLNCode = Trim(Mid(CboPlnCode, 1, InStr(1, CboPlnCode, "-") - 1))
                    Else
                        !PLNCode = Trim(Mid(CboPlnCode, 1, 4))
                    End If
                    
                    !INSCode = Mid(CboInsType, 1, 2)
                    !HLTCode = Mid(CboHltCode, 1, 2)
                    
                    If Len(CboPayorCode) Then
                        !PAYCode = Mid(CboPayorCode, 1, 2)
                    Else
                        !PAYCode = Empty
                    End If
                    
                    If Len(CboGrpCom) Then
                        !Grpcompany = CboGrpCom
                    Else
                        !Grpcompany = Empty
                    End If
                    !AnnLimit = TxtAnnualLimit
                    If TxtLifeTimeLimit <> "" Then
                        !LifeTimeLimit = TxtLifeTimeLimit
                    End If
                    
                    If Mid(Combo1, 1, 1) = "A" Then
                        !SuppLimitStatus = "A"
                        '!SuppLimit = TxtSupplyAL
                    ElseIf Mid(Combo1, 1, 1) = "B" Then
                        !SuppLimitStatus = "B"
                        !SuppLimit = TxtSupplyAL
                        !SuppLifeTimeLimit = TxtSuppLifeTimeLimit
                    ElseIf Mid(Combo1, 1, 1) = "C" Then
                        !SuppLimitStatus = "C"
                        !SuppLimit = TxtSupplyAL
                        !SuppLifeTimeLimit = TxtSuppLifeTimeLimit
                    End If
                                        
                    !AnnualEffDate = IIf(txtALEffDate = "__/__/____", "", CDate(txtALEffDate))
                    !AnnualLimitVersion = txtANNVersion
                    !AnnualLastUpdateDUser = Logon
                    !AnnualLastUpdateDDate = Date
                    If Len(txtDisabilityLmt) Then
                        !DisabilityLmt = txtDisabilityLmt
                    'Else
                        '!DisabilityLmt = Empty
                    End If
                End With
                SaveAnnual.Update
                Call ClearForm(Me)
                'Call AnnualListing(AnnualCriteria)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
 
                SaveAnnual.Close
                Set SaveAnnual = Nothing
                Call ClearForm(Me)
                Call AnnualListing
                Call DisableCmdButton(Me)
                Call EntriesEnable(True)
            End If
        EditFlag = False
        Save = False
   End If
  End If
  'End If
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub DeleteAnnual()
Dim DeleteAnnual As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If CboPlnCode = "" Then
   MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
   'CboPlnCode.SetFocus
ElseIf lstAnnualList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where PLNCode = '" & RemStr(CboPlnCode) & "'" 'ChkStr(lstAnnualList.SelectedItem.Text) & "'"
   
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
     
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From AnnualLimit Where AnnualIndex = '" & ChkStr(lstAnnualList.SelectedItem.SubItems(10)) & "'"
           DeleteAnnual.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call AnnualListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           CboPlnCode.SetFocus
         End If
     End If
End If
Exit Sub
End Sub

Private Sub SearchAnnual()
    Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(CboPlnCode)) Then
        strAppend = strAppend + " And PLNCode = '" & Trim(Mid(CboPlnCode, 1, IIf(InStr(1, CboPlnCode, "-") = 0, 4, InStr(1, CboPlnCode, "-") - 1))) & "'"
    End If
    If Len(Trim(CboHltCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Trim(Mid(CboHltCode, 1, IIf(InStr(1, CboHltCode, "-") = 0, 4, InStr(1, CboHltCode, "-") - 1))) & "'"
        'strAppend = strAppend + " And HLTCode = '" & Trim(Mid(CboHltCode, 1, 2)) & "'"
    End If
    
    If Len(Trim(CboInsType)) Then
        strAppend = strAppend + " And INSCode = '" & Trim(Mid(CboInsType, 1, IIf(InStr(1, CboInsType, "-") = 0, 4, InStr(1, CboInsType, "-") - 1))) & "'"
        'strAppend = strAppend + " And INSCode = '" & Trim(Mid(CboInsType, 1, 2)) & "'"
    End If
    
    If Len(Trim(CboPayorCode)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayorCode, 1, IIf(InStr(1, CboPayorCode, "-") = 0, 4, InStr(1, CboPayorCode, "-") - 1))) & "'"
    End If

    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend + " AND GRPCompany like '" & RemStr(Mid(CboGrpCom, 1, IIf(InStr(1, CboGrpCom, "-") = 0, 4, InStr(1, CboGrpCom, "-") - 1))) & "%'"
    End If
    
    AnnualCriteria = strAppend
    Call AnnualListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub AnnualListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim AnnualMaster As ListItem
    Dim sql As String
    Dim TotalAnnualCount As Integer
    
    TotalAnnualCount = 0
    lstAnnualList.ListItems.Clear
    
    sql = " Select AnnualIndex, PLNCode, INSCode, HLTCode, PAYCode, GRPCompany, " & _
          " AnnLimit, SuppLimit, AnnualEffDate, SuppLimitStatus, AnnualLimitVersion, LifeTimeLimit, " & _
          " SuppLifeTimeLimit, DisabilityLmt From AnnualLimit where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
    sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do While Not rst2.EOF
            Set AnnualMaster = lstAnnualList.ListItems.Add(, , IIf(Len(rst2!HLTCode), rst2!HLTCode, ""))
                AnnualMaster.SubItems(1) = IIf(Len(rst2!PLNCode), rst2!PLNCode, "")
                AnnualMaster.SubItems(2) = IIf(Len(rst2!INSCode), rst2!INSCode, "")
                AnnualMaster.SubItems(3) = Format(IIf(Len(rst2!AnnLimit), rst2!AnnLimit, ""), "#,##0.00")
                AnnualMaster.SubItems(4) = Format(IIf(Len(rst2!SuppLimit), rst2!SuppLimit, ""), "#,##0.00")
                AnnualMaster.SubItems(5) = IIf(Len(Trim(rst2!AnnualEffDate)) >= 1, rst2!AnnualEffDate, "")
                If Len(rst2!PAYCode) Then
                    AnnualMaster.SubItems(6) = rst2!PAYCode
                End If
                If Len(rst2!Grpcompany) Then
                   AnnualMaster.SubItems(7) = rst2!Grpcompany
                End If
                If Len(rst2!AnnualLimitVersion) Then
                   AnnualMaster.SubItems(8) = rst2!AnnualLimitVersion
                End If
                If Len(rst2!SuppLimitStatus) Then
                   AnnualMaster.SubItems(9) = rst2!SuppLimitStatus
                End If
                AnnualMaster.SubItems(10) = rst2!AnnualIndex
                AnnualMaster.SubItems(11) = Format(IIf(Len(rst2!LifeTimeLimit), rst2!LifeTimeLimit, ""), "#,##0.00")
                AnnualMaster.SubItems(12) = Format(IIf(Len(rst2!SuppLifeTimeLimit), rst2!SuppLifeTimeLimit, ""), "#,##0.00")
                AnnualMaster.SubItems(13) = Format(IIf(Len(rst2!DisabilityLmt), rst2!DisabilityLmt, ""), "#,##0.00")
                rst2.MoveNext
        TotalAnnualCount = TotalAnnualCount + 1
        txtTotalRecord = TotalAnnualCount
        Loop
        
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub AnnualComboListing()

    Dim INS As New ADODB.Recordset
    Dim HLT As New ADODB.Recordset
    'Dim PLN As New ADODB.Recordset
    'Dim PAY As New ADODB.Recordset
    'Dim GrpCom As New ADODB.Recordset
    
    'PLN.Open "select PLNCode, PLNDescription  from PLN", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    INS.Open "select INSCOde, INSDescription from INS", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    HLT.Open "select HLTCOde, HLTDescription from HLT", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    'PAY.Open "select PAYCOde, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    'GrpCom.Open "select Distinct GRPCompany from PLN ", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    CboPlnCode.Clear
    CboHltCode.Clear
    CboInsType.Clear
    CboPayorCode.Clear
    CboGrpCom.Clear
    Combo1.Clear
    
    Do Until INS.EOF
        CboInsType.AddItem INS!INSCode & " - " & INS!INSDescription
        INS.MoveNext
    Loop
    Do Until HLT.EOF
        CboHltCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
        HLT.MoveNext
    Loop
    'Do Until PLN.EOF
    '    CboPlnCode.AddItem PLN!PLNCode & " - " & PLN!PLNDescription
    '    PLN.MoveNext
    'Loop
    'Do Until PAY.EOF
    '    CboPayorCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
    '    PAY.MoveNext
    'Loop
    'Do Until GrpCom.EOF
        'If Len(GrpCom!GrpCompany) Then
            'If (GRPCOM!GRPCompany) <> CboGrpCom.ListIndex Then
            'GroupBy (GRPCOM!GRPCompany)
           'CboGrpCom.AddItem GrpCom!GrpCompany
           
        'End If
        'End If
        'GrpCom.MoveNext
    'Loop
    
    Combo1.AddItem "A" & " - " & "Share Limit For Family"
    Combo1.AddItem "B" & " - " & "Share Limit For Supplementary"
    Combo1.AddItem "C" & " - " & "Separate Limit For Supplementary"
    Combo1.Text = "A" & " - " & "Share Limit For Family"
        
    'PLN.Close
    'Set PLN = Nothing
    INS.Close
    Set INS = Nothing
    HLT.Close
    Set HLT = Nothing
    'PAY.Close
    'Set PAY = Nothing
    'GrpCom.Close
    'Set GrpCom = Nothing
End Sub

Private Sub SaveALB()
Dim ALB As New ADODB.Recordset
Dim SaveALB As New ADODB.Recordset
Dim RecordSaved
Dim strsql As String
Dim LifeTime As Boolean
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset

         
    If InStr(1, cboALBPlnCode, "-") > 0 Then
        strsqlPLN = " Select PLNCode, PLNLifeTimeStatus From PLN Where PLNCode = '" & Trim(Mid(cboALBPlnCode, 1, InStr(1, cboALBPlnCode, "-") - 1)) & "'" & _
                    " AND PLNLifeTimeStatus = 'Y'"
    Else
        strsqlPLN = "Select PLNCode, PLNLifeTimeStatus From PLN Where PLNCode = '" & Trim(Mid(cboALBPlnCode, 1, 4)) & "' AND PLNLifeTimeStatus = 'Y'"
    End If
    
    PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If PLN.recordCount > 0 Then
        LifeTime = True
    Else
        LifeTime = False
    End If
    
    PLN.Close
    Set PLN = Nothing
    
    If Len(Trim(cboALBPlnCode)) = 0 Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboALBPlnCode.SetFocus
    ElseIf Len(Trim(cboALBHltCode)) = 0 Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboALBHltCode.SetFocus
    ElseIf Len(Trim(cboALBInsType)) = 0 Then
        MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboALBInsType.SetFocus
    ElseIf txtALBEffDate = "__/__/____" Then
        MsgBox "Please Enter Annual Limit Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtALBEffDate.SetFocus
    Else
        
        strsql = "Select * From BnfAnnLmt "
        strsql = strsql & " Where " & _
                " INSCode = '" & Trim(Mid(cboALBInsType, 1, 2)) & "' " & _
                " AND HLTCode = '" & Trim(Mid(cboALBHltCode, 1, 2)) & "' " & _
                " AND PLNCode = '" & Trim(Mid(cboALBPlnCode, 1, 4)) & "' "
        ALB.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
      If ALB.recordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! " & vbNewLine & _
                " Please Change The Corresponding Field", vbCritical, "Error"
        ALB.Close
        cboALBInsType.SetFocus
        cboALBInsType.SelStart = 0
        cboALBInsType.SelLength = Len(cboALBInsType)
        Set ALB = Nothing
        Exit Sub
      Else
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            Save = True
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    SaveALB.Open "BnfAnnLmt", medixmasterconn, adOpenKeyset, adLockOptimistic
                    SaveALB.AddNew
                Else
                    strsql = "Select * From BnfAnnLmt Where ALBCode = '" & ChkStr(lstALB.SelectedItem.SubItems(8)) & "'"
                    SaveALB.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                End If
            
                With SaveALB
                    If InStr(1, cboALBPlnCode, "-") > 0 Then
                        !PLNCode = Trim(Mid(cboALBPlnCode, 1, InStr(1, cboALBPlnCode, "-") - 1))
                    Else
                        !PLNCode = Trim(Mid(cboALBPlnCode, 1, 4))
                    End If
                    
                    !INSCode = Mid(cboALBInsType, 1, 2)
                    !HLTCode = Mid(cboALBHltCode, 1, 2)
                   
                    If Mid(Combo4, 1, 1) = "A" Then
                        !SuppALBStatus = "A"
                        '!SuppLimit = TxtSupplyAL
                    ElseIf Mid(Combo4, 1, 1) = "B" Then
                        !SuppALBStatus = "B"
                    ElseIf Mid(Combo4, 1, 1) = "C" Then
                        !SuppALBStatus = "C"
                    End If
                                        
                    !ALBEffDate = IIf(txtALBEffDate = "__/__/____", "", CDate(txtALBEffDate))
                    !KidDialysisLmt = txtALBKidDiaLmt
                    !CancerLmt = txtALBCancerLmt
                    !HomeNursLmt = txtALBHomeNursLmt
                    !ALBLastUpdateUser = Logon
                    !ALBLastUpdateDate = Date
                End With
                SaveALB.Update
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
 
                SaveALB.Close
                Set SaveALB = Nothing
                Call ALBListing
                Call ClearForm(Me)
            End If
        EditFlag = False
        Save = False
   End If
  End If
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub SaveProCat()
Dim SaveProCat As New ADODB.Recordset
Dim strProCat
Dim strsql As String

    strsql = "Select ProductCode, ProductName From ProductCategory where ProductCode = '" & Trim(txtProCode) & "'"
    SaveProCat.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    If SaveProCat.recordCount > 0 And EditFlag = False Then
        MsgBox "Duplicate Data! Please select another Product Code.", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    ElseIf SaveProCat.recordCount > 0 And EditFlag = True Then
        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
    ElseIf SaveProCat.recordCount = 0 And EditFlag = False Then
        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
        SaveProCat.AddNew
    End If
    
        With SaveProCat
            !ProductCode = txtProCode
            !ProductName = txtProName
        End With
        SaveProCat.Update
        SaveProCat.Close
        Set SaveProCat = Nothing
        Call ProCatListing
        Call ClearForm(Me)
        EditFlag = False
End Sub

Private Sub DeleteALB()

Dim DeleteALB As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord
Dim List As String

EditFlag = False
If lstALB.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
    DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes Then
        strsql = "Delete From BnfAnnLmt Where ALBCode = '" & ChkStr(lstALB.SelectedItem.SubItems(8)) & "'"
        DeleteALB.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
            CmdSave.Enabled = True
            Call ClearForm(Me)
            Call ALBListing
            Call DisableCmdButton(Me)
            Call EntriesEnable(True)
            cboALBPlnCode.SetFocus
    End If
End If
Exit Sub
End Sub

Private Sub DeleteProCat()
Dim DeleteProCat As New ADODB.Recordset
Dim strsql As String
Dim DeleteRecord

    EditFlag = False
    If lstProductCat.ListItems.Count = 0 Then
        MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
            strsql = "Delete From ProductCategory Where ProductCode = '" & ChkStr(lstProductCat.SelectedItem.Text) & "'"
            DeleteProCat.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
                CmdSave.Enabled = True
                Call ClearForm(Me)
                Call ProCatListing
                Call DisableCmdButton(Me)
                Call EntriesEnable(True)
         End If
    End If
End Sub

Private Sub SearchALB()
    Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(cboALBPlnCode)) Then
        strAppend = strAppend + " And PLNCode = '" & Trim(Mid(cboALBPlnCode, 1, IIf(InStr(1, cboALBPlnCode, "-") = 0, 4, InStr(1, cboALBPlnCode, "-") - 1))) & "'"
    End If
    
    If Len(Trim(cboALBHltCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Trim(Mid(cboALBHltCode, 1, IIf(InStr(1, cboALBHltCode, "-") = 0, 4, InStr(1, cboALBHltCode, "-") - 1))) & "'"
    End If
    
    If Len(Trim(cboALBInsType)) Then
        strAppend = strAppend + " And INSCode = '" & Trim(Mid(cboALBInsType, 1, IIf(InStr(1, cboALBInsType, "-") = 0, 4, InStr(1, cboALBInsType, "-") - 1))) & "'"
    End If
    
    ALBCriteria = strAppend
    Call ALBListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub SearchProCat()
Dim strAppend As String
    
    lstProductCat.ListItems.Clear
    strAppend = ""
    If Len(txtProCode) Then
        strAppend = strAppend + " AND ProductCode = '" & Trim(txtProCode) & "'"
    End If
    
    ProCatCriteria = strAppend
    Call ProCatListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub ALBListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim AnnualMaster As ListItem
    Dim sql As String
    Dim TotalAnnualCount As Integer
    
    TotalAnnualCount = 0
    lstALB.ListItems.Clear
    
    sql = " Select ALBCode, PLNCode, INSCode, HLTCode, " & _
          " ALBEffDate, SuppALBStatus, KidDialysisLmt, CancerLmt, " & _
          " HomeNursLmt From VIEW_ANN_LMT_BNF Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
    sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set AnnualMaster = lstALB.ListItems.Add(, , rst2!HLTCode)
                AnnualMaster.SubItems(1) = rst2!PLNCode
                AnnualMaster.SubItems(2) = rst2!INSCode
                AnnualMaster.SubItems(3) = IIf(Len(Trim(rst2!ALBEffDate)) >= 1, rst2!ALBEffDate, "")
                                
                If Len(rst2!SuppALBStatus) Then
                   AnnualMaster.SubItems(7) = rst2!SuppALBStatus
                End If
                
                AnnualMaster.SubItems(8) = rst2!ALBCode
                AnnualMaster.SubItems(4) = rst2!KidDialysisLmt
                AnnualMaster.SubItems(5) = rst2!CancerLmt
                AnnualMaster.SubItems(6) = rst2!HomeNursLmt
                                
                rst2.MoveNext
        TotalAnnualCount = TotalAnnualCount + 1
        txtTotalRecord = TotalAnnualCount
        Loop
        
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub ALBComboListing()

    Dim INS As New ADODB.Recordset
    Dim HLT As New ADODB.Recordset
    Dim PLN As New ADODB.Recordset
        
    PLN.Open "select PLNCode, PLNDescription  from PLN", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    INS.Open "select INSCOde, INSDescription from INS", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    HLT.Open "select HLTCOde, HLTDescription from HLT", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    cboALBPlnCode.Clear
    cboALBHltCode.Clear
    cboALBInsType.Clear
    Combo4.Clear
    
    Do Until INS.EOF
        cboALBInsType.AddItem INS!INSCode & " - " & INS!INSDescription
        INS.MoveNext
    Loop
    Do Until HLT.EOF
        cboALBHltCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
        HLT.MoveNext
    Loop
    'Do Until PLN.EOF
    '    cboALBPlnCode.AddItem PLN!PLNCode & " - " & PLN!PLNDescription
    '    PLN.MoveNext
    'Loop
        
    Combo4.AddItem "A" & " - " & "Share Limit For Family"
    Combo4.AddItem "B" & " - " & "Share Limit For Supplementary"
    Combo4.AddItem "C" & " - " & "Separate Limit For Supplementary"
    Combo4.Text = "A" & " - " & "Share Limit For Family"
        
    PLN.Close
    Set PLN = Nothing
    INS.Close
    Set INS = Nothing
    HLT.Close
    Set HLT = Nothing
End Sub

Private Sub ProCatListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim lstProCat As ListItem
Dim strsql As String

    lstProductCat.ListItems.Clear
    strsql = "Select ProductCode, ProductName from ProductCategory where 0 = 0"
   
    If Len(strAppend) Then
        strsql = strsql + strAppend
    Else
        strsql = strsql
    End If
    
    rst2.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    If rst2.EOF Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set lstProCat = lstProductCat.ListItems.Add(, , rst2!ProductCode)
                lstProCat.SubItems(1) = rst2!ProductName
            rst2.MoveNext
        Loop
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub SaveMCOStatus()
Dim SaveMCOStat As New ADODB.Recordset
Dim strsql As String
Dim tmpCode As String

    strsql = "Select MCOCode, MCOStatus From SuppMCOStatus where MCOCode = '" & Trim(txtMCOCode) & "'"
    SaveMCOStat.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    If txtMCOStatus = "" Then
        MsgBox "Please enter MCO Status.", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    
    If SaveMCOStat.recordCount > 0 And EditFlag = False Then
        MsgBox "Duplicate Data! Please select another Product Code.", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    ElseIf SaveMCOStat.recordCount > 0 And EditFlag = True Then
        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
    ElseIf SaveMCOStat.recordCount = 0 And EditFlag = False Then
        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
        SaveMCOStat.AddNew
    End If
    
        With SaveMCOStat
            !MCOCode = txtMCOCode
            !MCOStatus = txtMCOStatus
        End With
        SaveMCOStat.Update
        SaveMCOStat.Close
        Set SaveMCOStat = Nothing
        Call MCOStatusListing
        Call ClearForm(Me)
        EditFlag = False
End Sub

Private Sub DeleteMCOStatus()
Dim DeleteMCOStatus As New ADODB.Recordset
Dim strsql As String
Dim DeleteRecord

    EditFlag = False
    CmdSave.Enabled = False
    If lstMCOStatus.ListItems.Count = 0 Then
        MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
            strsql = "Delete From SuppMCOStatus Where MCOCode = '" & ChkStr(lstMCOStatus.SelectedItem.Text) & "'"
            DeleteMCOStatus.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
                CmdSave.Enabled = True
                Call ClearForm(Me)
                Call MCOStatusListing
                Call DisableCmdButton(Me)
                Call EntriesEnable(True)
         End If
    End If
End Sub

Private Sub MCOStatusListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim lstMCOStat As ListItem
Dim strsql As String

    lstMCOStatus.ListItems.Clear
    strsql = "Select MCOCode, MCOStatus from SuppMCOStatus where 0 = 0"
   
    If Len(strAppend) Then
        strsql = strsql + strAppend
    Else
        strsql = strsql
    End If
    
    rst2.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    If rst2.EOF Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set lstMCOStat = lstMCOStatus.ListItems.Add(, , rst2!MCOCode)
                lstMCOStat.SubItems(1) = rst2!MCOStatus
            rst2.MoveNext
        Loop
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub SearchMCOStat()
    Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(txtMCOCode)) Then
        strAppend = strAppend + " And MCOCOde = '" & Trim(txtMCOCode) & "'"
    End If
        
    MCOStatCriteria = strAppend
    Call MCOStatusListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstProductCat_Click()
Dim strsql As String
    
    If lstProductCat.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtProCode = lstProductCat.SelectedItem.Text
        txtProCode.Enabled = False
        txtProName = lstProductCat.SelectedItem.SubItems(1)
        txtProName.Enabled = False
            
        'CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If

End Sub

Private Sub lstMCOStatus_Click()
Dim strsql As String
    
    CmdSave.Enabled = False
    If lstMCOStatus.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtMCOCode = lstMCOStatus.SelectedItem.Text
        txtMCOCode.Enabled = False
        txtMCOStatus = lstMCOStatus.SelectedItem.SubItems(1)
        txtMCOStatus.Enabled = False
    End If

End Sub

Private Sub lstAnnualList_lostfocus()
    chkPLN.Enabled = True
    chkPayor.Enabled = True
    chkGrpCom.Enabled = True
End Sub

Private Sub lstAnnualList_Click()
    
    chkPLN.Enabled = False
    chkPayor.Enabled = False
    chkGrpCom.Enabled = False
    If lstAnnualList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboHltCode = lstAnnualList.SelectedItem.Text
        CboPlnCode = lstAnnualList.SelectedItem.SubItems(1)
        CboInsType = lstAnnualList.SelectedItem.SubItems(2)
        TxtAnnualLimit = lstAnnualList.SelectedItem.SubItems(3)
        TxtSupplyAL = lstAnnualList.SelectedItem.SubItems(4)
        txtALEffDate = lstAnnualList.SelectedItem.SubItems(5)
        CboPayorCode = lstAnnualList.SelectedItem.SubItems(6)
        CboGrpCom = lstAnnualList.SelectedItem.SubItems(7)
        txtANNVersion = lstAnnualList.SelectedItem.SubItems(8)
        TxtLifeTimeLimit = lstAnnualList.SelectedItem.SubItems(11)
        TxtSuppLifeTimeLimit = lstAnnualList.SelectedItem.SubItems(12)
        txtDisabilityLmt = lstAnnualList.SelectedItem.SubItems(13)
        If lstAnnualList.SelectedItem.SubItems(9) = "A" Then
            Combo1.ListIndex = 0
        ElseIf lstAnnualList.SelectedItem.SubItems(9) = "B" Then
            Combo1.ListIndex = 1
        ElseIf lstAnnualList.SelectedItem.SubItems(9) = "C" Then
            Combo1.ListIndex = 2
        End If
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstAnnualList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstAnnualList, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstAnnualList, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView lstAnnualList, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    End Select
End Sub

Private Sub TxtIAAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtICAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtFAAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtFCAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSFAAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSFCAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtGAAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtGCAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtHAAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtHCAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSHAAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSHCAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtCoverID_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub txtEffDate_LostFocus()
    If IsDate(txtEffDate) Then
    Else
        If txtEffDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtEffDate.SetFocus
        End If
    End If
End Sub

Private Sub PLNDisEnable()
Dim AA As Integer
    For AA = 0 To txtPLN1.UBound
        txtPLN1(AA).Enabled = False
    Next AA
End Sub
Private Sub UpdatePlan()
Dim SaveMCOStat As New ADODB.Recordset
Dim strsql As String
Dim tmpCode As String

    strsql = "Update HISDB..pln set PLNLGPrivate='" & txtLGPrivate & "' where plncode='" & ChkStr(txtPLN1(0)) & "'"
    SaveMCOStat.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    MsgBox "GL no updated Successfully", vbOKOnly + vbInformation, DialogBoxTitle

End Sub

Private Sub cboPLNAnnLimit_Change()
    If InStr(cboPLNAnnLimit.Text, "null") Then
       cboPLNAnnLimit.Enabled = False
    End If
End Sub

Private Sub cboPLNAnnLimit_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNAnnLimit.Text, cboPLNAnnLimit.SelStart) + Chr(KeyAscii) + Mid(cboPLNAnnLimit.Text, cboPLNAnnLimit.SelStart + cboPLNAnnLimit.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNAnnLimit.ListCount - 1
 
        If NewText = LCase(Left(cboPLNAnnLimit.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNAnnLimit.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPLNAnnLimit.Text = ValidValue
               cboPLNAnnLimit.SelStart = 0
               cboPLNAnnLimit.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNPremium_Change()
    If InStr(cboPLNPremium.Text, "null") Then
       cboPLNPremium.Enabled = False
    End If
End Sub

Private Sub cboPLNPremium_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNPremium.Text, cboPLNPremium.SelStart) + Chr(KeyAscii) + Mid(cboPLNPremium.Text, cboPLNPremium.SelStart + cboPLNPremium.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNPremium.ListCount - 1
 
        If NewText = LCase(Left(cboPLNPremium.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNPremium.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPLNPremium.Text = ValidValue
               cboPLNPremium.SelStart = 0
               cboPLNPremium.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNMCO_Change()
    If InStr(cboPLNMCO.Text, "null") Then
       cboPLNMCO.Enabled = False
    End If
End Sub

Private Sub cboPLNMCO_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNMCO.Text, cboPLNMCO.SelStart) + Chr(KeyAscii) + Mid(cboPLNMCO.Text, cboPLNMCO.SelStart + cboPLNMCO.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNMCO.ListCount - 1
 
        If NewText = LCase(Left(cboPLNMCO.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNMCO.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboPLNMCO.Text = ValidValue
               cboPLNMCO.SelStart = 0
               cboPLNMCO.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboLifeTimeStatus_Change()
    If InStr(CboLifeTimeStatus.Text, "null") Then
       CboLifeTimeStatus.Enabled = False
    End If
End Sub

Private Sub CboLifeTimeStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboLifeTimeStatus.Text, CboLifeTimeStatus.SelStart) + Chr(KeyAscii) + Mid(CboLifeTimeStatus.Text, CboLifeTimeStatus.SelStart + CboLifeTimeStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboLifeTimeStatus.ListCount - 1
 
        If NewText = LCase(Left(CboLifeTimeStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboLifeTimeStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboLifeTimeStatus.Text = ValidValue
               CboLifeTimeStatus.SelStart = 0
               CboLifeTimeStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboSpecGPeriod_Click()
    If InStr(CboSpecGPeriod.Text, "null") Then
       CboSpecGPeriod.Enabled = False
    End If
    
    If Mid(CboSpecGPeriod, 1, 1) = "Y" Then
        TxtSpecGPeriod.Enabled = True
        TxtSpecGPeriod.SetFocus
    Else
        TxtSpecGPeriod.Text = ""
        TxtSpecGPeriod.Enabled = False
    End If
    
End Sub

Private Sub CboSpecGPeriod_KeyPress(KeyAscii As Integer)

Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSpecGPeriod.Text, CboSpecGPeriod.SelStart) + Chr(KeyAscii) + Mid(CboSpecGPeriod.Text, CboSpecGPeriod.SelStart + CboSpecGPeriod.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSpecGPeriod.ListCount - 1
 
        If NewText = LCase(Left(CboSpecGPeriod.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSpecGPeriod.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSpecGPeriod.Text = ValidValue
               CboSpecGPeriod.SelStart = 0
               CboSpecGPeriod.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub SpecialListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim SpecialMaster As ListItem
Dim sql As String
Dim TotalSpecialCount As Integer
    
    
    TotalSpecialCount = 0
    lstSpecialNoteList.ListItems.Clear
    
    sql = "Select distinct PLNCode, GrpCompany, PLNSpecialNote From PLN Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
    sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set SpecialMaster = lstSpecialNoteList.ListItems.Add(, , rst2!PLNCode)
                If Len(rst2!Grpcompany) Then
                    SpecialMaster.SubItems(1) = Trim(rst2!Grpcompany)
                End If
                If Len(rst2!PLNSpecialNote) Then
                    SpecialMaster.SubItems(2) = Trim(rst2!PLNSpecialNote)
                End If
                                
                rst2.MoveNext
        TotalSpecialCount = TotalSpecialCount + 1
        txtTotalRecord = TotalSpecialCount
        Loop
        
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub DeleteSpecialNote()
Dim rsDeletePLN As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim strsql As String
Dim sqlMBM As String
Dim DeleteRecord

EditFlag = False
    
    If CboSpecialPlan <> "" Then
    
      If lstSpecialNoteList.ListItems.Count = 0 Then
         MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
      Else
          sqlMBM = "Select MBMNumber From MBM where PLNCode = '" & RemStr(CboSpecialPlan) & "'"
          
          rsCheckMBM.MaxRecords = 1
          rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
          
          If rsCheckMBM.recordCount Then
              MsgBox "This record cannot be deleted because it is used in Membership table!", vbCritical, "Error in deleting record..."
              Call ClearForm(Me)
              CmdDelete.Enabled = False
              CmdEdit.Enabled = False
          Else
              DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
          
              If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
                   strsql = "Delete From PLN Where PLNCode = '" & ChkStr(lstSpecialNoteList.SelectedItem.Text) & "'"
                   rsDeletePLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                   MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
                   Call ClearForm(Me)
                   Call SpecialListing(SpecialCriteria)
                   Call DisableCmdButton(Me)
                   Call EntriesEnable(True)
                   CmdAdd.Enabled = True
                   CmdSave.Enabled = True
              Else
                   CmdDelete.Enabled = False
                   CmdEdit.Enabled = False
              End If
          End If
          Exit Sub
        End If
    End If
End Sub

Private Sub SaveSpecialNote()
Dim RecordSaved
Dim SPENote As New ADODB.Recordset
Dim SaveSPENote As New ADODB.Recordset
Dim strsql As String
   
    If CboSpecialPlan = "" Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        'CboSpecialPlan.SetFocus
    'ElseIf CboSpecialGrpCom = "" Then
        'MsgBox "Please Enter Group Company Name", vbOKOnly + vbCritical, DialogBoxTitle
        'CboSpecialGrpCom.SetFocus
    'ElseIf TxtSpecialNote = "" Then
        'MsgBox "Please Enter Special Note", vbOKOnly + vbCritical, DialogBoxTitle
        'TxtSpecialNote.SetFocus
    Else
        'strsql = " Select * From PLN where PLNCode =  '" & Trim(CboSpecialPlan) & "'" & _
                 " And GRPCompany = '" & Trim(CboSpecialGrpCom) & "'"
        'SPENote.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'If SPENote.recordCount > 0 And EditFlag = False Then
           'MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           'SPENote.Close
           'CboSpecialPlan.SetFocus
           'CboSpecialPlan.SelStart = 0
           'CboSpecialPlan.SelLength = Len(CboSpecialPlan)
           'Set SPENote = Nothing
           'Exit Sub
        'Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              'If EditFlag = False Then
                 'SaveSPENote.Open "PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
                 'SaveSPENote.AddNew
              'Else
              '" & Trim(Mid(CboPlnCode, 1, InStr(1, CboPlnCode, "-") - 1)) & "'"
        
                 strsql = " Select * From PLN Where PLNCode =  '" & Trim(CboSpecialPlan) & "'" '&
                          '" And GRPCompany = '" & Trim(RemStr(CboSpecialGrpCom)) & "'"
                 SaveSPENote.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
              'End If
              
              With SaveSPENote
                !PLNCode = Trim(CboSpecialPlan)
                !Grpcompany = Trim(RemStr(CboSpecialGrpCom))
                !PLNSpecialNote = Trim(ChkStr(TxtSpecialNote))
                
              End With
              SaveSPENote.Update
            
             'Call ClearForm(Me)
             Call SpecialListing(SpecialCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             'SPENote.Close
             Set SaveSPENote = Nothing
        'End If
        EditFlag = False
     End If
   End If
End Sub

Private Sub SearchSpecialNote()
Dim strAppend As String
        
    Screen.MousePointer = vbHourglass
    
    strAppend = ""
        
    If Len(Trim(CboSpecialPlan)) Then
        strAppend = strAppend + " And PLNCode = '" & Trim(CboSpecialPlan) & "'"
    End If
        
    If Len(Trim(CboSpecialGrpCom)) Then
        strAppend = strAppend + " And GRPCompany = '" & Trim(RemStr(CboSpecialGrpCom)) & "'"
    End If
        
    SpecialCriteria = strAppend
    Call SpecialListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub CboSpecialPlan_Change()
    If InStr(CboSpecialPlan.Text, "null") Then
       CboSpecialPlan.Enabled = False
    End If
End Sub


Private Sub CboSpecialPlan_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSpecialPlan.Text, CboSpecialPlan.SelStart) + Chr(KeyAscii) + Mid(CboSpecialPlan.Text, CboSpecialPlan.SelStart + CboSpecialPlan.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSpecialPlan.ListCount - 1
 
        If NewText = LCase(Left(CboSpecialPlan.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSpecialPlan.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSpecialPlan.Text = ValidValue
               CboSpecialPlan.SelStart = 0
               CboSpecialPlan.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboSpecialGrpCom_Change()
    If InStr(CboSpecialGrpCom.Text, "null") Then
       CboSpecialGrpCom.Enabled = False
    End If
End Sub

Private Sub CboSpecialGrpCom_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSpecialGrpCom.Text, CboSpecialGrpCom.SelStart) + Chr(KeyAscii) + Mid(CboSpecialGrpCom.Text, CboSpecialGrpCom.SelStart + CboSpecialGrpCom.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboSpecialGrpCom.ListCount - 1
 
        If NewText = LCase(Left(CboSpecialGrpCom.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSpecialGrpCom.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboSpecialGrpCom.Text = ValidValue
               CboSpecialGrpCom.SelStart = 0
               CboSpecialGrpCom.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub lstSpecialNoteList_Click()

    If lstSpecialNoteList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboSpecialPlan = lstSpecialNoteList.SelectedItem.Text
        CboSpecialGrpCom = lstSpecialNoteList.SelectedItem.SubItems(1)
        TxtSpecialNote = lstSpecialNoteList.SelectedItem.SubItems(2)
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstSpecialNoteList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstSpecialNoteList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstSpecialNoteList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstSpecialNoteList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    End Select
End Sub

Private Sub CboNoofPlan_Change()
    If InStr(CboNoofPlan.Text, "null") Then
       CboNoofPlan.Enabled = False
    End If
End Sub

Private Sub CboNoofPlan_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboNoofPlan.Text, CboNoofPlan.SelStart) + Chr(KeyAscii) + Mid(CboNoofPlan.Text, CboNoofPlan.SelStart + CboNoofPlan.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboNoofPlan.ListCount - 1
 
        If NewText = LCase(Left(CboNoofPlan.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboNoofPlan.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboNoofPlan.Text = ValidValue
               CboNoofPlan.SelStart = 0
               CboNoofPlan.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub MgmtFeeCombolisting()
Dim INS As New ADODB.Recordset
Dim HLT As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
        
    PLN.Open "select PLNCode, PLNDescription  from PLN Where PLNMgmtFee = 'Y'", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    INS.Open "select INSCOde, INSDescription from INS", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    HLT.Open "select HLTCOde, HLTDescription from HLT", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    CboMgmtHltCode.Clear
    CboMgmtPlan.Clear
    CboMgmtIns.Clear
        
    Do Until INS.EOF
        CboMgmtIns.AddItem INS!INSCode & " - " & INS!INSDescription
        INS.MoveNext
    Loop
    Do Until HLT.EOF
        CboMgmtHltCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
        HLT.MoveNext
    Loop
    'Do Until PLN.EOF
    '    CboMgmtPlan.AddItem PLN!PLNCode & " - " & PLN!PLNDescription
    '    PLN.MoveNext
    'Loop
            
    PLN.Close
    Set PLN = Nothing
    INS.Close
    Set INS = Nothing
    HLT.Close
    Set HLT = Nothing
    
End Sub

Private Sub SearchMgmtFee()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(CboMgmtPlan)) Then
        strAppend = strAppend + " And MgmtPLNCode = '" & Trim(Mid(CboMgmtPlan, 1, IIf(InStr(1, CboMgmtPlan, "-") = 0, 4, InStr(1, CboMgmtPlan, "-") - 1))) & "'"
    End If
    If Len(Trim(CboMgmtHltCode)) Then
        strAppend = strAppend + " And MgmtHLTCode = '" & Trim(Mid(CboMgmtHltCode, 1, IIf(InStr(1, CboMgmtHltCode, "-") = 0, 4, InStr(1, CboMgmtHltCode, "-") - 1))) & "'"
    End If
    If Len(Trim(CboMgmtIns)) Then
        strAppend = strAppend + " And MgmtINSCode = '" & Trim(Mid(CboMgmtIns, 1, IIf(InStr(1, CboMgmtIns, "-") = 0, 4, InStr(1, CboMgmtIns, "-") - 1))) & "'"
    End If
    
        
    Call MgmtFeeListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub MgmtFeeListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim MgmtFeeMaster As ListItem
Dim sql As String
Dim TotalMgmtFeeCount As Integer
    
    
    TotalMgmtFeeCount = 0
    lstMgmtFeeList.ListItems.Clear
    
    sql = " Select * From CaseMgmtFee Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
    sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set MgmtFeeMaster = lstMgmtFeeList.ListItems.Add(, , rst2!MgmtHLTCode)
                MgmtFeeMaster.SubItems(1) = rst2!MgmtPLNCode
                MgmtFeeMaster.SubItems(2) = rst2!MgmtINSCode
                If Len(rst2!MgmtFee) Then
                    MgmtFeeMaster.SubItems(3) = Format(rst2!MgmtFee, "#,##0.00")
                Else
                    MgmtFeeMaster.SubItems(3) = "0.00"
                End If
                
                rst2.MoveNext
                
            TotalMgmtFeeCount = TotalMgmtFeeCount + 1
            txtTotalRecord = TotalMgmtFeeCount
        Loop
        
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub SaveMgmtFee()
Dim MgmtFee As New ADODB.Recordset
Dim SaveMgmtFee As New ADODB.Recordset
Dim RecordSaved
Dim strsql As String

    
    If Len(Trim(CboMgmtPlan)) = 0 Then
        MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboPlnCode.SetFocus
    ElseIf Len(Trim(CboMgmtHltCode)) = 0 Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboHltCode.SetFocus
    ElseIf Len(Trim(CboMgmtIns)) = 0 Then
        MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboInsType.SetFocus
    ElseIf Len(Trim(TxtMgmtFee)) = 0 Then
        MsgBox "Please Enter Mgmt. Fee", vbOKOnly + vbCritical, DialogBoxTitle
        TxtAnnualLimit.SetFocus
    Else
        
        strsql = "Select * From CaseMgmtFee "
        strsql = strsql & " Where " & _
                " MgmtINSCode = '" & Trim(Mid(CboMgmtIns, 1, 2)) & "' " & _
                " AND MgmtHLTCode = '" & Trim(Mid(CboMgmtHltCode, 1, 2)) & "' " & _
                " AND MgmtPLNCode = '" & Trim(Mid(CboMgmtPlan, 1, 4)) & "' "
                
        MgmtFee.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
      If MgmtFee.recordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! " & vbNewLine & _
                " Please Change The Corresponding Field", vbCritical, "Error"
        MgmtFee.Close
        CboInsType.SetFocus
        CboInsType.SelStart = 0
        CboInsType.SelLength = Len(CboInsType)
        Set MgmtFee = Nothing
        Exit Sub
      Else
        Save = True
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    SaveMgmtFee.Open "CaseMgmtFee", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                    SaveMgmtFee.AddNew
                Else
                    strsql = " Select * From CaseMgmtFee Where MgmtINSCode = '" & Trim(Mid(CboMgmtIns, 1, 2)) & "' " & _
                             " AND MgmtHLTCode = '" & Trim(Mid(CboMgmtHltCode, 1, 2)) & "' " & _
                             " AND MgmtPLNCode = '" & Trim(Mid(CboMgmtPlan, 1, 4)) & "' "
                    SaveMgmtFee.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                End If
            
                With SaveMgmtFee
                    If InStr(1, CboMgmtPlan, "-") > 0 Then
                        !MgmtPLNCode = Trim(Mid(CboMgmtPlan, 1, InStr(1, CboMgmtPlan, "-") - 1))
                    Else
                        !MgmtPLNCode = Trim(Mid(CboMgmtPlan, 1, 4))
                    End If
                    
                    !MgmtINSCode = Mid(CboMgmtIns, 1, 2)
                    !MgmtHLTCode = Mid(CboMgmtHltCode, 1, 2)
                    
                    If TxtMgmtFee <> "" Then
                        !MgmtFee = TxtMgmtFee
                    End If
                    
                    !MgmtLastUpdateUser = Logon
                    !MgmtLastUpdateDate = Date
                End With
                
                SaveMgmtFee.Update
                
                Call ClearForm(Me)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
 
                SaveMgmtFee.Close
                Set SaveMgmtFee = Nothing
                Call ClearForm(Me)
                Call MgmtFeeListing
                Call DisableCmdButton(Me)
                Call EntriesEnable(True)
            End If
        EditFlag = False
        Save = False
   End If
  End If
  
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub DeleteCaseMgmtFee()
Dim DeleteMgmtFee As New ADODB.Recordset
Dim strsql As String
Dim DeleteRecord

    EditFlag = False
    If lstMgmtFeeList.ListItems.Count = 0 Then
        MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
            strsql = " Delete From CaseMgmtFee Where MgmtINSCode = '" & Trim(Mid(CboMgmtIns, 1, 2)) & "'" & _
                     " AND MgmtHLTCode = '" & Trim(Mid(CboMgmtHltCode, 1, 2)) & "'" & _
                     " AND MgmtPLNCode = '" & Trim(Mid(CboMgmtPlan, 1, 4)) & "'"
            DeleteMgmtFee.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
                CmdSave.Enabled = True
                Call ClearForm(Me)
                Call MgmtFeeListing
                Call DisableCmdButton(Me)
                Call EntriesEnable(True)
         End If
    End If
End Sub

Private Sub lstMgmtFeeList_Click()

    If lstMgmtFeeList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboMgmtHltCode = lstMgmtFeeList.SelectedItem.Text
        CboMgmtPlan = lstMgmtFeeList.SelectedItem.SubItems(1)
        CboMgmtIns = lstMgmtFeeList.SelectedItem.SubItems(2)
        TxtMgmtFee = lstMgmtFeeList.SelectedItem.SubItems(3)
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstMgmtFeeList_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstMgmtFeeList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboMgmtHltCode = lstMgmtFeeList.SelectedItem.Text
        CboMgmtPlan = lstMgmtFeeList.SelectedItem.SubItems(1)
        CboMgmtIns = lstMgmtFeeList.SelectedItem.SubItems(2)
        TxtMgmtFee = lstMgmtFeeList.SelectedItem.SubItems(3)
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstMgmtFeeList_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstMgmtFeeList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        CboMgmtHltCode = lstMgmtFeeList.SelectedItem.Text
        CboMgmtPlan = lstMgmtFeeList.SelectedItem.SubItems(1)
        CboMgmtIns = lstMgmtFeeList.SelectedItem.SubItems(2)
        TxtMgmtFee = lstMgmtFeeList.SelectedItem.SubItems(3)
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstMgmtFeeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstMgmtFeeList.SortKey = ColumnHeader.Index - 1
    lstMgmtFeeList.Sorted = True
    If lstMgmtFeeList.SortOrder = lvwAscending Then
       lstMgmtFeeList.SortOrder = lvwDescending
    Else
       lstMgmtFeeList.SortOrder = lvwAscending
    End If
    
End Sub

Private Sub CboMgmtHltCode_Change()
    If InStr(CboMgmtHltCode.Text, "null") Then
       CboMgmtHltCode.Enabled = False
    End If
End Sub

Private Sub CboMgmtHltCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboMgmtHltCode.Text, CboMgmtHltCode.SelStart) + Chr(KeyAscii) + Mid(CboMgmtHltCode.Text, CboMgmtHltCode.SelStart + CboMgmtHltCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboMgmtHltCode.ListCount - 1
 
        If NewText = LCase(Left(CboMgmtHltCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboMgmtHltCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboMgmtHltCode.Text = ValidValue
               CboMgmtHltCode.SelStart = 0
               CboMgmtHltCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboMgmtPlan_Change()
    If InStr(CboMgmtPlan.Text, "null") Then
       CboMgmtPlan.Enabled = False
    End If
End Sub

Private Sub CboMgmtPlan_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboMgmtPlan.Text, CboMgmtPlan.SelStart) + Chr(KeyAscii) + Mid(CboMgmtPlan.Text, CboMgmtPlan.SelStart + CboMgmtPlan.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboMgmtPlan.ListCount - 1
 
        If NewText = LCase(Left(CboMgmtPlan.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboMgmtPlan.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboMgmtPlan.Text = ValidValue
               CboMgmtPlan.SelStart = 0
               CboMgmtPlan.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboMgmtIns_Change()
    If InStr(CboMgmtIns.Text, "null") Then
       CboMgmtIns.Enabled = False
    End If
End Sub

Private Sub CboMgmtIns_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboMgmtIns.Text, CboMgmtIns.SelStart) + Chr(KeyAscii) + Mid(CboMgmtIns.Text, CboMgmtIns.SelStart + CboMgmtIns.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboMgmtIns.ListCount - 1
 
        If NewText = LCase(Left(CboMgmtIns.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboMgmtIns.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboMgmtIns.Text = ValidValue
               CboMgmtIns.SelStart = 0
               CboMgmtIns.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub TxtMgmtFee_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub lstPlanListShow()
Dim AA As Integer
    If lstPlanList.ListItems.Count <> 0 Then
        For AA = 0 To txtPLN1.UBound
            txtPLN1(AA) = ""
        Next AA
        For AA = 0 To txtDesc1.UBound
            txtDesc1(AA) = ""
        Next AA
        cboPLNPayor = ""
        TxtGrpCom = ""
        CboTopupStatus = ""
        cboPLNAnnLimit = ""
        cboPLNPremium = ""
        cboPLNMCO = ""
        CboCoPayment = ""
        txtCoPayEffdate = "__/__/____"
        'CboLGInd = ""
        CboSOF = ""
        CboSpecGPeriod = ""
        TxtSpecGPeriod = ""
        CboSMPlan = ""
        CboLifeTimeStatus = ""
        cboProCat = ""
        CboMeal = ""
        CboNursing = ""
        CboTax = ""
        CboMRI = ""
        CboPLNMgmtStatus = ""
        'cboPolicyWording = ""
        TxtPolicyWording.Text = ""
        
        cboPLNHLTCode = lstPlanList.SelectedItem.Text
        txtPLN1(0) = lstPlanList.SelectedItem.SubItems(1)
        txtDesc1(0) = lstPlanList.SelectedItem.SubItems(3)
        cboPLNPayor = lstPlanList.SelectedItem.SubItems(4)
        TxtGrpCom = lstPlanList.SelectedItem.SubItems(5)
        CboTopupStatus = lstPlanList.SelectedItem.SubItems(6)
        cboPLNAnnLimit = lstPlanList.SelectedItem.SubItems(8)
        cboPLNPremium = lstPlanList.SelectedItem.SubItems(9)
        cboPLNMCO = lstPlanList.SelectedItem.SubItems(10)
        CboCoPayment = lstPlanList.SelectedItem.SubItems(11)
        If Len(lstPlanList.SelectedItem.SubItems(12)) Then
            txtCoPayEffdate = lstPlanList.SelectedItem.SubItems(12)
        End If
        'CboLGInd = lstPlanList.SelectedItem.SubItems(13)
        CboSpecGPeriod = lstPlanList.SelectedItem.SubItems(13)
        TxtSpecGPeriod = lstPlanList.SelectedItem.SubItems(14)
        CboSMPlan = lstPlanList.SelectedItem.SubItems(15)
        CboLifeTimeStatus = lstPlanList.SelectedItem.SubItems(16)
        CboSOF = lstPlanList.SelectedItem.SubItems(17)
        cboProCat = lstPlanList.SelectedItem.SubItems(18)
        CboMeal = lstPlanList.SelectedItem.SubItems(19)
        CboNursing = lstPlanList.SelectedItem.SubItems(20)
        CboTax = lstPlanList.SelectedItem.SubItems(21)
        CboMRI = lstPlanList.SelectedItem.SubItems(22)
        CboPLNMgmtStatus = lstPlanList.SelectedItem.SubItems(23)
        'cboPolicyWording = lstPlanList.SelectedItem.SubItems(25)
        TxtPolicyWording.Text = lstPlanList.SelectedItem.SubItems(25)
        'Change combobox to textbox - By CHEng 2 Dec 2016
        cboClientPLN = lstPlanList.SelectedItem.SubItems(26)
        txtClientPolNo = lstPlanList.SelectedItem.SubItems(27)
        If IsNumeric(lstPlanList.SelectedItem.SubItems(28)) Then
            txtLGPrivate = lstPlanList.SelectedItem.SubItems(28)
        Else
            txtLGPrivate = ""
        End If
        If Len(lstPlanList.SelectedItem.SubItems(29)) Then
            cmbVIPInd = lstPlanList.SelectedItem.SubItems(29)
        End If
        cboPLNHLTCode.Enabled = False
        txtPLN1(0).Enabled = False
        txtDesc1(0).Enabled = False
        cboPLNPayor.Enabled = False
        TxtGrpCom.Enabled = False
        CboTopupStatus.Enabled = False
        CboCoPayment.Enabled = False
        'CboLGInd.Enabled = False
        CboSOF.Enabled = False
        cboPLNAnnLimit.Enabled = False
        cboPLNPremium.Enabled = False
        cboPLNMCO.Enabled = False
        cboDisIndicator.Enabled = False
        txtLGPrivate.Enabled = False
        cboClientPLN.Enabled = False
        cmbVIPInd.Enabled = False
        txtFileName.Enabled = False
        txtClientPolNo.Enabled = False
        txtCoPayEffdate.Enabled = False
        CboSpecGPeriod.Enabled = False
        TxtSpecGPeriod.Enabled = False
        Frame2.Enabled = False
        CboSMPlan.Enabled = False
        CboLifeTimeStatus.Enabled = False
        cboProCat.Enabled = False
        CboMeal.Enabled = False
        CboNursing.Enabled = False
        CboTax.Enabled = False
        CboMRI.Enabled = False
        CboPLNMgmtStatus.Enabled = False
        cboPolicyWording.Enabled = False
        CmdEdit.Enabled = True
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
    
End Sub






