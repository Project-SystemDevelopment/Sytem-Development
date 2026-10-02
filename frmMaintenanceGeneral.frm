VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenanceGeneral 
   Caption         =   "Entity Maintenance"
   ClientHeight    =   8880
   ClientLeft      =   2160
   ClientTop       =   1980
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   444
   ScaleMode       =   2  'Point
   ScaleWidth      =   594
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   735
      Left            =   360
      TabIndex        =   108
      Top             =   7440
      Width           =   10995
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5520
         TabIndex        =   104
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   6960
         TabIndex        =   105
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   4680
         TabIndex        =   103
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   8640
         TabIndex        =   107
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         TabIndex        =   106
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   101
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   100
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   1320
         TabIndex        =   99
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3840
         TabIndex        =   102
         Top             =   240
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8085
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   14261
      _Version        =   393216
      Tabs            =   12
      Tab             =   6
      TabsPerRow      =   6
      TabHeight       =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Access Fees"
      TabPicture(0)   =   "frmMaintenanceGeneral.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label31"
      Tab(0).Control(1)=   "lstPRVList"
      Tab(0).Control(2)=   "Frame7"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Hospital Info"
      TabPicture(1)   =   "frmMaintenanceGeneral.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "lstHospitalList"
      Tab(1).Control(2)=   "Label27"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Clinic Info"
      TabPicture(2)   =   "frmMaintenanceGeneral.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "lstClinicList"
      Tab(2).Control(2)=   "Label28"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Payor Info"
      TabPicture(3)   =   "frmMaintenanceGeneral.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label25"
      Tab(3).Control(1)=   "lstPayorList"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Payor Branch"
      TabPicture(4)   =   "frmMaintenanceGeneral.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label26"
      Tab(4).Control(1)=   "lstBranchList"
      Tab(4).Control(2)=   "Frame2"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Service Provider "
      TabPicture(5)   =   "frmMaintenanceGeneral.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).Control(1)=   "lstServiceProviderList"
      Tab(5).Control(2)=   "Label30"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Hospital Group"
      TabPicture(6)   =   "frmMaintenanceGeneral.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label47"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "lstHosGrpList"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Frame14"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Group Company"
      TabPicture(7)   =   "frmMaintenanceGeneral.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame11"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "lstCompanyGrpList"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Label46"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "Payor Group"
      TabPicture(8)   =   "frmMaintenanceGeneral.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame12"
      Tab(8).Control(1)=   "lstPAYGRPList"
      Tab(8).Control(2)=   "Label38"
      Tab(8).ControlCount=   3
      TabCaption(9)   =   "Agent Maintenance"
      TabPicture(9)   =   "frmMaintenanceGeneral.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame5"
      Tab(9).Control(1)=   "lstAgentList"
      Tab(9).Control(2)=   "Label34"
      Tab(9).ControlCount=   3
      TabCaption(10)  =   "Special Approval"
      TabPicture(10)  =   "frmMaintenanceGeneral.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame10"
      Tab(10).Control(1)=   "lstSpecialList"
      Tab(10).Control(2)=   "Label58"
      Tab(10).ControlCount=   3
      TabCaption(11)  =   "Invoice Rec. Mode"
      TabPicture(11)  =   "frmMaintenanceGeneral.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame8"
      Tab(11).Control(1)=   "lstRecModeList"
      Tab(11).Control(2)=   "Label81"
      Tab(11).ControlCount=   3
      Begin VB.Frame Frame8 
         Height          =   1575
         Left            =   -74760
         TabIndex        =   274
         Top             =   960
         Width           =   11055
         Begin VB.TextBox TxtRecModeCode 
            Height          =   285
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   276
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox TxtRecModeDesc 
            Height          =   840
            Left            =   2640
            MaxLength       =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   275
            Top             =   600
            Width           =   8280
         End
         Begin VB.Label Label80 
            Caption         =   "Rec. Mode Code*"
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
            Left            =   255
            TabIndex        =   278
            Top             =   240
            Width           =   2280
         End
         Begin VB.Label Label79 
            Caption         =   "Rec. Mode Description *"
            Height          =   255
            Left            =   240
            TabIndex        =   277
            Top             =   600
            Width           =   2295
         End
      End
      Begin VB.Frame Frame11 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   239
         Top             =   960
         Width           =   11175
         Begin VB.ComboBox cmbTims 
            Height          =   315
            ItemData        =   "frmMaintenanceGeneral.frx":0150
            Left            =   9795
            List            =   "frmMaintenanceGeneral.frx":015A
            Sorted          =   -1  'True
            TabIndex        =   289
            Top             =   2160
            Width           =   1150
         End
         Begin VB.ComboBox cmbChannel 
            Height          =   315
            Left            =   7440
            Sorted          =   -1  'True
            TabIndex        =   287
            Top             =   2520
            Width           =   1335
         End
         Begin VB.ComboBox CboCoGrpPayor 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            TabIndex        =   241
            Top             =   600
            Width           =   3975
         End
         Begin VB.TextBox txtCoAdd1 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   242
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox txtCompanyName 
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
            Left            =   5160
            MaxLength       =   250
            TabIndex        =   240
            Top             =   255
            Width           =   5775
         End
         Begin VB.TextBox txtGrpCoCode 
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
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   0
            Top             =   255
            Width           =   930
         End
         Begin VB.TextBox txtCoAdd2 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   243
            Top             =   1200
            Width           =   3975
         End
         Begin VB.TextBox txtCoAdd3 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   244
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox txtCoContact 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   245
            Top             =   1770
            Width           =   3975
         End
         Begin VB.TextBox txtCoAdmContact 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   246
            Top             =   2025
            Width           =   3975
         End
         Begin VB.TextBox txtCoCEO 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   247
            Top             =   2280
            Width           =   3975
         End
         Begin VB.TextBox txtCoPhone 
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
            Left            =   7440
            MaxLength       =   15
            TabIndex        =   248
            Top             =   555
            Width           =   1275
         End
         Begin VB.TextBox txtCoFax 
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
            Left            =   9600
            MaxLength       =   15
            TabIndex        =   249
            Top             =   555
            Width           =   1330
         End
         Begin VB.TextBox txtCoEmail 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   250
            Top             =   930
            Width           =   3500
         End
         Begin VB.TextBox txtCoHomepage 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   251
            Top             =   1305
            Width           =   3500
         End
         Begin VB.ComboBox cboCoConversion 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   253
            Top             =   1560
            Width           =   1150
         End
         Begin VB.ComboBox cboCoPrePost 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   255
            Top             =   1845
            Width           =   1150
         End
         Begin MSMask.MaskEdBox txtCoEffDate 
            Height          =   285
            Left            =   7440
            TabIndex        =   252
            Top             =   1560
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboCoTax 
            Height          =   315
            Left            =   7440
            Sorted          =   -1  'True
            TabIndex        =   254
            Top             =   1815
            Width           =   1335
         End
         Begin VB.ComboBox CboGrpComStatus 
            Height          =   315
            ItemData        =   "frmMaintenanceGeneral.frx":016D
            Left            =   7440
            List            =   "frmMaintenanceGeneral.frx":016F
            Sorted          =   -1  'True
            TabIndex        =   256
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "TIMS Client"
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
            Index           =   24
            Left            =   8880
            TabIndex        =   290
            Top             =   2175
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "Business Type *"
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
            Index           =   23
            Left            =   6120
            TabIndex        =   288
            Top             =   2580
            Width           =   1290
         End
         Begin VB.Label Label2 
            Caption         =   "Grp Com Status *"
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
            Index           =   21
            Left            =   6120
            TabIndex        =   273
            Top             =   2220
            Width           =   1290
         End
         Begin VB.Label Label2 
            Caption         =   "Pre&&Post"
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
            Index           =   20
            Left            =   8880
            TabIndex        =   271
            Top             =   1860
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "R&&B Tax"
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
            Index           =   19
            Left            =   6120
            TabIndex        =   270
            Top             =   1860
            Width           =   810
         End
         Begin VB.Label Label78 
            Caption         =   "Adm. Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   269
            Top             =   2025
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Eff. Date  *"
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
            Index           =   18
            Left            =   6120
            TabIndex        =   268
            Top             =   1575
            Width           =   1050
         End
         Begin VB.Label Label2 
            Caption         =   "Conversion"
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
            Index           =   17
            Left            =   8880
            TabIndex        =   267
            Top             =   1590
            Width           =   930
         End
         Begin VB.Label Label77 
            Caption         =   "CEO Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   266
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Grp. Company Code *"
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
            Index           =   15
            Left            =   120
            TabIndex        =   265
            Top             =   255
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Company Address    *"
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
            Index           =   10
            Left            =   240
            TabIndex        =   264
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Group Company Name *"
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
            Index           =   2
            Left            =   3120
            TabIndex        =   263
            Top             =   285
            Width           =   2055
         End
         Begin VB.Label Label76 
            Caption         =   "Fax No"
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
            TabIndex        =   262
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label75 
            Caption         =   "Tel No"
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
            Left            =   6120
            TabIndex        =   261
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label74 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6120
            TabIndex        =   260
            Top             =   945
            Width           =   495
         End
         Begin VB.Label Label73 
            Caption         =   "Homepage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6120
            TabIndex        =   259
            Top             =   1320
            Width           =   870
         End
         Begin VB.Label Label1 
            Caption         =   "Payor Code           *"
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
            Index           =   9
            Left            =   120
            TabIndex        =   258
            Top             =   555
            Width           =   1695
         End
         Begin VB.Label Label72 
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   257
            Top             =   1770
            Width           =   1695
         End
      End
      Begin VB.Frame Frame10 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   213
         Top             =   960
         Width           =   11175
         Begin VB.TextBox txtSpecDesc 
            Height          =   960
            Left            =   2640
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   215
            Top             =   600
            Width           =   8280
         End
         Begin VB.TextBox txtSpecApp 
            Height          =   285
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   214
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label57 
            Caption         =   "Special Approval Description *"
            Height          =   255
            Left            =   240
            TabIndex        =   217
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label37 
            Caption         =   "Special Approval Code*"
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
            Left            =   255
            TabIndex        =   216
            Top             =   240
            Width           =   2280
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   205
         Top             =   960
         Width           =   11175
         Begin VB.TextBox txtAGTCode 
            Height          =   285
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   196
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtAGTDesc 
            Height          =   960
            Left            =   1920
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   197
            Top             =   600
            Width           =   9000
         End
         Begin VB.Label Label29 
            Caption         =   "Agent Code*"
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
            Left            =   255
            TabIndex        =   207
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label14 
            Caption         =   "Agent Description *"
            Height          =   255
            Left            =   240
            TabIndex        =   206
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame12 
         Height          =   1935
         Left            =   -74760
         TabIndex        =   191
         Top             =   960
         Width           =   11055
         Begin VB.TextBox TxtPayGrpCode 
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
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   96
            Top             =   360
            Width           =   675
         End
         Begin VB.TextBox TxtPayGrpDesc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   2040
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   98
            Top             =   720
            Width           =   8835
         End
         Begin VB.TextBox TxtPayGrpName 
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
            Left            =   5520
            MaxLength       =   200
            TabIndex        =   97
            Top             =   360
            Width           =   5325
         End
         Begin VB.Label Label1 
            Caption         =   "Payor Grp. Code *"
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
            Index           =   8
            Left            =   240
            TabIndex        =   194
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Payor Grp. Description "
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
            Left            =   240
            TabIndex        =   193
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Payor Grp. Name *"
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
            Index           =   7
            Left            =   3840
            TabIndex        =   192
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   -74760
         TabIndex        =   176
         Top             =   960
         Width           =   10995
         Begin VB.ComboBox CboPAYCode 
            Height          =   315
            Left            =   1560
            TabIndex        =   71
            Top             =   360
            Width           =   3975
         End
         Begin VB.ComboBox CboBRN 
            Height          =   315
            Left            =   1560
            TabIndex        =   72
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox TxtBRNAdd3 
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   75
            Top             =   1800
            Width           =   3975
         End
         Begin VB.TextBox TxtBRNAdd2 
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   74
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox TxtBRNAdd1 
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   73
            Top             =   1080
            Width           =   3975
         End
         Begin VB.TextBox TxtBRNContact 
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
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   77
            Top             =   720
            Width           =   3795
         End
         Begin VB.TextBox TxtBRNFaxNo 
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
            Left            =   7080
            MaxLength       =   12
            TabIndex        =   79
            Top             =   1440
            Width           =   3795
         End
         Begin VB.TextBox txtBRNTelNo 
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
            Left            =   7080
            MaxLength       =   12
            TabIndex        =   78
            Top             =   1080
            Width           =   3795
         End
         Begin VB.TextBox TxtBRNEmail 
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
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   80
            Top             =   1800
            Width           =   3795
         End
         Begin VB.TextBox txtMNGname 
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
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   76
            Top             =   360
            Width           =   3780
         End
         Begin VB.Label Label41 
            Caption         =   "Payor             *"
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
            Left            =   360
            TabIndex        =   187
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label45 
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5880
            TabIndex        =   183
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label44 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5880
            TabIndex        =   182
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label43 
            Caption         =   "Telephone No. "
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
            Left            =   5880
            TabIndex        =   181
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label42 
            Caption         =   "Fax No."
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
            Left            =   5880
            TabIndex        =   180
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Payor Address *"
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
            Index           =   9
            Left            =   360
            TabIndex        =   179
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "Branch           *"
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
            Left            =   360
            TabIndex        =   178
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Manager Name"
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
            Left            =   5880
            TabIndex        =   177
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   163
         Top             =   960
         Width           =   11055
         Begin VB.ComboBox CboCancelStatus 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   281
            Top             =   2760
            Width           =   1095
         End
         Begin VB.ComboBox CboCoPay 
            Height          =   315
            Left            =   7440
            Sorted          =   -1  'True
            TabIndex        =   70
            Top             =   2760
            Width           =   1335
         End
         Begin VB.ComboBox CboPrePost 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   69
            Top             =   2400
            Width           =   1095
         End
         Begin VB.ComboBox CboRBTax 
            Height          =   315
            Left            =   7440
            Sorted          =   -1  'True
            TabIndex        =   68
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox TxtAdmContact 
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
            Left            =   1920
            MaxLength       =   60
            TabIndex        =   60
            Top             =   2100
            Width           =   3975
         End
         Begin VB.ComboBox CboConversion 
            Height          =   315
            Left            =   9795
            Sorted          =   -1  'True
            TabIndex        =   67
            Top             =   2040
            Width           =   1095
         End
         Begin VB.ComboBox cboPAYHLTCode 
            Height          =   315
            Left            =   1920
            TabIndex        =   52
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox TxtCEO 
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
            Left            =   1920
            MaxLength       =   60
            TabIndex        =   61
            Top             =   2400
            Width           =   3975
         End
         Begin VB.TextBox txtPAYcode 
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
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   53
            Top             =   600
            Width           =   690
         End
         Begin VB.TextBox txtPAYname 
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
            Left            =   5760
            MaxLength       =   50
            TabIndex        =   54
            Top             =   240
            Width           =   5175
         End
         Begin VB.TextBox txtPAYemail 
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
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   64
            Top             =   1320
            Width           =   3435
         End
         Begin VB.TextBox txtPAYhomepage 
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
            Left            =   7440
            MaxLength       =   150
            TabIndex        =   65
            Top             =   1680
            Width           =   3435
         End
         Begin VB.TextBox txtPAYphone 
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
            Left            =   7440
            MaxLength       =   12
            TabIndex        =   62
            Top             =   960
            Width           =   1275
         End
         Begin VB.TextBox txtPAYfax 
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
            Left            =   9600
            MaxLength       =   12
            TabIndex        =   63
            Top             =   960
            Width           =   1275
         End
         Begin VB.ComboBox CboPAYGRP 
            Height          =   315
            Left            =   5760
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   600
            Width           =   5175
         End
         Begin VB.TextBox TxtContact 
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
            Left            =   1920
            MaxLength       =   60
            TabIndex        =   59
            Top             =   1800
            Width           =   3975
         End
         Begin VB.TextBox TxtPayAdd1 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   56
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox TxtPayAdd2 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   57
            Top             =   1200
            Width           =   3975
         End
         Begin VB.TextBox TxtPayAdd3 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   58
            Top             =   1440
            Width           =   3975
         End
         Begin MSMask.MaskEdBox TxtPAYEffDate 
            Height          =   285
            Left            =   7440
            TabIndex        =   66
            Top             =   2040
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Caption         =   "MBM Cancel Status"
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
            Index           =   22
            Left            =   8880
            TabIndex        =   282
            Top             =   2760
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Co-Payment"
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
            Index           =   16
            Left            =   6360
            TabIndex        =   238
            Top             =   2760
            Width           =   930
         End
         Begin VB.Label Label2 
            Caption         =   "Pre&&Post"
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
            Index           =   15
            Left            =   8880
            TabIndex        =   212
            Top             =   2400
            Width           =   810
         End
         Begin VB.Label Label2 
            Caption         =   "R&&B Tax"
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
            Index           =   14
            Left            =   6360
            TabIndex        =   211
            Top             =   2400
            Width           =   810
         End
         Begin VB.Label Label35 
            Caption         =   "Adm. Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   209
            Top             =   2175
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Eff. Date  *"
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
            Index           =   13
            Left            =   6360
            TabIndex        =   204
            Top             =   2040
            Width           =   1050
         End
         Begin VB.Label Label2 
            Caption         =   "Conversion"
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
            Index           =   12
            Left            =   8880
            TabIndex        =   203
            Top             =   2040
            Width           =   930
         End
         Begin VB.Label Label13 
            Caption         =   "Health Type       *"
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
            Left            =   360
            TabIndex        =   190
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "CEO Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   173
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Payor Code         *"
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
            Index           =   0
            Left            =   360
            TabIndex        =   172
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Payor Address       *"
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
            Left            =   360
            TabIndex        =   171
            Top             =   1020
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Payor Name *"
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
            Index           =   0
            Left            =   4200
            TabIndex        =   170
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Fax No"
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
            Left            =   9000
            TabIndex        =   169
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Tel No"
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
            Left            =   6360
            TabIndex        =   168
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6360
            TabIndex        =   167
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Homepage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6360
            TabIndex        =   166
            Top             =   1680
            Width           =   870
         End
         Begin VB.Label Label1 
            Caption         =   "Payor Grp. Code*"
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
            Index           =   7
            Left            =   4200
            TabIndex        =   165
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label39 
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   164
            Top             =   1875
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   152
         Top             =   960
         Width           =   11175
         Begin VB.TextBox txtCLNCode 
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
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   1
            Top             =   240
            Width           =   930
         End
         Begin VB.TextBox txtCLNName 
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
            Left            =   4020
            MaxLength       =   50
            TabIndex        =   24
            Top             =   240
            Width           =   6795
         End
         Begin VB.ComboBox CboClinicGrp 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   25
            Top             =   500
            Width           =   3615
         End
         Begin VB.TextBox TxtCLNAdd1 
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
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   27
            Top             =   780
            Width           =   3615
         End
         Begin VB.TextBox TxtCLNAdd2 
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
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   28
            Top             =   1040
            Width           =   3615
         End
         Begin VB.TextBox TxtCLNAdd3 
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
            Left            =   1680
            MaxLength       =   200
            TabIndex        =   29
            Top             =   1280
            Width           =   3615
         End
         Begin VB.TextBox txtCLNpostcode 
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
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   30
            Top             =   1530
            Width           =   1455
         End
         Begin VB.TextBox txtMIXClinic 
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
            MaxLength       =   50
            TabIndex        =   26
            Top             =   500
            Width           =   4095
         End
         Begin VB.ComboBox cboCLNType 
            Height          =   315
            Left            =   6720
            Sorted          =   -1  'True
            TabIndex        =   40
            Top             =   735
            Width           =   1455
         End
         Begin VB.ComboBox cboCLNcity 
            Height          =   315
            Left            =   4080
            Sorted          =   -1  'True
            TabIndex        =   31
            Top             =   1530
            Width           =   1215
         End
         Begin VB.ComboBox cboCLNstate 
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   32
            Top             =   1790
            Width           =   1455
         End
         Begin VB.TextBox TxtCLNContact 
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   34
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox TxtCLNDocName 
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   36
            Top             =   2280
            Width           =   1455
         End
         Begin VB.ComboBox cboCLNBNFType 
            Height          =   315
            Left            =   9120
            Sorted          =   -1  'True
            TabIndex        =   41
            Top             =   720
            Width           =   1695
         End
         Begin VB.ComboBox cboCLNxray 
            Height          =   315
            Left            =   9120
            Sorted          =   -1  'True
            TabIndex        =   43
            Top             =   1005
            Width           =   1695
         End
         Begin VB.TextBox txtCLNformema 
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
            Left            =   9120
            MaxLength       =   15
            TabIndex        =   45
            Top             =   1290
            Width           =   1695
         End
         Begin VB.TextBox txtCLNEmail 
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
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   33
            Top             =   1800
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtCLNReceivedDate 
            Height          =   285
            Left            =   6720
            TabIndex        =   42
            Top             =   1020
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboCLNpanel 
            Height          =   315
            Left            =   6720
            Sorted          =   -1  'True
            TabIndex        =   44
            Top             =   1260
            Width           =   1455
         End
         Begin VB.TextBox txtCLNTelNo 
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
            Left            =   4080
            MaxLength       =   12
            TabIndex        =   35
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtCLNAccPerson 
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   38
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtCLNFaxNo 
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
            Left            =   4080
            MaxLength       =   12
            TabIndex        =   37
            Text            =   "  "
            Top             =   2265
            Width           =   1215
         End
         Begin VB.TextBox txtCLNnodoc 
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
            Left            =   4080
            MaxLength       =   3
            TabIndex        =   39
            Top             =   2520
            Width           =   1215
         End
         Begin VB.ComboBox cboCLN24hrs 
            Height          =   315
            Left            =   6720
            Sorted          =   -1  'True
            TabIndex        =   46
            Top             =   1530
            Width           =   1455
         End
         Begin VB.TextBox txtCLNbranch 
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
            Left            =   9120
            MaxLength       =   50
            TabIndex        =   47
            Top             =   1540
            Width           =   1695
         End
         Begin VB.TextBox txtCLNBankAcc 
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
            MaxLength       =   25
            TabIndex        =   48
            Top             =   1800
            Width           =   4095
         End
         Begin VB.TextBox txtCLNbank 
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
            MaxLength       =   50
            TabIndex        =   49
            Top             =   2040
            Width           =   4095
         End
         Begin VB.TextBox txtCLNremark 
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
            MaxLength       =   1500
            TabIndex        =   50
            Top             =   2280
            Width           =   4095
         End
         Begin VB.TextBox txtCLNHomePage 
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
            MaxLength       =   50
            TabIndex        =   51
            Top             =   2520
            Width           =   4095
         End
         Begin VB.Label Label1 
            Caption         =   "24 Hours *"
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
            Index           =   14
            Left            =   5400
            TabIndex        =   237
            Top             =   1580
            Width           =   1095
         End
         Begin VB.Label Label71 
            Caption         =   "Branch"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8280
            TabIndex        =   236
            Top             =   1580
            Width           =   1215
         End
         Begin VB.Label Label70 
            Caption         =   "Bank A/C No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5400
            TabIndex        =   235
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label69 
            Caption         =   "Name Of Bank"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5400
            TabIndex        =   234
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label68 
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5400
            TabIndex        =   233
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label66 
            Caption         =   "Received Date *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5400
            TabIndex        =   231
            Top             =   1035
            Width           =   1455
         End
         Begin VB.Label Label65 
            Caption         =   "Clinic Initial *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5400
            TabIndex        =   230
            Top             =   555
            Width           =   1215
         End
         Begin VB.Label Label64 
            Caption         =   "For. Ref."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8280
            TabIndex        =   229
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label62 
            Caption         =   "X-Ray"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8280
            TabIndex        =   228
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Panel *"
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
            Index           =   13
            Left            =   5400
            TabIndex        =   227
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label59 
            Caption         =   "No Of Doc"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3240
            TabIndex        =   226
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label48 
            Caption         =   "Doctor Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   225
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label49 
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   224
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label Label63 
            Caption         =   "State *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   223
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Clinic Type*"
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
            Index           =   3
            Left            =   5400
            TabIndex        =   220
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label Label36 
            Caption         =   "Bnf. Type*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8280
            TabIndex        =   210
            Top             =   795
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Clinic Grp. Code"
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
            Index           =   11
            Left            =   240
            TabIndex        =   160
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Clinic Code        "
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
            Index           =   6
            Left            =   240
            TabIndex        =   159
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Clinic Address       *"
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
            Index           =   6
            Left            =   240
            TabIndex        =   158
            Top             =   800
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Clinic Name *"
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
            Index           =   6
            Left            =   2760
            TabIndex        =   157
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Fax No.      "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3240
            TabIndex        =   156
            Top             =   2280
            Width           =   1005
         End
         Begin VB.Label Label22 
            Caption         =   "Tel No. "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3240
            TabIndex        =   155
            Top             =   2040
            Width           =   1230
         End
         Begin VB.Label Label23 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3240
            TabIndex        =   154
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label24 
            Caption         =   "Homepage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5400
            TabIndex        =   153
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label61 
            Caption         =   "Post Code *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   222
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label Label60 
            Caption         =   "City *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3240
            TabIndex        =   221
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label Label67 
            Caption         =   "Accounts Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   232
            Top             =   2520
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   138
         Top             =   960
         Width           =   11055
         Begin VB.ComboBox cboHOSType 
            Height          =   315
            Left            =   2160
            Sorted          =   -1  'True
            TabIndex        =   285
            Top             =   2400
            Width           =   4095
         End
         Begin VB.ComboBox cboHOState 
            Height          =   315
            Left            =   2160
            Sorted          =   -1  'True
            TabIndex        =   283
            Top             =   2040
            Width           =   4095
         End
         Begin VB.TextBox txtHOSAccCode 
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
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   12
            Top             =   600
            Width           =   930
         End
         Begin VB.TextBox TxtHosRemark 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2160
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   2760
            Width           =   8715
         End
         Begin VB.TextBox TxtHOSAdd3 
            Height          =   285
            Left            =   2160
            MaxLength       =   200
            TabIndex        =   16
            Top             =   1680
            Width           =   4095
         End
         Begin VB.TextBox TxtHOSAdd2 
            Height          =   285
            Left            =   2160
            MaxLength       =   200
            TabIndex        =   15
            Top             =   1320
            Width           =   4095
         End
         Begin VB.TextBox TxtHOSAdd1 
            Height          =   285
            Left            =   2160
            MaxLength       =   200
            TabIndex        =   14
            Top             =   960
            Width           =   4095
         End
         Begin VB.TextBox TxtHosContact 
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
            Left            =   7560
            MaxLength       =   200
            TabIndex        =   17
            Top             =   2400
            Width           =   3375
         End
         Begin VB.ComboBox CboHOSGRP 
            Height          =   315
            Left            =   4920
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   6015
         End
         Begin VB.TextBox TxtFaxNo 
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
            Left            =   7560
            MaxLength       =   40
            TabIndex        =   21
            Top             =   1320
            Width           =   3315
         End
         Begin VB.TextBox TxtTelNo 
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
            Left            =   7560
            MaxLength       =   40
            TabIndex        =   20
            Top             =   960
            Width           =   3315
         End
         Begin VB.TextBox TxtHomepage 
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
            Left            =   7560
            MaxLength       =   100
            TabIndex        =   23
            Top             =   2040
            Width           =   3315
         End
         Begin VB.TextBox TxtEmail 
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
            Left            =   7560
            MaxLength       =   100
            TabIndex        =   22
            Top             =   1680
            Width           =   3315
         End
         Begin VB.TextBox TxtHOSName 
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
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   13
            Top             =   600
            Width           =   6015
         End
         Begin VB.TextBox TxtHOSCode 
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
            Left            =   2160
            MaxLength       =   5
            TabIndex        =   10
            Top             =   240
            Width           =   930
         End
         Begin VB.TextBox TxtHosDocName 
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
            Left            =   7080
            MaxLength       =   200
            TabIndex        =   18
            Top             =   3480
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Label Label63 
            Caption         =   "Hospital Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   360
            TabIndex        =   286
            Top             =   2415
            Width           =   1455
         End
         Begin VB.Label Label63 
            Caption         =   "State *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   360
            TabIndex        =   284
            Top             =   2055
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Hospital Acc Code  *"
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
            Index           =   2
            Left            =   360
            TabIndex        =   202
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label56 
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
            Height          =   240
            Left            =   360
            TabIndex        =   149
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label55 
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6360
            TabIndex        =   148
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Hospital Grp. Code *"
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
            Index           =   12
            Left            =   3240
            TabIndex        =   147
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label54 
            Caption         =   "Homepage"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6360
            TabIndex        =   146
            Top             =   2040
            Width           =   870
         End
         Begin VB.Label Label53 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6360
            TabIndex        =   145
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label52 
            Caption         =   "Telephone No. "
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
            Left            =   6360
            TabIndex        =   144
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label51 
            Caption         =   "Fax No."
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
            Left            =   6360
            TabIndex        =   143
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Hospital Name     *"
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
            Index           =   1
            Left            =   3240
            TabIndex        =   142
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Hospital Address         *"
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
            Left            =   360
            TabIndex        =   141
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Hospital Code          *"
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
            Index           =   1
            Left            =   360
            TabIndex        =   140
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label50 
            Caption         =   "CEO Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5280
            TabIndex        =   139
            Top             =   3600
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   131
         Top             =   960
         Width           =   10935
         Begin VB.TextBox txtAccVer 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   9720
            MaxLength       =   8
            TabIndex        =   9
            Top             =   960
            Width           =   855
         End
         Begin VB.ComboBox txtHealthCode 
            Height          =   315
            ItemData        =   "frmMaintenanceGeneral.frx":0171
            Left            =   1680
            List            =   "frmMaintenanceGeneral.frx":0173
            TabIndex        =   3
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txtPRVProviderCharges 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1665
            MaxLength       =   8
            TabIndex        =   7
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txtPRVCaseCharges 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7200
            MaxLength       =   8
            TabIndex        =   6
            Top             =   600
            Width           =   3375
         End
         Begin VB.ComboBox txtINSCode1 
            Height          =   315
            ItemData        =   "frmMaintenanceGeneral.frx":0175
            Left            =   1680
            List            =   "frmMaintenanceGeneral.frx":0177
            TabIndex        =   5
            Top             =   600
            Width           =   2775
         End
         Begin VB.ComboBox txtPRVCode 
            Height          =   315
            ItemData        =   "frmMaintenanceGeneral.frx":0179
            Left            =   7200
            List            =   "frmMaintenanceGeneral.frx":017B
            TabIndex        =   4
            Top             =   240
            Width           =   3375
         End
         Begin MSMask.MaskEdBox txtAccessDate 
            Height          =   285
            Left            =   7200
            TabIndex        =   8
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Caption         =   "Access Version"
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
            Left            =   8520
            TabIndex        =   201
            Top             =   960
            Width           =   1410
         End
         Begin VB.Label Label2 
            Caption         =   "Access Eff. Date  *"
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
            Left            =   5400
            TabIndex        =   200
            Top             =   960
            Width           =   1410
         End
         Begin VB.Label Label20 
            Caption         =   "Health Type  *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   186
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label Label4 
            Caption         =   "Insured Type *"
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
            Index           =   5
            Left            =   360
            TabIndex        =   135
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Srv. Provider Code *"
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
            Index           =   5
            Left            =   5400
            TabIndex        =   134
            Top             =   240
            Width           =   1845
         End
         Begin VB.Label Label2 
            Caption         =   "Access Fee        *"
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
            Index           =   5
            Left            =   360
            TabIndex        =   133
            Top             =   960
            Width           =   1890
         End
         Begin VB.Label Label15 
            Caption         =   "Case Charges              *"
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
            TabIndex        =   132
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Frame14 
         Height          =   1935
         Left            =   240
         TabIndex        =   123
         Top             =   960
         Width           =   10935
         Begin VB.TextBox TxtHosGrpCode 
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
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   93
            Top             =   360
            Width           =   795
         End
         Begin VB.TextBox TxtHosGrpDesc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   2040
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   95
            Top             =   720
            Width           =   8595
         End
         Begin VB.TextBox TxtHosGrpName 
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
            Left            =   4920
            MaxLength       =   200
            TabIndex        =   94
            Top             =   360
            Width           =   5685
         End
         Begin VB.Label Label1 
            Caption         =   "Hospital Grp. Code *"
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
            Index           =   10
            Left            =   240
            TabIndex        =   126
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Hospital Grp. Description "
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
            Index           =   11
            Left            =   240
            TabIndex        =   125
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Hospital Grp. Name *"
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
            Index           =   9
            Left            =   3120
            TabIndex        =   124
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   110
         Top             =   960
         Width           =   11055
         Begin VB.TextBox txtFaxNo2 
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
            Left            =   7680
            MaxLength       =   12
            TabIndex        =   90
            Top             =   2040
            Width           =   3195
         End
         Begin VB.TextBox txtTollFree 
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
            Left            =   7680
            MaxLength       =   15
            TabIndex        =   91
            Top             =   2400
            Width           =   3195
         End
         Begin VB.TextBox txtTelNo1 
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
            Left            =   7680
            MaxLength       =   12
            TabIndex        =   87
            Top             =   960
            Width           =   3195
         End
         Begin VB.TextBox TxtSRVContact 
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
            Left            =   7680
            MaxLength       =   200
            TabIndex        =   86
            Top             =   600
            Width           =   3195
         End
         Begin VB.TextBox txtFaxNo1 
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
            Left            =   7680
            MaxLength       =   12
            TabIndex        =   89
            Top             =   1680
            Width           =   3195
         End
         Begin VB.TextBox txtTelNo2 
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
            Left            =   7680
            MaxLength       =   12
            TabIndex        =   88
            Top             =   1320
            Width           =   3195
         End
         Begin VB.TextBox TxtSRVAdd3 
            Height          =   285
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   85
            Top             =   1320
            Width           =   3855
         End
         Begin VB.TextBox TxtSRVAdd2 
            Height          =   285
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   84
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox TxtSRVAdd1 
            Height          =   285
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   83
            Top             =   600
            Width           =   3855
         End
         Begin VB.TextBox txtSRVcode 
            Height          =   285
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   81
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtRemark 
            Height          =   1035
            Left            =   2040
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   92
            Top             =   1680
            Width           =   3855
         End
         Begin VB.TextBox txtSRVname 
            Height          =   285
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   82
            Top             =   240
            Width           =   5955
         End
         Begin VB.Label Label40 
            Caption         =   "Fax No. 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6480
            TabIndex        =   120
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label32 
            Caption         =   "Toll Free"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6480
            TabIndex        =   119
            Top             =   2400
            Width           =   870
         End
         Begin VB.Label Label19 
            Caption         =   "Fax No. 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6480
            TabIndex        =   118
            Top             =   1680
            Width           =   870
         End
         Begin VB.Label Label18 
            Caption         =   "Tel No. 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6480
            TabIndex        =   117
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Contact Person"
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
            Left            =   6480
            TabIndex        =   116
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Tel No. 1"
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
            Left            =   6480
            TabIndex        =   115
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Remark"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   114
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Srv. Provider Code *"
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
            Index           =   4
            Left            =   240
            TabIndex        =   113
            Top             =   240
            Width           =   2160
         End
         Begin VB.Label Label2 
            Caption         =   "Srv. Provider Address   *"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   112
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Srv. Provider Name *"
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
            Index           =   4
            Left            =   3000
            TabIndex        =   111
            Top             =   240
            Width           =   2415
         End
      End
      Begin MSComctlLib.ListView lstServiceProviderList 
         Height          =   2040
         Left            =   -74760
         TabIndex        =   121
         Top             =   4200
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   3598
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Service Provider Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Service Provider Name"
            Object.Width           =   15875
         EndProperty
      End
      Begin MSComctlLib.ListView lstHosGrpList 
         Height          =   3000
         Left            =   240
         TabIndex        =   127
         Top             =   3240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5292
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Hospital Group Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Hospital Group Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hospital Group Description"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstCompanyGrpList 
         Height          =   2745
         Left            =   -74880
         TabIndex        =   129
         Top             =   4320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4842
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Company Group Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Company Group Name"
            Object.Width           =   14111
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Payor Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "BusinessChannel"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "TIMS Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstPRVList 
         Height          =   3555
         Left            =   -74760
         TabIndex        =   136
         Top             =   2700
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   6271
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Srv. Pro. Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Health Type"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Insured Type"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Provider Charges"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Case Charges"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Effective Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Index"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Version"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstHospitalList 
         Height          =   1440
         Left            =   -74760
         TabIndex        =   150
         Top             =   4800
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   2540
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Hospital Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Hospital Acc Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hospital Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Hospital Group Code"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView lstClinicList 
         Height          =   2160
         Left            =   -74760
         TabIndex        =   161
         Top             =   4080
         Width           =   10905
         _ExtentX        =   19235
         _ExtentY        =   3810
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
         NumItems        =   27
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Clinic Code"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Clinic Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clinic Group Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Clinic Initial"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Clinic Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Benefit Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Address 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Address 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Address 3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Post Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "City"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "State"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Received Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Panel"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "X-Ray"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Formema Ref"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "No Of Doctor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Tel. No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Fax No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Email"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Homepage"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Bank Account"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Bank"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Branch"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Account Person"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "24 Hours"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstPayorList 
         Height          =   1665
         Left            =   -74760
         TabIndex        =   174
         Top             =   4560
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   2937
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Health Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Payor Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Payor Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Payor Grp. Code"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lstBranchList 
         Height          =   2760
         Left            =   -74760
         TabIndex        =   184
         Top             =   3480
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4868
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Branch"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Branch Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Payor Address1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Payor Address2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Payor Address 3"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstPAYGRPList 
         Height          =   3000
         Left            =   -74760
         TabIndex        =   195
         Top             =   3240
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5292
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor Group Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Payor Group Name"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstAgentList 
         Height          =   3360
         Left            =   -74760
         TabIndex        =   198
         Top             =   3000
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   5927
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Agent Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstSpecialList 
         Height          =   3360
         Left            =   -74760
         TabIndex        =   218
         Top             =   3000
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   5927
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Approval Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstRecModeList 
         Height          =   3360
         Left            =   -74760
         TabIndex        =   279
         Top             =   2880
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   5927
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Approval Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin VB.Label Label81 
         Alignment       =   2  'Center
         Caption         =   "REC. MODE LIST"
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
         Height          =   225
         Left            =   -74640
         TabIndex        =   280
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         Caption         =   "SPECIAL APPROVE LIST"
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
         Height          =   225
         Left            =   -74880
         TabIndex        =   219
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "AGENT LIST"
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
         Height          =   225
         Left            =   -74880
         TabIndex        =   208
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "PAYOR GOUP LIST"
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
         Height          =   225
         Left            =   -74760
         TabIndex        =   199
         Top             =   3000
         Width           =   1515
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "BRANCH LIST"
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
         Height          =   225
         Left            =   -74760
         TabIndex        =   185
         Top             =   3240
         Width           =   1155
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "PAYOR INFO LIST"
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
         Height          =   225
         Left            =   -74760
         TabIndex        =   175
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "CLINIC LIST"
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
         Height          =   225
         Left            =   -74760
         TabIndex        =   162
         Top             =   3850
         Width           =   1035
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "HOSPITAL LIST"
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
         Height          =   225
         Left            =   -74760
         TabIndex        =   151
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "PROVIDER CHARGES LIST"
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
         Height          =   225
         Left            =   -74760
         TabIndex        =   137
         Top             =   2500
         Width           =   2235
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "COMPANY GROUP LIST"
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
         Height          =   225
         Left            =   -74760
         TabIndex        =   130
         Top             =   3960
         Width           =   1875
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         Caption         =   "HOSPITAL GOUP LIST"
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
         Height          =   225
         Left            =   240
         TabIndex        =   128
         Top             =   3000
         Width           =   1515
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "SERVICE PROVIDER LIST"
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
         Height          =   225
         Left            =   -74760
         TabIndex        =   122
         Top             =   3960
         Width           =   2175
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Entity Maintenance"
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
      Height          =   195
      Left            =   0
      TabIndex        =   272
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6720
      TabIndex        =   189
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "TOTAL RECORDS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5160
      TabIndex        =   188
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable        * - Mandatory Fields"
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
      Left            =   240
      TabIndex        =   109
      Top             =   8520
      Width           =   4095
   End
End
Attribute VB_Name = "frmMaintenanceGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim bExit As Boolean
Public PAYCriteria As String
Public BRNCriteria As String
Public HOSCriteria As String
Public CLNCriteria As String
Public AGTCriteria As String
Public PRVChargesCriteria As String
Public SPECCriteria As String
Public RECModeCriteria As String
Public SRVCriteria As String
Public PAYGRPCriteria As String
Public GRPCoCriteria As String
Public HosGrpCriteria As String
Public ListClick As Boolean
Private m_blnDirection(1 To 6) As Boolean

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

'**********************************Start Branch**************************************
Private Sub SaveBranch()
On Error Resume Next
Dim RecordSaved
Dim BRN As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim SaveBRN As New ADODB.Recordset
Dim strsql As String
Dim StrSql2 As String
Dim PAYName As String
    
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    If CboPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboPAYCode.SetFocus
    ElseIf CboBRN = "" Then
        MsgBox "Please Enter Branch Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboBRN.SetFocus
    'ElseIf txtBRNname = "" Then
        'MsgBox "Please Enter Branch Name", vbOKOnly + vbCritical, DialogBoxTitle
        'txtBRNname.SetFocus
    ElseIf TxtBRNAdd1 = "" Then
        MsgBox "Please Enter Branch Address", vbOKOnly + vbCritical, DialogBoxTitle
        TxtBRNAdd1.SetFocus
    
    Else
        
        strsql = "Select * From BRN Where PAYCode = '" & Trim(Mid(CboPAYCode, 1, 2)) & "' AND BRNCode = '" & Trim(Mid(CboBRN, 1, 2)) & "'"
        BRN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If BRN.recordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           BRN.Close
           CboPAYCode.SetFocus
           CboPAYCode.SelStart = 0
           CboPAYCode.SelLength = Len(CboPAYCode)
            
           Set BRN = Nothing
           'Exit Sub
        Else
            
            StrSql2 = "Select PAYCode, PAYCompanyName From PAY Where PAYCode = '" & Mid(CboPAYCode, 1, 2) & "'"
            PAY.Open StrSql2, medixmasterconn, adOpenKeyset, adLockOptimistic
        
            PAYName = PAY!PAYCompanyName
            
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    BRN.Open "BRN", medixmasterconn, adOpenKeyset, adLockOptimistic
                    'PAY.Open "PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
                    BRN.AddNew
                Else
                    strsql = "Select * From BRN Where PAYCode = '" & Trim(CboPAYCode) & "' AND BRNCode = '" & Trim(CboBRN) & "'"
                    BRN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                End If
            
                With BRN
                    !PAYCode = Trim(Mid(CboPAYCode, 1, 2))
                    !BRNCode = Trim(Mid(CboBRN, 1, 2))
                    !BRNName = ChkStr(PAYName)
                    !BRNAddress1 = ChkStr(TxtBRNAdd1)
                    !BRNAddress2 = ChkStr(TxtBRNAdd2)
                    !BRNAddress3 = ChkStr(TxtBRNAdd3)
                    !BRNTelNo = ChkStr(txtBRNTelNo)
                    !BRNFaxNo = ChkStr(TxtBRNFaxNo)
                    !BRNEmail = ChkStr(TxtBRNEmail)
                    !BRNMNGName = ChkStr(txtMNGname)
                    !BRNContact = ChkStr(TxtBRNContact)
                    !BRNLastUpdateUser = Logon
                    !BRNLastUpdateDate = Date
                    BRN.Update
                End With
                
                Call ClearForm(Me)
                Call BranchListing
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
      
                BRN.Close
                Set SaveBRN = Nothing
            End If
        EditFlag = False
        End If
    End If
    
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End Sub

Private Sub DeleteBranch()
Dim DeleteBRN As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

EditFlag = False
If CboBRN = "" Then
   MsgBox "Please Enter Branch Code", vbOKOnly + vbCritical, DialogBoxTitle
   'CboBRN.SetFocus
ElseIf lstBranchList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where BRNCode = '" & RemStr(CboBRN) & "' AND PAYCode = '" & RemStr(CboPAYCode) & "'"
   
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From BRN Where PAYCode = '" & ChkStr(lstBranchList.SelectedItem.Text) & _
                    "' AND BRNcode = '" & ChkStr(lstBranchList.SelectedItem.SubItems(1)) & "'"
           DeleteBRN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call BranchListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           CboBRN.SetFocus
         End If
     End If
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
Exit Sub

End Sub

Private Sub BranchListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim branchMaster As ListItem
Dim sql As String
Dim TotalBRNCount As Integer
     
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
     
    TotalBRNCount = 0
    lstBranchList.ListItems.Clear
        sql = " SELECT BRNCode, BRNName, PAYCode, BRNAddress1, BRNAddress2, " & _
              "BRNAddress3, BRNTelNo, BRNFaxNo, BRNEmail, BRNContact, BRNMNGName " & _
              " from BRN where 0=0 "
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set branchMaster = lstBranchList.ListItems.Add(, , rst2!PAYCode)
                branchMaster.SubItems(1) = rst2!BRNCode
                branchMaster.SubItems(2) = rst2!BRNName
                branchMaster.SubItems(3) = rst2!BRNAddress1
                branchMaster.SubItems(4) = rst2!BRNAddress2
                branchMaster.SubItems(5) = rst2!BRNAddress3
                rst2.MoveNext
        TotalBRNCount = TotalBRNCount + 1
        txtTotalRecord = TotalBRNCount
        Loop
    rst2.Close
    End If
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

'Private Sub PrintPAYBRN()
    'RptBRNList.Show
'End Sub

Private Sub CboBRN_Change()
If InStr(CboBRN.Text, "null") Then
       CboBRN.Enabled = False
    End If
End Sub

Private Sub CboBRN_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboBRN.Text, CboBRN.SelStart) + Chr(KeyAscii) + Mid(CboBRN.Text, CboBRN.SelStart + CboBRN.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboBRN.ListCount - 1
 
        If NewText = LCase(Left(CboBRN.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboBRN.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboBRN.Text = ValidValue
                CboBRN.SelStart = 0
                CboBRN.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cmdPrint_Click()
    EditFlag = False
    Call ClearForm(Me)
    Call EntriesEnable(True)
    Call PRVChargesComboListing
        
    'Select Case SSTab1.Tab
        'Case 0
            'Call PrintPRV
        'Case 1
            'Call PrintHOS
        'Case 2
            'If lstClinicList.listitems.count <> 0 Then
                'Call ExportListViewtoExcel(lstClinicList)
            'Else
                'MsgBox "Report cannot be printed without any record.", vbInformation
            'End If
            
        'Case 3
            'Call PrintPAY
        'Case 4
            'Call PrintPAYBRN
        'Case 5
            'Call PrintSRV
        'Case 6
            'Call PrintHOSGRP
        'Case 7
            'Call PrintCompanyGrp
        'Case 8
            'Call PrintPAYGRP
        'Case 9
            'Call PrintAgent
        'Case 10
            'Call PrintSpecialApproval
        'Case 11
            'Call PrintRecMode
    'End Select
    
    Select Case SSTab1.Tab
        Case 0
            If lstPRVList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstPRVList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 1
            If lstHospitalList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstHospitalList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 2
            If lstClinicList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstClinicList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 3
            If lstPayorList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstPayorList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 4
            If lstBranchList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstBranchList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 5
            If lstServiceProviderList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstServiceProviderList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 6
            If lstHosGrpList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstHosGrpList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 7
            If lstCompanyGrpList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstCompanyGrpList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 8
            If lstPAYGRPList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstPAYGRPList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 9
            If lstAgentList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstAgentList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 10
            If lstSpecialList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstSpecialList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 11
            If lstRecModeList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstRecModeList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
    End Select
    
End Sub

Private Sub lstAgentList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstAgentList.SortKey = ColumnHeader.Index - 1
lstAgentList.Sorted = True

    If lstAgentList.SortOrder = lvwAscending Then
        lstAgentList.SortOrder = lvwDescending
    Else
        lstAgentList.SortOrder = lvwAscending
    End If

End Sub

Private Sub lstAgentList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String

If lstAgentList.ListItems.Count Then
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    strsql = "SELECT AgentCode, Remarks " & _
    "from Agent Where AgentCode = '" & lstAgentList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    
    txtAGTCode = rst3!AgentCode
    txtAGTDesc = Trim(rst3!Remarks)
    Call EntriesEnable(False)
    cmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End If

End Sub

Private Sub lstAgentList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String

If lstAgentList.ListItems.Count Then
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    strsql = "SELECT AgentCode, Remarks " & _
    "from Agent Where AgentCode = '" & lstAgentList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    
    txtAGTCode = rst3!AgentCode
    txtAGTDesc = Trim(rst3!Remarks)
    Call EntriesEnable(False)
    cmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
  '
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If

End Sub

Private Sub lstBranchList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
        
    If lstBranchList.ListItems.Count <> 0 Then
        
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        
        strsql = "Select * from BRN Where PAYCode = '" & Trim(lstBranchList.SelectedItem.Text) & "' And BRNCode = '" & Trim(lstBranchList.SelectedItem.SubItems(1)) & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        CboBRN = rst3!BRNCode
        CboPAYCode = rst3!PAYCode
        TxtBRNAdd1 = rst3!BRNAddress1
        If Len(rst3!BRNAddress2) Then
           TxtBRNAdd2 = rst3!BRNAddress2
        End If
        If Len(rst3!BRNAddress3) Then
           TxtBRNAdd3 = rst3!BRNAddress3
        End If
        If Len(rst3!BRNContact) Then
           TxtBRNContact = rst3!BRNContact
        End If
        If Len(rst3!BRNMNGName) Then
           txtMNGname = rst3!BRNMNGName
        End If
        If Len(rst3!BRNTelNo) Then
           txtBRNTelNo = rst3!BRNTelNo
        End If
        If Len(rst3!BRNFaxNo) Then
           TxtBRNFaxNo = rst3!BRNFaxNo
        End If
        If Len(rst3!BRNEmail) Then
           TxtBRNEmail = rst3!BRNEmail
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
   '     medixmasterconn.Close
     '   Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub SearchBranch()
    Dim StrAppend As String
        
    Screen.MousePointer = vbHourglass
    StrAppend = ""
    
    If Len(Trim(CboBRN)) Then
        StrAppend = StrAppend + " AND BRNCode like '" & RemStr(Mid(CboBRN, 1, IIf(InStr(1, CboBRN, "-") = 0, 4, InStr(1, CboBRN, "-") - 1))) & "%'"
        'strAppend = strAppend + " And PAYGRPCode like '" & RemStr(Mid(CboPAYGRP, 1, IIf(InStr(1, CboPAYGRP, "-") = 0, 4, InStr(1, CboPAYGRP, "-") - 1))) & "%'"
    End If
    
    If Len(Trim(CboPAYCode)) Then
        StrAppend = StrAppend + " AND PAYCode like '" & RemStr(Mid(CboPAYCode, 1, IIf(InStr(1, CboPAYCode, "-") = 0, 4, InStr(1, CboPAYCode, "-") - 1))) & "%'"
        'strAppend = strAppend + "PAYCode = '" & Trim(Mid(CboPAYCode, 1, 2)) & "'"
    End If
    
    BRNCriteria = StrAppend
    Call BranchListing(StrAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstBranchList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

lstBranchList.SortKey = ColumnHeader.Index - 1
lstBranchList.Sorted = True
If lstBranchList.SortOrder = lvwAscending Then
lstBranchList.SortOrder = lvwDescending
Else
lstBranchList.SortOrder = lvwAscending
End If

End Sub

Private Sub BRNComboListing()
Dim BRN As New ADODB.Recordset
Dim BRNPAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cmd2 As New ADODB.Command

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    ' Set up a command object for the stored procedure.
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    BRNPAY.CursorLocation = adUseClient
    BRNPAY.CursorType = adOpenForwardOnly
    BRNPAY.CacheSize = 100
    
    Set cmd2.ActiveConnection = medixmasterconnMaint
    cmd2.CommandText = "SearchCity"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 15
    BRN.CursorLocation = adUseClient
    BRN.CursorType = adOpenForwardOnly
    BRN.CacheSize = 100
    
    ' Create a recordset by executing the command.
    Set BRNPAY = cmd.Execute
    Set BRN = cmd2.Execute
    
    CboPAYCode.Clear
    CboBRN.Clear
    
    Do Until BRNPAY.EOF
        With BRNPAY
            CboPAYCode.AddItem !PAYCode & " - " & Trim(!PAYCompanyName)
        End With
        BRNPAY.MoveNext
    Loop
    BRNPAY.Close
    Set BRNPAY = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
        
    Do Until BRN.EOF
        CboBRN.AddItem BRN!CTYCode & " - " & BRN!CTYDescription
        BRN.MoveNext
    Loop
    BRN.Close
    Set BRN = Nothing
 '   medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End Sub
'**********************************End Branch**************************************

'**********************************Start Hospital**************************************

Private Sub saveHospital()
On Error Resume Next
Dim RecordSaved
Dim HOS As New ADODB.Recordset
Dim saveHos As New ADODB.Recordset
Dim strsql As String
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If TxtHOSCode = "" Then
        MsgBox "Please Enter Hospital Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtHOSCode.SetFocus
    ElseIf TxtHOSName = "" Then
        MsgBox "Please Enter Hospital Name", vbOKOnly + vbCritical, DialogBoxTitle
        TxtHOSName.SetFocus
    ElseIf CboHOSGRP = "" Then
        MsgBox "Please Enter Hospital Group Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboHOSGRP.SetFocus
    ElseIf TxtHOSAdd1 = "" Then
        MsgBox "Please Enter Hospital Address", vbOKOnly + vbCritical, DialogBoxTitle
        TxtHOSAdd1.SetFocus
    Else
        strsql = "Select * From HOS Where HOSCode = '" & Trim(TxtHOSCode) & "'"
        HOS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If HOS.recordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
        HOS.Close
        Set HOS = Nothing
        
        TxtHOSCode.SetFocus
        TxtHOSCode.SelStart = 0
        TxtHOSCode.SelLength = Len(TxtHOSCode)
        Else
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                HOS.Open "HOS", medixmasterconn, adOpenKeyset, adLockOptimistic
                HOS.AddNew
            Else
                strsql = "Select * From HOS Where HOSCode = '" & Trim(TxtHOSCode) & "'"
                HOS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With HOS
                    !HOSCode = ChkStr(TxtHOSCode)
                    !HOSAccCode = ChkStr(txtHOSAccCode)
                    !HosName = ChkStr(TxtHOSName)
                    !HOSGRPCode = Trim(Mid(CboHOSGRP, 1, 6))
                    !HOSAdd1 = ChkStr(TxtHOSAdd1)
                    !HOSAdd2 = ChkStr(TxtHOSAdd2)
                    !HOSAdd3 = ChkStr(TxtHOSAdd3)
                    !HOSTelNo = ChkStr(TxtTelNo)
                    !HOSFaxNo = ChkStr(TxtFaxNo)
                    !HOSEmail = ChkStr(TxtEmail)
                    !HOSHomepage = ChkStr(TxtHomepage)
                    !HOSContact = ChkStr(TxtHosContact)
                    !HOSDocName = ChkStr(TxtHosDocName)
                    !HOSRemark = ChkStr(TxtHosRemark)
                    !HOSState = cboHOState
                    !HOSStatus = Mid(cboHOSType, 1, 1)
                    !HOSLastUpdateUser = Logon
                    !HOSLastUpdateDate = Date
                End With
                HOS.Update
                Call ClearForm(Me)
                Call HospitalListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                
            HOS.Close
            Set saveHos = Nothing
        End If
        EditFlag = False
    End If
End If

'medixmasterconn.Close
'Set medixmasterconn = Nothing

End Sub

Private Sub DeleteHospital()

Dim DeleteHOS As New ADODB.Recordset
Dim rsCheckHOS As New ADODB.Recordset
Dim strsql, sqlCLMINV As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

EditFlag = False
If TxtHOSCode = "" Then
   MsgBox "Please Enter Hospital Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtHOSCode.SetFocus
ElseIf lstHospitalList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCLMINV = "Select * from CAS where CASHOSCODE = '" & RemStr(TxtHOSCode) & "'"
   
   
   rsCheckHOS.MaxRecords = 1
   rsCheckHOS.Open sqlCLMINV, medixmasterconn, adOpenKeyset, adLockOptimistic
   
     
   If rsCheckHOS.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Claim Invoice Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckHOS.recordCount = 0 Then
           strsql = "Delete From HOS Where HOScode = '" & ChkStr(lstHospitalList.SelectedItem.Text) & "'"
           DeleteHOS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call HospitalListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           TxtHOSCode.SetFocus
         End If
     End If
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing

Exit Sub

End Sub

Private Sub HospitalListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim HospitalMaster As ListItem
Dim sql As String
Dim TotalHOSCount As Integer
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    sql = "Select HOSCode, HOSAccCode, HOSName, HOSGrpCode, HOSAdd1, HOSAdd2, HOSAdd3," & _
          "HOSTelNo, HOSFaxNo, HOSEmail, HOSHomepage, HOSContact, HOSDocName, HOSRemark " & _
          "from HOS where 0=0 "
          
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If
              
    TotalHOSCount = 0
    lstHospitalList.ListItems.Clear
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set HospitalMaster = lstHospitalList.ListItems.Add(, , rst2!HOSCode)
                If Len(rst2!HOSAccCode) Then
                    HospitalMaster.SubItems(1) = rst2!HOSAccCode
                End If
                HospitalMaster.SubItems(2) = rst2!HosName
                If Len(rst2!HOSGRPCode) Then
                   HospitalMaster.SubItems(3) = rst2!HOSGRPCode
                End If
                rst2.MoveNext
        TotalHOSCount = TotalHOSCount + 1
        txtTotalRecord = TotalHOSCount
        Loop
    rst2.Close
    End If
    
    'HOS.Close
    'Set HOS = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub lstBranchList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstBranchList.ListItems.Count <> 0 Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        
        strsql = "Select * from BRN Where PAYCode = '" & Trim(lstBranchList.SelectedItem.Text) & "' And BRNCode = '" & Trim(lstBranchList.SelectedItem.SubItems(1)) & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        CboBRN = rst3!BRNCode
        CboPAYCode = rst3!PAYCode
        TxtBRNAdd1 = rst3!BRNAddress1
        If Len(rst3!BRNAddress2) Then
           TxtBRNAdd2 = rst3!BRNAddress2
        End If
        If Len(rst3!BRNAddress3) Then
           TxtBRNAdd3 = rst3!BRNAddress3
        End If
        If Len(rst3!BRNContact) Then
           TxtBRNContact = rst3!BRNContact
        End If
        If Len(rst3!BRNMNGName) Then
           txtMNGname = rst3!BRNMNGName
        End If
        If Len(rst3!BRNTelNo) Then
           txtBRNTelNo = rst3!BRNTelNo
        End If
        If Len(rst3!BRNFaxNo) Then
           TxtBRNFaxNo = rst3!BRNFaxNo
        End If
        If Len(rst3!BRNEmail) Then
           TxtBRNEmail = rst3!BRNEmail
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstBranchList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstBranchList.ListItems.Count <> 0 Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        
        strsql = "Select * from BRN Where PAYCode = '" & Trim(lstBranchList.SelectedItem.Text) & "' And BRNCode = '" & Trim(lstBranchList.SelectedItem.SubItems(1)) & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        CboBRN = rst3!BRNCode
        CboPAYCode = rst3!PAYCode
        TxtBRNAdd1 = rst3!BRNAddress1
        If Len(rst3!BRNAddress2) Then
           TxtBRNAdd2 = rst3!BRNAddress2
        End If
        If Len(rst3!BRNAddress3) Then
           TxtBRNAdd3 = rst3!BRNAddress3
        End If
        If Len(rst3!BRNContact) Then
           TxtBRNContact = rst3!BRNContact
        End If
        If Len(rst3!BRNMNGName) Then
           txtMNGname = rst3!BRNMNGName
        End If
        If Len(rst3!BRNTelNo) Then
           txtBRNTelNo = rst3!BRNTelNo
        End If
        If Len(rst3!BRNFaxNo) Then
           TxtBRNFaxNo = rst3!BRNFaxNo
        End If
        If Len(rst3!BRNEmail) Then
           TxtBRNEmail = rst3!BRNEmail
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
      '  Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstClinicList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstClinicList.ListItems.Count <> 0 Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        strsql = "Select * from CLN Where CLNCode = '" & lstClinicList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtCLNCode = rst3!CLNCode
                
        If Trim(rst3!CLNname) <> " " Then
            txtCLNName = rst3!CLNname
        End If
        
        If Trim(rst3!ClinicGrpCode) <> " " Then
            CboClinicGrp = rst3!ClinicGrpCode
        End If
        
        If Trim(rst3!CLNAdd1) <> " " Then
            TxtCLNAdd1 = rst3!CLNAdd1
        End If
        
        If Trim(rst3!CLNAdd2) <> " " Then
            TxtCLNAdd2 = rst3!CLNAdd2
        End If
        
        If Trim(rst3!CLNAdd3) <> " " Then
            TxtCLNAdd3 = rst3!CLNAdd3
        End If
        
        If Trim(rst3!CLNPhone) <> " " Then
            txtCLNTelNo = rst3!CLNPhone
        End If
        
        If Trim(rst3!CLNFax) <> " " Then
            txtCLNFaxNo = rst3!CLNFax
        End If
        
        If Trim(rst3!CLNEmail) <> " " Then
            txtCLNEmail = rst3!CLNEmail
        End If
        
        If Trim(rst3!CLNHomepage) <> " " Then
            txtCLNHomePage = rst3!CLNHomepage
        End If
        
        If Trim(rst3!CLNRepresentative) <> " " Then
            TxtCLNContact = rst3!CLNRepresentative
        End If
        
        If Trim(rst3!CLNDocName) <> " " Then
            TxtCLNDocName = rst3!CLNDocName
        End If
        
        If Trim(rst3!CLNBNFType) <> " " Then
            cboCLNBNFType = rst3!CLNBNFType
        End If
        
        If Trim(rst3!CLNForRef) <> " " Then
            txtCLNformema = rst3!CLNForRef
        End If
        
        If Trim(rst3!CLNNoofDoc) <> " " Then
            txtCLNnodoc = rst3!CLNNoofDoc
        End If
               
        If Trim(rst3!CLNType) <> " " Then
            cboCLNType = rst3!CLNType
        End If
        
        If Trim(rst3!CLNXRay) <> " " Then
            cboCLNxray = rst3!CLNXRay
        End If
        
        If Trim(rst3!CLNPanel) <> " " Then
            cboCLNpanel = rst3!CLNPanel
        End If
        
        If Trim(rst3!CLNState) <> " " Then
            cboCLNstate = rst3!CLNState
        End If
        
        If Trim(rst3!CLNCity) <> " " Then
            cboCLNcity = rst3!CLNCity
        End If
        
        If Trim(rst3!CLNPostCode) <> " " Then
            txtCLNpostcode = rst3!CLNPostCode
        End If
        
        If Trim(rst3!CLNinitial) <> " " Then
            txtMIXClinic = rst3!CLNinitial
        End If
        
        If rst3!CLNReceivedDate <> "__/__/____" Then
            txtCLNReceivedDate = rst3!CLNReceivedDate
        End If
        
        If Trim(rst3!CLNBankAcc) <> " " Then
            txtCLNBankAcc = rst3!CLNBankAcc
        End If
        
        If Trim(rst3!CLNBankName) <> " " Then
            txtCLNbank = rst3!CLNBankName
        End If
        
        If Trim(rst3!CLNBankBrn) <> " " Then
            txtCLNbranch = rst3!CLNBankBrn
        End If
        
        If Trim(rst3!CLNAccPerson) <> " " Then
            txtCLNAccPerson = rst3!CLNAccPerson
        End If
        
        If Trim(rst3!CLN24Hrs) <> " " Then
            cboCLN24hrs = rst3!CLN24Hrs
        End If
        
        If Trim(rst3!CLNRemark) <> " " Then
            txtCLNremark = rst3!CLNRemark
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
   End If
End Sub

Private Sub lstClinicList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstClinicList.ListItems.Count <> 0 Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        strsql = "Select * from CLN Where CLNCode = '" & lstClinicList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtCLNCode = rst3!CLNCode
        
        If Trim(rst3!CLNname) <> " " Then
            txtCLNName = rst3!CLNname
        End If
        
        If Trim(rst3!ClinicGrpCode) <> " " Then
            CboClinicGrp = rst3!ClinicGrpCode
        End If
        
        If Trim(rst3!CLNAdd1) <> " " Then
            TxtCLNAdd1 = rst3!CLNAdd1
        End If
        
        If Trim(rst3!CLNAdd2) <> " " Then
            TxtCLNAdd2 = rst3!CLNAdd2
        End If
        
        If Trim(rst3!CLNAdd3) <> " " Then
            TxtCLNAdd3 = rst3!CLNAdd3
        End If
        
        If Trim(rst3!CLNPhone) <> " " Then
            txtCLNTelNo = rst3!CLNPhone
        End If
        
        If Trim(rst3!CLNFax) <> " " Then
            txtCLNFaxNo = rst3!CLNFax
        End If
        
        If Trim(rst3!CLNEmail) <> " " Then
            txtCLNEmail = rst3!CLNEmail
        End If
        
        If Trim(rst3!CLNHomepage) <> " " Then
            txtCLNHomePage = rst3!CLNHomepage
        End If
        
        If Trim(rst3!CLNRepresentative) <> " " Then
            TxtCLNContact = rst3!CLNRepresentative
        End If
        
        If Trim(rst3!CLNDocName) <> " " Then
            TxtCLNDocName = rst3!CLNDocName
        End If
        
        If Trim(rst3!CLNBNFType) <> " " Then
            cboCLNBNFType = rst3!CLNBNFType
        End If
        
        If Trim(rst3!CLNForRef) <> " " Then
            txtCLNformema = rst3!CLNForRef
        End If
        
        If Trim(rst3!CLNNoofDoc) <> " " Then
            txtCLNnodoc = rst3!CLNNoofDoc
        End If
               
        If Trim(rst3!CLNType) <> " " Then
            cboCLNType = rst3!CLNType
        End If
        
        If Trim(rst3!CLNXRay) <> " " Then
            cboCLNxray = rst3!CLNXRay
        End If
        
        If Trim(rst3!CLNPanel) <> " " Then
            cboCLNpanel = rst3!CLNPanel
        End If
        
        If Trim(rst3!CLNState) <> " " Then
            cboCLNstate = rst3!CLNState
        End If
        
        If Trim(rst3!CLNCity) <> " " Then
            cboCLNcity = rst3!CLNCity
        End If
        
        If Trim(rst3!CLNPostCode) <> " " Then
            txtCLNpostcode = rst3!CLNPostCode
        End If
        
        If Trim(rst3!CLNinitial) <> " " Then
            txtMIXClinic = rst3!CLNinitial
        End If
        
        If rst3!CLNReceivedDate <> "__/__/____" Then
            txtCLNReceivedDate = rst3!CLNReceivedDate
        End If
        
        If Trim(rst3!CLNBankAcc) <> " " Then
            txtCLNBankAcc = rst3!CLNBankAcc
        End If
        
        If Trim(rst3!CLNBankName) <> " " Then
            txtCLNbank = rst3!CLNBankName
        End If
        
        If Trim(rst3!CLNBankBrn) <> " " Then
            txtCLNbranch = rst3!CLNBankBrn
        End If
        
        If Trim(rst3!CLNAccPerson) <> " " Then
            txtCLNAccPerson = rst3!CLNAccPerson
        End If
        
        If Trim(rst3!CLN24Hrs) <> " " Then
            cboCLN24hrs = rst3!CLN24Hrs
        End If
        
        If Trim(rst3!CLNRemark) <> " " Then
            txtCLNremark = rst3!CLNRemark
        End If
       
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
     '   medixmasterconn.Close
    '    Set medixmasterconn = Nothing
        
   End If
End Sub

Private Sub lstHosGrpList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstHosGrpList.ListItems.Count <> 0 Then
    '    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from HOSGrp Where HOSGrpCode = '" & lstHosGrpList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        TxtHosGrpCode = rst3!HOSGRPCode
        TxtHosGrpName = rst3!HOSGRPName
        If Len(rst3!HOSGrpDesc) Then
            TxtHosGrpDesc = rst3!HOSGrpDesc
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
        'Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstHosGrpList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstHosGrpList.ListItems.Count <> 0 Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from HOSGrp Where HOSGrpCode = '" & lstHosGrpList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        TxtHosGrpCode = rst3!HOSGRPCode
        TxtHosGrpName = rst3!HOSGRPName
        If Len(rst3!HOSGrpDesc) Then
            TxtHosGrpDesc = rst3!HOSGrpDesc
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
    '    medixmasterconn.Close
      '  Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstHospitalList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstHospitalList.ListItems.Count <> 0 Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from HOS Where HOSCode = '" & lstHospitalList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        TxtHOSCode = rst3!HOSCode
        TxtHOSName = rst3!HosName
        
        If Len(rst3!HOSGRPCode) Then
           CboHOSGRP = rst3!HOSGRPCode
        End If
                
        If Len(rst3!HOSAccCode) Then
           txtHOSAccCode = rst3!HOSAccCode
        End If

        If Len(rst3!HOSAdd1) Then
           TxtHOSAdd1 = rst3!HOSAdd1
        End If
        
        If Len(rst3!HOSAdd2) Then
           TxtHOSAdd2 = rst3!HOSAdd2
        End If
        
        If Len(rst3!HOSAdd3) Then
           TxtHOSAdd3 = rst3!HOSAdd3
        End If
        
        If Len(rst3!HOSTelNo) Then
           TxtTelNo = rst3!HOSTelNo
        End If
        
        If Len(rst3!HOSFaxNo) Then
           TxtFaxNo = rst3!HOSFaxNo
        End If
        
        If Len(rst3!HOSEmail) Then
           TxtEmail = rst3!HOSEmail
        End If
        
        If Len(rst3!HOSHomepage) Then
           TxtHomepage = rst3!HOSHomepage
        End If
        
        If Len(rst3!HOSContact) Then
           TxtHosContact = rst3!HOSContact
        End If
        
        If Len(rst3!HOSDocName) Then
           TxtHosDocName = rst3!HOSDocName
        End If
        
        If Len(rst3!HOSRemark) Then
           TxtHosRemark = rst3!HOSRemark
        End If
        
        If Len(rst3!HOSState) Then
           cboHOState = rst3!HOSState
        End If
        
        If Len(rst3!HOSStatus) Then
           cboHOSType = rst3!HOSStatus
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
        'Set medixmasterconn = Nothing
    End If
End Sub

Private Sub SearchHOS()
    Dim StrAppend As String
   
    Screen.MousePointer = vbHourglass
    StrAppend = ""
    
    If Len(Trim(TxtHOSCode)) Then
        StrAppend = StrAppend + " And HOSCode like '%" & ChkStr(TxtHOSCode) & "%'"
    End If
    
    If Len(Trim(TxtHOSName)) Then
        StrAppend = StrAppend + " And HOSName like '%" & ChkStr(TxtHOSName) & "%'"
    End If
    
    If Len(Trim(CboHOSGRP)) Then
        StrAppend = StrAppend + " And HOSGrpCode like '" & RemStr(Mid(CboHOSGRP, 1, IIf(InStr(1, CboHOSGRP, "-") = 0, 4, InStr(1, CboHOSGRP, "-") - 1))) & "%'"
    End If
            
    HOSCriteria = StrAppend
    Call HospitalListing(StrAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstHospitalList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstHospitalList.SortKey = ColumnHeader.Index - 1
lstHospitalList.Sorted = True
If lstHospitalList.SortOrder = lvwAscending Then
lstHospitalList.SortOrder = lvwDescending
Else
lstHospitalList.SortOrder = lvwAscending
End If

End Sub

Private Sub HOSGRPComboListing()
Dim HosGrp As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim STA As New ADODB.Recordset
Dim cmd2 As New ADODB.Command

    Set cmd2.ActiveConnection = medixmasterconn
    cmd2.CommandText = "SearchState"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 15
    STA.CursorLocation = adUseClient
    STA.CursorType = adOpenForwardOnly
    STA.CacheSize = 100
    Set STA = cmd2.Execute
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHospGrp"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HosGrp.CursorLocation = adUseClient
    HosGrp.CursorType = adOpenForwardOnly
    HosGrp.CacheSize = 100
    Set HosGrp = cmd.Execute
    CboHOSGRP.Clear
        
    Do Until HosGrp.EOF
        CboHOSGRP.AddItem HosGrp!HOSGRPCode & " - " & HosGrp!HOSGRPName
        HosGrp.MoveNext
    Loop
        
    HosGrp.Close
    Set HosGrp = Nothing
    
    cboHOState.Clear
    Do Until STA.EOF
        cboHOState.AddItem STA!STADescription
        STA.MoveNext
    Loop
    STA.Close
    Set STA = Nothing
    
    cboHOSType.Clear
    cboHOSType.AddItem "P" & "-" & "Private Hospital"
    cboHOSType.AddItem "G" & "-" & "Government Hospital"
    
End Sub

'Private Sub PrintHOS()
    'RptHOSList.Show
'End Sub
'**********************************End Hospital**************************************

'**********************************Start Payor Info**************************************

Private Sub SavePayor()
On Error Resume Next
Dim RecordSaved
Dim PAY As New ADODB.Recordset
Dim SavePAY As New ADODB.Recordset
Dim strsql As String
    
'    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If cboPAYHLTCode = "" Then
       MsgBox "Please Enter Health Type", vbOKOnly + vbCritical, DialogBoxTitle
       cboPAYHLTCode.SetFocus
    ElseIf txtPAYcode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtPAYcode.SetFocus
    ElseIf CboPAYGRP = "" Then
        MsgBox "Please Enter Payor Group Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboPAYGRP.SetFocus
    ElseIf txtPAYname = "" Then
        MsgBox "Please Enter Payor Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtPAYname.SetFocus
    ElseIf TxtPayAdd1 = "" Then
        MsgBox "Please Enter Payor Address", vbOKOnly + vbCritical, DialogBoxTitle
        TxtPayAdd1.SetFocus
    Else
        strsql = "Select * From PAY Where PAYCode = '" & Trim(txtPAYcode) & "'"
        PAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
        If PAY.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            PAY.Close
            txtPAYcode.SetFocus
            txtPAYcode.SelStart = 0
            txtPAYcode.SelLength = Len(txtPAYcode)
            Set PAY = Nothing
            'Exit Sub
        Else
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                SavePAY.Open "PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
                SavePAY.AddNew
            Else
                strsql = "Select * From PAY Where PAYCode = '" & Trim(txtPAYcode) & "'"
                SavePAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With SavePAY
                    !HLTCode = Trim(Mid(cboPAYHLTCode, 1, 2))
                    !PAYCode = ChkStr(txtPAYcode)
                    !PayGrpCode = Trim(Mid(CboPAYGRP, 1, 3))
                    !PAYCompanyName = ChkStr(txtPAYname)
                    !PAYAddress1 = ChkStr(TxtPayAdd1)
                    !PAYAddress2 = ChkStr(TxtPayAdd2)
                    !PAYAddress3 = ChkStr(TxtPayAdd3)
                    !PAYContactCEO = ChkStr(TxtCEO)
                    !PAYMAINContact = ChkStr(TxtContact)
                    !PAYAdmissionContact = ChkStr(TxtAdmContact)
                    !PAYTELno = ChkStr(txtPAYphone)
                    !PAYFAXNo = ChkStr(txtPAYfax)
                    !PayEmail = ChkStr(txtPAYemail)
                    !PAYHomepage = ChkStr(txtPAYhomepage)
                    !PayEffDate = Format(TxtPAYEffDate, "dd/mm/yyyy")
                    !Conversion = Trim(Mid(CboConversion, 1, 1))
                    !PAYRBTax = Trim(Mid(CboRBTax, 1, 1))
                    !PAYPrePost = Trim(Mid(CboPrePost, 1, 1))
                    !PAYCoPayment = Trim(Mid(CboCoPay, 1, 1))
                    !MBMCancelStatus = Trim(Mid(CboCancelStatus, 1, 1))
                    !PAYLastUpdateUser = Logon
                    !PayLastUpdateDate = Date
                End With
                SavePAY.Update
                Call ClearForm(Me)
                Call PayorListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If

            SavePAY.Close
            Set SavePAY = Nothing
        End If
        EditFlag = False
    End If
End If

'medixmasterconn.Close
'Set medixmasterconn = Nothing

End Sub

Private Sub DeletePayor()
Dim DeletePAY As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim rsCheckPRM As New ADODB.Recordset
Dim strsql, sqlMBM, sqlPRM As String 'sqlPAYGRP As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

EditFlag = False
If txtPAYcode = "" Then
   MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtPAYcode.SetFocus
ElseIf lstPayorList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where PAYCode = '" & RemStr(txtPAYcode) & "'"
   sqlPRM = "select PRMCOde from PRM where PAYCode = '" & RemStr(txtPAYcode) & "'"
   'sqlPAYGRP = "select PAYGRPCode from PAY where PAYCode = '" & RemStr(txtPAYcode) & "'"
   
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   rsCheckPRM.MaxRecords = 1
   rsCheckPRM.Open sqlPRM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   'rsCheckPAYGRP.MaxRecords = 1
   'rsCheckPAYGRP.Open sqlPAYGRP, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   ElseIf rsCheckPRM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Premium Table!", vbOKOnly + vbCritical, "Error Message"
   'ElseIf rsCheckPAYGRP.recordCount Then
        'MsgBox "This record cannot be deleted because it is used in Payor Group Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From PAY Where PAYcode = '" & ChkStr(lstPayorList.SelectedItem.SubItems(1)) & "'"
           DeletePAY.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call PayorListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtPAYcode.SetFocus
         End If
     End If
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing

Exit Sub

End Sub

Private Sub SearchPAY()
    Dim StrAppend As String
   
    Screen.MousePointer = vbHourglass
    StrAppend = ""
    
    If Len(Trim(txtPAYcode)) Then
        StrAppend = StrAppend + " And PAYCode like '%" & Mid(ChkStr(txtPAYcode), 1, 2) & "%'"
    End If
    If Len(Trim(txtPAYname)) Then
        StrAppend = StrAppend + " And PAYCompanyName like '%" & ChkStr(txtPAYname) & "%'"
    End If
    
    If Len(Trim(CboPAYGRP)) Then
        StrAppend = StrAppend + " And PAYGRPCode like '" & RemStr(Mid(CboPAYGRP, 1, IIf(InStr(1, CboPAYGRP, "-") = 0, 4, InStr(1, CboPAYGRP, "-") - 1))) & "%'"
    End If
    
    If Len(Trim(cboPAYHLTCode)) Then
        StrAppend = StrAppend + " AND HLTCode like '" & RemStr(Mid(cboPAYHLTCode, 1, IIf(InStr(1, cboPAYHLTCode, "-") = 0, 4, InStr(1, cboPAYHLTCode, "-") - 1))) & "%'"
    End If
      
    PAYCriteria = StrAppend
    Call PayorListing(StrAppend)
    Screen.MousePointer = Default
End Sub

Private Sub PayorListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim PayMaster As ListItem
Dim sql As String
Dim TotalPayCount As Integer
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    sql = "Select PAYCode, HLTCode, PAYCompanyName, PAYGRPCode, PAYAddress1, " & _
          "PAYAddress2, PAYAddress3, PAYContactCEO, PAYTELno, " & _
          "PAYFAXno, PAYEmail, PAYMAINContact, PAYHOMEPage, PAYAdmissionContact " & _
          "FROM PAY where 0=0 "
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If

    TotalPayCount = 0
    lstPayorList.ListItems.Clear
    
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set PayMaster = lstPayorList.ListItems.Add(, , rst2!HLTCode)
                PayMaster.SubItems(1) = rst2!PAYCode
                PayMaster.SubItems(2) = rst2!PAYCompanyName
                PayMaster.SubItems(3) = IIf(Len(rst2!PayGrpCode), rst2!PayGrpCode, "")
                'PayMaster.SubItems(3) = rst2!PAYGRPName
                rst2.MoveNext
        TotalPayCount = TotalPayCount + 1
        txtTotalRecord = TotalPayCount
        Loop
    rst2.Close
    End If
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub lstHospitalList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstHospitalList.ListItems.Count <> 0 Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from HOS Where HOSCode = '" & lstHospitalList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        TxtHOSCode = rst3!HOSCode
        TxtHOSName = rst3!HosName
        
        If Len(rst3!HOSGRPCode) Then
           CboHOSGRP = rst3!HOSGRPCode
        End If
        
        If Len(rst3!HOSAccCode) Then
           txtHOSAccCode = rst3!HOSAccCode
        End If
        
        If Len(rst3!HOSAdd1) Then
           TxtHOSAdd1 = rst3!HOSAdd1
        End If
        
        If Len(rst3!HOSAdd2) Then
           TxtHOSAdd2 = rst3!HOSAdd2
        End If
        
        If Len(rst3!HOSAdd3) Then
           TxtHOSAdd3 = rst3!HOSAdd3
        End If
        
        If Len(rst3!HOSTelNo) Then
           TxtTelNo = rst3!HOSTelNo
        End If
        
        If Len(rst3!HOSFaxNo) Then
           TxtFaxNo = rst3!HOSFaxNo
        End If
        
        If Len(rst3!HOSEmail) Then
           TxtEmail = rst3!HOSEmail
        End If
        
        If Len(rst3!HOSHomepage) Then
           TxtHomepage = rst3!HOSHomepage
        End If
        
        If Len(rst3!HOSContact) Then
           TxtHosContact = rst3!HOSContact
        End If
        
        If Len(rst3!HOSDocName) Then
           TxtHosDocName = rst3!HOSDocName
        End If
        
        If Len(rst3!HOSRemark) Then
           TxtHosRemark = rst3!HOSRemark
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
        'medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstHospitalList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
        
    If lstHospitalList.ListItems.Count <> 0 Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from HOS Where HOSCode = '" & lstHospitalList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        TxtHOSCode = rst3!HOSCode
        TxtHOSName = rst3!HosName
        
        If Len(rst3!HOSGRPCode) Then
           CboHOSGRP = rst3!HOSGRPCode
        End If
        
        If Len(rst3!HOSAccCode) Then
           txtHOSAccCode = rst3!HOSAccCode
        End If
        
        If Len(rst3!HOSAdd1) Then
           TxtHOSAdd1 = rst3!HOSAdd1
        End If
        
        If Len(rst3!HOSAdd2) Then
           TxtHOSAdd2 = rst3!HOSAdd2
        End If
        
        If Len(rst3!HOSAdd3) Then
           TxtHOSAdd3 = rst3!HOSAdd3
        End If
        
        If Len(rst3!HOSTelNo) Then
           TxtTelNo = rst3!HOSTelNo
        End If
        
        If Len(rst3!HOSFaxNo) Then
           TxtFaxNo = rst3!HOSFaxNo
        End If
        
        If Len(rst3!HOSEmail) Then
           TxtEmail = rst3!HOSEmail
        End If
        
        If Len(rst3!HOSHomepage) Then
           TxtHomepage = rst3!HOSHomepage
        End If
        
        If Len(rst3!HOSContact) Then
           TxtHosContact = rst3!HOSContact
        End If
        
        If Len(rst3!HOSDocName) Then
           TxtHosDocName = rst3!HOSDocName
        End If
        
        If Len(rst3!HOSRemark) Then
           TxtHosRemark = rst3!HOSRemark
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
      '  Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstPAYGRPList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
        
    If lstPAYGRPList.ListItems.Count <> 0 Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from PAYGRP Where PAYGRPCode = '" & lstPAYGRPList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        TxtPayGrpCode = rst3!PayGrpCode
        TxtPayGrpName = rst3!PAYGRPName
        If Len(rst3!PAYGRPDesc) Then
            TxtPayGrpDesc = rst3!PAYGRPDesc
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstPAYGRPList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstPAYGRPList.ListItems.Count <> 0 Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from PAYGRP Where PAYGRPCode = '" & lstPAYGRPList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        TxtPayGrpCode = rst3!PayGrpCode
        TxtPayGrpName = rst3!PAYGRPName
        If Len(rst3!PAYGRPDesc) Then
            TxtPayGrpDesc = rst3!PAYGRPDesc
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstPayorList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstPayorList.SortKey = ColumnHeader.Index - 1
lstPayorList.Sorted = True
If lstPayorList.SortOrder = lvwAscending Then
lstPayorList.SortOrder = lvwDescending
Else
lstPayorList.SortOrder = lvwAscending
End If
End Sub

Private Sub PAYInfoComboListing()
Dim PAYGRP As New ADODB.Recordset
Dim HLT As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cmd2 As New ADODB.Command
        
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    ' Set up a command object for the stored procedure.
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayGrp"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAYGRP.CursorLocation = adUseClient
    PAYGRP.CursorType = adOpenForwardOnly
    PAYGRP.CacheSize = 100
    
    Set cmd2.ActiveConnection = medixmasterconn
    cmd2.CommandText = "SearchHLTType"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 15
    HLT.CursorLocation = adUseClient
    HLT.CursorType = adOpenForwardOnly
    HLT.CacheSize = 100
    ' Create a recordset by executing the command.
    Set PAYGRP = cmd.Execute
    Set HLT = cmd2.Execute
    
    CboPAYGRP.Clear
    cboPAYHLTCode.Clear
    CboCoGrpPayor.Clear
    
        
    Do Until PAYGRP.EOF
        CboPAYGRP.AddItem PAYGRP!PayGrpCode & " - " & PAYGRP!PAYGRPName
        CboCoGrpPayor.AddItem PAYGRP!PayGrpCode & " - " & PAYGRP!PAYGRPName
        PAYGRP.MoveNext
    Loop
        
    
    Do Until HLT.EOF
        cboPAYHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
        HLT.MoveNext
    Loop
            
    PAYGRP.Close
    Set PAYGRP = Nothing
    HLT.Close
    Set HLT = Nothing
    
    CboConversion.Clear
    CboConversion.AddItem "N" & " - " & "No"
    CboConversion.AddItem "Y" & " - " & "Yes"
    
    CboRBTax.Clear
    CboRBTax.AddItem "N" & " - " & "No"
    CboRBTax.AddItem "Y" & " - " & "Yes"
    
    CboPrePost.Clear
    CboPrePost.AddItem "N" & " - " & "No"
    CboPrePost.AddItem "Y" & " - " & "Yes"
    
    CboCoPay.Clear
    CboCoPay.AddItem "N" & " - " & "No"
    CboCoPay.AddItem "Y" & " - " & "Yes"
    
    CboCancelStatus.Clear
    CboCancelStatus.AddItem "N" & " - " & "No"
    CboCancelStatus.AddItem "Y" & " - " & "Yes"
    
'    medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End Sub


Private Sub lstPayorList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
        
    If lstPayorList.ListItems.Count <> 0 Then
    '    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        strsql = "Select * from PAY Where PAYCode = '" & lstPayorList.SelectedItem.SubItems(1) & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ' , adOpenKeyset, adLockOptimistic
        
        cboPAYHLTCode = rst3!HLTCode
        txtPAYcode = rst3!PAYCode
        CboPAYGRP = rst3!PayGrpCode
        txtPAYname = rst3!PAYCompanyName
        
        If Len(Trim(rst3!PAYAddress1)) Then
           TxtPayAdd1 = rst3!PAYAddress1
        End If
                
        If Len(Trim(rst3!PAYAddress2)) Then
           TxtPayAdd2 = rst3!PAYAddress2
        End If
        
        If Len(Trim(rst3!PAYAddress3)) Then
        TxtPayAdd3 = rst3!PAYAddress3
        End If
        
        If Len(Trim(rst3!PAYTELno)) Then
           txtPAYphone = rst3!PAYTELno
        End If
        
        If Len(Trim(rst3!PAYFAXNo)) Then
           txtPAYfax = rst3!PAYFAXNo
        End If
        
        If Len(Trim(rst3!PAYHomepage)) Then
           txtPAYhomepage = rst3!PAYHomepage
        End If
        
        If Len(Trim(rst3!PayEmail)) Then
        txtPAYemail = rst3!PayEmail
        End If
        
        If Len(Trim(rst3!PAYContactCEO)) Then
           TxtCEO = Trim(rst3!PAYContactCEO)
        End If
        
        If Len(Trim(rst3!PAYMAINContact)) Then
           TxtContact = Trim(rst3!PAYMAINContact)
        End If
        
        If Len(Trim(rst3!PAYAdmissionContact)) Then
           TxtAdmContact = Trim(rst3!PAYAdmissionContact)
        End If
        
        If Len(Trim(rst3!PayEffDate)) Then
            TxtPAYEffDate = Format(rst3!PayEffDate, "dd/mm/yyyy")
        End If
        
        If Len(Trim(rst3!Conversion)) Then
            CboConversion = rst3!Conversion
        End If
        
        If Len(Trim(rst3!PAYRBTax)) Then
            CboRBTax = rst3!PAYRBTax
        End If
        
        If Len(Trim(rst3!PAYPrePost)) Then
            CboPrePost = rst3!PAYPrePost
        End If
        
        If Len(Trim(rst3!PAYCoPayment)) Then
            CboCoPay = rst3!PAYCoPayment
        End If
        
        If Len(Trim(rst3!MBMCancelStatus)) Then
            CboCancelStatus = rst3!MBMCancelStatus
        End If
        
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    
        rst3.Close
        Set rst3 = Nothing
        
    '    medixmasterconn.Close
    '    Set medixmasterconn = Nothing
        
    End If
End Sub

'Private Sub PrintPAY()
    'RptPAYList.Show
'End Sub
'**********************************End Payor Info**************************************

'**********************************Start Clinic**************************************
Private Sub SaveClinic()
'On Error Resume Next
Dim RecordSaved
Dim CLN As New ADODB.Recordset
Dim SaveCLN As New ADODB.Recordset
Dim City As New ADODB.Recordset
Dim state As New ADODB.Recordset
Dim strsql As String
Dim str2 As String
Dim str3 As String
Dim CLNSeq As String
Dim StrAppend As String
Dim CityName As String
Dim StateName As String
Dim CLNname As String
Dim DocName As String
Dim CLNinitial As String
Dim contact As String

'If txtCLNCode = "" Then
'    MsgBox "Please Enter Clinic Code", vbOKOnly + vbCritical, DialogBoxTitle
'    txtCLNCode.SetFocus

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

If txtCLNName = "" Then
    MsgBox "Please Enter Clinic Name", vbOKOnly + vbCritical, DialogBoxTitle
    txtCLNName.SetFocus
ElseIf txtMIXClinic = "" Then
    MsgBox "Please Enter Clinic Initial", vbOKOnly + vbCritical, DialogBoxTitle
    txtMIXClinic.SetFocus
ElseIf cboCLNType = "" Then
    MsgBox "Please Enter Clinic Type", vbOKOnly + vbCritical, DialogBoxTitle
    cboCLNType.SetFocus
ElseIf TxtCLNAdd1 = "" Then
    MsgBox "Please Enter Clinic Address", vbOKOnly + vbCritical, DialogBoxTitle
    TxtCLNAdd1.SetFocus
ElseIf cboCLNBNFType = "" Then
    MsgBox " Please Enter Benefit Type", vbOKOnly + vbCritical, DialogBoxTitle
    cboCLNBNFType.SetFocus
ElseIf txtCLNReceivedDate = "__/__/____" Or IsDate(txtCLNReceivedDate) = False Then
    MsgBox "Please Enter Received Date!", vbOKOnly + vbCritical, DialogBoxTitle
    txtCLNReceivedDate.SetFocus
ElseIf cboCLNstate = "" Then
    MsgBox " Please Enter State", vbOKOnly + vbCritical, DialogBoxTitle
    cboCLNstate.SetFocus
ElseIf txtCLNpostcode = "" Then
    MsgBox " Please Enter PostCode", vbOKOnly + vbCritical, DialogBoxTitle
    txtCLNpostcode.SetFocus
ElseIf cboCLNcity = "" Then
    MsgBox " Please Enter City", vbOKOnly + vbCritical, DialogBoxTitle
    cboCLNcity.SetFocus
ElseIf cboCLNpanel = "" Then
    MsgBox " Please Enter Panel", vbOKOnly + vbCritical, DialogBoxTitle
    cboCLNpanel.SetFocus
ElseIf cboCLN24hrs = "" Then
    MsgBox " Please Enter Clinic Operating Hours", vbOKOnly + vbCritical, DialogBoxTitle
    cboCLN24hrs.SetFocus
Else
'      strsql = "Select * From CLN Where CLNCode = '" & Trim(txtCLNCode) & "'"
'      CLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
      
'      If CLN.recordCount > 0 And EditFlag = False Then
'      MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
'      CLN.Close
'      txtCLNCode.SetFocus
'      txtCLNCode.SelStart = 0
'      txtCLNCode.SelLength = Len(txtCLNCode)
'      Set CLN = Nothing
'      Exit Sub
'      Else
      
'        str2 = "Select * from CTY Where CTYCode = '" & Mid(cboCLNcity, 1, 3) & "'"
'        City.Open str2, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
       
'        CityName = City!CTYDescription
        
'        City.Close
'        Set City = Nothing
      
'        str3 = "Select * from STA Where STACode = '" & Mid(cboCLNstate, 1, 2) & "'"
'        State.Open str3, medixmasterconn, adOpenKeyset, adLockOptimistic
        
'        StateName = State!STADescription
        
'        State.Close
'        Set State = Nothing
      
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                CLN.Open "CLN", medixmasterconn, adOpenKeyset, adLockOptimistic
                CLN.AddNew
                CLNSeq = GenerateCLNSeq(Mid(txtMIXClinic, 1, 1))
            Else
                strsql = "Select * From CLN Where CLNCode = '" & Trim(txtCLNCode) & "'"
                CLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                CLNSeq = txtCLNCode
            End If
            
                With CLN
                    'CLNSeq = Mid(txtMIXClinic, 1, 1)
                                       
                    !CLNCode = ChkStr(CLNSeq)
                    'CLNname = Replace(txtCLNName, "'", "''")
                    !CLNname = ChkStr(txtCLNName)
                    !ClinicGrpCode = Trim(Mid(CboClinicGrp, 1, 5))
                    !CLNPhone = ChkStr(txtCLNTelNo)
                    !CLNFax = ChkStr(txtCLNFaxNo)
                    !CLNAdd1 = ChkStr(TxtCLNAdd1)
                    !CLNAdd2 = ChkStr(TxtCLNAdd2)
                    !CLNAdd3 = ChkStr(TxtCLNAdd3)
                    !CLNEmail = ChkStr(txtCLNEmail)
                    !CLNHomepage = ChkStr(txtCLNHomePage)
                    'contact = Replace(TxtCLNContact, "'", "''")
                    !CLNRepresentative = ChkStr(TxtCLNContact)
                    'DocName = Replace(TxtCLNDocName, "'", "''")
                    !CLNDocName = ChkStr(TxtCLNDocName)
                    !CLNBNFType = Trim(Mid(cboCLNBNFType, 1, 1))
                    !CLNForRef = ChkStr(txtCLNformema)
                    If txtCLNnodoc <> "" Then
                        !CLNNoofDoc = txtCLNnodoc
                    End If
                    !CLNType = Trim(Mid(cboCLNType, 1, 1))
                    !CLNXRay = Trim(Mid(cboCLNxray, 1, 1))
                    !CLNPanel = Trim(Mid(cboCLNpanel, 1, 1))
                    !CLN24Hrs = Trim(Mid(cboCLN24hrs, 1, 1))
                    !CLNState = ChkStr(cboCLNstate)
                    !CLNCity = ChkStr(cboCLNcity)
                    !CLNPostCode = ChkStr(txtCLNpostcode)
                    'CLNinitial = Replace(txtMIXClinic, "'", "''")
                    !CLNinitial = ChkStr(txtMIXClinic)
                    !CLNReceivedDate = txtCLNReceivedDate
                    !CLNBankAcc = ChkStr(txtCLNBankAcc)
                    !CLNBankName = ChkStr(txtCLNbank)
                    !CLNBankBrn = ChkStr(txtCLNbranch)
                    !CLNAccPerson = ChkStr(txtCLNAccPerson)
                    !CLNRemark = ChkStr(txtCLNremark)
                    !CLNLastUpdateUser = Logon
                    !CLNLastUpdateDate = Date
                End With
                CLN.Update
                'Call ClearForm(Me)
                Call SearchCLN
                'Call CLNListing(strAppend)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
 
            CLN.Close
            Set SaveCLN = Nothing
        End If
        EditFlag = False
    End If
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
'End If
End Sub

Private Sub DeleteClinic()
Dim DeleteCLN As New ADODB.Recordset
Dim rsCheckMBMOther As New ADODB.Recordset
Dim strsql, sqlMBMOther As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

EditFlag = False
If txtCLNName = "" Then
   MsgBox "Please Enter Clinic Name", vbOKOnly + vbCritical, DialogBoxTitle
   'txtCLNCode.SetFocus
ElseIf lstClinicList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   'sqlMBMOther = "Select MBMNumber from MBMOthers where CLNCode = '" & RemStr(txtCLNCode) & "'"
   
   
   'rsCheckMBMOther.MaxRecords = 1
   'rsCheckMBMOther.Open sqlMBMOther, medixmasterconn, adOpenKeyset, adLockOptimistic
   
     
   'If rsCheckMBMOther.recordCount Then
   '     MsgBox "This record cannot be deleted because it is used in Member Others Table!", vbOKOnly + vbCritical, "Error Message"
   'Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes Then
           strsql = "Delete From CLN Where CLNcode = '" & ChkStr(lstClinicList.SelectedItem.Text) & "'"
           DeleteCLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           lstClinicList.ListItems.Clear
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtCLNName.SetFocus
     End If
     'End If
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
Exit Sub

End Sub

Private Sub SearchCLN()

    Dim StrAppend As String
    Dim cliname As String
    Dim clinitial As String
   
    Screen.MousePointer = vbHourglass
    StrAppend = ""
    cliname = ""
    clinitial = ""
    
    If Len(Trim(txtMIXClinic)) Then
    clinitial = Replace(txtMIXClinic, "'", "''")
        StrAppend = StrAppend + " And CLNInitial like '%" & ChkStr(clinitial) & "%'"
    End If
    
    If Len(Trim(txtCLNName)) Then
        cliname = Replace(txtCLNName, "'", "''")
        StrAppend = StrAppend + " And CLNName like '%" & ChkStr(cliname) & "%'"
    End If
    
    If Len(Trim(cboCLNpanel)) Then
        StrAppend = StrAppend + " AND CLNPanel = '" & RemStr(Mid(cboCLNpanel, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboClinicGrp)) Then
        StrAppend = StrAppend + " AND ClinicGrpCode like '" & RemStr(Mid(CboClinicGrp, 1, IIf(InStr(1, CboClinicGrp, "-") = 0, 4, InStr(1, CboClinicGrp, "-") - 1))) & "%'"
    End If
    
    If Len(Trim(cboCLNType)) Then
        StrAppend = StrAppend + " AND CLNType = '" & RemStr(Mid(cboCLNType, 1, 1)) & "'"
    End If
    
    If Len(Trim(cboCLN24hrs)) Then
        StrAppend = StrAppend + " AND CLN24Hrs = '" & RemStr(Mid(cboCLN24hrs, 1, 1)) & "'"
    End If
    
    CLNCriteria = StrAppend
    Call CLNListing(StrAppend)
    Screen.MousePointer = Default
End Sub

Private Sub CLNListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim CLNMaster As ListItem
Dim sql As String
Dim TotalCLNCount As Integer
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    sql = "Select * from CLN where 0=0"
    
    'sql = "Select CLNCode, CLNName, ClinicGrpCode, CLNAdd1, CLNAdd2, CLNAdd3," & _
          "CLNPhone, CLNFAX, CLNEmail, CLNHOMEPage, CLNRepresentative, CLNDocName, CLNBNFType " & _
          "from CLN where 0=0 "
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If

    TotalCLNCount = 0
    lstClinicList.ListItems.Clear
    
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set CLNMaster = lstClinicList.ListItems.Add(, , rst2!CLNCode)
                If Trim(rst2!CLNname) <> " " Then
                    CLNMaster.SubItems(1) = rst2!CLNname
                End If
                
                If Trim(rst2!ClinicGrpCode) <> " " Then
                    CLNMaster.SubItems(2) = rst2!ClinicGrpCode
                End If
                
                If Trim(rst2!CLNinitial) <> " " Then
                    CLNMaster.SubItems(3) = rst2!CLNinitial
                End If
                
                If Trim(rst2!CLNType) <> " " Then
                    CLNMaster.SubItems(4) = rst2!CLNType
                End If
                
                If Trim(rst2!CLNBNFType) <> " " Then
                    CLNMaster.SubItems(5) = rst2!CLNBNFType
                End If
                
                If Trim(rst2!CLNAdd1) <> " " Then
                    CLNMaster.SubItems(6) = rst2!CLNAdd1
                End If
                
                If Trim(rst2!CLNAdd2) <> " " Then
                    CLNMaster.SubItems(7) = rst2!CLNAdd2
                End If
                
                If Trim(rst2!CLNAdd3) <> " " Then
                    CLNMaster.SubItems(8) = rst2!CLNAdd3
                End If
                
                If Trim(rst2!CLNPostCode) <> " " Then
                    CLNMaster.SubItems(9) = rst2!CLNPostCode
                End If
                
                If Trim(rst2!CLNCity) <> " " Then
                    CLNMaster.SubItems(10) = rst2!CLNCity
                End If
                
                If Trim(rst2!CLNState) <> " " Then
                    CLNMaster.SubItems(11) = rst2!CLNState
                End If
                
                If Trim(rst2!CLNReceivedDate) <> " " Then
                    CLNMaster.SubItems(12) = rst2!CLNReceivedDate
                End If
                
                If Trim(rst2!CLNPanel) <> " " Then
                    CLNMaster.SubItems(13) = rst2!CLNPanel
                End If
                
                If Trim(rst2!CLNXRay) <> " " Then
                    CLNMaster.SubItems(14) = rst2!CLNXRay
                End If
                
                If Trim(rst2!CLNForRef) <> " " Then
                    CLNMaster.SubItems(15) = rst2!CLNForRef
                End If
                
                If Trim(rst2!CLNNoofDoc) <> " " Then
                    CLNMaster.SubItems(16) = rst2!CLNNoofDoc
                End If
                
                If Trim(rst2!CLNPhone) <> " " Then
                    CLNMaster.SubItems(17) = rst2!CLNPhone
                End If
                
                If Trim(rst2!CLNFax) <> " " Then
                    CLNMaster.SubItems(18) = rst2!CLNFax
                End If
                
                If Trim(rst2!CLNEmail) <> " " Then
                    CLNMaster.SubItems(19) = rst2!CLNEmail
                End If
                
                If Trim(rst2!CLNHomepage) <> " " Then
                    CLNMaster.SubItems(20) = rst2!CLNHomepage
                End If
                
                If Trim(rst2!CLNBankAcc) <> " " Then
                    CLNMaster.SubItems(21) = rst2!CLNBankAcc
                End If
                
                If Trim(rst2!CLNBankName) <> " " Then
                    CLNMaster.SubItems(22) = rst2!CLNBankName
                End If
                
                If Trim(rst2!CLNBankBrn) <> " " Then
                    CLNMaster.SubItems(23) = rst2!CLNBankBrn
                End If
                
                If Trim(rst2!CLNAccPerson) <> " " Then
                    CLNMaster.SubItems(24) = rst2!CLNAccPerson
                End If
                
                If Trim(rst2!CLN24Hrs) <> " " Then
                    CLNMaster.SubItems(25) = rst2!CLN24Hrs
                End If
                
                If Trim(rst2!CLNRemark) <> " " Then
                    CLNMaster.SubItems(26) = rst2!CLNRemark
                End If
                
                rst2.MoveNext
        TotalCLNCount = TotalCLNCount + 1
        txtTotalRecord = TotalCLNCount
        Loop
    rst2.Close
    End If
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub CLNComboListing()
Dim CLNGRP As New ADODB.Recordset
Dim CTY As New ADODB.Recordset
Dim STA As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    ' Set up a command object for the stored procedure.
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchState"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    STA.CursorLocation = adUseClient
    STA.CursorType = adOpenForwardOnly
    STA.CacheSize = 100
    
    Set cmd2.ActiveConnection = medixmasterconnMaint
    cmd2.CommandText = "SearchCity"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 15
    CTY.CursorLocation = adUseClient
    CTY.CursorType = adOpenForwardOnly
    CTY.CacheSize = 100
    
    Set cmd3.ActiveConnection = medixmasterconn
    cmd3.CommandText = "SearchClinicGrp"
    cmd3.CommandType = adCmdStoredProc
    cmd3.CommandTimeout = 15
    CLNGRP.CursorLocation = adUseClient
    CLNGRP.CursorType = adOpenForwardOnly
    CLNGRP.CacheSize = 100
    
    ' Create a recordset by executing the command.
    Set STA = cmd.Execute
    Set CTY = cmd2.Execute
    Set CLNGRP = cmd3.Execute
        
    CboClinicGrp.Clear
    cboCLNBNFType.Clear
    cboCLNType.Clear
    cboCLNpanel.Clear
    cboCLNxray.Clear
    cboCLNcity.Clear
    cboCLNstate.Clear
    cboCLN24hrs.Clear
    
    Do Until CLNGRP.EOF
        CboClinicGrp.AddItem CLNGRP!ClinicGrpCode & " - " & CLNGRP!ClinicGrpName
        CLNGRP.MoveNext
    Loop
        
    Do Until CTY.EOF
        cboCLNcity.AddItem CTY!CTYDescription
        CTY.MoveNext
    Loop
        
    Do Until STA.EOF
        cboCLNstate.AddItem STA!STADescription
        STA.MoveNext
    Loop
    
    cboCLNBNFType.AddItem "1 - Standard"
    cboCLNBNFType.AddItem "2 - Clinic"
    cboCLNBNFType.AddItem "3 - Plan"
    cboCLNBNFType.ListIndex = 0
        
    cboCLNType.AddItem "G - General Practisional"
    cboCLNType.AddItem "S - Specialised"
        
    cboCLNxray.AddItem "Y - Yes"
    cboCLNxray.AddItem "N - No"
    
    cboCLNpanel.AddItem "Y - Yes"
    cboCLNpanel.AddItem "N - No"
    cboCLNpanel.AddItem "D - Declined"
    cboCLNpanel.AddItem "C - Confirmed"
    
    cboCLN24hrs.AddItem "Y - Yes"
    cboCLN24hrs.AddItem "N - No"
        
    CLNGRP.Close
    CTY.Close
    STA.Close
    
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub lstClinicList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstClinicList.ListItems.Count <> 0 Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        strsql = "Select * from CLN Where CLNCode = '" & lstClinicList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtCLNCode = rst3!CLNCode
        
        If Trim(rst3!CLNname) <> " " Then
            txtCLNName = rst3!CLNname
        End If
        
        If Trim(rst3!ClinicGrpCode) <> " " Then
            CboClinicGrp = rst3!ClinicGrpCode
        End If
        
        If Trim(rst3!CLNAdd1) <> " " Then
            TxtCLNAdd1 = rst3!CLNAdd1
        End If
        
        If Trim(rst3!CLNAdd2) <> " " Then
            TxtCLNAdd2 = rst3!CLNAdd2
        End If
        
        If Trim(rst3!CLNAdd3) <> " " Then
            TxtCLNAdd3 = rst3!CLNAdd3
        End If
        
        If Trim(rst3!CLNPhone) <> " " Then
            txtCLNTelNo = rst3!CLNPhone
        End If
        
        If Trim(rst3!CLNFax) <> " " Then
            txtCLNFaxNo = rst3!CLNFax
        End If
        
        If Trim(rst3!CLNEmail) <> " " Then
            txtCLNEmail = rst3!CLNEmail
        End If
        
        If Trim(rst3!CLNHomepage) <> " " Then
            txtCLNHomePage = rst3!CLNHomepage
        End If
        
        If Trim(rst3!CLNRepresentative) <> " " Then
            TxtCLNContact = rst3!CLNRepresentative
        End If
        
        If Trim(rst3!CLNDocName) <> " " Then
            TxtCLNDocName = rst3!CLNDocName
        End If
        
        If Trim(rst3!CLNBNFType) <> " " Then
            cboCLNBNFType = rst3!CLNBNFType
        End If
        
        If Trim(rst3!CLNForRef) <> " " Then
            txtCLNformema = rst3!CLNForRef
        End If
        
        If Trim(rst3!CLNNoofDoc) <> " " Then
            txtCLNnodoc = rst3!CLNNoofDoc
        End If
               
        If Trim(rst3!CLNType) <> " " Then
            cboCLNType = rst3!CLNType
        End If
        
        If Trim(rst3!CLNXRay) <> " " Then
            cboCLNxray = rst3!CLNXRay
        End If
        
        If Trim(rst3!CLNPanel) <> " " Then
            cboCLNpanel = rst3!CLNPanel
        End If
        
        If Trim(rst3!CLNState) <> " " Then
            cboCLNstate = rst3!CLNState
        End If
        
        If Trim(rst3!CLNCity) <> " " Then
            cboCLNcity = rst3!CLNCity
        End If
        
        If Trim(rst3!CLNPostCode) <> " " Then
            txtCLNpostcode = rst3!CLNPostCode
        End If
               
        If Trim(rst3!CLNinitial) <> " " Then
            txtMIXClinic = rst3!CLNinitial
        End If
                
        If rst3!CLNReceivedDate <> "__/__/____" Then
            txtCLNReceivedDate = rst3!CLNReceivedDate
        End If
                
        If Trim(rst3!CLNBankAcc) <> " " Then
            txtCLNBankAcc = rst3!CLNBankAcc
        End If
        
        If Trim(rst3!CLNBankName) <> " " Then
            txtCLNbank = rst3!CLNBankName
        End If
        
        If Trim(rst3!CLNBankBrn) <> " " Then
            txtCLNbranch = rst3!CLNBankBrn
        End If
        
        If Trim(rst3!CLNAccPerson) <> " " Then
            txtCLNAccPerson = rst3!CLNAccPerson
        End If
        
        If Trim(rst3!CLN24Hrs) <> " " Then
            cboCLN24hrs = rst3!CLN24Hrs
        End If
        
        If Trim(rst3!CLNRemark) <> " " Then
            txtCLNremark = rst3!CLNRemark
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconn.Close
        'Set medixmasterconn = Nothing
        
   End If
End Sub

Private Sub lstClinicList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstClinicList.SortKey = ColumnHeader.Index - 1
lstClinicList.Sorted = True
If lstClinicList.SortOrder = lvwAscending Then
lstClinicList.SortOrder = lvwDescending
Else
lstClinicList.SortOrder = lvwAscending
End If
End Sub

'Private Sub PrintCLN()
    'RptCLNList.Show
'End Sub
'**********************************End Clinic**************************************

'**********************************Start MCOFees**************************************
'Private Sub SaveMCOFees()
'On Error GoTo SaveErr
'Dim RecordSaved
'Dim MCO As New ADODB.Recordset
'Dim SaveMCO As New ADODB.Recordset
'Dim strsql As String
        
    'If txtINScode = "" Then
        'MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
        'txtINScode.SetFocus
    'ElseIf txtHLTcode = "" Then
        'MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        'txtHLTcode.SetFocus
    'ElseIf txtMCOdescription = "" Then
        'MsgBox "Please Enter MCO Description", vbOKOnly + vbCritical, DialogBoxTitle
        'txtMCOdescription.SetFocus
    'ElseIf txtMCOAmountAdult = "" Then
        'MsgBox "Please Enter MCO Fees Adult", vbOKOnly + vbCritical, DialogBoxTitle
        'txtMCOAmountAdult.SetFocus
    'ElseIf txtMCOAmountChild = "" Then
        'MsgBox "Please Enter MCO Fees Child", vbOKOnly + vbCritical, DialogBoxTitle
        'txtMCOAmountChild.SetFocus
    'ElseIf txtMCOAmountAdult.Text = "" Then
        'MsgBox "Invalid MCO Fees", vbOKOnly + vbCritical, DialogBoxTitle
        'txtMCOAmountAdult.SetFocus
    'ElseIf txtEffdate = "__/__/____" Or Not IsDate(txtEffdate) Then
        'MsgBox "Please validate your MCO Effective date", vbOKOnly + vbCritical, DialogBoxTitle
        'txtEffdate.SetFocus
        
    'Else
      
      'strsql = "Select INSCode, HLTCode From MCO "
      'strsql = strsql & " where (INSCode = '" & Trim(Mid(txtINScode, 1, 1)) & "')  and (HLTCode = '" & Trim(Mid(txtHLTcode, 1, 1)) & "')" & _
                " and MCOEffDate = '" & Format(temptxtEffDate, "mm/dd/yyyy") & "' "
      'MCO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
      'If MCO.recordCount > 0 And EditFlag = False Then
        'MsgBox "Record Duplicated! " & vbNewLine & _
                '" Please Change The Corresponding Field", vbCritical, "Error"
        'MCO.Close
        'txtINScode.SetFocus
        'txtINScode.SelStart = 0
        'txtINScode.SelLength = Len(txtINScode)
        'Set MCO = Nothing
        'Exit Sub
      'Else
        'RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            'If RecordSaved = vbYes Then
                'If EditFlag = False Then
                    'SaveMCO.Open "MCO", medixmasterconn, adOpenKeyset, adLockOptimistic
                    'SaveMCO.AddNew
                'Else
                    'strsql = "Select * From MCO Where HLTCode = '" & Trim(txtHLTcode) & "' And INSCode ='" & Trim(txtINScode) & "'"
                    'If Trim(temptxtEffDate) <> "" Then
                        'strsql = strsql & " and MCOEffDate = '" & Format(temptxtEffDate, "mm/dd/yyyy") & "' "
                    'Else
                        'strsql = strsql & " and (MCOEffDate) IS  NULL "
                    'End If
                    'SaveMCO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                'End If
            
                'With SaveMCO
                    '!INSCode = Mid(txtINScode, 1, 2)
                    '!HLTCode = Mid(txtHLTcode, 1, 2)
                    '!MCODescription = txtMCOdescription
                    'If EditFlag = False Then
                        '!MCOAmountAdult = txtMCOAmountAdult
                        '!MCOAmountChild = txtMCOAmountChild
                        '!MCOEffDate = IIf(txtEffdate = "__/__/____", "", CDate(txtEffdate))
                    'End If
                    '!MCOLastUpdateUser = Logon
                    '!MCOLastUpdateDate = Date
                    
                   ' !MCOExpDate = IIf(txtExpDate = "__/__/____", "", CDate(txtExpDate))
                'End With
                'SaveMCO.Update
                'Call ClearForm(Me)
                'Call MCOListing
                    'If EditFlag = False Then
                        'MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    'Else
                        'MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    'End If
 
                'SaveMCO.Close
                'Set SaveMCO = Nothing
            'End If
        'EditFlag = False
   'End If
  'End If
  'Exit Sub
'SaveErr:
    'MsgBox ERR.Description
'End Sub

'Private Sub DeleteMCOFees()

'Dim DeleteMCO As New ADODB.Recordset
'Dim rsCheckMBM As New ADODB.Recordset
'Dim strsql, sqlMBM As String
'Dim DeleteRecord

'EditFlag = False
'If txtINScode = "" Then
   'MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtINScode.SetFocus
'ElseIf lstMCOList.listitems.Count = 0 Then
   'MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
'Else
   'sqlMBM = "Select MBMNumber from MBM where INSCode = '" & ChkStr(lstMCOList.SelectedItem.Text) & "'" & _
            " and MBMPayorEffdate >='" & Format(lstMCOList.SelectedItem.SubItems(6), "mm/dd/yyyy") & "'"
   
   'rsCheckMBM.MaxRecords = 1
    'rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
     
   'If rsCheckMBM.recordCount Then
        'MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
   'Else
   'DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     'If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           'strsql = "Delete From MCO Where MCOcode = '" & ChkStr(lstMCOList.SelectedItem.Text) & "'"
           'DeleteMCO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           'MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           'cmdSave.Enabled = True
           'Call ClearForm(Me)
           'Call MCOListing
           'Call DisableCmdButton(Me)
           'Call EntriesEnable(True)
           'txtINScode.SetFocus
         'End If
     'End If
'End If
'Exit Sub
'End Sub

'Private Sub SearchMCO()
    'Dim strAppend As String
   
    'strAppend = ""
    
    'If Len(Trim(txtINScode)) Then
        'strAppend = strAppend + " And INSCode = '" & Mid(txtINScode, 1, 2) & "'"
    'End If
    'If Len(Trim(txtHLTcode)) Then
        'strAppend = strAppend + " And HLTCode = '" & Mid(txtHLTcode, 1, 2) & "'"
    'End If

    'MCOCriteria = strAppend
    'Call MCOListing(strAppend)
'End Sub

'Private Sub MCOListing(Optional strAppend As String)

    'Dim rst2 As New ADODB.Recordset
    'Dim MCOMaster As ListItem
    'Dim sql As String
        
    
    'lstMCOList.listitems.clear
    'sql = "Select MCOCOde , INSCode, HLTCode, " & _
      "MCODescription, MCOAmountAdult, MCOAmountChild , MCOEffDate " & _
      "from MCO where 0=0 "
    
    'If Len(Trim(strAppend)) Then
    'sql = sql + strAppend
    'End If

    'rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'If rst2.recordCount = 0 Then
        'MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    'Else
        'Do Until rst2.EOF
            'Set MCOMaster = lstMCOList.listitems.Add(, , rst2!INSCode)
                'MCOMaster.SubItems(1) = rst2!HLTCode
                'MCOMaster.SubItems(2) = rst2!MCODescription
                'MCOMaster.SubItems(3) = Format(rst2!MCOAmountAdult, "#,##0.00")
                'MCOMaster.SubItems(4) = Format(rst2!MCOAmountChild, "#,##0.00")
                'MCOMaster.SubItems(5) = IIf(Len(Trim(rst2!MCOEffDate)) >= 1, rst2!MCOEffDate, "")
                'rst2.MoveNext
        'Loop
    'rst2.Close
    'Set rst2 = Nothing
    'End If
'End Sub

'Private Sub MCOComboListing()

    'Dim INS As New ADODB.Recordset
    'Dim HLT As New ADODB.Recordset
    
    'INS.Open "INS", medixmasterconn, adOpenKeyset, adLockOptimistic
    'HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'txtINScode.clear
    'txtHLTcode.clear
    
    'Do Until INS.EOF
        'txtINScode.AddItem INS!INSCode & " - " & INS!INSDescription
        'INS.MoveNext
    'Loop
    'Do Until HLT.EOF
        'txtHLTcode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
        'HLT.MoveNext
    'Loop
    
    'INS.Close
    'HLT.Close
'End Sub

'Private Sub lstMCOList_Click()
'On Error Resume Next
   
    'Call EnableCmdButton(Me)
    'Call ClearForm(Me)
    
  
    'txtHLTcode = lstMCOList.SelectedItem.SubItems(1)
    'txtINScode = lstMCOList.SelectedItem.Text
    'txtMCOdescription = lstMCOList.SelectedItem.SubItems(2)
    'txtMCOAmountAdult = lstMCOList.SelectedItem.SubItems(3)
    'txtMCOAmountChild = lstMCOList.SelectedItem.SubItems(4)
    'txtEffdate = lstMCOList.SelectedItem.SubItems(5)
    'temptxtEffDate = lstMCOList.SelectedItem.SubItems(5)
    'txtExpDate = rst3!MCOExpDate
    'Call EntriesEnable(False)
    'cmdSave.Enabled = False
'End Sub

'Private Sub lstMCOList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'lstMCOList.SortKey = ColumnHeader.Index - 1
'lstMCOList.Sorted = True
'If lstMCOList.SortOrder = lvwAscending Then
'lstMCOList.SortOrder = lvwDescending
'Else
'lstMCOList.SortOrder = lvwAscending
'End If

'End Sub
'**********************************End MCOFees**************************************

'**********************************Start Service Provider**********************************

Private Sub SaveServiceProvider()
On Error Resume Next
Dim RecordSaved
Dim SRV As New ADODB.Recordset
Dim SaveSRV As New ADODB.Recordset
Dim strsql As String
    
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
   If txtSRVcode = "" Then
       MsgBox "Please Enter Service Provider Code", vbOKOnly + vbCritical, DialogBoxTitle
       txtSRVcode.SetFocus
    Else
    If txtSRVname = "" Then
        MsgBox "Please Enter Service Provider Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtSRVname.SetFocus
    Else
    If TxtSRVAdd1 = "" Then
        MsgBox "Please Enter Service Provider Address", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSRVAdd1.SetFocus
        Else
      strsql = "Select * From SRV Where SRVCode = '" & Trim(txtSRVcode) & "'"
                SRV.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
      If SRV.recordCount > 0 And EditFlag = False Then
      MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
      SRV.Close
      txtSRVcode.SetFocus
      txtSRVcode.SelStart = 0
      txtSRVcode.SelLength = Len(txtSRVcode)
      Set SRV = Nothing
      'Exit Sub
      Else
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                SRV.Open "SRV", medixmasterconn, adOpenKeyset, adLockOptimistic
                SRV.AddNew
            Else
                strsql = "Select * From SRV Where SRVCode = '" & Trim(txtSRVcode) & "'"
                SRV.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With SRV
                    !SRVCode = ChkStr(txtSRVcode)
                    !SRVname = ChkStr(txtSRVname)
                    !SRVaddress1 = ChkStr(TxtSRVAdd1)
                    !SRVaddress2 = ChkStr(TxtSRVAdd2)
                    !SRVaddress3 = ChkStr(TxtSRVAdd3)
                    !ContactPerson = ChkStr(TxtSRVContact)
                    !SRVTelNo1 = ChkStr(txtTelNo1)
                    !SRVTelNo2 = ChkStr(txtTelNo2)
                    !SRVFaxNo1 = ChkStr(txtFaxNo1)
                    !SRVFaxNo2 = ChkStr(txtFaxNo2)
                    !TollFree = ChkStr(txtTollFree)
                    !SRVRemark = ChkStr(txtRemark)
                    !SRVLastUpdateUser = Logon
                    !SRVLastUpdateDate = Date
                End With
                SRV.Update
                Call ClearForm(Me)
                txtSRVcode.Enabled = True
                Call ServiceProviderListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
 
            SRV.Close
            Set SaveSRV = Nothing
        End If
        EditFlag = False
    End If
    End If
    End If
    End If
    
 '   medixmasterconn.Close
 '   Set medixmasterconn = Nothing
    
End Sub

Private Sub DeleteServiceProvider()
Dim rsDeleteSRV As New ADODB.Recordset
Dim rsCheckCAS As New ADODB.Recordset
Dim rsCheckPRV As New ADODB.Recordset
Dim rscheckUSR As New ADODB.Recordset
Dim strsql, sqlCAS, sqlPRV, sqlUSR As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

EditFlag = False

If txtSRVcode = "" Then
   MsgBox "Please Enter Service Provider Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtSRVcode.SetFocus
ElseIf lstServiceProviderList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else

sqlCAS = "Select CASNumber from CAS where SRVCode = '" & RemStr(txtSRVcode) & "'"
sqlPRV = "Select PRVIndex from PRV where SRVCode = '" & RemStr(txtSRVcode) & "'"
sqlUSR = "Select USRCode from USR where SRVCode = '" & RemStr(txtSRVcode) & "'"


rsCheckCAS.MaxRecords = 1
rsCheckCAS.Open sqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic

rsCheckPRV.MaxRecords = 1
rsCheckPRV.Open sqlPRV, medixmasterconn, adOpenKeyset, adLockOptimistic

rscheckUSR.MaxRecords = 1
rscheckUSR.Open sqlUSR, medixmasterconn, adOpenKeyset, adLockOptimistic

If rsCheckCAS.recordCount Then
    MsgBox "This record cannot be deleted because it is used in CAS Table!", vbOKOnly + vbCritical, "Error Message"
ElseIf rsCheckPRV.recordCount Then
    MsgBox "This record cannot be deleted because it is used in PRV Table!", vbOKOnly + vbCritical, "Error Message"
Else
    DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes And rsCheckCAS.recordCount = 0 Then
        strsql = "Delete from SRV where SRVcode = '" & ChkStr(lstServiceProviderList.SelectedItem.Text) & "'"
        rsDeleteSRV.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        Call ClearForm(Me)
        Call ServiceProviderListing(SRVCriteria)
        Call setCmdButton(Me, False)
        Call EntriesEnable(True)
        txtSRVcode.SetFocus
        cmdAdd.Enabled = True
    End If
  End If
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
Exit Sub



End Sub

Private Sub SearchServiceProvider()
    Dim StrAppend As String
   
    Screen.MousePointer = vbHourglass
    StrAppend = ""
    
    If Len(Trim(txtSRVcode)) Then
        StrAppend = StrAppend + " And SRVCode like '%" & Mid(ChkStr(txtSRVcode), 1, 2) & "%'"
    End If
    If Len(Trim(txtSRVname)) Then
        StrAppend = StrAppend + " And SRVName LIKE '%" & txtSRVname & "%'"
    End If
    
    SRVCriteria = StrAppend
    Call ServiceProviderListing(StrAppend)
    Screen.MousePointer = Default
End Sub

Private Sub ServiceProviderListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim SRVMaster As ListItem
Dim sql As String
Dim TotalServiceCount As Integer
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
    sql = "Select SRVCode, SRVName, SRVAddress1, SRVAddress2, SRVAddress3, " & _
          "ContactPerson, SRVTelNo1, SRVTelNo2, SRVFaxNo1, SRVFaxNo2, " & _
          "TollFree, SRVRemark " & _
          "from SRV where 0=0 "
          
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If
    
    TotalServiceCount = 0
    lstServiceProviderList.ListItems.Clear
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set SRVMaster = lstServiceProviderList.ListItems.Add(, , rst2!SRVCode)
                SRVMaster.SubItems(1) = rst2!SRVname
                rst2.MoveNext
        TotalServiceCount = TotalServiceCount + 1
        txtTotalRecord = TotalServiceCount
        Loop
    rst2.Close
    End If
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub lstPayorList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstPayorList.ListItems.Count <> 0 Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        strsql = "Select * from PAY Where PAYCode = '" & lstPayorList.SelectedItem.SubItems(1) & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ' , adOpenKeyset, adLockOptimistic
        
        cboPAYHLTCode = rst3!HLTCode
        txtPAYcode = rst3!PAYCode
        CboPAYGRP = rst3!PayGrpCode
        txtPAYname = rst3!PAYCompanyName
        
        If Len(Trim(rst3!PAYAddress1)) Then
            TxtPayAdd1 = rst3!PAYAddress1
        End If
                
        If Len(Trim(rst3!PAYAddress2)) Then
           TxtPayAdd2 = rst3!PAYAddress2
        End If
        
        If Len(Trim(rst3!PAYAddress3)) Then
        TxtPayAdd3 = rst3!PAYAddress3
        End If
        
        If Len(Trim(rst3!PAYTELno)) Then
           txtPAYphone = rst3!PAYTELno
        End If
        
        If Len(Trim(rst3!PAYFAXNo)) Then
           txtPAYfax = rst3!PAYFAXNo
        End If
        
        If Len(Trim(rst3!PAYHomepage)) Then
           txtPAYhomepage = rst3!PAYHomepage
        End If
        
        If Len(Trim(rst3!PayEmail)) Then
        txtPAYemail = rst3!PayEmail
        End If
        
        If Len(Trim(rst3!PAYContactCEO)) Then
           TxtCEO = rst3!PAYContactCEO
        End If
        
        If Len(Trim(rst3!PAYMAINContact)) Then
           TxtContact = rst3!PAYMAINContact
        End If
        
        If Len(Trim(rst3!PAYAdmissionContact)) Then
           TxtAdmContact = rst3!PAYAdmissionContact
        End If
        
        If Len(Trim(rst3!PayEffDate)) Then
            TxtPAYEffDate = Format(rst3!PayEffDate, "dd/mm/yyyy")
        End If
        
        If Len(Trim(rst3!Conversion)) Then
            CboConversion = rst3!Conversion
        End If
        
        If Len(Trim(rst3!PAYRBTax)) Then
            CboRBTax = rst3!PAYRBTax
        End If
        
        If Len(Trim(rst3!PAYPrePost)) Then
            CboPrePost = rst3!PAYPrePost
        End If
        
        If Len(Trim(rst3!PAYCoPayment)) Then
            CboCoPay = rst3!PAYCoPayment
        End If
        
        If Len(Trim(rst3!MBMCancelStatus)) Then
            CboCancelStatus = rst3!MBMCancelStatus
        End If
        
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstPayorList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
        
    If lstPayorList.ListItems.Count <> 0 Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
           
        strsql = "Select * from PAY Where PAYCode = '" & lstPayorList.SelectedItem.SubItems(1) & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        cboPAYHLTCode = rst3!HLTCode
        txtPAYcode = rst3!PAYCode
        CboPAYGRP = rst3!PayGrpCode
        txtPAYname = rst3!PAYCompanyName
        
        If Len(Trim(rst3!PAYAddress1)) Then
           TxtPayAdd1 = rst3!PAYAddress1
        End If
        
        If Len(Trim(rst3!PAYAddress2)) Then
           TxtPayAdd2 = rst3!PAYAddress2
        End If
        
        If Len(Trim(rst3!PAYAddress3)) Then
        TxtPayAdd3 = rst3!PAYAddress3
        End If
        
        If Len(Trim(rst3!PAYTELno)) Then
           txtPAYphone = rst3!PAYTELno
        End If
        
        If Len(Trim(rst3!PAYFAXNo)) Then
           txtPAYfax = rst3!PAYFAXNo
        End If
        
        If Len(Trim(rst3!PAYHomepage)) Then
           txtPAYhomepage = rst3!PAYHomepage
        End If
        
        If Len(Trim(rst3!PayEmail)) Then
        txtPAYemail = rst3!PayEmail
        End If
        
        If Len(Trim(rst3!PAYContactCEO)) Then
           TxtCEO = rst3!PAYContactCEO
        End If
        
        If Len(Trim(rst3!PAYMAINContact)) Then
           TxtContact = rst3!PAYMAINContact
        End If
        
        If Len(Trim(rst3!PAYAdmissionContact)) Then
           TxtAdmContact = rst3!PAYAdmissionContact
        End If
        
        If Len(Trim(rst3!PayEffDate)) Then
            TxtPAYEffDate = Format(rst3!PayEffDate, "dd/mm/yyyy")
        End If
        
        If Len(Trim(rst3!Conversion)) Then
            CboConversion = rst3!Conversion
        End If
        
        If Len(Trim(rst3!PAYRBTax)) Then
            CboRBTax = rst3!PAYRBTax
        End If
        
        If Len(Trim(rst3!PAYPrePost)) Then
            CboPrePost = rst3!PAYPrePost
        End If
        
        If Len(Trim(rst3!PAYCoPayment)) Then
            CboCoPay = rst3!PAYCoPayment
        End If
        
        If Len(Trim(rst3!MBMCancelStatus)) Then
            CboCancelStatus = rst3!MBMCancelStatus
        End If
        
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    
        rst3.Close
        Set rst3 = Nothing
        
     '   medixmasterconn.Close
     '   Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstPRVList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String

    If lstPRVList.ListItems.Count <> 0 Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        
        'strsql = "Select * from PRV Where SRVCode = '" & lstPRVList.SelectedItem.Text & "' And INSCode = '" & lstPRVList.SelectedItem.SubItems(2) & "' And HLTCode = '" & lstPRVList.SelectedItem.SubItems(1) & "'"
        strsql = "Select * from PRV Where PRVIndex  = '" & lstPRVList.SelectedItem.SubItems(6) & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtPRVCode = rst3!SRVCode
        'txtINSCode1 = lstPRVList.SelectedItem.SubItems(2)
        txtINSCode1 = Trim(rst3!INSCode)
        txtHealthCode = rst3!HLTCode
        txtPRVProviderCharges = Format(rst3!PRVProviderCharges, "#,##0.00")
        txtPRVCaseCharges = Format(rst3!PRVCaseCharges, "#,##.00")
        txtAccessDate = Format(rst3!AccessDate, "dd/mm/yyyy")
        If Len(rst3!AccessVersion) Then
            txtAccVer = rst3!AccessVersion
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstPRVList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String

    If lstPRVList.ListItems.Count <> 0 Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        
        'strsql = "Select * from PRV Where SRVCode = '" & lstPRVList.SelectedItem.Text & "' And INSCode = '" & lstPRVList.SelectedItem.SubItems(2) & "' And HLTCode = '" & lstPRVList.SelectedItem.SubItems(1) & "'"
        strsql = "Select * from PRV Where PRVIndex  = '" & lstPRVList.SelectedItem.SubItems(6) & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtPRVCode = rst3!SRVCode
        'txtINSCode1 = lstPRVList.SelectedItem.SubItems(2)
        txtINSCode1 = Trim(rst3!INSCode)
        txtHealthCode = rst3!HLTCode
        txtPRVProviderCharges = Format(rst3!PRVProviderCharges, "#,##0.00")
        txtPRVCaseCharges = Format(rst3!PRVCaseCharges, "#,##.00")
        txtAccessDate = Format(rst3!AccessDate, "dd/mm/yyyy")
        If Len(rst3!AccessVersion) Then
            txtAccVer = rst3!AccessVersion
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
        'medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstServiceProviderList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstServiceProviderList.SortKey = ColumnHeader.Index - 1
    lstServiceProviderList.Sorted = True
If lstServiceProviderList.SortOrder = lvwAscending Then
    lstServiceProviderList.SortOrder = lvwDescending
Else
    lstServiceProviderList.SortOrder = lvwAscending
End If
End Sub

Private Sub lstServiceProviderList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstServiceProviderList.ListItems.Count <> 0 Then
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from SRV Where SRVCode = '" & lstServiceProviderList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtSRVcode = rst3!SRVCode
        txtSRVname = rst3!SRVname
        TxtSRVAdd1 = rst3!SRVaddress1
        If Len(rst3!SRVaddress2) Then
           TxtSRVAdd2 = rst3!SRVaddress2
        End If
        If Len(rst3!SRVaddress3) Then
           TxtSRVAdd3 = rst3!SRVaddress3
        End If
        If Len(rst3!ContactPerson) Then
           TxtSRVContact = rst3!ContactPerson
        End If
        If Len(rst3!SRVTelNo1) Then
           txtTelNo1 = rst3!SRVTelNo1
        End If
        If Len(rst3!SRVTelNo2) Then
           txtTelNo2 = rst3!SRVTelNo2
        End If
        If Len(rst3!SRVFaxNo1) Then
           txtFaxNo1 = rst3!SRVFaxNo1
        End If
        If Len(rst3!SRVFaxNo2) Then
           txtFaxNo2 = rst3!SRVFaxNo2
        End If
        If Len(rst3!TollFree) Then
           txtTollFree = rst3!TollFree
        End If
        If Len(rst3!SRVRemark) Then
           txtRemark = rst3!SRVRemark
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
        
    End If
End Sub

'Private Sub PrintSRV()
    'RptSRVList.Show
'End Sub
'**********************************End Service Provider******************************

'**********************************Start Provider Charge**************************************

Private Sub SaveProviderCharge()
On Error Resume Next
Dim RecordSaved
Dim PRV As New ADODB.Recordset
Dim SavePRV As New ADODB.Recordset
Dim strsql As String
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If txtPRVCode = "" Then
        MsgBox "Please Enter Service Provider Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtSRVcode.SetFocus
    ElseIf txtHealthCode = "" Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtHealthCode.SetFocus
    ElseIf txtINSCode1 = "" Then
        MsgBox "Please Enter Insured Type", vbOKOnly + vbCritical, DialogBoxTitle
        txtINSCode1.SetFocus
    ElseIf txtPRVProviderCharges = "" Then
        MsgBox "Please Enter Provider Charges", vbOKOnly + vbCritical, DialogBoxTitle
        txtPRVProviderCharges.SetFocus
    ElseIf txtPRVCaseCharges = "" Then
        MsgBox "Please Enter Case Charges", vbOKOnly + vbCritical, DialogBoxTitle
        txtPRVCaseCharges.SetFocus
    ElseIf txtAccessDate = "__/__/____" Then
        MsgBox "Please Enter Access Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtAccessDate.SetFocus
    Else
      
      strsql = "Select SRVCode, INSCode, HLTCode, PRVProviderCharges, PRVCaseCharges From PRV"
      strsql = strsql & " Where (SRVCode = '" & Trim(Mid(txtPRVCode, 1, 2)) & "') and (HLTCode = '" & Trim(Mid(txtHealthCode, 1, 1)) & "')"
      strsql = strsql & " and (INSCode = '" & Trim(Mid(txtINSCode1, 1, 1)) & "')"
      strsql = strsql & " and (AccessDate = '" & Trim(txtAccessDate) & "')"
      
      PRV.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
      If PRV.recordCount > 0 And EditFlag = False Then
      MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
      PRV.Close
      txtPRVCode.SetFocus
      txtPRVCode.SelStart = 0
      txtPRVCode.SelLength = Len(txtPRVCode)
      Set PRV = Nothing
      'Exit Sub
      Else
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                PRV.Open "PRV", medixmasterconn, adOpenKeyset, adLockOptimistic
                PRV.AddNew
            Else
                strsql = "Select * From PRV Where SRVCode = '" & Trim(txtPRVCode) & "' And INSCode = '" & Trim(txtINSCode1) & "' And HLTCode = '" & Trim(txtHealthCode) & "'"
                PRV.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With PRV
                    !SRVCode = Trim(Mid(txtPRVCode, 1, 2))
                    !HLTCode = Trim(Mid(txtHealthCode, 1, 1))
                    !INSCode = Trim(Mid(txtINSCode1, 1, 2))
                    !PRVProviderCharges = Format(txtPRVProviderCharges, "#,##0.00")
                    !PRVCaseCharges = Format(txtPRVCaseCharges, "#,##0.00")
                    !AccessDate = txtAccessDate
                    !PRVLastUpdateUser = Logon
                    !PRVLastUpdateDate = Date
                End With
                PRV.Update
                Call ClearForm(Me)
                Call PRVListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                
            PRV.Close
            Set SavePRV = Nothing
        End If
        EditFlag = False
    End If
  End If
  
'medixmasterconn.Close
'Set medixmasterconn = Nothing
  
End Sub

Private Sub DeleteProviderCharge()
Dim rsDeletePRV As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord
EditFlag = False

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
If txtPRVCode = "" Then
   MsgBox "Please Enter Service Provider Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtPRVCode.SetFocus
ElseIf lstPRVList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
sqlMBM = "Select MBMNumber from MBM where SRVCodeList = '" & ChkStr(lstPRVList.SelectedItem.Text) & "'" 'SubItems(1)) & "'"
'& _ " and INSCode = '" & ChkStr(lstPRVList.SelectedItem.SubItems(2)) & "'"

rsCheckMBM.MaxRecords = 1
rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic

If rsCheckMBM.recordCount Then
    MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
Else
    DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
        strsql = "Delete from PRV where PRVIndex = " & Trim(lstPRVList.SelectedItem.SubItems(6))
        rsDeletePRV.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        Call ClearForm(Me)
        Call PRVListing(BRNCriteria)
        Call setCmdButton(Me, False)
        Call EntriesEnable(True)
        txtPRVCode.SetFocus

        cmdAdd.Enabled = True
    End If
  End If
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
Exit Sub

End Sub

Private Sub PRVListing(Optional StrAppend As String)
'On Error Resume Next
Dim rst2 As New ADODB.Recordset
Dim PRVMaster As ListItem
Dim sql As String
Dim TotalPRVCount As Integer

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
    sql = "Select PRVIndex, SRVCode, INSCode, HLTCode, " & _
        "PRVProviderCharges, PRVCaseCharges, AccessDate, AccessVersion " & _
        "from PRV where 0=0 "
          
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If

    TotalPRVCount = 0
    lstPRVList.ListItems.Clear
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set PRVMaster = lstPRVList.ListItems.Add(, , rst2!SRVCode)
                PRVMaster.SubItems(1) = rst2!HLTCode
                PRVMaster.SubItems(2) = rst2!INSCode
                PRVMaster.SubItems(3) = Format(rst2!PRVProviderCharges, "#,##0.00")
                PRVMaster.SubItems(4) = Format(rst2!PRVCaseCharges, "#,##.00")
                PRVMaster.SubItems(5) = Format(rst2!AccessDate, "dd/mm/yyyy")
                PRVMaster.SubItems(6) = rst2!PRVIndex
                PRVMaster.SubItems(7) = rst2!AccessVersion
                rst2.MoveNext
        TotalPRVCount = TotalPRVCount + 1
        txtTotalRecord = TotalPRVCount
        Loop
    rst2.Close
    End If
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub PRVChargesComboListing()
Dim SRV As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim HLT As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
    ' Set up a command object for the stored procedure.
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchSRV"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    SRV.CursorLocation = adUseClient
    SRV.CursorType = adOpenForwardOnly
    SRV.CacheSize = 100
    
    Set cmd2.ActiveConnection = medixmasterconn
    cmd2.CommandText = "SearchHLTType"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 15
    HLT.CursorLocation = adUseClient
    HLT.CursorType = adOpenForwardOnly
    HLT.CacheSize = 100
    
    Set cmd3.ActiveConnection = medixmasterconn
    cmd3.CommandText = "SearchINS"
    cmd3.CommandType = adCmdStoredProc
    cmd3.CommandTimeout = 15
    INS.CursorLocation = adUseClient
    INS.CursorType = adOpenForwardOnly
    INS.CacheSize = 100
    
    ' Create a recordset by executing the command.
    Set SRV = cmd.Execute
    Set HLT = cmd2.Execute
    Set INS = cmd3.Execute
    
    txtPRVCode.Clear
    txtINSCode1.Clear
    txtHealthCode.Clear
    
    Do Until SRV.EOF
        txtPRVCode.AddItem SRV!SRVCode & " - " & SRV!SRVname
        SRV.MoveNext
    Loop
       
    Do Until INS.EOF
        txtINSCode1.AddItem INS!INSCode & " - " & INS!INSDescription
        INS.MoveNext
    Loop
    
    Do Until HLT.EOF
        txtHealthCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
        HLT.MoveNext
    Loop
    
    SRV.Close
    Set SRV = Nothing
    INS.Close
    Set INS = Nothing
    HLT.Close
    Set HLT = Nothing
        
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing

End Sub

Private Sub lstPRVList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String

    If lstPRVList.ListItems.Count <> 0 Then
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        
        'strsql = "Select * from PRV Where SRVCode = '" & lstPRVList.SelectedItem.Text & "' And INSCode = '" & lstPRVList.SelectedItem.SubItems(2) & "' And HLTCode = '" & lstPRVList.SelectedItem.SubItems(1) & "'"
        strsql = "Select * from PRV Where PRVIndex  = '" & lstPRVList.SelectedItem.SubItems(6) & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtPRVCode = rst3!SRVCode
        'txtINSCode1 = lstPRVList.SelectedItem.SubItems(2)
        txtINSCode1 = Trim(rst3!INSCode)
        txtHealthCode = rst3!HLTCode
        txtPRVProviderCharges = Format(rst3!PRVProviderCharges, "#,##0.00")
        txtPRVCaseCharges = Format(rst3!PRVCaseCharges, "#,##.00")
        txtAccessDate = Format(rst3!AccessDate, "dd/mm/yyyy")
        If Len(rst3!AccessVersion) Then
            txtAccVer = rst3!AccessVersion
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub SearchPRVCharges()
    Dim StrAppend As String
    
    Screen.MousePointer = vbHourglass
    If Len(Trim(txtPRVCode)) Then
        StrAppend = StrAppend + " AND SRVCode like '" & RemStr(Mid(txtPRVCode, 1, IIf(InStr(1, txtPRVCode, "-") = 0, 4, InStr(1, txtPRVCode, "-") - 1))) & "%'"
        'strAppend = strAppend + " And SRVCode like '%" & Mid(ChkStr(txtPRVCode), 1, 2) & "%'"
    End If
    If Len(Trim(txtINSCode1)) Then
        StrAppend = StrAppend + " AND INSCode like '" & RemStr(Mid(txtINSCode1, 1, IIf(InStr(1, txtINSCode1, "-") = 0, 4, InStr(1, txtINSCode1, "-") - 1))) & "%'"
        'strAppend = strAppend + " And INSCode = '" & Mid(txtINSCode1, 1, 2) & "'"
    End If
    If Len(Trim(txtHealthCode)) Then
        StrAppend = StrAppend + " AND HLTCode like '" & RemStr(Mid(txtHealthCode, 1, IIf(InStr(1, txtHealthCode, "-") = 0, 4, InStr(1, txtHealthCode, "-") - 1))) & "%'"
        'strAppend = strAppend + " And HLTCode = '" & Mid(txtHealthCode, 1, 2) & "'"
    End If

    PRVChargesCriteria = StrAppend
    Call PRVListing(StrAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstPRVList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstPRVList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPRVList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPRVList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPRVList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstPRVList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstPRVList, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    End Select
End Sub

'Private Sub PrintPRV()
    'RptPRVList.Show
'End Sub
'**********************************End Provider Charge**************************************

'**********************************Start Payor Group**************************************
Private Sub SavePAYGRP()
'On Error Resume Next
Dim RecordSaved
Dim PAYGRP As New ADODB.Recordset
Dim SavePAYGRP As New ADODB.Recordset
Dim strsql As String
            
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            
    If TxtPayGrpCode = "" Then
        MsgBox "Please Enter Payor Group Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtPayGrpCode.SetFocus
    ElseIf TxtPayGrpName = "" Then
        MsgBox "Please Enter Payor Group Name", vbOKOnly + vbCritical, DialogBoxTitle
        TxtPayGrpName.SetFocus
    Else
        
        strsql = "Select * From PAYGRP Where PAYGRPCode = '" & Trim(TxtPayGrpCode) & "'"
        PAYGRP.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
      
        If PAYGRP.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            PAYGRP.Close
            TxtPayGrpCode.SetFocus
            TxtPayGrpCode.SelStart = 0
            TxtPayGrpCode.SelLength = Len(TxtPayGrpCode)
            Set PAYGRP = Nothing
            'Exit Sub
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                'If PAYGRP.recordCount = 0 Then
                    SavePAYGRP.Open "PAYGRP", medixmasterconn, adOpenKeyset, adLockOptimistic
                    SavePAYGRP.AddNew
                Else
                    strsql = "Select * From PAYGRP Where PAYGRPCode = '" & Trim(TxtPayGrpCode) & "'"
                    SavePAYGRP.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                End If
            
                With SavePAYGRP
                    !PayGrpCode = ChkStr(TxtPayGrpCode)
                    !PAYGRPName = ChkStr(TxtPayGrpName)
                    !PAYGRPDesc = ChkStr(TxtPayGrpDesc)
                    !PAYGRPLastUpdatedUser = Logon
                    !PAYGRPLastUpdatedDate = Date
                End With
                SavePAYGRP.Update
                Call ClearForm(Me)
                Call PAYGRPListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                
            SavePAYGRP.Close
            Set SavePAYGRP = Nothing
        End If
        EditFlag = False
    End If
End If

''medixmasterconn.Close
'Set medixmasterconn = Nothing

End Sub

Private Sub DeletePAYGRP()

Dim DeletePAYGRP As New ADODB.Recordset
Dim rsCheckPAY As New ADODB.Recordset
Dim strsql, sqlPAY As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

EditFlag = False
If TxtPayGrpCode = "" Then
   MsgBox "Please Enter Payor Group Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtPayGrpCode.SetFocus
ElseIf lstPAYGRPList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlPAY = "Select PAYCode from PAY where PAYGRPCode = '" & RemStr(TxtPayGrpCode) & "'"
   
   rsCheckPAY.MaxRecords = 1
   rsCheckPAY.Open sqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   If rsCheckPAY.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Payor Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckPAY.recordCount = 0 Then
           strsql = "Delete From PAYGRP Where PAYGRPCode = '" & ChkStr(lstPAYGRPList.SelectedItem.Text) & "'"
           DeletePAYGRP.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call PAYGRPListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           TxtPayGrpCode.SetFocus
         End If
     End If
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
Exit Sub
    
End Sub

Private Sub PAYGRPListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim PAYGRPMaster As ListItem
Dim sql As String
Dim TotalPayGrpCount As Integer
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    sql = "Select PAYGRPCode, PAYGRPName, PAYGRPDesc " & _
          "from PAYGRP where 0=0 "
          
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If
              
    TotalPayGrpCount = 0
    lstPAYGRPList.ListItems.Clear
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set PAYGRPMaster = lstPAYGRPList.ListItems.Add(, , rst2!PayGrpCode)
                PAYGRPMaster.SubItems(1) = rst2!PAYGRPName
                rst2.MoveNext
        TotalPayGrpCount = TotalPayGrpCount + 1
        txtTotalRecord = TotalPayGrpCount
        Loop
    rst2.Close
    End If
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub lstPAYGRPList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
        
    If lstPAYGRPList.ListItems.Count <> 0 Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from PAYGRP Where PAYGRPCode = '" & lstPAYGRPList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        TxtPayGrpCode = rst3!PayGrpCode
        TxtPayGrpName = rst3!PAYGRPName
        If Len(rst3!PAYGRPDesc) Then
            TxtPayGrpDesc = rst3!PAYGRPDesc
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
        'Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub SearchPAYGRP()
    Dim StrAppend As String
   
    Screen.MousePointer = vbHourglass
    StrAppend = ""
    
    If Len(Trim(TxtPayGrpCode)) Then
        StrAppend = StrAppend + " And PAYGRPCode like '%" & ChkStr(TxtPayGrpCode) & "%'"
    End If
    If Len(Trim(TxtPayGrpName)) Then
        StrAppend = StrAppend + " And PAYGRPName like '%" & ChkStr(TxtPayGrpName) & "%'"
    End If
    
    PAYGRPCriteria = StrAppend
    Call PAYGRPListing(StrAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstPAYGRPList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

lstPAYGRPList.SortKey = ColumnHeader.Index - 1
lstPAYGRPList.Sorted = True
If lstPAYGRPList.SortOrder = lvwAscending Then
lstPAYGRPList.SortOrder = lvwDescending
Else
lstPAYGRPList.SortOrder = lvwAscending
End If
End Sub

'Private Sub PrintPAYGRP()
    'RptPayGrpList.Show
'End Sub

'Private Sub PrintAgent()
    'RptAgentList.Show
'End Sub
'**********************************End Payor Group**************************************

'**********************************Start Group Company**************************************
Private Sub SaveCompanyGrp()
Dim RecordSaved
Dim CompanyGrp As New ADODB.Recordset
Dim SaveCompanyGrp As New ADODB.Recordset
Dim CISCompanyGrp As New ADODB.Recordset
Dim SaveCISCompanyGrp As New ADODB.Recordset
Dim strsql As String
Dim StrSql2 As String
Dim COMSeq As String

  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            
    If txtCompanyName = "" Then
        MsgBox "Please Enter Company Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtCompanyName.SetFocus
    'ElseIf CboCoGrpPayor = "" Then
        'MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        'CboCoGrpPayor.SetFocus
    ElseIf txtCoAdd1 = "" Then
        MsgBox "Please Enter Group Company Address", vbOKOnly + vbCritical, DialogBoxTitle
        txtCoAdd1.SetFocus
    ElseIf txtCoEffDate = "__/__/____" Then
        MsgBox "Please Enter Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtCoEffDate.SetFocus
    ElseIf CboGrpComStatus = "" Then
        MsgBox "Please Enter Status", vbOKOnly + vbCritical, DialogBoxTitle
        CboGrpComStatus.SetFocus
    Else
        strsql = "Select * From CompanyGRP Where CoGRPName = '" & Trim(txtCompanyName) & "'"
        CompanyGrp.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        StrSql2 = "Select * From GRPCompany Where GRPCoName = '" & Trim(txtCompanyName) & "'"
        CISCompanyGrp.Open StrSql2, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If (CompanyGrp.recordCount > 0 Or CISCompanyGrp.recordCount > 0) And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            CompanyGrp.Close
            CISCompanyGrp.Close
            txtCompanyName.SetFocus
            txtCompanyName.SelStart = 0
            txtCompanyName.SelLength = Len(txtCompanyName)
            Set CompanyGrp = Nothing
            Set CISCompanyGrp = Nothing
            'Exit Sub
        Else
            CompanyGrp.Close
            Set CompanyGrp = Nothing
            CISCompanyGrp.Close
            Set CISCompanyGrp = Nothing
            
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    If Mid(CboGrpComStatus, 1, 1) = "B" Then
                        CompanyGrp.Open "CompanyGRP", medixmasterconn, adOpenKeyset, adLockOptimistic
                        CompanyGrp.AddNew
                        CISCompanyGrp.Open "GRPCompany", MediConnCISDB, adOpenKeyset, adLockOptimistic
                        CISCompanyGrp.AddNew
                    ElseIf Mid(CboGrpComStatus, 1, 1) = "I" Then
                        CompanyGrp.Open "CompanyGRP", medixmasterconn, adOpenKeyset, adLockOptimistic
                        CompanyGrp.AddNew
                    ElseIf Mid(CboGrpComStatus, 1, 1) = "O" Then
                        CISCompanyGrp.Open "GRPCompany", MediConnCISDB, adOpenKeyset, adLockOptimistic
                        CISCompanyGrp.AddNew
                    End If
                    COMSeq = GenerateCoGrpSeq(Mid(txtCompanyName, 1, 1))
                Else
                    If Mid(CboGrpComStatus, 1, 1) = "O" Then
                        StrSql2 = "Select * From GRPCompany Where GRPCoCode = '" & Trim(txtGrpCoCode) & "'"
                        CISCompanyGrp.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                        COMSeq = txtGrpCoCode
                    ElseIf Mid(CboGrpComStatus, 1, 1) = "I" Then
                        strsql = "Select * From CompanyGRP Where CoGRPCode = '" & Trim(txtGrpCoCode) & "'"
                        CompanyGrp.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        COMSeq = txtGrpCoCode
                    Else
                        strsql = "Select * From CompanyGRP Where CoGRPCode = '" & Trim(txtGrpCoCode) & "'"
                        CompanyGrp.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        COMSeq = txtGrpCoCode
                        
                        StrSql2 = "Select * From GRPCompany Where GRPCoCode = '" & Trim(txtGrpCoCode) & "'"
                        CISCompanyGrp.Open StrSql2, MediConnCISDB, adOpenKeyset, adLockOptimistic
                        COMSeq = txtGrpCoCode
                    End If
                End If
                
                    If Mid(CboGrpComStatus, 1, 1) = "B" Then
                    
                        With CompanyGrp
                            !CoGrpCode = ChkStr(COMSeq)
                            !CoGRPName = ChkStr(txtCompanyName)
                            !CoGRPPaycode = Mid(CboCoGrpPayor, 1, 2)
                            !CoGRPAdd1 = IIf(Len(txtCoAdd1), ChkStr(txtCoAdd1), Null)
                            !CoGRPAdd2 = IIf(Len(txtCoAdd2), ChkStr(txtCoAdd2), Null)
                            !CoGRPAdd3 = IIf(Len(txtCoAdd3), ChkStr(txtCoAdd3), Null)
                            !CoGRPAdmContact = IIf(Len(txtCoAdmContact), ChkStr(txtCoAdmContact), Null)
                            !CoGRPContact = IIf(Len(txtCoContact), ChkStr(txtCoContact), Null)
                            !CoGRPCEO = IIf(Len(txtCoCEO), ChkStr(txtCoCEO), Null)
                            !CoGRPPhone = IIf(Len(txtCoPhone), ChkStr(txtCoPhone), Null)
                            !CoGRPFax = IIf(Len(txtCoFax), ChkStr(txtCoFax), Null)
                            !CoGRPEmail = IIf(Len(txtCoEmail), ChkStr(txtCoEmail), Null)
                            !CoGRPHomepage = IIf(Len(txtCoHomepage), ChkStr(txtCoHomepage), Null)
                            !CoGRPEffDate = IIf(IsDate(txtCoEffDate), txtCoEffDate, Null)
                            !CoGRPTax = Mid(cboCoTax, 1, 1)
                            !CoGRPConversion = Mid(cboCoConversion, 1, 1)
                            !CoGRPPrePost = Mid(cboCoConversion, 1, 1)
                            !CoGRPStatus = Mid(CboGrpComStatus, 1, 1)
                            !BusinessChanel = IIf(Len(cmbChannel), Mid(cmbChannel, 1, 1), Null)
                            !CoGrpTISM = IIf(Len(cmbTims), Mid(cmbTims, 1, 1), Null)
                            !CoGRPLastUpdateUser = Logon
                            !CoGRPLastUpdateDate = Date
                        End With
                        CompanyGrp.Update
                        CompanyGrp.Close
                        Set CompanyGrp = Nothing
                        With CISCompanyGrp
                            !GRPCoCode = ChkStr(COMSeq)
                            !GRPCoName = ChkStr(txtCompanyName)
                            !GRPCoPaycode = Mid(CboCoGrpPayor, 1, 2)
                            !GRPCoAdd1 = IIf(Len(txtCoAdd1), ChkStr(txtCoAdd1), Null)
                            !GRPCoAdd2 = IIf(Len(txtCoAdd2), ChkStr(txtCoAdd2), Null)
                            !GRPCoAdd3 = IIf(Len(txtCoAdd3), ChkStr(txtCoAdd3), Null)
                            !GRPCoAdmContact = IIf(Len(txtCoAdmContact), ChkStr(txtCoAdmContact), Null)
                            !GRPCoContact = IIf(Len(txtCoContact), ChkStr(txtCoContact), Null)
                            !GRPCoCEO = IIf(Len(txtCoCEO), ChkStr(txtCoCEO), Null)
                            !GRPCoPhone = IIf(Len(txtCoPhone), ChkStr(txtCoPhone), Null)
                            !GRPCoFax = IIf(Len(txtCoFax), ChkStr(txtCoFax), Null)
                            !GRPCoEmail = IIf(Len(txtCoEmail), ChkStr(txtCoEmail), Null)
                            !GRPCoHomepage = IIf(Len(txtCoHomepage), ChkStr(txtCoHomepage), Null)
                            !GRPCoEffDate = IIf(IsDate(txtCoEffDate), txtCoEffDate, Null)
                            !GRPCoTax = Mid(cboCoTax, 1, 1)
                            !GRPCoConversion = Mid(cboCoConversion, 1, 1)
                            !GRPCoPrePost = Mid(cboCoConversion, 1, 1)
                            !GRPCoStatus = Mid(CboGrpComStatus, 1, 1)
                            !BusinessChanel = IIf(Len(cmbChannel), Mid(cmbChannel, 1, 1), Null)
                            !GRPTIMS = IIf(Len(cmbTims), Mid(cmbTims, 1, 1), Null)
                            !GRPCoLastUpdateUser = Logon
                            !GRPCoLastUpdateDate = Date
                        End With
                        CISCompanyGrp.Update
                        CISCompanyGrp.Close
                        Set CISCompanyGrp = Nothing
                    ElseIf Mid(CboGrpComStatus, 1, 1) = "I" Then
                        With CompanyGrp
                            !CoGrpCode = ChkStr(COMSeq)
                            !CoGRPName = ChkStr(txtCompanyName)
                            !CoGRPPaycode = Mid(CboCoGrpPayor, 1, 2)
                            !CoGRPAdd1 = IIf(Len(txtCoAdd1), ChkStr(txtCoAdd1), Null)
                            !CoGRPAdd2 = IIf(Len(txtCoAdd2), ChkStr(txtCoAdd2), Null)
                            !CoGRPAdd3 = IIf(Len(txtCoAdd3), ChkStr(txtCoAdd3), Null)
                            !CoGRPAdmContact = IIf(Len(txtCoAdmContact), ChkStr(txtCoAdmContact), Null)
                            !CoGRPContact = IIf(Len(txtCoContact), ChkStr(txtCoContact), Null)
                            !CoGRPCEO = IIf(Len(txtCoCEO), ChkStr(txtCoCEO), Null)
                            !CoGRPPhone = IIf(Len(txtCoPhone), ChkStr(txtCoPhone), Null)
                            !CoGRPFax = IIf(Len(txtCoFax), ChkStr(txtCoFax), Null)
                            !CoGRPEmail = IIf(Len(txtCoEmail), ChkStr(txtCoEmail), Null)
                            !CoGRPHomepage = IIf(Len(txtCoHomepage), ChkStr(txtCoHomepage), Null)
                            !CoGRPEffDate = IIf(IsDate(txtCoEffDate), txtCoEffDate, Null)
                            !CoGRPTax = Mid(cboCoTax, 1, 1)
                            !CoGRPConversion = Mid(cboCoConversion, 1, 1)
                            !CoGRPPrePost = Mid(cboCoConversion, 1, 1)
                            !CoGRPStatus = Mid(CboGrpComStatus, 1, 1)
                            !BusinessChanel = IIf(Len(cmbChannel), Mid(cmbChannel, 1, 1), Null)
                            !CoGrpTISM = IIf(Len(cmbTims), Mid(cmbTims, 1, 1), Null)
                            !CoGRPLastUpdateUser = Logon
                            !CoGRPLastUpdateDate = Date
                        End With
                        CompanyGrp.Update
                        CompanyGrp.Close
                        Set CompanyGrp = Nothing
                    ElseIf Mid(CboGrpComStatus, 1, 1) = "O" Then
                        With CISCompanyGrp
                            !GRPCoCode = ChkStr(COMSeq)
                            !GRPCoName = ChkStr(txtCompanyName)
                            !GRPCoPaycode = Mid(CboCoGrpPayor, 1, 2)
                            !GRPCoAdd1 = IIf(Len(txtCoAdd1), ChkStr(txtCoAdd1), Null)
                            !GRPCoAdd2 = IIf(Len(txtCoAdd2), ChkStr(txtCoAdd2), Null)
                            !GRPCoAdd3 = IIf(Len(txtCoAdd3), ChkStr(txtCoAdd3), Null)
                            !GRPCoAdmContact = IIf(Len(txtCoAdmContact), ChkStr(txtCoAdmContact), Null)
                            !GRPCoContact = IIf(Len(txtCoContact), ChkStr(txtCoContact), Null)
                            !GRPCoCEO = IIf(Len(txtCoCEO), ChkStr(txtCoCEO), Null)
                            !GRPCoPhone = IIf(Len(txtCoPhone), ChkStr(txtCoPhone), Null)
                            !GRPCoFax = IIf(Len(txtCoFax), ChkStr(txtCoFax), Null)
                            !GRPCoEmail = IIf(Len(txtCoEmail), ChkStr(txtCoEmail), Null)
                            !GRPCoHomepage = IIf(Len(txtCoHomepage), ChkStr(txtCoHomepage), Null)
                            !GRPCoEffDate = IIf(IsDate(txtCoEffDate), txtCoEffDate, Null)
                            !GRPCoTax = Mid(cboCoTax, 1, 1)
                            !GRPCoConversion = Mid(cboCoConversion, 1, 1)
                            !GRPCoPrePost = Mid(cboCoConversion, 1, 1)
                            !GRPCoStatus = Mid(CboGrpComStatus, 1, 1)
                            !BusinessChanel = IIf(Len(cmbChannel), Mid(cmbChannel, 1, 1), Null)
                            !GRPTIMS = IIf(Len(cmbTims), Mid(cmbTims, 1, 1), Null)
                            !GRPCoLastUpdateUser = Logon
                            !GRPCoLastUpdateDate = Date
                        End With
                        CISCompanyGrp.Update
                        CISCompanyGrp.Close
                        Set CISCompanyGrp = Nothing
                        
                    End If
                        Call ClearForm(Me)
                        Call CompanyGRPListing
                        If EditFlag = False Then
                            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                        Else
                            MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                        End If
            End If
            EditFlag = False
        End If
    End If
End Sub

Private Sub DeleteCompanyGrp()
Dim DeleteCoGrp As New ADODB.Recordset
Dim rsChkCo As New ADODB.Recordset
Dim strsql, sqlCompany As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

EditFlag = False
If txtGrpCoCode = "" Then
   MsgBox "Please Enter Company Group Code", vbOKOnly + vbCritical, DialogBoxTitle
   txtGrpCoCode.SetFocus
ElseIf lstCompanyGrpList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCompany = "Select CoGRPCode from CompanyGRP where CoGRPCode = '" & RemStr(txtGrpCoCode) & "'"
   
   rsChkCo.MaxRecords = 1
   rsChkCo.Open sqlCompany, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   If rsChkCo.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Company Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsChkCo.recordCount = 0 Then
           strsql = "Delete From CompanyGRP Where CoGRPCode = '" & ChkStr(lstCompanyGrpList.SelectedItem.Text) & "'"
           DeleteCoGrp.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call CompanyGRPListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtGrpCoCode.SetFocus
         End If
     End If
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing
Exit Sub

End Sub

Private Sub CompanyGRPListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim CompanyGrpMaster As ListItem
Dim sql As String
Dim TotalGRPCOMCount As Integer
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
    sql = "Select DISTINCT CoGrpCode, CoGRPName, CoGRPPaycode,BusinessChanel ,Decr,TIMSClinet from VIEW_GRPCompany where 0 = 0 "
   
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If
              
    TotalGRPCOMCount = 0
    lstCompanyGrpList.ListItems.Clear
    
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set CompanyGrpMaster = lstCompanyGrpList.ListItems.Add(, , rst2!CoGrpCode)
                CompanyGrpMaster.SubItems(1) = rst2!CoGRPName
                If Len(rst2!CoGRPPaycode) Then
                    CompanyGrpMaster.SubItems(2) = rst2!CoGRPPaycode
                End If
                If Len(rst2!BusinessChanel) Then
                    CompanyGrpMaster.SubItems(3) = rst2!BusinessChanel + " - " + rst2!Decr
                End If
                If Len(rst2!TIMSClinet) Then
                    CompanyGrpMaster.SubItems(4) = rst2!TIMSClinet
                Else
                    CompanyGrpMaster.SubItems(4) = "N"
                End If
                rst2.MoveNext
        TotalGRPCOMCount = TotalGRPCOMCount + 1
        txtTotalRecord = TotalGRPCOMCount
        Loop
    rst2.Close
    End If
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing

End Sub

Private Sub CompanyGRPComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
        
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    ' Set up a command object for the stored procedure.
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
        
    ' Create a recordset by executing the command.
    Set PAY = cmd.Execute
        
    CboCoGrpPayor.Clear

    Do Until PAY.EOF
        CboCoGrpPayor.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        PAY.MoveNext
    Loop
            
    PAY.Close
    Set PAY = Nothing
    
    cboCoConversion.Clear
    cboCoConversion.AddItem "N" & " - " & "No"
    cboCoConversion.AddItem "Y" & " - " & "Yes"
    
    cboCoTax.Clear
    cboCoTax.AddItem "N" & " - " & "No"
    cboCoTax.AddItem "Y" & " - " & "Yes"
    
    cboCoPrePost.Clear
    cboCoPrePost.AddItem "N" & " - " & "No"
    cboCoPrePost.AddItem "Y" & " - " & "Yes"
    
    CboGrpComStatus.Clear
    CboGrpComStatus.AddItem "B" & " - " & "Both"
    CboGrpComStatus.AddItem "I" & " - " & "In-Patient"
    CboGrpComStatus.AddItem "O" & " - " & "Out-Patient"
    
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End Sub

Private Sub lstCompanyGrpList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstCompanyGrpList.ListItems.Count <> 0 Then
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from VIEW_GRPCompany Where CoGRPCode = '" & lstCompanyGrpList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtGrpCoCode = rst3!CoGrpCode
        txtCompanyName = rst3!CoGRPName
        
        If Len(rst3!CoGRPPaycode) Then
            CboCoGrpPayor = rst3!CoGRPPaycode
        End If
        If Len(rst3!CoGRPAdd1) Then
            txtCoAdd1 = rst3!CoGRPAdd1
        End If
        If Len(rst3!CoGRPAdd2) Then
            txtCoAdd2 = rst3!CoGRPAdd2
        End If
        If Len(rst3!CoGRPAdd3) Then
            txtCoAdd3 = rst3!CoGRPAdd3
        End If
        If Len(rst3!CoGRPAdmContact) Then
            txtCoAdmContact = rst3!CoGRPAdmContact
        End If
        If Len(rst3!CoGRPContact) Then
            txtCoContact = rst3!CoGRPContact
        End If
        If Len(rst3!CoGRPCEO) Then
            txtCoCEO = rst3!CoGRPCEO
        End If
        If Len(rst3!CoGRPPhone) Then
            txtCoPhone = rst3!CoGRPPhone
        End If
        If Len(rst3!CoGRPFax) Then
            txtCoFax = rst3!CoGRPFax
        End If
        If Len(rst3!CoGRPEmail) Then
            txtCoEmail = rst3!CoGRPEmail
        End If
        If Len(rst3!CoGRPHomepage) Then
            txtCoHomepage = rst3!CoGRPHomepage
        End If
        If Len(rst3!CoGRPEffDate) Then
            txtCoEffDate = rst3!CoGRPEffDate
        End If
        If Len(rst3!CoGRPTax) Then
            cboCoTax = rst3!CoGRPTax
        End If
        If Len(rst3!CoGRPConversion) Then
            cboCoConversion = rst3!CoGRPConversion
        End If
        If Len(rst3!CoGRPPrePost) Then
            cboCoPrePost = rst3!CoGRPPrePost
        End If
        If Len(rst3!CoGRPStatus) Then
            CboGrpComStatus = rst3!CoGRPStatus
        End If
        If Len(rst3!BusinessChanel) Then
            cmbChannel = rst3!BusinessChanel + " - " + rst3!Decr
        End If
        If Len(rst3!TIMSClinet) Then
            cmbTims = rst3!TIMSClinet
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstCompanyGrpList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstCompanyGrpList.ListItems.Count <> 0 Then
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from VIEW_GRPCompany Where CoGRPCode = '" & lstCompanyGrpList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtGrpCoCode = rst3!CoGrpCode
        txtCompanyName = rst3!CoGRPName
        
        If Len(rst3!CoGRPPaycode) Then
            CboCoGrpPayor = rst3!CoGRPPaycode
        End If
'        txtCompanyName = rst3!GRPCoInitial
        If Len(rst3!CoGRPAdd1) Then
            txtCoAdd1 = rst3!CoGRPAdd1
        End If
        If Len(rst3!CoGRPAdd2) Then
            txtCoAdd2 = rst3!CoGRPAdd2
        End If
        If Len(rst3!CoGRPAdd3) Then
            txtCoAdd3 = rst3!CoGRPAdd3
        End If
        If Len(rst3!CoGRPAdmContact) Then
            txtCoAdmContact = rst3!CoGRPAdmContact
        End If
        If Len(rst3!CoGRPContact) Then
            txtCoContact = rst3!CoGRPContact
        End If
        If Len(rst3!CoGRPCEO) Then
            txtCoCEO = rst3!CoGRPCEO
        End If
        If Len(rst3!CoGRPPhone) Then
            txtCoPhone = rst3!CoGRPPhone
        End If
        If Len(rst3!CoGRPFax) Then
            txtCoFax = rst3!CoGRPFax
        End If
        If Len(rst3!CoGRPEmail) Then
            txtCoEmail = rst3!CoGRPEmail
        End If
        If Len(rst3!CoGRPHomepage) Then
            txtCoHomepage = rst3!CoGRPHomepage
        End If
        If Len(rst3!CoGRPEffDate) Then
            txtCoEffDate = rst3!CoGRPEffDate
        End If
        If Len(rst3!CoGRPTax) Then
            cboCoTax = rst3!CoGRPTax
        End If
        If Len(rst3!CoGRPConversion) Then
            cboCoConversion = rst3!CoGRPConversion
        End If
        If Len(rst3!CoGRPPrePost) Then
            cboCoPrePost = rst3!CoGRPPrePost
        End If
        If Len(rst3!CoGRPStatus) Then
            CboGrpComStatus = rst3!CoGRPStatus
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
End Sub

Private Sub lstCompanyGrpList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstCompanyGrpList.ListItems.Count <> 0 Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from VIEW_GRPCompany Where CoGRPCode = '" & lstCompanyGrpList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtGrpCoCode = rst3!CoGrpCode
        txtCompanyName = rst3!CoGRPName
        
        If Len(rst3!CoGRPPaycode) Then
            CboCoGrpPayor = rst3!CoGRPPaycode
        End If
'        txtCompanyName = rst3!GRPCoInitial
        If Len(rst3!CoGRPAdd1) Then
            txtCoAdd1 = rst3!CoGRPAdd1
        End If
        If Len(rst3!CoGRPAdd2) Then
            txtCoAdd2 = rst3!CoGRPAdd2
        End If
        If Len(rst3!CoGRPAdd3) Then
            txtCoAdd3 = rst3!CoGRPAdd3
        End If
        If Len(rst3!CoGRPAdmContact) Then
            txtCoAdmContact = rst3!CoGRPAdmContact
        End If
        If Len(rst3!CoGRPContact) Then
            txtCoContact = rst3!CoGRPContact
        End If
        If Len(rst3!CoGRPCEO) Then
            txtCoCEO = rst3!CoGRPCEO
        End If
        If Len(rst3!CoGRPPhone) Then
            txtCoPhone = rst3!CoGRPPhone
        End If
        If Len(rst3!CoGRPFax) Then
            txtCoFax = rst3!CoGRPFax
        End If
        If Len(rst3!CoGRPEmail) Then
            txtCoEmail = rst3!CoGRPEmail
        End If
        If Len(rst3!CoGRPHomepage) Then
            txtCoHomepage = rst3!CoGRPHomepage
        End If
        If Len(rst3!CoGRPEffDate) Then
            txtCoEffDate = rst3!CoGRPEffDate
        End If
        If Len(rst3!CoGRPTax) Then
            cboCoTax = rst3!CoGRPTax
        End If
        If Len(rst3!CoGRPConversion) Then
            cboCoConversion = rst3!CoGRPConversion
        End If
        If Len(rst3!CoGRPPrePost) Then
            cboCoPrePost = rst3!CoGRPPrePost
        End If
        If Len(rst3!CoGRPStatus) Then
            CboGrpComStatus = rst3!CoGRPStatus
        End If
        
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconn.Close
     '   Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub SearchCompanyGrp()
Dim StrAppend As String
   
    Screen.MousePointer = vbHourglass
    StrAppend = ""
    
    If Len(Trim(txtGrpCoCode)) Then
        StrAppend = StrAppend + " And CoGRPCode like '%" & ChkStr(txtGrpCoCode) & "%'"
    End If
    
    If Len(Trim(txtCompanyName)) Then
        StrAppend = StrAppend + " And CoGRPName like '%" & ChkStr(txtCompanyName) & "%'"
    End If
    
    GRPCoCriteria = StrAppend
    Call CompanyGRPListing(StrAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstCompanyGrpList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

lstCompanyGrpList.SortKey = ColumnHeader.Index - 1
lstCompanyGrpList.Sorted = True
    If lstCompanyGrpList.SortOrder = lvwAscending Then
        lstCompanyGrpList.SortOrder = lvwDescending
    Else
        lstCompanyGrpList.SortOrder = lvwAscending
    End If
End Sub

'Private Sub PrintCompanyGrp()
'    RptCompanyGrpList.show
'End Sub
'**********************************End Group Company**************************************

'**********************************Start Hospital Group**************************************
Private Sub SaveHOSGrp()
Dim RecordSaved
Dim HosGrp As New ADODB.Recordset
Dim SaveHOSGrp As New ADODB.Recordset
Dim strsql As String
        
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If TxtHosGrpCode = "" Then
        MsgBox "Please Enter Hospital Group Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtHosGrpCode.SetFocus
    ElseIf TxtHosGrpName = "" Then
        MsgBox "Please Enter Hospital Group Name", vbOKOnly + vbCritical, DialogBoxTitle
        TxtHosGrpName.SetFocus
    Else
        strsql = "Select * From HOSGrp Where HOSGrpCode = '" & Trim(TxtHosGrpCode) & "'"
        HosGrp.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
      If HosGrp.recordCount > 0 And EditFlag = False Then
      MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
      HosGrp.Close
      TxtHosGrpCode.SetFocus
      TxtHosGrpCode.SelStart = 0
      TxtHosGrpCode.SelLength = Len(TxtHosGrpCode)
      Set HosGrp = Nothing
      'Exit Sub
      Else
      HosGrp.Close
      Set HosGrp = Nothing
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                HosGrp.Open "HOSGrp", medixmasterconn, adOpenKeyset, adLockOptimistic
                HosGrp.AddNew
            Else
                strsql = "Select * From HOSGrp Where HOSGrpCode = '" & Trim(TxtHosGrpCode) & "'"
                HosGrp.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With HosGrp
                    !HOSGRPCode = ChkStr(TxtHosGrpCode)
                    !HOSGRPName = ChkStr(TxtHosGrpName)
                    !HOSGrpDesc = ChkStr(TxtHosGrpDesc)
                    !HOSGrpLastUpdateUser = Logon
                    !HOSGrpLastUpdateDate = Date
                End With
                HosGrp.Update
                Call ClearForm(Me)
                Call HOSGrpListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                
            HosGrp.Close
            Set SaveHOSGrp = Nothing
        End If
        EditFlag = False
    End If
End If

'medixmasterconn.Close
'Set medixmasterconn = Nothing

End Sub

Private Sub DeleteHOSGrp()
Dim DeleteHOSGrp As New ADODB.Recordset
Dim rsCheckHOS As New ADODB.Recordset
Dim strsql, sqlHOS As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
EditFlag = False
If TxtHosGrpCode = "" Then
   MsgBox "Please Enter Hospital Group Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtHosGrpCode.SetFocus
ElseIf lstHosGrpList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlHOS = "Select HOSCode from HOS where HOSGrpCode = '" & RemStr(TxtHosGrpCode) & "'"
   
   rsCheckHOS.MaxRecords = 1
   rsCheckHOS.Open sqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   If rsCheckHOS.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Hospital Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckHOS.recordCount = 0 Then
           strsql = "Delete From HOSGrp Where HOSGrpCode = '" & ChkStr(lstHosGrpList.SelectedItem.Text) & "'"
           DeleteHOSGrp.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call HOSGrpListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           TxtHosGrpCode.SetFocus
         End If
     End If
End If
'medixmasterconn.Close
'Set medixmasterconn = Nothing

Exit Sub

End Sub

Private Sub HOSGrpListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim HOSGrpMaster As ListItem
Dim sql As String
Dim TotalHOSGRPCount As Integer
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
    sql = "Select HOSGrpCode, HOSGrpName, HOSGrpDesc " & _
          "from HOSGrp where 0=0 "
          
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If
              
    TotalHOSGRPCount = 0
    lstHosGrpList.ListItems.Clear
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set HOSGrpMaster = lstHosGrpList.ListItems.Add(, , rst2!HOSGRPCode)
                HOSGrpMaster.SubItems(1) = rst2!HOSGRPName
                HOSGrpMaster.SubItems(2) = rst2!HOSGrpDesc
                rst2.MoveNext
        TotalHOSGRPCount = TotalHOSGRPCount + 1
        txtTotalRecord = TotalHOSGRPCount
        Loop
    rst2.Close
    End If
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub lstHOSGrpList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstHosGrpList.ListItems.Count <> 0 Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from HOSGrp Where HOSGrpCode = '" & lstHosGrpList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        TxtHosGrpCode = rst3!HOSGRPCode
        TxtHosGrpName = rst3!HOSGRPName
        If Len(rst3!HOSGrpDesc) Then
            TxtHosGrpDesc = rst3!HOSGrpDesc
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
        'Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub SearchHOSGrp()
    Dim StrAppend As String
   
    Screen.MousePointer = vbHourglass
    StrAppend = ""
    
    If Len(Trim(TxtHosGrpCode)) Then
        StrAppend = StrAppend + " And HOSGrpCode like '%" & ChkStr(TxtHosGrpCode) & "%'"
    End If
    If Len(Trim(TxtHosGrpName)) Then
        StrAppend = StrAppend + " And HOSGrpName like '%" & ChkStr(TxtHosGrpName) & "%'"
    End If
    
    HosGrpCriteria = StrAppend
    Call HOSGrpListing(StrAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstHOSGrpList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    lstHosGrpList.SortKey = ColumnHeader.Index - 1
    lstHosGrpList.Sorted = True
        If lstHosGrpList.SortOrder = lvwAscending Then
            lstHosGrpList.SortOrder = lvwDescending
        Else
            lstHosGrpList.SortOrder = lvwAscending
        End If
End Sub

'Private Sub PrintHOSGRP()
    'RptHosGrpList.Show
'End Sub
'**********************************End Hospital Group**************************************

'********************Start Add, Delete, Edit, Save, Search, Cancel buttons******************

Private Sub CmdAdd_Click()
EditFlag = False
Call ClearForm(Me)
Call DisableCmdButton(Me)
Call EntriesEnable(True)
cmdSave.Enabled = True
cmdAdd.Enabled = True
    Select Case SSTab1.Tab
        Case 0
            txtHealthCode.SetFocus
        Case 1
            TxtHOSCode.SetFocus
        Case 2
            txtCLNCode.Enabled = False
            txtCLNName.SetFocus
            Call CLNComboListing
        Case 3
            cboPAYHLTCode.SetFocus
        Case 4
            CboPAYCode.SetFocus
        Case 5
            txtSRVcode.SetFocus
        Case 6
            TxtHosGrpCode.SetFocus
        Case 7
            txtGrpCoCode.Enabled = False
            txtCompanyName.SetFocus
        Case 8
            TxtPayGrpCode.SetFocus
        Case 9
            txtAGTCode.SetFocus
        Case 10
            txtSpecApp.SetFocus
        Case 11
            TxtRecModeCode.SetFocus
    End Select

End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass
    If SSTab1.Tab = 0 Then
        Call DeleteProviderCharge
    ElseIf SSTab1.Tab = 1 Then
        Call DeleteHospital
    ElseIf SSTab1.Tab = 2 Then
        Call DeleteClinic
    ElseIf SSTab1.Tab = 3 Then
        Call DeletePayor
    ElseIf SSTab1.Tab = 4 Then
        Call DeleteBranch
    ElseIf SSTab1.Tab = 5 Then
        Call DeleteServiceProvider
    ElseIf SSTab1.Tab = 6 Then
        Call DeleteHOSGrp
    ElseIf SSTab1.Tab = 7 Then
        Call DeleteCompanyGrp
    ElseIf SSTab1.Tab = 8 Then
        Call DeletePAYGRP
    ElseIf SSTab1.Tab = 9 Then
        Call DeleteAgent
    ElseIf SSTab1.Tab = 10 Then
        Call DeleteSpec
    ElseIf SSTab1.Tab = 11 Then
        Call DeleteRecMode
    End If
cmdAdd.Enabled = True
Screen.MousePointer = Default
End Sub

Private Sub CmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    'txtMCOAmountAdult.Enabled = False
    'txtMCOAmountChild.Enabled = False
    
    txtPAYname.SetFocus
    CboBRN.SetFocus
    TxtHOSName.SetFocus
    txtCLNName.SetFocus
    txtSRVname.SetFocus
    txtAGTDesc.SetFocus
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    'txtEffdate.Enabled = False
    
    If cboPAYHLTCode <> "" Then
       cboPAYHLTCode.Enabled = False
       txtPAYcode.Enabled = False
       CboPAYGRP.Enabled = False
       'Frame1.Enabled = True
       txtPAYname.SetFocus
    End If
    
    If CboPAYCode <> "" Then
       CboPAYCode.Enabled = False
       CboBRN.Enabled = False
       TxtBRNAdd1.SetFocus
    End If
    
    If TxtHOSCode <> "" Then
       TxtHOSCode.Enabled = False
       TxtHOSName.SetFocus
    End If
    
    If txtCLNCode <> "" Then
       txtCLNCode.Enabled = False
       txtMIXClinic.Enabled = False
       txtCLNName.SetFocus
    End If
    
    If txtGrpCoCode <> "" Then
       txtGrpCoCode.Enabled = False
       txtCompanyName.Enabled = False
       txtCoAdd1.SetFocus
    End If
    
    'If txtINScode <> "" Or txtHLTcode <> "" Then
       'txtINScode.Enabled = False
       'txtHLTcode.Enabled = False
       'txtMCOdescription.SetFocus
    'End If
    
    If txtSRVcode <> "" Then
       txtSRVcode.Enabled = False
       txtSRVname.SetFocus
    End If
    
    If txtAGTCode <> "" Then
       txtAGTCode.Enabled = False
       txtAGTDesc.SetFocus
    End If
    
    If txtPRVCode <> "" Or txtHealthCode <> "" Or txtINSCode1 <> "" Then
       txtPRVCode.Enabled = False
       txtHealthCode.Enabled = False
       txtINSCode1.Enabled = False
       txtAccessDate.Enabled = False
       txtPRVCaseCharges.SetFocus
    End If
    
    If TxtRecModeCode <> "" Then
       TxtRecModeCode.Enabled = False
       TxtRecModeDesc.SetFocus
    End If
    
    If txtSpecApp <> "" Then
       txtSpecApp.Enabled = False
       txtSpecApp.SetFocus
    End If
    
    If TxtPayGrpCode <> "" Then
       TxtPayGrpCode.Enabled = False
       TxtPayGrpName.SetFocus
    End If
    
    If txtGrpCoCode <> "" Then
       txtGrpCoCode.Enabled = False
       txtCompanyName.Enabled = False
       txtCoAdd1.SetFocus
    End If
    
    If TxtHosGrpCode <> "" Then
       TxtHosGrpCode.Enabled = False
       TxtHosGrpName.SetFocus
    End If
    
End Sub

Private Sub cmdSave_Click()
   Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SaveProviderCharge
            'Call SavePayor
        Case 1
            Call saveHospital
            'Call SaveBranch
        Case 2
            Call SaveClinic
        Case 3
            Call SavePayor
            'Call SaveClinic
        Case 4
            Call SaveBranch
            'Call SaveMCOFees
        Case 5
            Call SaveServiceProvider
        Case 6
            Call SaveHOSGrp
            'Call SaveProviderCharge
        Case 7
            Call SaveCompanyGrp
        Case 8
            Call SavePAYGRP
            'Call SaveGroup
        Case 9
            Call SaveAgent
        Case 10
            Call SaveSPEC
        Case 11
            Call SaveRecMode
    End Select
    cmdAdd.Enabled = True
    cmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

Private Sub CmdSearch_Click()
    
   cmdSearch.Enabled = False
 ''  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
   
   Screen.MousePointer = vbHourglass
   Select Case SSTab1.Tab
        Case 0
            Call SearchPRVCharges
            txtPRVCode.Enabled = True
            txtHealthCode.Enabled = True
            txtINSCode1.Enabled = True
            txtPRVCaseCharges.Enabled = False
            txtPRVProviderCharges.Enabled = False
            txtPRVCode.SetFocus

            'Call SearchPAY
            'txtPAYcode.Enabled = True
            'txtPAYcode.SetFocus
            'CboPAYGRP.Enabled = True
            'txtPAYname.Enabled = True
                                   
        Case 1
            Call SearchHOS
            TxtHOSCode.Enabled = True
            TxtHOSName.Enabled = True
            TxtHOSCode.SetFocus

            'Call SearchBranch
            'txtBRNcode.Enabled = True
            'txtBRNname.Enabled = True
            'txtBRNcode.SetFocus
            
        Case 2
            Call SearchCLN
            Call EntriesEnable(False)
                        
        Case 3
            Call SearchPAY
            cboPAYHLTCode.Enabled = True
            cboPAYHLTCode.SetFocus
            txtPAYcode.Enabled = True
            CboPAYGRP.Enabled = True
            txtPAYname.Enabled = True

            'Call SearchCLN
            'txtCLNCode.Enabled = True
            'txtCLNName.Enabled = True
            'CboClinicGrp.Enabled = True
            'txtCLNTelNo.Enabled = False
            'txtCLNFaxNo.Enabled = False
            'TxtCLNAdd1.Enabled = False
            'TxtCLNAdd2.Enabled = False
            'TxtCLNAdd3.Enabled = False
            'txtCLNEmail.Enabled = False
            'txtCLNHomePage.Enabled = False
            'TxtCLNDocName.Enabled = False
            'TxtCLNContact.Enabled = False
            'txtCLNCode.SetFocus
        Case 4
            Call SearchBranch
            CboBRN.Enabled = True
            CboPAYCode.Enabled = True
            CboPAYCode.SetFocus

            'Call SearchMCO
            'txtINScode.Enabled = True
            'txtHLTcode.Enabled = True
            'txtMCOdescription.Enabled = False
            'txtMCOAmountAdult.Enabled = False
            'txtMCOAmountChild.Enabled = False
            'txtEffDate.Enabled = False
            'txtExpDate.Enabled = False
            'txtINScode.SetFocus
        Case 5
            Call SearchServiceProvider
            txtSRVcode.Enabled = True
            txtSRVname.Enabled = True
            TxtSRVAdd1.Enabled = False
            txtSRVcode.SetFocus
        Case 6
            Call SearchHOSGrp
            TxtHosGrpCode.Enabled = True
            TxtHosGrpName.Enabled = True
            TxtHosGrpDesc.Enabled = False
            TxtHosGrpCode.SetFocus

            'Call SearchPRVCharges
            'txtPRVCode.Enabled = True
            'txtHealthCode.Enabled = True
            'txtINSCode1.Enabled = True
            'txtPRVCaseCharges.Enabled = False
            'txtPRVProviderCharges.Enabled = False
            'txtPRVCode.SetFocus
        Case 7
            Call SearchCompanyGrp
            Call ClearForm(Me)
            txtGrpCoCode.Enabled = True
            txtCompanyName.Enabled = True
            txtGrpCoCode.SetFocus
        Case 8
            Call SearchPAYGRP
            TxtPayGrpCode.Enabled = True
            TxtPayGrpName.Enabled = True
            TxtPayGrpDesc.Enabled = False
            TxtPayGrpCode.SetFocus
        
        Case 9
            Call SearchAgent
            
            'Call SearchGroup
            txtAGTCode.Enabled = True
            txtAGTDesc.Enabled = False
            'TxtGDescription.Enabled = True
            txtAGTCode.SetFocus
         Case 10
            Call SearchSpec
            txtSpecApp.Enabled = True
            txtSpecDesc.Enabled = False
            txtSpecApp.SetFocus
         Case 11
            Call SearchRecMode
            TxtRecModeCode.Enabled = True
            TxtRecModeDesc.Enabled = False
            TxtRecModeCode.SetFocus
    End Select
    Screen.MousePointer = Default
    cmdAdd.Enabled = True
    
 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    cmdSearch.Enabled = True

End Sub

Private Sub CmdRefresh_Click()

Call ClearForm(Me)
Call EntriesEnable(True)

Select Case SSTab1.Tab

    Case 0
        Call PRVChargesComboListing
        
    Case 1
        Call HOSGRPComboListing
        
    Case 2
        Call CLNComboListing
      
    Case 3
        Call PAYInfoComboListing
      
    Case 4
        Call BRNComboListing
     
    'Case 5
        'Call ServiceProviderListing
    'Case 6
       'Call PRVChargesComboListing
       'Call PRVListing
    Case 7
        Call CompanyGRPComboListing
    'Case 8
     '   Call ICDListing
    
End Select
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lstPRVList.ListItems.Clear
    lstHospitalList.ListItems.Clear
    lstClinicList.ListItems.Clear
    lstPayorList.ListItems.Clear
    lstBranchList.ListItems.Clear
    lstServiceProviderList.ListItems.Clear
    lstHosGrpList.ListItems.Clear
    lstCompanyGrpList.ListItems.Clear
    lstPAYGRPList.ListItems.Clear
    lstAgentList.ListItems.Clear
    lstSpecialList.ListItems.Clear
    lstRecModeList.ListItems.Clear
    txtTotalRecord = 0
    cmdEdit.Enabled = False
End Sub

Private Sub cmdCancel_Click()
'    Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub
'********************End Add, Delete, Edit, Save, Search, Cancel buttons******************


Private Sub lstServiceProviderList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstServiceProviderList.ListItems.Count <> 0 Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from SRV Where SRVCode = '" & lstServiceProviderList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtSRVcode = rst3!SRVCode
        txtSRVname = rst3!SRVname
        TxtSRVAdd1 = rst3!SRVaddress1
        If Len(rst3!SRVaddress2) Then
           TxtSRVAdd2 = rst3!SRVaddress2
        End If
        If Len(rst3!SRVaddress3) Then
           TxtSRVAdd3 = rst3!SRVaddress3
        End If
        If Len(rst3!ContactPerson) Then
           TxtSRVContact = rst3!ContactPerson
        End If
        If Len(rst3!SRVTelNo1) Then
           txtTelNo1 = rst3!SRVTelNo1
        End If
        If Len(rst3!SRVTelNo2) Then
           txtTelNo2 = rst3!SRVTelNo2
        End If
        If Len(rst3!SRVFaxNo1) Then
           txtFaxNo1 = rst3!SRVFaxNo1
        End If
        If Len(rst3!SRVFaxNo2) Then
           txtFaxNo2 = rst3!SRVFaxNo2
        End If
        If Len(rst3!TollFree) Then
           txtTollFree = rst3!TollFree
        End If
        If Len(rst3!SRVRemark) Then
           txtRemark = rst3!SRVRemark
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
       '
      '  medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstServiceProviderList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstServiceProviderList.ListItems.Count <> 0 Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from SRV Where SRVCode = '" & lstServiceProviderList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtSRVcode = rst3!SRVCode
        txtSRVname = rst3!SRVname
        TxtSRVAdd1 = rst3!SRVaddress1
        If Len(rst3!SRVaddress2) Then
           TxtSRVAdd2 = rst3!SRVaddress2
        End If
        If Len(rst3!SRVaddress3) Then
           TxtSRVAdd3 = rst3!SRVaddress3
        End If
        If Len(rst3!ContactPerson) Then
           TxtSRVContact = rst3!ContactPerson
        End If
        If Len(rst3!SRVTelNo1) Then
           txtTelNo1 = rst3!SRVTelNo1
        End If
        If Len(rst3!SRVTelNo2) Then
           txtTelNo2 = rst3!SRVTelNo2
        End If
        If Len(rst3!SRVFaxNo1) Then
           txtFaxNo1 = rst3!SRVFaxNo1
        End If
        If Len(rst3!SRVFaxNo2) Then
           txtFaxNo2 = rst3!SRVFaxNo2
        End If
        If Len(rst3!TollFree) Then
           txtTollFree = rst3!TollFree
        End If
        If Len(rst3!SRVRemark) Then
           txtRemark = rst3!SRVRemark
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
     '   medixmasterconn.Close
      '  Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstSpecialList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstSpecialList.SortKey = ColumnHeader.Index - 1
lstSpecialList.Sorted = True

    If lstSpecialList.SortOrder = lvwAscending Then
        lstSpecialList.SortOrder = lvwDescending
    Else
        lstSpecialList.SortOrder = lvwAscending
    End If

End Sub

Private Sub lstSpecialList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
        
    If lstSpecialList.ListItems.Count <> 0 Then
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from SpecApprove Where SpecAppCode = '" & lstSpecialList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly ' , adOpenKeyset, adLockOptimistic
        
        txtSpecApp = rst3!SpecAppCode
        If Len(rst3!SpecAppDesc) Then
            txtSpecDesc = rst3!SpecAppDesc
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconnMaint.Close
        'Set medixmasterconnMaint = Nothing
        
    End If
End Sub

Private Sub lstSpecialList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String

    If lstSpecialList.ListItems.Count <> 0 Then
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
    
        strsql = "Select * from SpecApprove Where SpecAppCode = '" & lstSpecialList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        txtSpecApp = rst3!SpecAppCode
        If Len(rst3!SpecAppDesc) Then
            txtSpecDesc = rst3!SpecAppDesc
        End If
        Call EntriesEnable(False)
        cmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconnMaint.Close
     '   Set medixmasterconnMaint = Nothing
        
    End If
End Sub


'*****************************Start Tab Clicking**********************************
Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
    Case 0 'Access Fees
         'Call PayorListing
         CmdRefresh.Enabled = True
         lstPRVList.ListItems.Clear
         txtTotalRecord = 0
         Call PRVChargesComboListing
         Call ClearForm(Me)
    Case 1 'Hospital Info
         'Call BRNPAYComboListing
         Call HOSGRPComboListing
         CmdRefresh.Enabled = True
         lstHospitalList.ListItems.Clear
         txtTotalRecord = 0
         Call ClearForm(Me)
    Case 2 'Clinic Info
         CmdRefresh.Enabled = True
         lstClinicList.ListItems.Clear
         txtTotalRecord = 0
         Call ClearForm(Me)
         Call CLNComboListing
         txtCLNCode.Enabled = False
    Case 3 'Payor Info
         Call PAYInfoComboListing
         CmdRefresh.Enabled = True
         lstPayorList.ListItems.Clear
         txtTotalRecord = 0
         Call ClearForm(Me)
    Case 4 'Payor Branch
         CmdRefresh.Enabled = True
         Call BRNComboListing
         lstBranchList.ListItems.Clear
         txtTotalRecord = 0
         Call ClearForm(Me)
    Case 5 'Service Provider
         CmdRefresh.Enabled = False
         'Call ServiceProviderListing
         lstServiceProviderList.ListItems.Clear
         txtTotalRecord = 0
         Call ClearForm(Me)
    Case 6 'Hospital Group
         CmdRefresh.Enabled = False
         lstHosGrpList.ListItems.Clear
         txtTotalRecord = 0
         'txtPRVCode.Clear
         'txtINSCode1.Clear
         'txtHealthCode.Clear
         Call ClearForm(Me)
    Case 7 'Company Group
         lstCompanyGrpList.ListItems.Clear
         txtTotalRecord = 0
         CmdRefresh.Enabled = True
         Call ClearForm(Me)
         Call CompanyGRPComboListing
    Case 8 'Payor Group
         lstPAYGRPList.ListItems.Clear
         txtTotalRecord = 0
         CmdRefresh.Enabled = False
         'txtPRVCode.Clear
         'txtINSCode1.Clear
         'txtHealthCode.Clear
         Call ClearForm(Me)
         'End If
    Case 9 'Agent Maintenance
         lstAgentList.ListItems.Clear
         txtTotalRecord = 0
         CmdRefresh.Enabled = False
         Call ClearForm(Me)
    Case 10 'Special Approval
         lstSpecialList.ListItems.Clear
         txtTotalRecord = 0
         CmdRefresh.Enabled = False
         Call ClearForm(Me)
    Case 11 'Special Approval
         lstRecModeList.ListItems.Clear
         txtTotalRecord = 0
         CmdRefresh.Enabled = False
         Call ClearForm(Me)
    End Select
    Screen.MousePointer = Default
End Sub
'*****************************End Tab Clicking**********************************

Private Sub txtCLNnodoc_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtCLNpostcode_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtCLNReceivedDate_LostFocus()
    If IsDate(txtCLNReceivedDate) Then
    Else
        If txtCLNReceivedDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtCLNReceivedDate.SetFocus
        End If
    End If
End Sub


Private Sub txtCoEffDate_LostFocus()
    If txtCoEffDate <> "__/__/____" Then
        If IsDate(txtCoEffDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtCoEffDate.SetFocus
        ElseIf Format(txtCoEffDate, "yyyymmdd") > Format(Date, "yyyymmdd") Then
            MsgBox "Date cannot greater than system date ! ", vbCritical
            txtCoEffDate.SetFocus
        End If
    End If
End Sub

Private Sub txtCoFax_KeyPress(KeyAscii As Integer)
Clipboard.Clear
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtCoPhone_KeyPress(KeyAscii As Integer)
Clipboard.Clear
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtFaxNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtFaxNo1_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtFaxNo2_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

'**********************Start Controlling of ComboList**************************

Private Sub txtHealthCode_Change()
    If InStr(txtHealthCode.Text, "null") Then
       txtHealthCode.Enabled = False
    End If
End Sub

Private Sub txtHealthCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtHealthCode.Text, txtHealthCode.SelStart) + Chr(KeyAscii) + Mid(txtHealthCode.Text, txtHealthCode.SelStart + txtHealthCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To txtHealthCode.ListCount - 1
 
        If NewText = LCase(Left(txtHealthCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = txtHealthCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                txtHealthCode.Text = ValidValue
                txtHealthCode.SelStart = 0
                txtHealthCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub txtINSCode1_Change()
    If InStr(txtINSCode1.Text, "null") Then
       txtINSCode1.Enabled = False
    End If
End Sub

Private Sub txtINSCode1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtINSCode1.Text, txtINSCode1.SelStart) + Chr(KeyAscii) + Mid(txtINSCode1.Text, txtINSCode1.SelStart + txtINSCode1.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To txtINSCode1.ListCount - 1
 
        If NewText = LCase(Left(txtINSCode1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = txtINSCode1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               txtINSCode1.Text = ValidValue
               txtINSCode1.SelStart = 0
               txtINSCode1.SelLength = Len(ValidValue)
             End If
        End If
End Sub
Private Sub txtCLNFaxNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtCLNTelNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPRVProviderCharges_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPAYfax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPAYphone_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPRVCaseCharges_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPRVCode_Change()
    If InStr(txtHealthCode.Text, "null") Then
       txtHealthCode.Enabled = False
    End If
End Sub

Private Sub txtPRVCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtPRVCode.Text, txtPRVCode.SelStart) + Chr(KeyAscii) + Mid(txtPRVCode.Text, txtPRVCode.SelStart + txtPRVCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To txtPRVCode.ListCount - 1
 
        If NewText = LCase(Left(txtPRVCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = txtPRVCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               txtPRVCode.Text = ValidValue
               txtPRVCode.SelStart = 0
               txtPRVCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboPAYGRP_Change()
    If InStr(CboPAYGRP.Text, "null") Then
       CboPAYGRP.Enabled = False
    End If
End Sub

Private Sub CboPAYGRP_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPAYGRP.Text, CboPAYGRP.SelStart) + Chr(KeyAscii) + Mid(CboPAYGRP.Text, CboPAYGRP.SelStart + CboPAYGRP.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPAYGRP.ListCount - 1
 
        If NewText = LCase(Left(CboPAYGRP.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPAYGRP.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPAYGRP.Text = ValidValue
               CboPAYGRP.SelStart = 0
               CboPAYGRP.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboCLNType_Change()
    If InStr(cboCLNType.Text, "null") Then
       cboCLNType.Enabled = False
    End If
End Sub

Private Sub cboCLNType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNType.Text, cboCLNType.SelStart) + Chr(KeyAscii) + Mid(cboCLNType.Text, cboCLNType.SelStart + cboCLNType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCLNType.ListCount - 1
 
        If NewText = LCase(Left(cboCLNType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCLNType.Text = ValidValue
               cboCLNType.SelStart = 0
               cboCLNType.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboCLNstate_Change()
    If InStr(cboCLNstate.Text, "null") Then
       cboCLNstate.Enabled = False
    End If
End Sub

Private Sub cboCLNstate_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNstate.Text, cboCLNstate.SelStart) + Chr(KeyAscii) + Mid(cboCLNstate.Text, cboCLNstate.SelStart + cboCLNstate.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCLNstate.ListCount - 1
 
        If NewText = LCase(Left(cboCLNstate.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNstate.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCLNstate.Text = ValidValue
                cboCLNstate.SelStart = 0
                cboCLNstate.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCLNcity_Change()
    If InStr(cboCLNcity.Text, "null") Then
       cboCLNcity.Enabled = False
    End If
End Sub

Private Sub cboCLNcity_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNcity.Text, cboCLNcity.SelStart) + Chr(KeyAscii) + Mid(cboCLNcity.Text, cboCLNcity.SelStart + cboCLNcity.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCLNcity.ListCount - 1
 
        If NewText = LCase(Left(cboCLNcity.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNcity.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCLNcity.Text = ValidValue
               cboCLNcity.SelStart = 0
               cboCLNcity.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboCLNpanel_Change()
    If InStr(cboCLNpanel.Text, "null") Then
       cboCLNpanel.Enabled = False
    End If
End Sub

Private Sub cboCLNpanel_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNpanel.Text, cboCLNpanel.SelStart) + Chr(KeyAscii) + Mid(cboCLNpanel.Text, cboCLNpanel.SelStart + cboCLNpanel.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCLNpanel.ListCount - 1
 
        If NewText = LCase(Left(cboCLNpanel.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNpanel.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCLNpanel.Text = ValidValue
               cboCLNpanel.SelStart = 0
               cboCLNpanel.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboCLN24hrs_Change()
    If InStr(cboCLN24hrs.Text, "null") Then
       cboCLN24hrs.Enabled = False
    End If
End Sub

Private Sub cboCLN24hrs_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLN24hrs.Text, cboCLN24hrs.SelStart) + Chr(KeyAscii) + Mid(cboCLN24hrs.Text, cboCLN24hrs.SelStart + cboCLN24hrs.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCLN24hrs.ListCount - 1
 
        If NewText = LCase(Left(cboCLN24hrs.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLN24hrs.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCLN24hrs.Text = ValidValue
               cboCLN24hrs.SelStart = 0
               cboCLN24hrs.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboCLNBNFType_Change()
    If InStr(cboCLNBNFType.Text, "null") Then
       cboCLNBNFType.Enabled = False
    End If
End Sub

Private Sub cboCLNBNFType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNBNFType.Text, cboCLNBNFType.SelStart) + Chr(KeyAscii) + Mid(cboCLNBNFType.Text, cboCLNBNFType.SelStart + cboCLNBNFType.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCLNBNFType.ListCount - 1
 
        If NewText = LCase(Left(cboCLNBNFType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNBNFType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCLNBNFType.Text = ValidValue
               cboCLNBNFType.SelStart = 0
               cboCLNBNFType.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboCLNxray_Change()
    If InStr(cboCLNxray.Text, "null") Then
       cboCLNxray.Enabled = False
    End If
End Sub

Private Sub cboCLNxray_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCLNxray.Text, cboCLNxray.SelStart) + Chr(KeyAscii) + Mid(cboCLNxray.Text, cboCLNxray.SelStart + cboCLNxray.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCLNxray.ListCount - 1
 
        If NewText = LCase(Left(cboCLNxray.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCLNxray.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCLNxray.Text = ValidValue
               cboCLNxray.SelStart = 0
               cboCLNxray.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboConversion_Change()
    If InStr(CboConversion.Text, "null") Then
       CboConversion.Enabled = False
    End If
End Sub

Private Sub CboConversion_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboConversion.Text, CboConversion.SelStart) + Chr(KeyAscii) + Mid(CboConversion.Text, CboConversion.SelStart + CboConversion.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboConversion.ListCount - 1
 
        If NewText = LCase(Left(CboConversion.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboConversion.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboConversion.Text = ValidValue
               CboConversion.SelStart = 0
               CboConversion.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboClinicGrp_Change()
    If InStr(CboClinicGrp.Text, "null") Then
       CboClinicGrp.Enabled = False
    End If
End Sub

Private Sub CboClinicGrp_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClinicGrp.Text, CboClinicGrp.SelStart) + Chr(KeyAscii) + Mid(CboClinicGrp.Text, CboClinicGrp.SelStart + CboClinicGrp.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboClinicGrp.ListCount - 1
 
        If NewText = LCase(Left(CboClinicGrp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClinicGrp.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboClinicGrp.Text = ValidValue
               CboClinicGrp.SelStart = 0
               CboClinicGrp.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboCoGrpPayor_Change()
    If InStr(CboCoGrpPayor.Text, "null") Then
       CboCoGrpPayor.Enabled = False
    End If
End Sub

Private Sub CboCoGrpPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCoGrpPayor.Text, CboCoGrpPayor.SelStart) + Chr(KeyAscii) + Mid(CboCoGrpPayor.Text, CboCoGrpPayor.SelStart + CboCoGrpPayor.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboCoGrpPayor.ListCount - 1
 
        If NewText = LCase(Left(CboCoGrpPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCoGrpPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboCoGrpPayor.Text = ValidValue
               CboCoGrpPayor.SelStart = 0
               CboCoGrpPayor.SelLength = Len(ValidValue)
             End If
        End If
End Sub
Private Sub cboCoTax_Change()
    If InStr(cboCoTax.Text, "null") Then
       cboCoTax.Enabled = False
    End If
End Sub

Private Sub cboCoTax_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCoTax.Text, cboCoTax.SelStart) + Chr(KeyAscii) + Mid(cboCoTax.Text, cboCoTax.SelStart + cboCoTax.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCoTax.ListCount - 1
 
        If NewText = LCase(Left(cboCoTax.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCoTax.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCoTax.Text = ValidValue
               cboCoTax.SelStart = 0
               cboCoTax.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboGrpComStatus_Change()
    If InStr(CboGrpComStatus.Text, "null") Then
       CboGrpComStatus.Enabled = False
    End If
End Sub

Private Sub CboGrpComStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGrpComStatus.Text, CboGrpComStatus.SelStart) + Chr(KeyAscii) + Mid(CboGrpComStatus.Text, CboGrpComStatus.SelStart + CboGrpComStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboGrpComStatus.ListCount - 1
 
        If NewText = LCase(Left(CboGrpComStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGrpComStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboGrpComStatus.Text = ValidValue
               CboGrpComStatus.SelStart = 0
               CboGrpComStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboCoConversion_Change()
    If InStr(cboCoConversion.Text, "null") Then
       cboCoConversion.Enabled = False
    End If
End Sub

Private Sub cboCoConversion_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCoConversion.Text, cboCoConversion.SelStart) + Chr(KeyAscii) + Mid(cboCoConversion.Text, cboCoConversion.SelStart + cboCoConversion.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCoConversion.ListCount - 1
 
        If NewText = LCase(Left(cboCoConversion.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCoConversion.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCoConversion.Text = ValidValue
               cboCoConversion.SelStart = 0
               cboCoConversion.SelLength = Len(ValidValue)
             End If
        End If
End Sub
Private Sub cboCoPrePost_Change()
    If InStr(cboCoPrePost.Text, "null") Then
       cboCoPrePost.Enabled = False
    End If
End Sub

Private Sub cboCoPrePost_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCoPrePost.Text, cboCoPrePost.SelStart) + Chr(KeyAscii) + Mid(cboCoPrePost.Text, cboCoPrePost.SelStart + cboCoPrePost.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboCoPrePost.ListCount - 1
 
        If NewText = LCase(Left(cboCoPrePost.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCoPrePost.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboCoPrePost.Text = ValidValue
               cboCoPrePost.SelStart = 0
               cboCoPrePost.SelLength = Len(ValidValue)
             End If
        End If
End Sub
'**********************End Controlling of ComboList**************************

'************************Start EntriesEnable*********************************
Private Sub EntriesEnable(status As Boolean)
    cboPAYHLTCode.Enabled = status
    txtPAYcode.Enabled = status
    txtPAYname.Enabled = status
    CboPAYGRP.Enabled = status
    txtPAYphone.Enabled = status
    txtPAYfax.Enabled = status
    TxtPayAdd1.Enabled = status
    TxtPayAdd2.Enabled = status
    TxtPayAdd2.Enabled = status
    TxtPayAdd3.Enabled = status
    txtPAYemail.Enabled = status
    txtPAYhomepage.Enabled = status
    TxtContact.Enabled = status
    TxtAdmContact.Enabled = status
    TxtCEO.Enabled = status
    TxtPAYEffDate.Enabled = status
    CboConversion.Enabled = status
    CboRBTax.Enabled = status
    CboPrePost.Enabled = status
    CboCoPay.Enabled = status
    CboCancelStatus.Enabled = status
    txtHOSAccCode.Enabled = status
    txtAGTCode.Enabled = status
    txtAGTDesc.Enabled = status
    txtSRVcode.Enabled = status
    txtSRVname.Enabled = status
    TxtSRVAdd1.Enabled = status
    TxtSRVAdd2.Enabled = status
    TxtSRVAdd3.Enabled = status
    txtRemark.Enabled = status
    TxtSRVContact.Enabled = status
    txtTelNo1.Enabled = status
    txtTelNo2.Enabled = status
    txtFaxNo1.Enabled = status
    txtFaxNo2.Enabled = status
    txtTollFree.Enabled = status
    txtPRVCode.Enabled = status
    txtHealthCode.Enabled = status
    txtINSCode1.Enabled = status
    txtPRVCaseCharges.Enabled = status
    txtAccessDate.Enabled = status
    txtPRVProviderCharges.Enabled = status
    CboBRN.Enabled = status
    TxtBRNAdd1.Enabled = status
    TxtBRNAdd2.Enabled = status
    TxtBRNAdd3.Enabled = status
    txtBRNTelNo.Enabled = status
    TxtBRNFaxNo.Enabled = status
    TxtBRNContact.Enabled = status
    TxtBRNEmail.Enabled = status
    txtMNGname.Enabled = status
    CboPAYCode.Enabled = status
    TxtHOSCode.Enabled = status
    TxtHOSName.Enabled = status
    CboHOSGRP.Enabled = status
    TxtHOSAdd1.Enabled = status
    TxtHOSAdd2.Enabled = status
    TxtHOSAdd3.Enabled = status
    TxtTelNo.Enabled = status
    TxtFaxNo.Enabled = status
    TxtEmail.Enabled = status
    TxtHomepage.Enabled = status
    TxtHosContact.Enabled = status
    TxtHosDocName.Enabled = status
    TxtHosRemark.Enabled = status
    txtCLNCode.Enabled = status
    txtCLNName.Enabled = status
    txtMIXClinic.Enabled = status
    CboClinicGrp.Enabled = status
    txtCLNTelNo.Enabled = status
    txtCLNFaxNo.Enabled = status
    TxtCLNAdd1.Enabled = status
    TxtCLNAdd2.Enabled = status
    TxtCLNAdd3.Enabled = status
    txtCLNEmail.Enabled = status
    txtCLNHomePage.Enabled = status
    TxtCLNContact.Enabled = status
    TxtCLNDocName.Enabled = status
    txtCLNpostcode.Enabled = status
    cboCLNcity.Enabled = status
    cboCLNstate.Enabled = status
    cboCLNType.Enabled = status
    cboCLNBNFType.Enabled = status
    cboCLNpanel.Enabled = status
    cboCLNxray.Enabled = status
    txtCLNnodoc.Enabled = status
    txtCLNformema.Enabled = status
    txtCLNReceivedDate.Enabled = status
    txtCLNBankAcc.Enabled = status
    txtCLNbank.Enabled = status
    txtCLNbranch.Enabled = status
    txtCLNAccPerson.Enabled = status
    cboCLN24hrs.Enabled = status
    txtCLNremark.Enabled = status
    TxtPayGrpCode.Enabled = status
    TxtPayGrpName.Enabled = status
    TxtPayGrpDesc.Enabled = status
    txtGrpCoCode.Enabled = status
    txtCompanyName.Enabled = status
    CboCoGrpPayor.Enabled = status
    txtCoAdd1.Enabled = status
    txtCoAdd2.Enabled = status
    txtCoAdd3.Enabled = status
    txtCoContact.Enabled = status
    txtCoAdmContact.Enabled = status
    txtCoCEO.Enabled = status
    txtCoPhone.Enabled = status
    txtCoFax.Enabled = status
    txtCoEmail.Enabled = status
    txtCoHomepage.Enabled = status
    txtCoEffDate.Enabled = status
    cboCoTax.Enabled = status
    cboCoConversion.Enabled = status
    cboCoPrePost.Enabled = status
    CboGrpComStatus.Enabled = status
    TxtHosGrpCode.Enabled = status
    TxtHosGrpName.Enabled = status
    TxtHosGrpDesc.Enabled = status
    TxtRecModeCode.Enabled = status
    TxtRecModeDesc.Enabled = status
End Sub
'************************End EntriesEnable*********************************

'************************Start Form*********************************
'Private Sub Form_Activate()
    'SSTab1.Tab = 0
    'CmdRefresh.Enabled = True
    'Call PRVChargesComboListing
    'Call CompanyGRPComboListing
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub

Private Sub Form_Load()
    'cmdDelete.Enabled = True
    cmdSave.Enabled = True
    'Me.WindowState = 0
    PAYCriteria = ""
    BRNCriteria = ""
    HOSCriteria = ""
    CLNCriteria = ""
    PRVChargesCriteria = ""
    'MCOCriteria = ""
    SRVCriteria = ""
    'WorkCriteria = ""
    'GroupCriteria = ""
    PAYGRPCriteria = ""
    GRPCoCriteria = ""
    HosGrpCriteria = ""
    Call BindChannel
End Sub
'************************Start Form*********************************
Private Sub BindChannel()
Dim strsql As String
Dim CompanyGrp As New ADODB.Recordset
    cmbChannel.Clear
    strsql = "Select PreFix,Description from HISMaintenance..Tbl_BusinessChaneel "
    CompanyGrp.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until CompanyGrp.EOF
       cmbChannel.AddItem (CompanyGrp!prefix + " - " + CompanyGrp!Description)
    CompanyGrp.MoveNext
    Loop
    CompanyGrp.Close
    Set CompanyGrp = Nothing
End Sub

Private Sub txtAccessDate_LostFocus()
    If IsDate(txtAccessDate) Then
    Else
        If txtAccessDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtAccessDate.SetFocus
        End If
    End If
End Sub

Private Sub TxtPAYEffDate_LostFocus()
    If IsDate(TxtPAYEffDate) Then
    Else
        If TxtPAYEffDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPAYEffDate.SetFocus
        End If
    End If
End Sub

Private Sub txtTelNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtTelNo1_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtTelNo2_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtTollFree_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub
'*******************************Start AgentCode *******************************
Private Sub SaveAgent()
'On Error Resume Next
Dim RecordSaved
Dim AGT As New ADODB.Recordset
Dim SaveAGT As New ADODB.Recordset
Dim strsql As String
      
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If txtAGTCode = "" Then
        MsgBox "Please Enter Agent Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtAGTCode.SetFocus
    ElseIf txtAGTDesc = "" Then
        MsgBox "Please Enter Agent Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtAGTDesc.SetFocus
    
    Else
          strsql = "Select * From Agent Where AgentCode = '" & Trim(txtAGTCode) & "'"
          
          AGT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
          
            If AGT.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                AGT.Close
                txtAGTCode.SetFocus
                txtAGTCode.SelStart = 0
                txtAGTCode.SelLength = Len(txtAGTCode)
                Set AGT = Nothing
                'Exit Sub
            Else
            
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        
         If RecordSaved = vbYes Then
            If EditFlag = False Then
                SaveAGT.Open "Agent", medixmasterconn, adOpenKeyset, adLockOptimistic
                SaveAGT.AddNew
            Else
                strsql = "Select * From Agent Where AgentCode = '" & Trim(txtAGTCode) & "'"
                SaveAGT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With SaveAGT
                    !AgentCode = ChkStr(txtAGTCode)
                    !Remarks = ChkStr(txtAGTDesc)
                    !LastUpdateUser = Logon
                    !LastUpdateDate = Date
                End With
                SaveAGT.Update
                
                Call ClearForm(Me)
                Call AGTListing(AGTCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        SaveAGT.Close
        Set SaveAGT = Nothing
        End If
        EditFlag = False
    End If
End If

'medixmasterconn.Close
'Set medixmasterconn = Nothing

End Sub

Private Sub DeleteAgent()
Dim DeleteAGT As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
EditFlag = False
If txtAGTCode = "" Then
   MsgBox "Please Enter Agent Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtAGTCode.SetFocus
ElseIf lstAgentList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBMOthers where MBMAgentCode = '" & RemStr(txtAGTCode) & "'"
      
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From Agent Where AgentCode = '" & Trim(txtAGTCode) & "'" 'ChkStr(lstRaceList.SelectedItem.Text) & "'"
           DeleteAGT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call AGTListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtAGTCode.SetFocus
         End If
     End If
End If

'medixmasterconn.Close
'Set medixmasterconn = Nothing

End Sub

Private Sub SearchAgent()
Dim StrAppend As String

Screen.MousePointer = vbHourglass
StrAppend = ""

If Len(Trim(txtAGTCode)) Then
StrAppend = StrAppend + " And AgentCode = '" & Trim(ChkStr(txtAGTCode)) & "'"
End If

'If Len(Trim(txtAGTDesc)) Then
'strAppend = strAppend + " And Remarks like '%" & ChkStr(txtAGTDesc) & "%'"
'End If

AGTCriteria = StrAppend
Call AGTListing(StrAppend)
Screen.MousePointer = Default
End Sub

Private Sub AGTListing(Optional StrAppend As String)
Dim AGT3 As New ADODB.Recordset
Dim AGTMaster As ListItem
Dim sql As String
Dim TotalAGTCount As Integer


TotalAGTCount = 0
lstAgentList.ListItems.Clear

    sql = "SELECT AgentCode, " & _
    " Remarks" & _
    " from Agent where 0=0 "
    
    If Len(Trim(StrAppend)) Then
    sql = sql + StrAppend
    End If
    
    AGT3.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If AGT3.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until AGT3.EOF
        Set AGTMaster = lstAgentList.ListItems.Add(, , AGT3!AgentCode)
            AGTMaster.SubItems(1) = Trim(AGT3!Remarks)
            AGT3.MoveNext
        TotalAGTCount = TotalAGTCount + 1
        txtTotalRecord = TotalAGTCount
        Loop
    AGT3.Close
    Set AGT3 = Nothing
End If

End Sub

Private Sub lstAgentList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String

If lstAgentList.ListItems.Count Then
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    strsql = "SELECT AgentCode, Remarks " & _
    "from Agent Where AgentCode = '" & lstAgentList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    
    txtAGTCode = rst3!AgentCode
    txtAGTDesc = Trim(rst3!Remarks)
    Call EntriesEnable(False)
    cmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If
End Sub
'*******************************End AgentCode *******************************

'*******************************Start SpecialApproval *******************************
Private Sub SaveSPEC()
On Error Resume Next
Dim RecordSaved
Dim SPEC As New ADODB.Recordset
Dim SaveSPEC As New ADODB.Recordset
Dim strsql As String
      
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If txtSpecApp = "" Then
        MsgBox "Please Enter Special Approval Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtSpecApp.SetFocus
    ElseIf txtSpecDesc = "" Then
        MsgBox "Please Enter Special Approval Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtSpecDesc.SetFocus
    
    Else
          strsql = "Select * From SpecApprove Where SpecAppCode = '" & Trim(txtSpecApp) & "'"
          
          SPEC.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
          
            If SPEC.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                SPEC.Close
                txtSpecApp.SetFocus
                txtSpecApp.SelStart = 0
                txtSpecApp.SelLength = Len(txtSpecApp)
                Set SPEC = Nothing
                'Exit Sub
    Else
      SPEC.Close
      Set SPEC = Nothing
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
         If RecordSaved = vbYes Then
            If EditFlag = False Then
                SPEC.Open "SpecApprove", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                SPEC.AddNew
            Else
                strsql = "Select * From SpecApprove Where SpecAppCode = '" & Trim(txtSpecApp) & "'"
                SPEC.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            End If
            
                With SPEC
                    !SpecAppCode = ChkStr(txtSpecApp)
                    !SpecAppDesc = ChkStr(txtSpecDesc)
                    !SpecLastUpdateUser = Logon
                    !SpecLastUpdateDate = Date
                End With
                SPEC.Update
                Call ClearForm(Me)
                Call SpecListing(SPECCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        SPEC.Close
        Set SaveSPEC = Nothing
        End If
        EditFlag = False
    End If
End If

'medixmasterconnMaint.Close
'Set medixmasterconnMaint = Nothing

End Sub

Private Sub DeleteSpec()
Dim DeleteSpec As New ADODB.Recordset
Dim rsCheckCAS As New ADODB.Recordset
Dim strsql, sqlCAS As String
Dim DeleteRecord

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
EditFlag = False
If txtSpecApp = "" Then
   MsgBox "Please Enter Special Approval Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtAGTCode.SetFocus
ElseIf lstSpecialList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCAS = "Select CasNumber from CasOthers where CasInd = '" & RemStr(txtSpecApp) & "'"
      
   rsCheckCAS.MaxRecords = 1
   rsCheckCAS.Open sqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckCAS.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Case Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckCAS.recordCount = 0 Then
           strsql = "Delete From SpecApprove Where SpecAppCode = '" & Trim(txtSpecApp) & "'" 'ChkStr(lstRaceList.SelectedItem.Text) & "'"
           DeleteSpec.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call SpecListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtSpecApp.SetFocus
         End If
     End If
End If

'medixmasterconn.Close
'Set medixmasterconn = Nothing
'medixmasterconnMaint.Close
'Set medixmasterconnMaint = Nothing

End Sub

Private Sub SearchSpec()
Dim StrAppend As String

Screen.MousePointer = vbHourglass
StrAppend = ""

If Len(Trim(txtSpecApp)) Then
StrAppend = StrAppend + " And SpecAppCode = '" & Trim(ChkStr(txtSpecApp)) & "'"
End If

'If Len(Trim(txtAGTDesc)) Then
'strAppend = strAppend + " And Remarks like '%" & ChkStr(txtAGTDesc) & "%'"
'End If

SPECCriteria = StrAppend
Call SpecListing(StrAppend)
Screen.MousePointer = Default
End Sub

Private Sub SpecListing(Optional StrAppend As String)
Dim Spec3 As New ADODB.Recordset
Dim SpecMaster As ListItem
Dim sql As String
Dim TotalSpecCount As Integer

'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

TotalSpecCount = 0
lstSpecialList.ListItems.Clear


    sql = "SELECT SpecAppCode, " & _
    " SpecAppDesc" & _
    " from SpecApprove where 0=0 "
    
    If Len(Trim(StrAppend)) Then
    sql = sql + StrAppend
    End If
    
    Spec3.Open sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    
    If Spec3.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until Spec3.EOF
            Set SpecMaster = lstSpecialList.ListItems.Add(, , Spec3!SpecAppCode)
                SpecMaster.SubItems(1) = Trim(Spec3!SpecAppDesc)
                Spec3.MoveNext
            TotalSpecCount = TotalSpecCount + 1
            txtTotalRecord = TotalSpecCount
            Loop
        Spec3.Close
        Set Spec3 = Nothing
    End If

    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing

End Sub

Private Sub lstSpecialList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String

If lstSpecialList.ListItems.Count Then
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    strsql = "SELECT SpecAppCode, SpecAppDesc " & _
    "from SpecApprove Where SpecAppCode = '" & lstSpecialList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    
    txtSpecApp = rst3!SpecAppCode
    txtSpecDesc = Trim(rst3!SpecAppDesc)
    Call EntriesEnable(False)
    cmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End If
End Sub

'Private Sub PrintSpecialApproval()
    'RptSpecialApprovedList.Show
'End Sub

'*******************************End SpecialApproval *******************************

'Function GenerateClinicCode() As String
'    Dim CLNQ As New ADODB.Recordset
'    Dim strsql As String
'    Dim SeqNo As String
    
'    strsql = "Select * from CLNSeq
    
'End Function

Private Sub CboRBTax_Change()
    If InStr(CboRBTax.Text, "null") Then
       CboRBTax.Enabled = False
    End If
End Sub

Private Sub CboRBTax_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRBTax.Text, CboRBTax.SelStart) + Chr(KeyAscii) + Mid(CboRBTax.Text, CboRBTax.SelStart + CboRBTax.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboRBTax.ListCount - 1
 
        If NewText = LCase(Left(CboRBTax.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboRBTax.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboRBTax.Text = ValidValue
               CboRBTax.SelStart = 0
               CboRBTax.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboPrePost_Change()
    If InStr(CboPrePost.Text, "null") Then
       CboPrePost.Enabled = False
    End If
End Sub

Private Sub CboPrePost_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPrePost.Text, CboPrePost.SelStart) + Chr(KeyAscii) + Mid(CboPrePost.Text, CboPrePost.SelStart + CboPrePost.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPrePost.ListCount - 1
 
        If NewText = LCase(Left(CboPrePost.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPrePost.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboPrePost.Text = ValidValue
               CboPrePost.SelStart = 0
               CboPrePost.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboCancelStatus_Change()
    If InStr(CboCancelStatus.Text, "null") Then
       CboCancelStatus.Enabled = False
    End If
End Sub

Private Sub CboCancelStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCancelStatus.Text, CboCancelStatus.SelStart) + Chr(KeyAscii) + Mid(CboCancelStatus.Text, CboCancelStatus.SelStart + CboCancelStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboCancelStatus.ListCount - 1
 
        If NewText = LCase(Left(CboCancelStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCancelStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboCancelStatus.Text = ValidValue
               CboCancelStatus.SelStart = 0
               CboCancelStatus.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CboCoPay_Change()
    If InStr(CboCoPay.Text, "null") Then
       CboCoPay.Enabled = False
    End If
End Sub

Private Sub CboCoPay_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCoPay.Text, CboCoPay.SelStart) + Chr(KeyAscii) + Mid(CboCoPay.Text, CboCoPay.SelStart + CboCoPay.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboCoPay.ListCount - 1
 
        If NewText = LCase(Left(CboCoPay.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCoPay.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboCoPay.Text = ValidValue
               CboCoPay.SelStart = 0
               CboCoPay.SelLength = Len(ValidValue)
             End If
        End If
End Sub

'Private Sub PrintCompanyGrp()
    'RptGrpCompanyList.Show
'End Sub

Private Sub SaveRecMode()
'On Error Resume Next
Dim RecordSaved
Dim RECMode As New ADODB.Recordset
Dim strsql As String
      
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If TxtRecModeCode = "" Then
        MsgBox "Please Enter Rec. Mode Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtRecModeCode.SetFocus
    ElseIf TxtRecModeDesc = "" Then
        MsgBox "Please Enter Rec. Mode Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtRecModeDesc.SetFocus
    
    Else
          strsql = "Select * From InvRecMode Where RecModeCode = '" & Trim(TxtRecModeCode) & "'"
          
          RECMode.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
          
            If RECMode.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                RECMode.Close
                TxtRecModeCode.SetFocus
                TxtRecModeCode.SelStart = 0
                TxtRecModeCode.SelLength = Len(TxtRecModeCode)
                Set RECMode = Nothing
                'Exit Sub
    Else
      RECMode.Close
      Set RECMode = Nothing
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
         If RecordSaved = vbYes Then
            If EditFlag = False Then
                RECMode.Open "InvRecMode", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                RECMode.AddNew
            Else
                strsql = "Select * From InvRecMode Where RecModeCode = '" & Trim(TxtRecModeCode) & "'"
                RECMode.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            End If
            
                With RECMode
                    !RecModeCode = ChkStr(TxtRecModeCode)
                    !RecModeDesc = ChkStr(TxtRecModeDesc)
                    !RecModeLastUpdateUser = Logon
                    !RecModeLastUpdateDate = Date
                End With
                RECMode.Update
                Call ClearForm(Me)
                Call RecModeListing(RECModeCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        RECMode.Close
        Set RECMode = Nothing
        End If
        EditFlag = False
    End If
End If

'medixmasterconnMaint.Close
'Set medixmasterconnMaint = Nothing

End Sub

Private Sub DeleteRecMode()
Dim DeleteRecMode As New ADODB.Recordset
Dim rsCheckCAS As New ADODB.Recordset
Dim strsql, sqlCAS As String
Dim DeleteRecord

'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

EditFlag = False
If TxtRecModeCode = "" Then
   MsgBox "Please Enter Rec. Mode Code", vbOKOnly + vbCritical, DialogBoxTitle
   
ElseIf lstRecModeList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCAS = "Select CASNumber from ClaimInvoice where CLMInvRecMode = '" & RemStr(TxtRecModeCode) & "'"
      
   rsCheckCAS.MaxRecords = 1
   rsCheckCAS.Open sqlCAS, medixmasterconnWS, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckCAS.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Claim Invoice Table!", vbOKOnly + vbCritical, "Error Message"
    
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckCAS.recordCount = 0 Then
           strsql = "Delete From InvRecMode Where RecModeCode = '" & Trim(TxtRecModeCode) & "'"
           DeleteRecMode.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call RecModeListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           TxtRecModeCode.SetFocus
         End If
     End If
End If

'medixmasterconnWS.Close
'Set medixmasterconnWS = Nothing
'medixmasterconnMaint.Close
'Set medixmasterconnMaint = Nothing

End Sub

Private Sub SearchRecMode()
Dim StrAppend As String

Screen.MousePointer = vbHourglass
StrAppend = ""

If Len(Trim(TxtRecModeCode)) Then
    StrAppend = StrAppend + " And RecModeCode = '" & Trim(ChkStr(TxtRecModeCode)) & "'"
End If


RECModeCriteria = StrAppend
Call RecModeListing(StrAppend)
Screen.MousePointer = Default
End Sub

Private Sub RecModeListing(Optional StrAppend As String)
Dim RECMode As New ADODB.Recordset
Dim RecModeMaster As ListItem
Dim sql As String
Dim TotalRecModeCount As Integer

'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

TotalRecModeCount = 0
lstRecModeList.ListItems.Clear

    sql = "SELECT RecModeCode, RecModeDesc From InvRecMode Where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If
    
    RECMode.Open sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    
    If RECMode.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until RECMode.EOF
        Set RecModeMaster = lstRecModeList.ListItems.Add(, , RECMode!RecModeCode)
            If Len(RECMode!RecModeDesc) Then
                RecModeMaster.SubItems(1) = Trim(RECMode!RecModeDesc)
            End If
            RECMode.MoveNext
        TotalRecModeCount = TotalRecModeCount + 1
        txtTotalRecord = TotalRecModeCount
        Loop
    RECMode.Close
    Set RECMode = Nothing
End If

'medixmasterconnMaint.Close
'Set medixmasterconnMaint = Nothing

End Sub

Private Sub lstRecModeList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String

If lstRecModeList.ListItems.Count Then
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    strsql = "SELECT RecModeCode, RecModeDesc " & _
    "from InvRecMode Where RecModeCode = '" & lstRecModeList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    
    TxtRecModeCode = rst3!RecModeCode
    If Len(rst3!RecModeDesc) Then
        TxtRecModeDesc = Trim(rst3!RecModeDesc)
    End If
    Call EntriesEnable(False)
    cmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
'    medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End If
End Sub

Private Sub lstRecModeList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String

If lstRecModeList.ListItems.Count Then
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    strsql = "SELECT RecModeCode, RecModeDesc " & _
    "from InvRecMode Where RecModeCode = '" & lstRecModeList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    
    TxtRecModeCode = rst3!RecModeCode
    If Len(rst3!RecModeDesc) Then
        TxtRecModeDesc = Trim(rst3!RecModeDesc)
    End If
    Call EntriesEnable(False)
    cmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  'medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    
End If
End Sub

Private Sub lstRecModeList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String

If lstRecModeList.ListItems.Count Then
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    strsql = "SELECT RecModeCode, RecModeDesc " & _
    "from InvRecMode Where RecModeCode = '" & lstRecModeList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
    
    TxtRecModeCode = rst3!RecModeCode
    If Len(rst3!RecModeDesc) Then
        TxtRecModeDesc = Trim(rst3!RecModeDesc)
    End If
    Call EntriesEnable(False)
    cmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End If
End Sub

Private Sub lstRecModeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstRecModeList.SortKey = ColumnHeader.Index - 1
lstRecModeList.Sorted = True

    If lstRecModeList.SortOrder = lvwAscending Then
        lstRecModeList.SortOrder = lvwDescending
    Else
        lstRecModeList.SortOrder = lvwAscending
    End If

End Sub

'Private Sub PrintRecMode()
    'RptRecModeList.Show
'End Sub
