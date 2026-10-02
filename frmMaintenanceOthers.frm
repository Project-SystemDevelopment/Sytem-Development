VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenanceOthers 
   Caption         =   "Member Maintenance"
   ClientHeight    =   8340
   ClientLeft      =   2220
   ClientTop       =   1875
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      Height          =   705
      Left            =   360
      TabIndex        =   51
      Top             =   6480
      Width           =   11055
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   4920
         TabIndex        =   36
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   4080
         TabIndex        =   35
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Default         =   -1  'True
         Height          =   315
         Left            =   7320
         TabIndex        =   37
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9000
         TabIndex        =   39
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   8160
         TabIndex        =   38
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   33
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   32
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   720
         TabIndex        =   31
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   315
         Left            =   3240
         TabIndex        =   34
         Top             =   285
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7125
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   12568
      _Version        =   393216
      Tabs            =   13
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Age Band"
      TabPicture(0)   =   "frmMaintenanceOthers.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstAgeBandList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Occupation"
      TabPicture(1)   =   "frmMaintenanceOthers.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label31"
      Tab(1).Control(1)=   "lstOccupationList"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Relationship"
      TabPicture(2)   =   "frmMaintenanceOthers.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label23"
      Tab(2).Control(1)=   "lstRelationshipList"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Nationality"
      TabPicture(3)   =   "frmMaintenanceOthers.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label24"
      Tab(3).Control(1)=   "lstNationalityList"
      Tab(3).Control(2)=   "Frame4"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Race"
      TabPicture(4)   =   "frmMaintenanceOthers.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label25"
      Tab(4).Control(1)=   "lstRaceList"
      Tab(4).Control(2)=   "Frame5"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Language"
      TabPicture(5)   =   "frmMaintenanceOthers.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label26"
      Tab(5).Control(1)=   "lstLanguageList"
      Tab(5).Control(2)=   "Frame6"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "City"
      TabPicture(6)   =   "frmMaintenanceOthers.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label27"
      Tab(6).Control(1)=   "lstCityList"
      Tab(6).Control(2)=   "Frame7"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "State"
      TabPicture(7)   =   "frmMaintenanceOthers.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label28"
      Tab(7).Control(1)=   "lstStateList"
      Tab(7).Control(2)=   "Frame8"
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "Work Nature"
      TabPicture(8)   =   "frmMaintenanceOthers.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label37"
      Tab(8).Control(1)=   "lstWorkList"
      Tab(8).Control(2)=   "Frame11"
      Tab(8).ControlCount=   3
      TabCaption(9)   =   "Endorsement  Status"
      TabPicture(9)   =   "frmMaintenanceOthers.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label39"
      Tab(9).Control(1)=   "lstStatusList"
      Tab(9).Control(2)=   "Frame12"
      Tab(9).ControlCount=   3
      TabCaption(10)  =   "Policy Status"
      TabPicture(10)  =   "frmMaintenanceOthers.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Label42"
      Tab(10).Control(1)=   "lstPolicyList"
      Tab(10).Control(2)=   "Frame13"
      Tab(10).ControlCount=   3
      TabCaption(11)  =   "CTP"
      TabPicture(11)  =   "frmMaintenanceOthers.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Label45"
      Tab(11).Control(1)=   "lstCPTList"
      Tab(11).Control(2)=   "Frame9"
      Tab(11).ControlCount=   3
      TabCaption(12)  =   "Receipt Type"
      TabPicture(12)  =   "frmMaintenanceOthers.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Label48"
      Tab(12).Control(1)=   "lstRecTypeList"
      Tab(12).Control(2)=   "Frame14"
      Tab(12).ControlCount=   3
      Begin VB.Frame Frame14 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   115
         Top             =   720
         Width           =   11175
         Begin VB.TextBox TxtRecTypeCode 
            Height          =   285
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   29
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox TxtRecTypeDesc 
            Height          =   960
            Left            =   1920
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   600
            Width           =   9000
         End
         Begin VB.Label Label47 
            Caption         =   "Rec. Type Code *"
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
            TabIndex        =   117
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label46 
            Caption         =   "Rec. Type Desc. *"
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
            TabIndex        =   116
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   110
         Top             =   720
         Width           =   11175
         Begin VB.TextBox TxtCPT 
            Height          =   285
            Left            =   2280
            MaxLength       =   5
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox TxtCPTDesc 
            Height          =   735
            Left            =   2280
            MaxLength       =   2500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   600
            Width           =   8580
         End
         Begin VB.Label Label44 
            Caption         =   "CPT Code            *"
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
            TabIndex        =   112
            Top             =   240
            Width           =   2100
         End
         Begin VB.Label Label43 
            Caption         =   "CPT Description         * "
            Height          =   255
            Left            =   360
            TabIndex        =   111
            Top             =   645
            Width           =   2280
         End
      End
      Begin VB.Frame Frame13 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   104
         Top             =   720
         Width           =   11175
         Begin VB.TextBox TxtPStatusDesc 
            Height          =   735
            Left            =   2280
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   600
            Width           =   8580
         End
         Begin VB.TextBox TxtPStatusCode 
            Height          =   285
            Left            =   2280
            MaxLength       =   1
            TabIndex        =   25
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label41 
            Caption         =   "Status Description         * "
            Height          =   255
            Left            =   360
            TabIndex        =   106
            Top             =   645
            Width           =   2280
         End
         Begin VB.Label Label40 
            Caption         =   "Status Code           *"
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
            TabIndex        =   105
            Top             =   240
            Width           =   2100
         End
      End
      Begin VB.Frame Frame12 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   99
         Top             =   720
         Width           =   11175
         Begin VB.TextBox TxtStatusCode 
            Height          =   285
            Left            =   2280
            MaxLength       =   1
            TabIndex        =   23
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox TxtStatusDesc 
            Height          =   735
            Left            =   2280
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   600
            Width           =   8580
         End
         Begin VB.Label Label38 
            Caption         =   "Status Code           *"
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
            TabIndex        =   101
            Top             =   240
            Width           =   2100
         End
         Begin VB.Label Label34 
            Caption         =   "Status Description         * "
            Height          =   255
            Left            =   360
            TabIndex        =   100
            Top             =   645
            Width           =   2280
         End
      End
      Begin VB.Frame Frame11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   -74760
         TabIndex        =   91
         Top             =   720
         Width           =   11085
         Begin VB.TextBox TxtWorkDescription 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   2520
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   720
            Width           =   8265
         End
         Begin VB.TextBox TxtWorkCode 
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
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   21
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label36 
            Caption         =   "WorkNature  Code             *"
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
            TabIndex        =   93
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label35 
            Caption         =   "WorkNature  Description *"
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
            TabIndex        =   92
            Top             =   720
            Width           =   2235
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   85
         Top             =   720
         Width           =   11175
         Begin VB.TextBox TxtOCPDesc 
            Height          =   1080
            Left            =   2640
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   555
            Width           =   8295
         End
         Begin VB.TextBox TxtOCPcode 
            Height          =   285
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   6
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label30 
            Caption         =   "Occupation Description *"
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
            TabIndex        =   87
            Top             =   555
            Width           =   2580
         End
         Begin VB.Label Label29 
            Caption         =   "Occupation Code          *"
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
            TabIndex        =   86
            Top             =   240
            Width           =   2220
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   74
         Top             =   720
         Width           =   11175
         Begin VB.ComboBox CboStateCode 
            Height          =   315
            Left            =   5160
            TabIndex        =   17
            Top             =   240
            Width           =   5655
         End
         Begin VB.TextBox txtCTYdesc 
            Height          =   735
            Left            =   2040
            MaxLength       =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   600
            Width           =   8820
         End
         Begin VB.TextBox txtCTYcode 
            Height          =   285
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   16
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label33 
            Caption         =   "State Description *"
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
            TabIndex        =   98
            Top             =   240
            Width           =   1980
         End
         Begin VB.Label Label13 
            Caption         =   "City Description  * "
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
            TabIndex        =   76
            Top             =   645
            Width           =   1620
         End
         Begin VB.Label Label12 
            Caption         =   "City Code           *"
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
            TabIndex        =   75
            Top             =   240
            Width           =   1740
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   11175
         Begin VB.ComboBox CboLifeTimePln 
            Height          =   315
            Left            =   5160
            TabIndex        =   2
            Top             =   240
            Width           =   5655
         End
         Begin VB.TextBox TxtAgeTo 
            Height          =   285
            Left            =   6240
            MaxLength       =   5
            TabIndex        =   5
            Top             =   1335
            Width           =   540
         End
         Begin VB.TextBox TxtAgeFr 
            Height          =   285
            Left            =   2160
            MaxLength       =   5
            TabIndex        =   4
            Top             =   1335
            Width           =   540
         End
         Begin VB.TextBox TxtAgeCode 
            Height          =   285
            Left            =   2160
            MaxLength       =   5
            TabIndex        =   1
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox TxtDesc 
            Height          =   600
            Left            =   2160
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   645
            Width           =   8685
         End
         Begin VB.Label Label49 
            Caption         =   "Life Time Plan Code"
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
            Left            =   3360
            TabIndex        =   120
            Top             =   240
            Width           =   1860
         End
         Begin VB.Label Label17 
            Caption         =   "Age To    *"
            Height          =   255
            Left            =   5040
            TabIndex        =   73
            Top             =   1335
            Width           =   1050
         End
         Begin VB.Label Label16 
            Caption         =   "Age From              *"
            Height          =   255
            Left            =   525
            TabIndex        =   72
            Top             =   1335
            Width           =   1530
         End
         Begin VB.Label Label15 
            Caption         =   "Age Code        *"
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
            TabIndex        =   71
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label Label14 
            Caption         =   "Age Description     *    "
            Height          =   255
            Left            =   480
            TabIndex        =   70
            Top             =   645
            Width           =   1455
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   67
         Top             =   720
         Width           =   11175
         Begin VB.TextBox txtSTAcode 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   19
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtSTAdesc 
            Height          =   735
            Left            =   2160
            MaxLength       =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   555
            Width           =   8835
         End
         Begin VB.Label Label20 
            Caption         =   "State Code          *"
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
            TabIndex        =   69
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label Label19 
            Caption         =   "State Description  *"
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
            TabIndex        =   68
            Top             =   645
            Width           =   1740
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   64
         Top             =   720
         Width           =   11175
         Begin VB.TextBox txtLGNdesc 
            Height          =   735
            Left            =   2520
            MaxLength       =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   600
            Width           =   8340
         End
         Begin VB.TextBox txtLGNcode 
            Height          =   285
            Left            =   2520
            MaxLength       =   2
            TabIndex        =   14
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   "Language Description * "
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
            TabIndex        =   66
            Top             =   645
            Width           =   2280
         End
         Begin VB.Label Label11 
            Caption         =   "Language Code         *"
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
            TabIndex        =   65
            Top             =   240
            Width           =   2100
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   61
         Top             =   720
         Width           =   11175
         Begin VB.TextBox txtRACDesc 
            Height          =   960
            Left            =   1920
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   600
            Width           =   9000
         End
         Begin VB.TextBox txtRACcode 
            Height          =   285
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   12
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "Race Description *"
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
            TabIndex        =   63
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Race Code         *"
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
            TabIndex        =   62
            Top             =   240
            Width           =   1800
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   58
         Top             =   720
         Width           =   11175
         Begin VB.TextBox txtNATcode 
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
            TabIndex        =   10
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtNATname 
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
            MaxLength       =   50
            TabIndex        =   11
            Top             =   720
            Width           =   8700
         End
         Begin VB.Label Label6 
            Caption         =   "Nationality Code    *"
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
            Left            =   495
            TabIndex        =   60
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label Label7 
            Caption         =   "Nationality Name   *"
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
            Left            =   480
            TabIndex        =   59
            Top             =   720
            Width           =   1755
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   55
         Top             =   720
         Width           =   11175
         Begin VB.TextBox txtRELcode 
            Height          =   285
            Left            =   2640
            MaxLength       =   1
            TabIndex        =   8
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtRELdesc 
            Height          =   960
            Left            =   2640
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   555
            Width           =   8295
         End
         Begin VB.Label Label2 
            Caption         =   "Relationship Code         *"
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
            TabIndex        =   57
            Top             =   240
            Width           =   2220
         End
         Begin VB.Label Label5 
            Caption         =   "Relationship Description *"
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
            TabIndex        =   56
            Top             =   555
            Width           =   2340
         End
      End
      Begin MSComctlLib.ListView lstAgeBandList 
         Height          =   3360
         Left            =   240
         TabIndex        =   43
         Top             =   2880
         Width           =   11115
         _ExtentX        =   19606
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Age Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Age From"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Age To"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descriptions"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Plan Code"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstOccupationList1 
         Height          =   3120
         Left            =   -74640
         TabIndex        =   44
         Top             =   2640
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   5503
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
            Text            =   "Occupation Code"
            Object.Width           =   2611
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstRelationshipList 
         Height          =   3480
         Left            =   -74760
         TabIndex        =   45
         Top             =   2760
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6138
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
            Text            =   "Relationship Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Relationship Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstNationalityList 
         Height          =   3960
         Left            =   -74760
         TabIndex        =   46
         Top             =   2280
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6985
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
            Text            =   "Nationality Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nationality Name"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstRaceList 
         Height          =   3480
         Left            =   -74760
         TabIndex        =   47
         Top             =   2760
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6138
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
            Text            =   "Race Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstLanguageList 
         Height          =   3675
         Left            =   -74760
         TabIndex        =   48
         Top             =   2520
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6482
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
            Text            =   "Language Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstCityList 
         Height          =   3720
         Left            =   -74760
         TabIndex        =   49
         Top             =   2520
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6562
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
            Text            =   "City Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "State Description"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstStateList 
         Height          =   3720
         Left            =   -74760
         TabIndex        =   50
         Top             =   2520
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6562
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
            Text            =   "State Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin VB.Frame Frame 
         Height          =   1215
         Left            =   -74640
         TabIndex        =   52
         Top             =   960
         Width           =   6375
         Begin VB.TextBox txtOCPDesc1 
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
            Left            =   1995
            MaxLength       =   50
            TabIndex        =   41
            Top             =   675
            Width           =   3300
         End
         Begin VB.TextBox txtOCPcode1 
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
            Left            =   1995
            MaxLength       =   2
            TabIndex        =   40
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Occupation Description"
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
            TabIndex        =   54
            Top             =   675
            Width           =   1770
         End
         Begin VB.Label Label1 
            Caption         =   "Occupation Code"
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
            Left            =   255
            TabIndex        =   53
            Top             =   360
            Width           =   1320
         End
      End
      Begin MSComctlLib.ListView lstOccupationList 
         Height          =   3360
         Left            =   -74760
         TabIndex        =   88
         Top             =   2880
         Width           =   11160
         _ExtentX        =   19685
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
            Text            =   "Occupation Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Occupation Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorkList 
         Height          =   3195
         Left            =   -74760
         TabIndex        =   94
         Top             =   3000
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   5636
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
            Text            =   "WorkNature Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "WorkNature Description"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstStatusList 
         Height          =   3675
         Left            =   -74760
         TabIndex        =   102
         Top             =   2520
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6482
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
            Text            =   "Endorsement Status Code"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Endorsement Status Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstPolicyList 
         Height          =   3675
         Left            =   -74760
         TabIndex        =   107
         Top             =   2520
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6482
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
            Text            =   "Policy Status Code"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Policy Status Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstCPTList 
         Height          =   3675
         Left            =   -74760
         TabIndex        =   113
         Top             =   2520
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6482
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
            Text            =   "CPT Code"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CPT Descriptions"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstRecTypeList 
         Height          =   3480
         Left            =   -74760
         TabIndex        =   118
         Top             =   2760
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6138
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
            Text            =   "Rec. Type Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   14111
         EndProperty
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         Caption         =   "REC. TYPE LIST"
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
         TabIndex        =   119
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Caption         =   "CPT LIST"
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
         TabIndex        =   114
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         Caption         =   "POLICY STATUS LIST"
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
         TabIndex        =   108
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Caption         =   "ENDORSEMENT STATUS LIST"
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
         TabIndex        =   103
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Caption         =   "WORKNATURE LIST"
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
         TabIndex        =   95
         Top             =   2760
         Width           =   1635
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "OCCUPATION LIST"
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
         TabIndex        =   89
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "STATE LIST"
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
         TabIndex        =   84
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "CITY LIST"
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
         TabIndex        =   83
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "LANGUAGE LIST"
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
         TabIndex        =   82
         Top             =   2230
         Width           =   1455
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "RACE LIST"
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
         TabIndex        =   81
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "NATIONALITY LIST"
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
         TabIndex        =   80
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "RELATIONSHIP LIST"
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
         TabIndex        =   79
         Top             =   2520
         Width           =   1755
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "OCCUPATIOP LIST"
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
         TabIndex        =   78
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "AGE BAND LIST"
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
         TabIndex        =   77
         Top             =   2640
         Width           =   1395
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Member Maintenance"
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
      TabIndex        =   109
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   97
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label32 
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
      Height          =   255
      Left            =   5160
      TabIndex        =   96
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label18 
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
      Left            =   120
      TabIndex        =   90
      Top             =   7440
      Width           =   4215
   End
End
Attribute VB_Name = "frmMaintenanceOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim bExit As Boolean
Public AGECriteria As String
Public OCPCriteria As String
Public RELCriteria As String
Public NATCriteria As String
Public RACCriteria As String
Public LGNCriteria As String
Public CTYCriteria As String
Public STACriteria As String
Public WorkCriteria As String
Public StatusCriteria As String
Public PolicyStatusCriteria As String
Public CPTCriteria As String
Public RecTypeCriteria As String


Private Sub CboLifeTimePln_Change()
    
    If InStr(CboLifeTimePln.Text, "null") Then
       CboLifeTimePln.Enabled = False
    End If
    
End Sub

Private Sub CboLifeTimePln_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboLifeTimePln.Text, CboLifeTimePln.SelStart) + Chr(KeyAscii) + Mid(CboLifeTimePln.Text, CboLifeTimePln.SelStart + CboLifeTimePln.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboLifeTimePln.ListCount - 1
 
        If NewText = LCase(Left(CboLifeTimePln.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboLifeTimePln.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboLifeTimePln.Text = ValidValue
               CboLifeTimePln.SelStart = 0
               CboLifeTimePln.SelLength = Len(ValidValue)
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

'***********************Start EntriesEnable***************************
Private Sub EntriesEnable(status As Boolean)

Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        End If
    Next ctl
    
End Sub
'***********************Start EntriesEnable***************************

'**********************Start Controlling of ComboList**************************

Private Sub CboStateCode_Change()
If InStr(CboStateCode.Text, "null") Then
       CboStateCode.Enabled = False
    End If
End Sub

Private Sub CboStateCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboStateCode.Text, CboStateCode.SelStart) + Chr(KeyAscii) + Mid(CboStateCode.Text, CboStateCode.SelStart + CboStateCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboStateCode.ListCount - 1
 
        If NewText = LCase(Left(CboStateCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboStateCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CboStateCode.Text = ValidValue
               CboStateCode.SelStart = 0
               CboStateCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CmdClear_Click()
Call ClearForm(Me)
    lstAgeBandList.listitems.Clear
    lstOccupationList.listitems.Clear
    lstRelationshipList.listitems.Clear
    lstNationalityList.listitems.Clear
    lstRaceList.listitems.Clear
    lstLanguageList.listitems.Clear
    lstCityList.listitems.Clear
    lstStateList.listitems.Clear
    lstWorkList.listitems.Clear
    lstStatusList.listitems.Clear
    lstPolicyList.listitems.Clear
    lstCPTList.listitems.Clear
    lstRecTypeList.listitems.Clear
    txtTotalRecord = 0
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
End Sub

Private Sub cmdPrint_Click()
  
 Screen.MousePointer = vbHourglass
    
    EditFlag = False
    Call ClearForm(Me)
    Call EntriesEnable(True)
       
    Select Case SSTab1.Tab
        Case 0
            If lstAgeBandList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstAgeBandList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 1
            If lstOccupationList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstOccupationList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 2
            If lstRelationshipList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstRelationshipList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 3
            If lstNationalityList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstNationalityList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 4
            If lstRaceList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstRaceList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 5
            If lstLanguageList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstLanguageList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 6
            If lstCityList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstCityList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 7
            If lstStateList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstStateList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 8
            If lstWorkList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstWorkList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 9
            If lstStatusList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstStatusList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 10
            If lstPolicyList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstPolicyList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 11
            If lstCPTList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstCPTList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 12
            If lstRecTypeList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstRecTypeList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
    End Select
    'Select Case SSTab1.Tab
        'Case 0
            'Call PrintAgeBand
        'Case 1
            'Call PrintOccupation
        'Case 2
            'Call PrintRelationship
        'Case 3
            'Call PrintNationality
        'Case 4
            'Call PrintRace
        'Case 5
            'Call PrintLanguage
        'Case 6
            'Call PrintCity
        'Case 7
            'Call PrintState
        'Case 8
            'Call PrintWork
        'Case 9
            'Call PrintStatus
        'Case 10
            'Call PrintPStatus
        'Case 11
            'Call PrintCPT
        'Case 12
            'Call PrintRecType
    'End Select
Screen.MousePointer = vbDefault
End Sub

Private Sub CmdRefresh_Click()
    
Call ClearForm(Me)
Call EntriesEnable(True)

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
Call CTYComboListing
'medixmasterconn.Close
'Set medixmasterconn = Nothing

End Sub

Private Sub lstAgeBandList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstAgeBandList.SortKey = ColumnHeader.Index - 1
    lstAgeBandList.Sorted = True
    If lstAgeBandList.SortOrder = lvwAscending Then
        lstAgeBandList.SortOrder = lvwDescending
    Else
        lstAgeBandList.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstAgeBandList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

    If lstAgeBandList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtAgeCode = lstAgeBandList.SelectedItem.Text
        TxtAgeFr = lstAgeBandList.SelectedItem.SubItems(1)
        TxtAgeTo = lstAgeBandList.SelectedItem.SubItems(2)
        TxtDesc = lstAgeBandList.SelectedItem.SubItems(3)
        CboLifeTimePln = lstAgeBandList.SelectedItem.SubItems(4)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If

'If lstAgeBandList.listitems.count Then
    'Call EnableCmdButton(Me)
    'Call ClearForm(Me)

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    'StrSql = "SELECT agecode, hltcode, agedescription, agefrom, ageto " & _
    "from AGE Where agecode = '" & lstAgeBandList.SelectedItem.Text & "'"
    'rst3.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'TxtAgeCode = rst3!AgeCode
    'TxtDesc = rst3!AGEDescription
    'TxtAgeFr = rst3!AGEFrom
    'TxtAgeTo = rst3!AgeTo
    'Call EntriesEnable(False)
    'CmdSave.Enabled = False
    'rst3.Close
    'Set rst3 = Nothing
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
'End If
End Sub

Private Sub lstAgeBandList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

    If lstAgeBandList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtAgeCode = lstAgeBandList.SelectedItem.Text
        TxtAgeFr = lstAgeBandList.SelectedItem.SubItems(1)
        TxtAgeTo = lstAgeBandList.SelectedItem.SubItems(2)
        TxtDesc = lstAgeBandList.SelectedItem.SubItems(3)
        CboLifeTimePln = lstAgeBandList.SelectedItem.SubItems(4)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
'If lstAgeBandList.listitems.count Then
    'Call EnableCmdButton(Me)
    'Call ClearForm(Me)

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    'StrSql = "SELECT agecode, hltcode, agedescription, agefrom, ageto " & _
    "from AGE Where agecode = '" & lstAgeBandList.SelectedItem.Text & "'"
    'rst3.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'TxtAgeCode = rst3!AgeCode
    'TxtDesc = rst3!AGEDescription
    'TxtAgeFr = rst3!AGEFrom
    'TxtAgeTo = rst3!AgeTo
    'Call EntriesEnable(False)
    'CmdSave.Enabled = False
    'rst3.Close
    'Set rst3 = Nothing
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
'End If
End Sub

Private Sub lstCityList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCityList.SortKey = ColumnHeader.Index - 1
    lstCityList.Sorted = True
    If lstCityList.SortOrder = lvwAscending Then
        lstCityList.SortOrder = lvwDescending
    Else
        lstCityList.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCityList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String
    
        
    If lstCityList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
    '    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "SELECT CTYcode, CTYDescription, STADescription " & _
        "from CTY Where CTYcode = '" & lstCityList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        txtCTYcode = rst3!CTYCode
        txtCTYdesc = rst3!CTYDescription
        CboStateCode = rst3!STADescription
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
     '   medixmasterconnMaint.Close
     '   Set medixmasterconnMaint = Nothing
        
    End If
End Sub

Private Sub lstCityList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String
    
        
    If lstCityList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
       ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "SELECT CTYcode, CTYDescription, STADescription " & _
                 "from CTY Where CTYcode = '" & lstCityList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        txtCTYcode = rst3!CTYCode
        txtCTYdesc = rst3!CTYDescription
        CboStateCode = rst3!STADescription
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
    '    medixmasterconnMaint.Close
    '    Set medixmasterconnMaint = Nothing
        
    End If
End Sub


Private Sub lstLanguageList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstLanguageList.SortKey = ColumnHeader.Index - 1
lstLanguageList.Sorted = True
If lstLanguageList.SortOrder = lvwAscending Then
    lstLanguageList.SortOrder = lvwDescending
Else
    lstLanguageList.SortOrder = lvwAscending
End If
End Sub

Private Sub lstLanguageList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstLanguageList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    strsql = "SELECT LGNcode, LGNDescription " & _
             "from LGN Where LGNcode = '" & lstLanguageList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    txtLGNcode = rst3!LGNCode
    txtLGNdesc = rst3!LGNDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End If
End Sub

Private Sub lstLanguageList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstLanguageList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
 '   medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    strsql = "SELECT LGNcode, LGNDescription " & _
             "from LGN Where LGNcode = '" & lstLanguageList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    txtLGNcode = rst3!LGNCode
    txtLGNdesc = rst3!LGNDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
 '   medixmasterconnMaint.Close
'    Set medixmasterconnMaint = Nothing
    
End If
End Sub

Private Sub lstNationalityList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstNationalityList.SortKey = ColumnHeader.Index - 1
lstNationalityList.Sorted = True
If lstNationalityList.SortOrder = lvwAscending Then
    lstNationalityList.SortOrder = lvwDescending
Else
    lstNationalityList.SortOrder = lvwAscending
End If
End Sub

Private Sub lstNationalityList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstNationalityList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT NATCode, NATName " & _
             "from NAT Where NATCode = '" & lstNationalityList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    txtNATcode = rst3!NATCode
    txtNATname = rst3!NATName
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconn.Close
 '   Set medixmasterconn = Nothing
    
End If
End Sub

Private Sub lstNationalityList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstNationalityList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT NATCode, NATName " & _
             "from NAT Where NATCode = '" & lstNationalityList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    txtNATcode = rst3!NATCode
    txtNATname = rst3!NATName
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End If
End Sub

Private Sub lstOccupationList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstOccupationList.SortKey = ColumnHeader.Index - 1
lstOccupationList.Sorted = True
If lstOccupationList.SortOrder = lvwAscending Then
    lstOccupationList.SortOrder = lvwDescending
Else
    lstOccupationList.SortOrder = lvwAscending
End If
End Sub

Private Sub lstOccupationList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstOccupationList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT OCPcode, OCPdescription " & _
    "from OCP Where OCPcode = '" & lstOccupationList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TxtOCPcode = rst3!OCPCode
    TxtOCPDesc = rst3!OCPDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If
End Sub

Private Sub lstOccupationList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstOccupationList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT OCPcode, OCPdescription " & _
    "from OCP Where OCPcode = '" & lstOccupationList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TxtOCPcode = rst3!OCPCode
    TxtOCPDesc = rst3!OCPDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
   '
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End If
End Sub

Private Sub lstOccupationList1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstOccupationList1.SortKey = ColumnHeader.Index - 1
lstOccupationList1.Sorted = True
If lstOccupationList1.SortOrder = lvwAscending Then
    lstOccupationList1.SortOrder = lvwDescending
Else
    lstOccupationList1.SortOrder = lvwAscending
End If
End Sub

Private Sub lstPolicyList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

        
    If lstPolicyList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        strsql = "SELECT PStatuscode, PStatusDesc " & _
        "from PolicyStatus Where PStatuscode = '" & lstPolicyList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        TxtPStatusCode = rst3!pstatuscode
        TxtPStatusDesc = rst3!pstatusDesc
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstPolicyList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

        
    If lstPolicyList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        strsql = "SELECT PStatuscode, PStatusDesc " & _
        "from PolicyStatus Where PStatuscode = '" & lstPolicyList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        TxtPStatusCode = rst3!pstatuscode
        TxtPStatusDesc = rst3!pstatusDesc
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
    End If
End Sub

Private Sub lstRaceList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstRaceList.SortKey = ColumnHeader.Index - 1
lstRaceList.Sorted = True
If lstRaceList.SortOrder = lvwAscending Then
    lstRaceList.SortOrder = lvwDescending
Else
    lstRaceList.SortOrder = lvwAscending
End If
End Sub

Private Sub lstRaceList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstRaceList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   '
    strsql = "SELECT RACcode, RACName " & _
    "from RAC Where RACcode = '" & lstRaceList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    txtRACcode = rst3!RACCode
    txtRACDesc = rst3!RACName
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End If
End Sub

Private Sub lstRaceList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstRaceList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT RACcode, RACName " & _
    "from RAC Where RACcode = '" & lstRaceList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    txtRACcode = rst3!RACCode
    txtRACDesc = rst3!RACName
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If
End Sub

Private Sub lstRelationshipList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstRelationshipList.SortKey = ColumnHeader.Index - 1
lstRelationshipList.Sorted = True
If lstRelationshipList.SortOrder = lvwAscending Then
    lstRelationshipList.SortOrder = lvwDescending
Else
    lstRelationshipList.SortOrder = lvwAscending
End If
End Sub

Private Sub lstRelationshipList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstRelationshipList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT RELcode, RELdescription " & _
    "from REL Where RELcode = '" & lstRelationshipList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    txtRELcode = rst3!RELCODE
    txtRELdesc = rst3!RELDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If

End Sub

Private Sub lstRelationshipList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstRelationshipList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT RELcode, RELdescription " & _
    "from REL Where RELcode = '" & lstRelationshipList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    txtRELcode = rst3!RELCODE
    txtRELdesc = rst3!RELDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End If

End Sub

Private Sub lstStateList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstStateList.SortKey = ColumnHeader.Index - 1
lstStateList.Sorted = True
If lstStateList.SortOrder = lvwAscending Then
    lstStateList.SortOrder = lvwDescending
Else
    lstStateList.SortOrder = lvwAscending
End If
End Sub

Private Sub lstStateList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String
Dim color As ColorConstants

color = vbRed

    
If lstStateList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT STAcode, STADescription " & _
    "from STA Where STAcode = '" & lstStateList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
    txtSTAcode = rst3!STACode
    txtSTAdesc = rst3!STADescription
        
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    
    rst3.Close
    Set rst3 = Nothing
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End If
End Sub

Private Sub lstStateList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String
Dim color As ColorConstants

color = vbRed

    
If lstStateList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT STAcode, STADescription " & _
    "from STA Where STAcode = '" & lstStateList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
    txtSTAcode = rst3!STACode
    txtSTAdesc = rst3!STADescription
        
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    
    rst3.Close
    Set rst3 = Nothing
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End If
End Sub

Private Sub lstStatusList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

    
If lstStatusList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT Statuscode, StatusDesc " & _
    "from Status Where Statuscode = '" & lstStatusList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TxtStatusCode = rst3!StatusCode
    TxtStatusDesc = rst3!StatusDesc
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If
End Sub

Private Sub lstStatusList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

    
If lstStatusList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT Statuscode, StatusDesc " & _
    "from Status Where Statuscode = '" & lstStatusList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TxtStatusCode = rst3!StatusCode
    TxtStatusDesc = rst3!StatusDesc
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If
End Sub

Private Sub lstWorkList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String
    
    If lstWorkList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "Select WorkID, WorkDescription " & _
        "from WorkNature Where WorkID = '" & lstWorkList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        TxtWorkCode = Trim(rst3!WorkID)
        TxtWorkDescription = Trim(rst3!WorkDescription)
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconnMaint.Close
      '  Set medixmasterconnMaint = Nothing
        
    End If
End Sub

Private Sub lstWorkList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String
    
    If lstWorkList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
     '   medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "Select WorkID, WorkDescription " & _
        "from WorkNature Where WorkID = '" & lstWorkList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        TxtWorkCode = Trim(rst3!WorkID)
        TxtWorkDescription = Trim(rst3!WorkDescription)
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
    '    medixmasterconnMaint.Close
      '  Set medixmasterconnMaint = Nothing
      '
    End If
End Sub

Private Sub TxtAgeCode_KeyPress(KeyAscii As Integer)
KeyAscii = KeyNumber(KeyAscii)
End Sub
'**********************End Controlling of ComboList**************************

'**********************Start Form *******************************************
'Private Sub Form_Activate()
    'SSTab1.Tab = 0
    'CmdRefresh.Enabled = False
'End Sub

Private Sub Form_Load()

    Me.Width = SSTab1.Width + 500

    AGECriteria = ""
    OCPCriteria = ""
    RELCriteria = ""
    NATCriteria = ""
    RACCriteria = ""
    LGNCriteria = ""
    CTYCriteria = ""
    STACriteria = ""
    WorkCriteria = ""
    CPTCriteria = ""
    
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Call AgebandComboListing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Call EnterToTab(KeyAscii)
End Sub
'**********************End Form *******************************************

'*****************************Start Tab Clicking**********************************

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Screen.MousePointer = vbHourglass
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    Select Case SSTab1.Tab
    Case 0
        CmdRefresh.Enabled = False
        lstAgeBandList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
         
    Case 1
        CmdRefresh.Enabled = False
        lstOccupationList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
         
    Case 2
        CmdRefresh.Enabled = False
        lstRelationshipList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
        
    Case 3
        CmdRefresh.Enabled = False
        lstNationalityList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
        
    Case 4
        CmdRefresh.Enabled = False
        lstRaceList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
        
    Case 5
        CmdRefresh.Enabled = False
        lstLanguageList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
        
    Case 6
        CmdRefresh.Enabled = True
        lstCityList.listitems.Clear
        txtTotalRecord = 0
        Call CTYComboListing
        Call ClearForm(Me)
        
    Case 7
        CmdRefresh.Enabled = False
        lstStateList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
         
    Case 8
        CmdRefresh.Enabled = False
        lstWorkList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
        
    Case 9
        CmdRefresh.Enabled = False
        lstStatusList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
        
    Case 10
        CmdRefresh.Enabled = False
        lstPolicyList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
    
    Case 11
        CmdRefresh.Enabled = False
        lstCPTList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
    Case 12
        CmdRefresh.Enabled = False
        lstRecTypeList.listitems.Clear
        txtTotalRecord = 0
        Call ClearForm(Me)
    End Select
 '   medixmasterconn.Close
 '   Set medixmasterconn = Nothing
 '   medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing

    Screen.MousePointer = Default
End Sub
'*****************************End Tab Clicking**********************************

'*******************************Start Exit, Edit, Add, Delete, Search *******************************
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdEdit_Click()
EditFlag = True
    Call EntriesEnable(True)
    TxtAgeCode.SetFocus
    TxtOCPcode.SetFocus
    txtRELcode.SetFocus
    CmdAdd.Enabled = True
    CmdDelete.Enabled = False
    CmdSave.Enabled = True
    
    If TxtAgeCode <> "" Then
       TxtAgeCode.Enabled = False
       Frame1.Enabled = True
       TxtDesc.SetFocus
    End If
    
    If TxtOCPcode <> "" Then
       TxtOCPcode.Enabled = False
       TxtOCPDesc.SetFocus
    End If
    
    If txtRELcode <> "" Then
       txtRELcode.Enabled = False
       txtRELdesc.SetFocus
    End If
    
    If txtNATcode <> "" Then
       txtNATcode.Enabled = False
       txtNATname.SetFocus
    End If
    
    If txtRACcode <> "" Then
       txtRACcode.Enabled = False
       txtRACDesc.SetFocus
    End If
    
    If txtLGNcode <> "" Then
       txtLGNcode.Enabled = False
       txtLGNdesc.SetFocus
    End If
    
    If txtCTYcode <> "" Then
       txtCTYcode.Enabled = False
       txtCTYdesc.SetFocus
    End If
    
    If txtSTAcode <> "" Then
       txtSTAcode.Enabled = False
       txtSTAdesc.SetFocus
    End If
    
    If TxtWorkCode <> "" Then
       TxtWorkCode.Enabled = False
       TxtWorkDescription.SetFocus
    End If
    
    If TxtStatusCode <> "" Then
       TxtStatusCode.Enabled = False
       TxtStatusDesc.SetFocus
    End If
    
    If TxtPStatusCode <> "" Then
       TxtPStatusCode.Enabled = False
       TxtPStatusDesc.SetFocus
    End If
    
    If TxtCPT <> "" Then
       TxtCPT.Enabled = False
       TxtCPTDesc.SetFocus
    End If
    
    If TxtRecTypeCode <> "" Then
       TxtRecTypeCode.Enabled = False
       TxtRecTypeDesc.SetFocus
    End If
    
End Sub

Private Sub CmdAdd_Click()
    EditFlag = False
    Call ClearForm(Me)
    Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    
    Select Case SSTab1.Tab
        Case 0
            TxtAgeCode.SetFocus
        Case 1
            TxtOCPcode.SetFocus
        Case 2
            txtRELcode.SetFocus
        Case 3
            txtNATcode.SetFocus
        Case 4
            txtRACcode.SetFocus
        Case 5
            txtLGNcode.SetFocus
        Case 6
            txtCTYcode.SetFocus
        Case 7
            txtSTAcode.SetFocus
        Case 8
            TxtWorkCode.SetFocus
        Case 9
            TxtStatusCode.SetFocus
        Case 10
            TxtPStatusCode.SetFocus
        Case 11
            TxtCPT.SetFocus
        Case 12
            TxtRecTypeCode.SetFocus
    End Select

End Sub

Private Sub cmdDelete_Click()
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

   If SSTab1.Tab = 0 Then
      Call DeleteAgeBand
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 1 Then
      Call DeleteOccupation
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 2 Then
      Call DeleteRelationship
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 3 Then
      Call DeleteNationality
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 4 Then
      Call DeleteRace
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 5 Then
      Call DeleteLanguage
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 6 Then
      Call DeleteCity
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 7 Then
      Call DeleteState
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 8 Then
      Call DeleteWork
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 9 Then
      Call DeleteStatus
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 10 Then
      Call DeletePStatus
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 11 Then
      Call DeleteCPT
      CmdAdd.Enabled = True
   ElseIf SSTab1.Tab = 12 Then
      Call DeleteRecType
      CmdAdd.Enabled = True
   End If
 
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing

End Sub

Private Sub CmdSearch_Click()
    CmdSearch.Enabled = False
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

Select Case SSTab1.Tab
        Case 0
            Call SearchAge
            TxtAgeCode.Enabled = True
        
        Case 1
            Call SearchOCP
            TxtOCPcode.Enabled = True
            TxtOCPDesc.Enabled = True
        
        Case 2
            Call SearchRelationship
            txtRELcode.Enabled = True
            txtRELdesc.Enabled = True
        
        Case 3
            Call SearchNationality
            txtNATcode.Enabled = True
            txtNATname.Enabled = True
        
        Case 4
            Call SearchRace
            txtRACcode.Enabled = True
            txtRACDesc.Enabled = True
        
        Case 5
            Call SearchLanguage
            txtLGNcode.Enabled = True
            txtLGNdesc.Enabled = True
            
        Case 6
            Call SearchCity
            txtCTYcode.Enabled = True
            txtCTYdesc.Enabled = True
            
        Case 7
            Call SearchState
            txtSTAcode.Enabled = True
            txtSTAdesc.Enabled = True
            
        Case 8
            Call SearchWork
            TxtWorkCode.Enabled = True
            TxtWorkDescription.Enabled = True
            
        Case 9
            Call SearchStatus
            TxtStatusCode.Enabled = True
            TxtStatusDesc.Enabled = False
            
        Case 10
            Call SearchPStatus
            TxtPStatusCode.Enabled = True
            TxtPStatusDesc.Enabled = False
        
        Case 11
            Call SearchCPT
            TxtCPT.Enabled = True
            TxtCPTDesc.Enabled = False
        
        Case 12
            Call SearchRecType
            TxtRecTypeCode.Enabled = True
            TxtRecTypeDesc.Enabled = False
        End Select
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
 '   medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    
    CmdSearch.Enabled = True
    CmdAdd.Enabled = True
End Sub

Private Sub cmdSave_Click()
'    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Select Case SSTab1.Tab
        Case 0
            Call SaveAgeBand
        Case 1
            Call SaveOccupation
        Case 2
            Call SaveRelationship
        Case 3
            Call SaveNationality
        Case 4
            Call SaveRace
        Case 5
            Call SaveLanguage
        Case 6
            Call SaveCity
        Case 7
            Call SaveState
        Case 8
            Call SaveWork
        Case 9
            Call SaveStatus
        Case 10
            Call SavePStatus
        Case 11
            Call SaveCPT
        Case 12
            Call SaveRecType
    End Select
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing

    CmdAdd.Enabled = True
    CmdEdit.Enabled = False

 
End Sub
'*******************************End Exit, Edit, Add, Delete, Search *******************************

'*******************************Start Age Band *******************************
Private Sub SaveAgeBand()
'On Error Resume Next
On Error GoTo SaveErr
Dim RecordSaved
Dim AGE As New adodb.Recordset
Dim SaveAGE As New adodb.Recordset
Dim strsql As String
    
    If Len(Trim(TxtAgeCode)) = 0 Then
        MsgBox "Please Enter Age Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtAgeCode.Enabled = True
        TxtAgeCode.SetFocus
 '   ElseIf CblHealth = "" Then
  '      MsgBox "Please Enter Health Plan", vbOKOnly + vbCritical, DialogBoxTitle
   '     CblHealth.Enabled = True
    '    CblHealth.SetFocus
    ElseIf TxtDesc = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtDesc.Enabled = True
        TxtDesc.SetFocus
    ElseIf TxtAgeFr = "" Then
        MsgBox "Please Enter Age From", vbOKOnly + vbCritical, DialogBoxTitle
        TxtAgeFr.Enabled = True
        TxtAgeFr.SetFocus
    ElseIf TxtAgeTo = "" Then
        MsgBox "Please Enter Age To", vbOKOnly + vbCritical, DialogBoxTitle
        TxtAgeTo.Enabled = True
        TxtAgeTo.SetFocus
    ElseIf TxtAgeFr.Text = TxtAgeTo.Text Then
           MsgBox "Invalid Age", vbCritical + vbOKOnly, "Error"
           TxtAgeFr.SetFocus
           TxtAgeFr.SelStart = 0
           TxtAgeFr.SelLength = Len(TxtAgeFr)
    Else
           
        strsql = "Select AGECode, PLNCode From AGE "
        
        strsql = strsql & "where (AGECode = '" & Trim(TxtAgeCode) & "') AND (PLNCode = '" & Trim(CboLifeTimePln) & "')"
        
        AGE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
           If AGE.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                AGE.Close
                TxtAgeCode.SetFocus
                TxtAgeCode.SelStart = 0
                TxtAgeCode.SelLength = Len(TxtAgeCode)
                Set AGE = Nothing
                Exit Sub
           Else
              RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           
              If RecordSaved = vbYes Then
                 If EditFlag = False Then
                    SaveAGE.Open "AGE", medixmasterconn, adOpenKeyset, adLockOptimistic
                    SaveAGE.AddNew
                 Else
                    strsql = "Select * From AGE Where AgeCode = '" & Trim(TxtAgeCode) & "' AND PLNCode = '" & Trim(CboLifeTimePln) & "'"
                    SaveAGE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                 End If
            
                 With SaveAGE
                    !AgeCode = ChkStr(TxtAgeCode)
                    !PLNCode = ChkStr(CboLifeTimePln)
                    !AGEDescription = ChkStr(TxtDesc)
                    !AGEFrom = ChkStr(TxtAgeFr)
                    !AgeTo = ChkStr(TxtAgeTo)
                    !AGELAstUpdateUser = Logon
                    !AGELastUpdateDate = Date
                 End With
                 SaveAGE.Update
                 Call ClearForm(Me)
                 Call AgeListing(AGECriteria)
                    If EditFlag = False Then
                       MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                       MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
                 SaveAGE.Close
                 Set SaveAGE = Nothing
              End If
        EditFlag = False
        
    End If
End If
Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub DeleteAgeBand()
Dim DeleteAge As New adodb.Recordset
Dim rsCheckMBM As New adodb.Recordset
Dim rsCheckPRM As New adodb.Recordset
Dim strsql, sqlMBM, sqlPRM As String
Dim DeleteRecord

EditFlag = False
If TxtAgeCode = "" Then
   MsgBox "Please Enter Age Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtAgeCode.SetFocus
ElseIf lstAgeBandList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select AGECode  From MBM"
   sqlMBM = sqlMBM & " Where (AGECode = '" & Trim(TxtAgeCode) & "')"
   
   sqlPRM = "Select AGECode From PRM"
   sqlPRM = sqlPRM & " Where (AGECode = '" & Trim(TxtAgeCode) & "') "
   
   
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   rsCheckPRM.MaxRecords = 1
   rsCheckPRM.Open sqlPRM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        Call ClearForm(Me)
   ElseIf rsCheckPRM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Premium Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From AGE Where Agecode = '" & Trim(TxtAgeCode) & "' AND PLNCode = '" & Trim(CboLifeTimePln) & "'" 'ChkStr(lstAgeBandList.SelectedItem.Text) & "'"
           DeleteAge.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call AgeListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           TxtAgeCode.SetFocus
         End If
     End If
End If
End Sub

Private Sub SearchAge()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(TxtAgeCode)) Then
    strAppend = strAppend + " And AgeCode like '" & Mid(ChkStr(TxtAgeCode), 1, 2) & "%'"
End If

If Len(Trim(CboLifeTimePln)) Then
    strAppend = strAppend + " And PLNCode = '" & Trim(CboLifeTimePln) & "'"
End If

AGECriteria = strAppend
Call AgeListing(strAppend)
Screen.MousePointer = Default
End Sub

Private Sub AgebandComboListing()
Dim LifeTimePLN As New adodb.Recordset
Dim cmd As New adodb.Command
    
    ' Set up a command object for the stored procedure.
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchLifeTimePlan"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    LifeTimePLN.CursorLocation = adUseClient
    LifeTimePLN.CursorType = adOpenForwardOnly
    LifeTimePLN.CacheSize = 100

    Set LifeTimePLN = cmd.Execute
    
    CboLifeTimePln.Clear
        
    Do Until LifeTimePLN.EOF
        CboLifeTimePln.AddItem LifeTimePLN!PLNCode
        LifeTimePLN.MoveNext
    Loop
    
    LifeTimePLN.Close
    Set LifeTimePLN = Nothing
    
End Sub

Private Sub AgeListing(Optional strAppend As String)
Dim rst2 As New adodb.Recordset
Dim AgeMaster As ListItem
Dim Sql As String
Dim TotalAgeCount As String

TotalAgeCount = 0
lstAgeBandList.listitems.Clear

    Sql = " SELECT AGECode, AGEDescription, AGEFrom, AgeTo, PLNCode " & _
          " From AGE Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
    
    rst2.Open Sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until rst2.EOF
        Set AgeMaster = lstAgeBandList.listitems.Add(, , rst2!AgeCode)
            AgeMaster.SubItems(1) = Trim(rst2!AGEFrom)
            AgeMaster.SubItems(2) = Trim(rst2!AgeTo)
            AgeMaster.SubItems(3) = Trim(rst2!AGEDescription)
            If Len(rst2!PLNCode) Then
                AgeMaster.SubItems(4) = Trim(rst2!PLNCode)
            End If
            rst2.MoveNext
            
        TotalAgeCount = TotalAgeCount + 1
        txtTotalRecord = TotalAgeCount
        Loop
            
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub lstAgeBandList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String

    If lstAgeBandList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtAgeCode = lstAgeBandList.SelectedItem.Text
        TxtAgeFr = lstAgeBandList.SelectedItem.SubItems(1)
        TxtAgeTo = lstAgeBandList.SelectedItem.SubItems(2)
        TxtDesc = lstAgeBandList.SelectedItem.SubItems(3)
        CboLifeTimePln = lstAgeBandList.SelectedItem.SubItems(4)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
    End If
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    'StrSql = "SELECT agecode, hltcode, agedescription, agefrom, ageto " & _
    "from AGE Where agecode = '" & lstAgeBandList.SelectedItem.Text & "'"
    'rst3.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'TxtAgeCode = rst3!AgeCode
    'TxtDesc = rst3!AGEDescription
    'TxtAgeFr = rst3!AGEFrom
    'TxtAgeTo = rst3!AgeTo
    'Call EntriesEnable(False)
    'CmdSave.Enabled = False
    'rst3.Close
    'Set rst3 = Nothing
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
'End If
End Sub

'Private Sub PrintAgeBand()
    'RptAgebandList.Show
'End Sub
'*******************************End Age Band *******************************

'*******************************Start Occupation *******************************
Private Sub SaveOccupation()
On Error Resume Next
Dim RecordSaved
Dim OCP As New adodb.Recordset
Dim SaveOCP As New adodb.Recordset
Dim strsql As String
    

    If TxtOCPcode = "" Then
       MsgBox "Please Enter Occupation Code", vbOKOnly + vbCritical, DialogBoxTitle
       TxtOCPcode.SetFocus
    ElseIf TxtOCPDesc = "" Then
       MsgBox "Please Enter Occupation Description", vbOKOnly + vbCritical, DialogBoxTitle
       TxtOCPDesc.SetFocus
    Else
       strsql = "Select * From OCP Where OCPCode = '" & Trim(TxtOCPcode) & "'"
       OCP.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
      
              If OCP.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                OCP.Close
                TxtOCPcode.SetFocus
                TxtOCPcode.SelStart = 0
                TxtOCPcode.SelLength = Len(TxtOCPcode)
                Set OCP = Nothing
                Exit Sub
            Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
         If RecordSaved = vbYes Then
            If EditFlag = False Then
                OCP.Open "OCP", medixmasterconn, adOpenKeyset, adLockOptimistic
                OCP.AddNew
            Else
                strsql = "Select * From OCP Where OCPCode = '" & Trim(TxtOCPcode) & "'"
                OCP.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With OCP
                    !OCPCode = ChkStr(TxtOCPcode)
                    !OCPDescription = ChkStr(TxtOCPDesc)
                    !OCPLastUpdateUser = Logon
                    !OCPLAstUpdateTime = Date
                End With
                OCP.Update
                Call ClearForm(Me)
                Call OCPListing(OCPCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
            OCP.Close
            Set SaveOCP = Nothing
        End If
        EditFlag = False
    End If
End If

End Sub

Private Sub DeleteOccupation()
Dim DeleteOCP As New adodb.Recordset
Dim rsCheckMBMOther As New adodb.Recordset
Dim strsql, sqlMBMOther As String
Dim DeleteRecord

EditFlag = False
If TxtOCPcode = "" Then
   MsgBox "Please Enter Occupation Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtOCPcode.SetFocus
ElseIf lstOccupationList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBMOther = "Select MBMNumber from MBMOthersTwo where MBMOccupation = '" & RemStr(TxtOCPcode) & "'"
   
      
   rsCheckMBMOther.MaxRecords = 1
   rsCheckMBMOther.Open sqlMBMOther, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckMBMOther.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBMOther.recordCount = 0 Then
           strsql = "Delete From OCP Where OCPcode = '" & Trim(TxtOCPcode) & "'" 'ChkStr(lstOccupationList.SelectedItem.Text) & "'"
           DeleteOCP.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call OCPListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           TxtOCPcode.SetFocus
         End If
     End If
End If
End Sub

Private Sub SearchOCP()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(TxtOCPcode)) Then
strAppend = strAppend + " And OCPCode like '%" & Mid(ChkStr(TxtOCPcode), 1, 2) & "%'"
End If

If Len(Trim(TxtOCPDesc)) Then
strAppend = strAppend + " And OCPDescription like '%" & ChkStr(TxtOCPDesc) & "%'"
End If

OCPCriteria = strAppend
Call OCPListing(strAppend)
Screen.MousePointer = Default
End Sub

Private Sub OCPListing(Optional strAppend As String)
Dim rst2 As New adodb.Recordset
Dim OcpMaster As ListItem
Dim Sql As String
Dim TotalOCPCount As String

TotalOCPCount = 0
lstOccupationList.listitems.Clear
    Sql = "SELECT OCPcode," & _
    " OCPDescription" & _
    " from OCP where 0=0 "
    
    If Len(Trim(strAppend)) Then
    Sql = Sql + strAppend
    End If
    
    rst2.Open Sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until rst2.EOF
        Set OcpMaster = lstOccupationList.listitems.Add(, , rst2!OCPCode)
            OcpMaster.SubItems(1) = rst2!OCPDescription
            rst2.MoveNext
        TotalOCPCount = TotalOCPCount + 1
        txtTotalRecord = TotalOCPCount
        Loop
    rst2.Close
    Set rst2 = Nothing
End If
End Sub

Private Sub lstOccupationList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstOccupationList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT OCPcode, OCPdescription " & _
    "from OCP Where OCPcode = '" & lstOccupationList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TxtOCPcode = rst3!OCPCode
    TxtOCPDesc = rst3!OCPDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
'    medixmasterconn.Close
 '   Set medixmasterconn = Nothing
    
End If
End Sub

'Private Sub PrintOccupation()
    'RptOcpList.Show
'End Sub
'*******************************End Occupation *******************************

'*******************************Start Relationship *******************************
Private Sub SaveRelationship()
On Error Resume Next
Dim RecordSaved
Dim REL1 As New adodb.Recordset
Dim SaveREL As New adodb.Recordset
Dim strsql As String
 
    If txtRELcode = "" Then
        MsgBox "Please Enter Relationship Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtRELcode.SetFocus
    ElseIf txtRELdesc = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtRELdesc.SetFocus
    
    Else
         strsql = "Select * From REL Where RELCode = '" & Trim(txtRELcode) & "'"
          REL1.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If REL1.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                REL1.Close
                txtRELcode.SetFocus
                txtRELcode.SelStart = 0
                txtRELcode.SelLength = Len(txtRELcode)
                Set REL1 = Nothing
                Exit Sub
            Else
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
         If RecordSaved = vbYes Then
            If EditFlag = False Then
                REL1.Open "REL", medixmasterconn, adOpenKeyset, adLockOptimistic
                REL1.AddNew
            Else
                strsql = "Select * From REL Where RELCode = '" & Trim(txtRELcode) & "'"
                REL1.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With REL1
                    !RELCODE = ChkStr(txtRELcode)
                    !RELDescription = ChkStr(txtRELdesc)
                    !RELLastUpdateUser = Logon
                    !RELLastUpdateDate = Date
                End With
                REL1.Update
                Call ClearForm(Me)
                Call RELListing(RELCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        REL1.Close
        Set SaveREL = Nothing
        End If
        EditFlag = False
    End If
End If

End Sub

Private Sub DeleteRelationship()
Dim DeleteREL As New adodb.Recordset
Dim rsCheckMBMOther As New adodb.Recordset
Dim strsql, sqlMBMOther As String
Dim DeleteRecord

EditFlag = False
If txtRELcode = "" Then
   MsgBox "Please Enter Relationship Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtRELcode.SetFocus
ElseIf lstRelationshipList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBMOther = "Select MBMNumber from MBMOthers where RELCode = '" & RemStr(txtRELcode) & "'"
      
   rsCheckMBMOther.MaxRecords = 1
   rsCheckMBMOther.Open sqlMBMOther, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckMBMOther.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBMOther.recordCount = 0 Then
           strsql = "Delete From REL Where RELcode = '" & Trim(txtRELcode) & "'" 'ChkStr(lstRelationshipList.SelectedItem.Text) & "'"
           DeleteREL.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call RELListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtRELcode.SetFocus
         End If
     End If
End If
End Sub

Private Sub SearchRelationship()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(txtRELcode)) Then
strAppend = strAppend + " And RELCode like '%" & Mid(ChkStr(txtRELcode), 1, 2) & "%'"
End If

If Len(Trim(txtRELdesc)) Then
strAppend = strAppend + " And RELDescription like '%" & ChkStr(txtRELdesc) & "%'"
End If

RELCriteria = strAppend
Call RELListing(strAppend)
Screen.MousePointer = Default
End Sub

Private Sub RELListing(Optional strAppend As String)
Dim REL2 As New adodb.Recordset
Dim RELMaster As ListItem
Dim Sql As String
Dim TotalRELCount As String

TotalRELCount = 0
lstRelationshipList.listitems.Clear


    Sql = "SELECT RELcode, " & _
    " RELDescription" & _
    " from REL where 0=0 "
    
    If Len(Trim(strAppend)) Then
    Sql = Sql + strAppend
    End If
    
    REL2.Open Sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If REL2.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until REL2.EOF
        Set RELMaster = lstRelationshipList.listitems.Add(, , REL2!RELCODE)
            RELMaster.SubItems(1) = REL2!RELDescription
            REL2.MoveNext
        TotalRELCount = TotalRELCount + 1
        txtTotalRecord = TotalRELCount
        Loop
    REL2.Close
    Set REL2 = Nothing
End If
End Sub

Private Sub lstRelationshipList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstRelationshipList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT RELcode, RELdescription " & _
    "from REL Where RELcode = '" & lstRelationshipList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    txtRELcode = rst3!RELCODE
    txtRELdesc = rst3!RELDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If

End Sub

'Private Sub PrintRelationship()
    'RptRelList.Show
'End Sub
'*******************************End Relationship *******************************

'*******************************Start Nationality *******************************
Private Sub SaveNationality()
On Error Resume Next
Dim RecordSaved
Dim NAT As New adodb.Recordset
Dim SaveNAT As New adodb.Recordset
Dim strsql As String
   
      
    If txtNATcode = "" Then
        MsgBox "Please Enter Nationality Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtNATcode.SetFocus
    ElseIf txtNATname = "" Then
        MsgBox "Please Enter Nationality Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtNATname.SetFocus
    
    Else
          strsql = "Select * From NAT Where NATCode = '" & Trim(txtNATcode) & "'"
          NAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             If NAT.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                NAT.Close
                txtNATcode.SetFocus
                txtNATcode.SelStart = 0
                txtNATcode.SelLength = Len(txtNATcode)
                Set NAT = Nothing
                Exit Sub
            Else
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
         If RecordSaved = vbYes Then
            If EditFlag = False Then
                NAT.Open "NAT", medixmasterconn, adOpenKeyset, adLockOptimistic
                NAT.AddNew
            Else
                strsql = "Select * From NAT Where NATCode = '" & Trim(txtNATcode) & "'"
                NAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With NAT
                    !NATCode = ChkStr(txtNATcode)
                    !NATName = ChkStr(txtNATname)
                    !NATLastUpdateUser = Logon
                    !NATLastUpdateDate = Date
                End With
                NAT.Update
                Call ClearForm(Me)
                Call NATListing(NATCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        NAT.Close
        Set SaveNAT = Nothing
        End If
        EditFlag = False
    End If
End If

End Sub

Private Sub DeleteNationality()
Dim DeleteNAT As New adodb.Recordset
Dim rsCheckMBM As New adodb.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If txtNATcode = "" Then
   MsgBox "Please Enter Nationality Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtNATcode.SetFocus
ElseIf lstNationalityList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where NATCode = '" & RemStr(txtNATcode) & "'"
         
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From NAT Where NATcode = '" & Trim(txtNATcode) & "'" 'ChkStr(lstNationalityList.SelectedItem.Text) & "'"
           DeleteNAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call NATListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtNATcode.SetFocus
         End If
     End If
End If
End Sub

Private Sub SearchNationality()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(txtNATcode)) Then
strAppend = strAppend + " And NATCode like '%" & Mid(ChkStr(txtNATcode), 1, 2) & "%'"
End If

If Len(Trim(txtNATname)) Then
strAppend = strAppend + " And NATName like '%" & txtNATname & "%'"
End If

NATCriteria = strAppend
Call NATListing(strAppend)
Screen.MousePointer = Default
End Sub

Private Sub NATListing(Optional strAppend As String)
Dim NAT3 As New adodb.Recordset
Dim NATMaster As ListItem
Dim Sql As String
Dim TotalNATCount As String

TotalNATCount = 0
lstNationalityList.listitems.Clear


    Sql = "SELECT NATCode, " & _
    " NATName" & _
    " from NAT where 0=0 "
    
    If Len(Trim(strAppend)) Then
    Sql = Sql + strAppend
    End If
    
    NAT3.Open Sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If NAT3.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until NAT3.EOF
        Set NATMaster = lstNationalityList.listitems.Add(, , NAT3!NATCode)
            NATMaster.SubItems(1) = NAT3!NATName
            NAT3.MoveNext
        TotalNATCount = TotalNATCount + 1
        txtTotalRecord = TotalNATCount
        Loop
    NAT3.Close
    Set NAT3 = Nothing
End If
End Sub

Private Sub lstNationalityList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstNationalityList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT NATCode, NATName " & _
    "from NAT Where NATCode = '" & lstNationalityList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    txtNATcode = rst3!NATCode
    txtNATname = rst3!NATName
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '
End If
End Sub

'Private Sub PrintNationality()
    'RptNatList.Show
'End Sub
'*******************************End Nationality *******************************

'*******************************Start Race *******************************
Private Sub SaveRace()
On Error Resume Next
Dim RecordSaved
Dim RAC As New adodb.Recordset
Dim MBM As New adodb.Recordset
Dim SaveRAC As New adodb.Recordset
Dim strsql As String
      
    If txtRACcode = "" Then
        MsgBox "Please Enter Race Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtRACcode.SetFocus
    ElseIf txtRACDesc = "" Then
        MsgBox "Please Enter Race Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtRACDesc.SetFocus
    
    Else
          strsql = "Select * From RAC Where RACCode = '" & Trim(txtRACcode) & "'"
          
          RAC.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
          
            If RAC.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                RAC.Close
                txtRACcode.SetFocus
                txtRACcode.SelStart = 0
                txtRACcode.SelLength = Len(txtRACcode)
                Set RAC = Nothing
                Exit Sub
            Else
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
         If RecordSaved = vbYes Then
            If EditFlag = False Then
                RAC.Open "Rac", medixmasterconn, adOpenKeyset, adLockOptimistic
                RAC.AddNew
            Else
                strsql = "Select * From RAC Where RACCode = '" & Trim(txtRACcode) & "'"
                RAC.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With RAC
                    !RACCode = ChkStr(txtRACcode)
                    !RACName = ChkStr(txtRACDesc)
                    !RACLastUpdateUser = Logon
                    !RACLastUpdateDate = Date
                End With
                RAC.Update
                Call ClearForm(Me)
                Call RACListing(RACCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        RAC.Close
        Set SaveRAC = Nothing
        End If
        EditFlag = False
    End If
End If

End Sub

Private Sub DeleteRace()
Dim DeleteRAC As New adodb.Recordset
Dim rsCheckMBM As New adodb.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If txtRACcode = "" Then
   MsgBox "Please Enter Race Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtRACcode.SetFocus
ElseIf lstRaceList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where RACCode = '" & RemStr(txtRACcode) & "'"
      
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From RAC Where RACcode = '" & Trim(txtRACcode) & "'" 'ChkStr(lstRaceList.SelectedItem.Text) & "'"
           DeleteRAC.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call RACListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtRACcode.SetFocus
         End If
     End If
End If
End Sub

Private Sub SearchRace()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(txtRACcode)) Then
strAppend = strAppend + " And RACCode like '%" & Mid(ChkStr(txtRACcode), 1, 2) & "%'"
End If

If Len(Trim(txtRACDesc)) Then
strAppend = strAppend + " And RACName like '%" & ChkStr(txtRACDesc) & "%'"
End If

RACCriteria = strAppend
Call RACListing(strAppend)
Screen.MousePointer = Default
End Sub

Private Sub RACListing(Optional strAppend As String)
Dim RAC3 As New adodb.Recordset
Dim RACMaster As ListItem
Dim Sql As String
Dim TotalRACCount As String

TotalRACCount = 0
lstRaceList.listitems.Clear


    Sql = "SELECT RACcode, " & _
    " RACName" & _
    " from RAC where 0=0 "
    
    If Len(Trim(strAppend)) Then
    Sql = Sql + strAppend
    End If
    
    RAC3.Open Sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If RAC3.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until RAC3.EOF
        Set RACMaster = lstRaceList.listitems.Add(, , RAC3!RACCode)
            RACMaster.SubItems(1) = RAC3!RACName
            RAC3.MoveNext
        TotalRACCount = TotalRACCount + 1
        txtTotalRecord = TotalRACCount
        Loop
    RAC3.Close
    Set RAC3 = Nothing
End If
End Sub

Private Sub lstRaceList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstRaceList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT RACcode, RACName " & _
    "from RAC Where RACcode = '" & lstRaceList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    txtRACcode = rst3!RACCode
    txtRACDesc = rst3!RACName
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If
End Sub

'Private Sub PrintRace()
    'RptRaceList.Show
'End Sub
'*******************************End Race *******************************

'*******************************Start Language *******************************
Private Sub SaveLanguage()
On Error Resume Next
Dim RecordSaved
Dim LGN As New adodb.Recordset
Dim SaveLGN As New adodb.Recordset
Dim strsql As String

      
 If txtLGNcode = "" Then
    MsgBox "Please Enter Language Code", vbOKOnly + vbCritical, DialogBoxTitle
    txtLGNcode.SetFocus
    ElseIf txtLGNdesc = "" Then
           MsgBox "Please Enter Language Description", vbOKOnly + vbCritical, DialogBoxTitle
           txtLGNdesc.SetFocus
    Else
       strsql = "Select * From LGN Where LGNCode = '" & Trim(txtLGNcode) & "'"
       LGN.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic

       If LGN.recordCount > 0 And EditFlag = False Then
          MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
          LGN.Close
          txtLGNcode.SetFocus
          txtLGNcode.SelStart = 0
          txtLGNcode.SelLength = Len(txtLGNcode)
          Set LGN = Nothing
          Exit Sub
       Else
          RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
          If RecordSaved = vbYes Then
             If EditFlag = False Then
                LGN.Open "LGN", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                LGN.AddNew
             Else
                strsql = "Select * From LGN Where LGNCode = '" & Trim(txtLGNcode) & "'"
                LGN.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
             End If
            
             With LGN
                !LGNCode = ChkStr(txtLGNcode)
                !LGNDescription = ChkStr(txtLGNdesc)
                !LGNLastUpdateUser = Logon
                !LGNLastUpdateDate = Date
             End With
             LGN.Update
             Call ClearForm(Me)
             Call LGNListing(LGNCriteria)
               If EditFlag = False Then
                  MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
               Else
                  MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
               End If
               LGN.Close
               Set SaveLGN = Nothing
          End If
          EditFlag = False
    End If
 End If

End Sub

Private Sub DeleteLanguage()
Dim DeleteLGN As New adodb.Recordset
Dim rsCheckMBM As New adodb.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If txtLGNcode = "" Then
   MsgBox "Please Enter Language Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtLGNcode.SetFocus
ElseIf lstLanguageList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where LGNCode = '" & RemStr(txtLGNcode) & "'" 'ChkStr(lstLanguageList.SelectedItem.Text) & "'"
      
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From LGN Where LGNcode = '" & Trim(txtLGNcode) & "'" 'ChkStr(lstLanguageList.SelectedItem.Text) & "'"
           DeleteLGN.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call LGNListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtLGNcode.SetFocus
         End If
     End If
End If

End Sub

Private Sub SearchLanguage()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(txtLGNcode)) Then
strAppend = strAppend + " And LGNCode like '%" & ChkStr(txtLGNcode) & "%'"
End If

If Len(Trim(txtLGNdesc)) Then
strAppend = strAppend + " And LGNDescription like '%" & ChkStr(txtLGNdesc) & "%'"
End If

LGNCriteria = strAppend
Call LGNListing(strAppend)
Screen.MousePointer = Default
End Sub

Private Sub LGNListing(Optional strAppend As String)
Dim LGN3 As New adodb.Recordset
Dim LGNMaster As ListItem
Dim Sql As String
Dim TotalLGNCount As String

TotalLGNCount = 0
lstLanguageList.listitems.Clear


    Sql = "SELECT LGNcode, " & _
    " LGNDescription" & _
    " from LGN where 0=0 "
    
    If Len(Trim(strAppend)) Then
    Sql = Sql + strAppend
    End If
    
    LGN3.Open Sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    
    If LGN3.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until LGN3.EOF
        Set LGNMaster = lstLanguageList.listitems.Add(, , LGN3!LGNCode)
            LGNMaster.SubItems(1) = LGN3!LGNDescription
            LGN3.MoveNext
        TotalLGNCount = TotalLGNCount + 1
        txtTotalRecord = TotalLGNCount
        Loop
    LGN3.Close
    Set LGN3 = Nothing
End If
End Sub

Private Sub lstLanguageList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstLanguageList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    strsql = "SELECT LGNcode, LGNDescription " & _
    "from LGN Where LGNcode = '" & lstLanguageList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    txtLGNcode = rst3!LGNCode
    txtLGNdesc = rst3!LGNDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End If
End Sub

'Private Sub PrintLanguage()
    'RptLgnList.Show
'End Sub
'*******************************End Language *******************************

'*******************************Start City *******************************
Private Sub SaveCity()
On Error Resume Next
Dim RecordSaved
Dim CTY As New adodb.Recordset
Dim STA As New adodb.Recordset
Dim SaveCTY As New adodb.Recordset
Dim strsql As String
Dim strsql2 As String
Dim STANAME As String

    
 If txtCTYcode = "" Then
    MsgBox "Please Enter City Code", vbOKOnly + vbCritical, DialogBoxTitle
    txtCTYcode.SetFocus
 ElseIf CboStateCode = "" Then
    MsgBox "Please Enter State Description", vbOKOnly + vbCritical, DialogBoxTitle
    CboStateCode.SetFocus
 ElseIf txtCTYdesc = "" Then
    MsgBox "Please Enter City Description", vbOKOnly + vbCritical, DialogBoxTitle
    txtCTYdesc.SetFocus
      Else
        strsql = "Select * From CTY Where CTYCode = '" & Trim(txtCTYcode) & "'"
        CTY.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
         If CTY.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            CTY.Close
            txtCTYcode.SetFocus
            txtCTYcode.SelStart = 0
            txtCTYcode.SelLength = Len(txtCTYcode)
            Set CTY = Nothing
            Exit Sub

         Else
         
        strsql2 = "Select STACode, STADescription From STA Where STACode = '" & Mid(CboStateCode, 1, 2) & "'"
        STA.Open strsql2, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        STANAME = STA!STADescription
        
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 CTY.Open "CTY", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                 CTY.AddNew
              Else
                 strsql = "Select * From CTY Where CTYCode = '" & Trim(txtCTYcode) & "'"
                 CTY.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
              End If
                    
              With CTY
                 !CTYCode = ChkStr(txtCTYcode)
                 !CTYDescription = ChkStr(txtCTYdesc)
                 !STADescription = ChkStr(STANAME)
                 !CTYLastUpdateUser = Logon
                 !CTYLastUpdateDate = Date
              End With
              CTY.Update
              Call ClearForm(Me)
              Call CTYListing(CTYCriteria)
                If EditFlag = False Then
                   MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                   MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                CTY.Close
                Set SaveCTY = Nothing
           End If
                 
           EditFlag = False
      End If
End If
End Sub

Private Sub DeleteCity()
Dim DeleteCTY As New adodb.Recordset
Dim rsCheckMBM As New adodb.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If txtCTYcode = "" Then
   MsgBox "Please Enter City Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtCTYcode.SetFocus
ElseIf lstCityList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where CTYCode = '" & RemStr(txtCTYcode) & "'"
          
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
           
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From CTY Where CTYcode = '" & Trim(txtCTYcode) & "'"
           DeleteCTY.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call CTYListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtCTYcode.SetFocus
         End If
     End If
End If

End Sub

Private Sub SearchCity()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(txtCTYcode)) Then
   strAppend = strAppend + " And CTYCode like '%" & Mid(ChkStr(txtCTYcode), 1, 2) & "%'"
End If

If Len(Trim(txtCTYdesc)) Then
   strAppend = strAppend + " And CTYDescription like '%" & ChkStr(txtCTYdesc) & "%'"
End If

If Len(Trim(CboStateCode)) Then
    strAppend = strAppend + " AND STADescription like '" & RemStr(Mid(CboStateCode, 6, IIf(InStr(6, CboStateCode, "-") = 0, 4, InStr(6, CboStateCode, "-") - 6))) & "%'"
   'strAppend = strAppend + " And STADescription like '%" & Mid(ChkStr(CboStateCode), 6, 4) & "%'"
End If

CTYCriteria = strAppend
Call CTYListing(strAppend)
Screen.MousePointer = Default

End Sub

Private Sub CTYListing(Optional strAppend As String)
Dim CTY3 As New adodb.Recordset
Dim CTYMaster As ListItem
Dim Sql As String
Dim TotalCTYCount As String

TotalCTYCount = 0
lstCityList.listitems.Clear


    Sql = "SELECT CTYcode, " & _
    " CTYDescription, STADescription" & _
    " from CTY where 0=0 "
    
    If Len(Trim(strAppend)) Then
    Sql = Sql + strAppend
    End If
    
    CTY3.Open Sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    
    If CTY3.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until CTY3.EOF
       Set CTYMaster = lstCityList.listitems.Add(, , CTY3!CTYCode)
           CTYMaster.SubItems(1) = CTY3!CTYDescription
           If Len(Trim(CTY3!STADescription)) Then
              CTYMaster.SubItems(2) = CTY3!STADescription
           End If
           CTY3.MoveNext
        TotalCTYCount = TotalCTYCount + 1
        txtTotalRecord = TotalCTYCount
       Loop
    CTY3.Close
    Set CTY3 = Nothing
End If
End Sub

Private Sub lstCityList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String
    
        
    If lstCityList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "SELECT CTYcode, CTYDescription, STADescription " & _
        "from CTY Where CTYcode = '" & lstCityList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        txtCTYcode = rst3!CTYCode
        txtCTYdesc = rst3!CTYDescription
        CboStateCode = rst3!STADescription
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
        'medixmasterconnMaint.Close
        'Set medixmasterconnMaint = Nothing
        
    End If
End Sub

Private Sub CTYComboListing()

    Dim STA As New adodb.Recordset
        
    STA.Open "STA", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
    CboStateCode.Clear
        
    Do Until STA.EOF
        CboStateCode.AddItem STA!STACode & " - " & STA!STADescription
        STA.MoveNext
    Loop
        
    STA.Close
    
End Sub

'Private Sub PrintCity()
   'RptCityList.Show
'End Sub
'*******************************End City *******************************

'*******************************Start State *******************************
Private Sub SaveState()
On Error Resume Next
Dim RecordSaved
Dim STA As New adodb.Recordset
Dim SaveSTA As New adodb.Recordset
Dim strsql As String
    
If txtSTAcode = "" Then
   MsgBox "Please Enter State Code", vbOKOnly + vbCritical, DialogBoxTitle
   txtSTAcode.SetFocus
  ElseIf txtSTAdesc = "" Then
      MsgBox "Please Enter State Description", vbOKOnly + vbCritical, DialogBoxTitle
      txtSTAdesc.SetFocus
    Else
       strsql = "Select * From STA Where STACode = '" & Trim(txtSTAcode) & "'"
       STA.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

          If STA.recordCount > 0 And EditFlag = False Then
             MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
             STA.Close
             txtSTAcode.SetFocus
             txtSTAcode.SelStart = 0
             txtSTAcode.SelLength = Len(txtSTAcode)
             Set STA = Nothing
             Exit Sub
          Else
             RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
             If RecordSaved = vbYes Then
                If EditFlag = False Then
                   STA.Open "STA", medixmasterconn, adOpenKeyset, adLockOptimistic
                   STA.AddNew
                Else
                   strsql = "Select * From STA Where STACode = '" & Trim(txtSTAcode) & "'"
                   STA.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                End If
            
                With STA
                  !STACode = ChkStr(txtSTAcode)
                  !STADescription = ChkStr(txtSTAdesc)
                  !STALastUpdateUser = Logon
                  !STALastUpdateDate = Date
                End With
                STA.Update
                Call ClearForm(Me)
                Call STAListing(STACriteria)
                
                If EditFlag = False Then
                   MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                STA.Close
                Set SaveSTA = Nothing
              End If
              EditFlag = False
    End If
End If

End Sub

Private Sub DeleteState()
Dim DeleteSTA As New adodb.Recordset
Dim rsCheckMBM As New adodb.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If txtSTAcode = "" Then
   MsgBox "Please Enter State Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtSTAcode.SetFocus
ElseIf lstStateList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where STACode = '" & RemStr(txtSTAcode) & "'"
       
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From STA Where STAcode = '" & Trim(txtSTAcode) & "'" 'ChkStr(lstStateList.SelectedItem.Text) & "' "
           DeleteSTA.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call STAListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtSTAcode.SetFocus
         End If
     End If
End If

End Sub

Private Sub SearchState()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(txtSTAcode)) Then
strAppend = strAppend + " And STACode like '%" & Mid(ChkStr(txtSTAcode), 1, 2) & "%'"
End If

If Len(Trim(txtSTAdesc)) Then
strAppend = strAppend + " And STAdescription like '%" & ChkStr(txtSTAdesc) & "%'"
End If

STACriteria = strAppend
Call STAListing(strAppend)
Screen.MousePointer = Default

End Sub

Private Sub STAListing(Optional strAppend As String)
Dim STA3 As New adodb.Recordset
Dim STAMaster As ListItem
Dim Sql As String
Dim TotalSTACount As String

TotalSTACount = 0
lstStateList.listitems.Clear


    Sql = "SELECT STAcode, " & _
    " STADescription" & _
    " from STA where 0=0 "
    
    If Len(Trim(strAppend)) Then
    Sql = Sql + strAppend
    End If
    
    STA3.Open Sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If STA3.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until STA3.EOF
        Set STAMaster = lstStateList.listitems.Add(, , STA3!STACode)
            STAMaster.SubItems(1) = STA3!STADescription
            STA3.MoveNext
        TotalSTACount = TotalSTACount + 1
        txtTotalRecord = TotalSTACount
        Loop
    STA3.Close
    Set STA3 = Nothing
End If
End Sub

Private Sub lstStateList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String
Dim color As ColorConstants

color = vbRed

    
If lstStateList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT STAcode, STADescription " & _
    "from STA Where STAcode = '" & lstStateList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
    txtSTAcode = rst3!STACode
    txtSTAdesc = rst3!STADescription
        
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If

End Sub

'Private Sub PrintState()
    'RptStateList.Show
'End Sub

'**********************************Start WorkNature**************************************
Public Sub SaveWork()
On Error Resume Next
Dim RecordSaved
Dim Work As New adodb.Recordset
Dim SaveWork As New adodb.Recordset
Dim strsql As String
    
    
    If TxtWorkCode = "" Then
        MsgBox "Please Enter Worknature Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtWorkCode.SetFocus
    ElseIf TxtWorkDescription = "" Then
        MsgBox "Please Enter Worknature Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtWorkDescription.SetFocus
    Else
        strsql = "Select * From WorkNature Where WorkID = '" & Trim(TxtWorkCode) & "'"
        Work.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            If Work.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            Work.Close
            TxtWorkCode.SetFocus
            TxtWorkCode.SelStart = 0
            TxtWorkCode.SelLength = Len(TxtWorkCode)
            Set Work = Nothing
            Exit Sub
            Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
            If EditFlag = False Then
      
             SaveWork.Open "WorkNature", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
             SaveWork.AddNew
             Else
                strsql = "Select * From WorkNature Where WorkID = '" & Trim(TxtWorkCode) & "'"
                SaveWork.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            End If
            
                With SaveWork
                    !WorkID = ChkStr(TxtWorkCode)
                    !WorkDescription = ChkStr(TxtWorkDescription)
                    !WorkLastUpdateUser = Logon
                    !WorkLastUpdateTime = Date
                End With
                SaveWork.Update
                Call ClearForm(Me)
                Call WorkListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
      
            SaveWork.Close
            Set SaveWork = Nothing
        End If
        EditFlag = False
    End If
End If
End Sub

Private Sub DeleteWork()
Dim rsDeleteWork As New adodb.Recordset
Dim rsCheckWork As New adodb.Recordset
Dim strsql, sqlWork As String
Dim DeleteRecord
EditFlag = False

If TxtWorkCode = "" Then
   MsgBox "Please Enter Worknature Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtWorkCode.SetFocus
ElseIf lstWorkList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else

sqlWork = "Select MBMNumber from MBMOthersTwo where MBMWorkNature = '" & RemStr(TxtWorkCode) & "'"

rsCheckWork.MaxRecords = 1
rsCheckWork.Open sqlWork, medixmasterconn, adOpenKeyset, adLockOptimistic

If rsCheckWork.recordCount Then
    MsgBox "This record cannot be deleted because it is used in MBMOthers Table!", vbOKOnly + vbCritical, "Error Message"
Else
    DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes And rsCheckWork.recordCount = 0 Then
        strsql = "Delete from WorkNature where WorkID = '" & ChkStr(lstWorkList.SelectedItem.Text) & "'"
        rsDeleteWork.Open strsql, medixmasterconnMaint, adOpenDynamic, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        Call ClearForm(Me)
        Call WorkListing(WorkCriteria)
        Call setCmdButton(Me, False)
        Call EntriesEnable(True)
        TxtWorkCode.SetFocus
        CmdAdd.Enabled = True
    End If
  End If
End If
Exit Sub

End Sub

Private Sub SearchWork()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(TxtWorkCode)) Then
        strAppend = strAppend + " And WorkID like '%" & ChkStr(TxtWorkCode) & "%'"
    End If
    If Len(Trim(TxtWorkDescription)) Then
        strAppend = strAppend + " And WorkDescription like '%" & ChkStr(TxtWorkDescription) & "%'"
    End If
    
    WorkCriteria = strAppend
    Call WorkListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub WorkListing(Optional strAppend As String)
    Dim rst2 As New adodb.Recordset
    Dim WorkMaster As ListItem
    Dim Sql As String
    Dim TotalWorkCount As String
     
    TotalWorkCount = 0
    lstWorkList.listitems.Clear
        Sql = " SELECT WorkID, WorkDescription " & _
        " from WorkNature where 0=0 "
    
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If

    rst2.Open Sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set WorkMaster = lstWorkList.listitems.Add(, , rst2!WorkID)
                WorkMaster.SubItems(1) = rst2!WorkDescription
                rst2.MoveNext
        TotalWorkCount = TotalWorkCount + 1
        txtTotalRecord = TotalWorkCount
        Loop
    rst2.Close
    End If
End Sub

Private Sub lstWorkList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lstWorkList.SortKey = ColumnHeader.Index - 1
lstWorkList.Sorted = True
If lstWorkList.SortOrder = lvwAscending Then
lstWorkList.SortOrder = lvwDescending
Else
lstWorkList.SortOrder = lvwAscending
End If
End Sub

Private Sub lstWorkList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String
    
    If lstWorkList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "Select WorkID, WorkDescription " & _
        "from WorkNature Where WorkID = '" & lstWorkList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        TxtWorkCode = Trim(rst3!WorkID)
        TxtWorkDescription = Trim(rst3!WorkDescription)
        
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconnMaint.Close
       ' Set medixmasterconnMaint = Nothing
        
    End If
End Sub

'Private Sub PrintWork()
    'RptWorkList.Show
'End Sub
'**********************************End WorkNature**************************************

'*******************************Start Status *******************************
Private Sub SaveStatus()
On Error Resume Next
Dim RecordSaved
Dim status As New adodb.Recordset
Dim SaveStatus As New adodb.Recordset
Dim strsql As String

      
 If TxtStatusCode = "" Then
    MsgBox "Please Enter Status Code", vbOKOnly + vbCritical, DialogBoxTitle
    TxtStatusCode.SetFocus
    ElseIf TxtStatusDesc = "" Then
           MsgBox "Please Enter Status Description", vbOKOnly + vbCritical, DialogBoxTitle
           TxtStatusDesc.SetFocus
    Else
       strsql = "Select * From Status Where StatusCode = '" & Trim(TxtStatusCode) & "'"
       status.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

       If status.recordCount > 0 And EditFlag = False Then
          MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
          status.Close
          TxtStatusCode.SetFocus
          TxtStatusCode.SelStart = 0
          TxtStatusCode.SelLength = Len(TxtStatusCode)
          Set status = Nothing
          Exit Sub
       Else
          RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
          If RecordSaved = vbYes Then
             If EditFlag = False Then
                status.Open "Status", medixmasterconn, adOpenKeyset, adLockOptimistic
                status.AddNew
             Else
                strsql = "Select * From Status Where StatusCode = '" & Trim(TxtStatusCode) & "'"
                status.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             End If
            
             With status
                !StatusCode = ChkStr(TxtStatusCode)
                !StatusDesc = ChkStr(TxtStatusDesc)
                !StatusLastUpdateUser = Logon
                !StatusLastUpdateDate = Date
             End With
             status.Update
             Call ClearForm(Me)
             Call StatusListing(StatusCriteria)
               If EditFlag = False Then
                  MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
               Else
                  MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
               End If
               status.Close
               Set SaveStatus = Nothing
          End If
          EditFlag = False
    End If
 End If

End Sub

Private Sub DeleteStatus()
Dim DeleteStatus As New adodb.Recordset
Dim rsCheckMBM As New adodb.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If TxtStatusCode = "" Then
   MsgBox "Please Enter Status Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtStatusCode.SetFocus
ElseIf lstStatusList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where MBMLastEndorseStatus = '" & RemStr(TxtStatusCode) & "'" 'ChkStr(lstStatusList.SelectedItem.Text) & "'"
      
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From Status Where Statuscode = '" & Trim(TxtStatusCode) & "'" 'ChkStr(lstStatusList.SelectedItem.Text) & "'"
           DeleteStatus.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call StatusListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           TxtStatusCode.SetFocus
         End If
     End If
End If

End Sub

Private Sub SearchStatus()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(TxtStatusCode)) Then
strAppend = strAppend + " And StatusCode like '%" & ChkStr(TxtStatusCode) & "%'"
End If

StatusCriteria = strAppend
Call StatusListing(strAppend)
Screen.MousePointer = Default
End Sub

Private Sub StatusListing(Optional strAppend As String)
Dim status As New adodb.Recordset
Dim StatusMaster As ListItem
Dim Sql As String
Dim TotalStatusCount As String

TotalStatusCount = 0
lstStatusList.listitems.Clear


    Sql = "SELECT Statuscode, " & _
    " StatusDesc" & _
    " from Status where 0=0 "
    
    If Len(Trim(strAppend)) Then
    Sql = Sql + strAppend
    End If
    
    status.Open Sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If status.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until status.EOF
        Set StatusMaster = lstStatusList.listitems.Add(, , status!StatusCode)
            StatusMaster.SubItems(1) = status!StatusDesc
            status.MoveNext
        TotalStatusCount = TotalStatusCount + 1
        txtTotalRecord = TotalStatusCount
        Loop
    status.Close
    Set status = Nothing
End If
End Sub

Private Sub lstStatusList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String

    
If lstStatusList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT Statuscode, StatusDesc " & _
    "from Status Where Statuscode = '" & lstStatusList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TxtStatusCode = rst3!StatusCode
    TxtStatusDesc = rst3!StatusDesc
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End If
End Sub

Private Sub lstStatusList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstStatusList.SortKey = ColumnHeader.Index - 1
    lstStatusList.Sorted = True
    If lstStatusList.SortOrder = lvwAscending Then
        lstStatusList.SortOrder = lvwDescending
    Else
        lstStatusList.SortOrder = lvwAscending
    End If
End Sub

'Private Sub PrintStatus()
    'RptStatusList.Show
'End Sub
'*******************************End Status *******************************


'*******************************Start Status *******************************
Private Sub SavePStatus()
On Error Resume Next
Dim RecordSaved
Dim PStatus As New adodb.Recordset
Dim SavePStatus As New adodb.Recordset
Dim strsql As String

      
 If TxtPStatusCode = "" Then
    MsgBox "Please Enter Policy Status Code", vbOKOnly + vbCritical, DialogBoxTitle
    TxtPStatusCode.SetFocus
    ElseIf TxtPStatusDesc = "" Then
           MsgBox "Please Enter Policy Status Description", vbOKOnly + vbCritical, DialogBoxTitle
           TxtPStatusDesc.SetFocus
    Else
       strsql = "Select * From PolicyStatus Where PStatusCode = '" & Trim(TxtPStatusCode) & "'"
       PStatus.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

       If PStatus.recordCount > 0 And EditFlag = False Then
          MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
          PStatus.Close
          TxtPStatusCode.SetFocus
          TxtPStatusCode.SelStart = 0
          TxtPStatusCode.SelLength = Len(TxtPStatusCode)
          Set PStatus = Nothing
          Exit Sub
       Else
          RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
          If RecordSaved = vbYes Then
             If EditFlag = False Then
                PStatus.Open "PolicyStatus", medixmasterconn, adOpenKeyset, adLockOptimistic
                PStatus.AddNew
             Else
                strsql = "Select * From PolicyStatus Where PStatusCode = '" & Trim(TxtPStatusCode) & "'"
                PStatus.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             End If
            
             With PStatus
                !pstatuscode = ChkStr(TxtPStatusCode)
                !pstatusDesc = ChkStr(TxtPStatusDesc)
                !PStatusLastUpdateUser = Logon
                !PStatusLastUpdateDate = Date
             End With
             PStatus.Update
             Call ClearForm(Me)
             Call PStatusListing(StatusCriteria)
               If EditFlag = False Then
                  MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
               Else
                  MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
               End If
               PStatus.Close
               Set SavePStatus = Nothing
          End If
          EditFlag = False
    End If
 End If

End Sub

Private Sub DeletePStatus()
Dim DeletePStatus As New adodb.Recordset
Dim rsCheckMBM As New adodb.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

EditFlag = False
If TxtPStatusCode = "" Then
   MsgBox "Please Enter Policy Status Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtPStatusCode.SetFocus
ElseIf lstPolicyList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlMBM = "Select MBMNumber from MBM where MBMStatus = '" & RemStr(TxtPStatusCode) & "'" 'ChkStr(lstStatusList.SelectedItem.Text) & "'"
      
   rsCheckMBM.MaxRecords = 1
   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
   If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
           strsql = "Delete From PolicyStatus Where PStatuscode = '" & Trim(TxtPStatusCode) & "'" 'ChkStr(lstpolicyList.SelectedItem.Text) & "'"
           DeletePStatus.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call PStatusListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           TxtPStatusCode.SetFocus
         End If
     End If
End If

End Sub

Private Sub SearchPStatus()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(TxtPStatusCode)) Then
    strAppend = strAppend + " And PStatusCode like '%" & ChkStr(TxtPStatusCode) & "%'"
End If

PolicyStatusCriteria = strAppend
Call PStatusListing(strAppend)
Screen.MousePointer = Default
End Sub

Private Sub PStatusListing(Optional strAppend As String)
Dim PStatus As New adodb.Recordset
Dim PStatusMaster As ListItem
Dim Sql As String
Dim TotalPStatusCount As String

TotalPStatusCount = 0
lstPolicyList.listitems.Clear


    Sql = "SELECT PStatusCode, " & _
    " PStatusDesc" & _
    " from PolicyStatus where 0=0 "
    
    If Len(Trim(strAppend)) Then
    Sql = Sql + strAppend
    End If
    
    PStatus.Open Sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If PStatus.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until PStatus.EOF
        Set PStatusMaster = lstPolicyList.listitems.Add(, , PStatus!pstatuscode)
            PStatusMaster.SubItems(1) = PStatus!pstatusDesc
            PStatus.MoveNext
        TotalPStatusCount = TotalPStatusCount + 1
        txtTotalRecord = TotalPStatusCount
        Loop
    PStatus.Close
    Set PStatus = Nothing
End If
End Sub

Private Sub lstPolicyList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String

        
    If lstPolicyList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        strsql = "SELECT PStatuscode, PStatusDesc " & _
        "from PolicyStatus Where PStatuscode = '" & lstPolicyList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        TxtPStatusCode = rst3!pstatuscode
        TxtPStatusDesc = rst3!pstatusDesc
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
    End If

End Sub

Private Sub lstPolicyList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstPolicyList.SortKey = ColumnHeader.Index - 1
    lstPolicyList.Sorted = True
    If lstPolicyList.SortOrder = lvwAscending Then
        lstPolicyList.SortOrder = lvwDescending
    Else
        lstPolicyList.SortOrder = lvwAscending
    End If
End Sub

'Private Sub PrintPStatus()
    'RptPStatusList.Show
'End Sub
'*******************************End Policy Status *******************************


Private Sub lstCPTList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCPTList.SortKey = ColumnHeader.Index - 1
    lstCPTList.Sorted = True
    If lstCPTList.SortOrder = lvwAscending Then
        lstCPTList.SortOrder = lvwDescending
    Else
        lstCPTList.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCPTList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String

        
    If lstCPTList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "SELECT CPTCode, CPTDesc " & _
        "From CPT Where CPTCode = '" & lstCPTList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        TxtCPT = rst3!CPTCode
        TxtCPTDesc = rst3!CPTDesc
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
      '  medixmasterconnMaint.Close
      '  Set medixmasterconnMaint = Nothing
        
    End If

End Sub

Private Sub lstCPTList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

        
    If lstCPTList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "SELECT CPTCode, CPTDesc " & _
        "From CPT Where CPTCode = '" & lstCPTList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        TxtCPT = rst3!CPTCode
        TxtCPTDesc = rst3!CPTDesc
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
       ' medixmasterconnMaint.Close
       ' Set medixmasterconnMaint = Nothing
        
    End If
End Sub

Private Sub lstCPTList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

        
    If lstCPTList.listitems.Count Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
            
    '    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "SELECT CPTCode, CPTDesc " & _
        "From CPT Where CPTCode = '" & lstCPTList.SelectedItem.Text & "'"
        rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        TxtCPT = rst3!CPTCode
        TxtCPTDesc = rst3!CPTDesc
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        rst3.Close
        Set rst3 = Nothing
        
     '   medixmasterconnMaint.Close
   '     Set medixmasterconnMaint = Nothing
        
    End If
End Sub

Private Sub SaveCPT()
Dim RecordSaved
Dim CPT As New adodb.Recordset
Dim strsql As String

      
 If TxtCPT = "" Then
    MsgBox "Please Enter CPT Code", vbOKOnly + vbCritical, DialogBoxTitle
    TxtCPT.SetFocus
    ElseIf TxtCPTDesc = "" Then
           MsgBox "Please Enter CPT Description", vbOKOnly + vbCritical, DialogBoxTitle
           TxtCPTDesc.SetFocus
    Else
       strsql = "Select * From CPT Where CPTCode = '" & Trim(TxtCPT) & "'"
       CPT.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic

       If CPT.recordCount > 0 And EditFlag = False Then
          MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
          CPT.Close
          TxtCPT.SetFocus
          TxtCPT.SelStart = 0
          TxtCPT.SelLength = Len(TxtCPT)
          Set CPT = Nothing
          Exit Sub
       Else
          RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
          If RecordSaved = vbYes Then
             If EditFlag = False Then
                CPT.Open "CPT", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                CPT.AddNew
             Else
                strsql = "Select * From CPT Where CPTCode = '" & Trim(TxtCPT) & "'"
                CPT.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
             End If
            
             With CPT
                !CPTCode = ChkStr(TxtCPT)
                !CPTDesc = ChkStr(TxtCPTDesc)
                !CPTLastUpdateUser = Logon
                !CPTLastUpdateDate = Date
             End With
             CPT.Update
             Call ClearForm(Me)
             Call CPTListing(CPTCriteria)
               If EditFlag = False Then
                  MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
               Else
                  MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
               End If
               CPT.Close
               Set CPT = Nothing
          End If
          EditFlag = False
    End If
 End If

End Sub

Private Sub DeleteCPT()
Dim DeleteCPT As New adodb.Recordset
Dim strsql As String
Dim DeleteRecord

EditFlag = False
    If TxtCPT = "" Then
       MsgBox "Please Enter CPT Code", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf lstCPTList.listitems.Count = 0 Then
       MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
       DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
         If DeleteRecord = vbYes Then
               strsql = "Delete From CPT Where CPTCode = '" & Trim(TxtCPT) & "'"
               DeleteCPT.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
               MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
               CmdSave.Enabled = True
               Call ClearForm(Me)
               Call CPTListing
               Call DisableCmdButton(Me)
               Call EntriesEnable(True)
               TxtCPT.SetFocus
            End If
    End If
End Sub

Private Sub SearchCPT()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(TxtCPT)) Then
    strAppend = strAppend + " And CPTCode = '" & ChkStr(TxtCPT) & "'"
End If

CPTCriteria = strAppend
Call CPTListing(strAppend)
Screen.MousePointer = Default
End Sub

Private Sub CPTListing(Optional strAppend As String)
Dim CPT As New adodb.Recordset
Dim CPTMaster As ListItem
Dim Sql As String
Dim TotalCPTCount As String

TotalCPTCount = 0
lstCPTList.listitems.Clear


    Sql = " SELECT CPTCode, CPTDesc From CPT where 0=0 "
    
    If Len(Trim(strAppend)) Then
    Sql = Sql + strAppend
    End If
    
    CPT.Open Sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    
    If CPT.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until CPT.EOF
        Set CPTMaster = lstCPTList.listitems.Add(, , CPT!CPTCode)
            CPTMaster.SubItems(1) = CPT!CPTDesc
            CPT.MoveNext
        TotalCPTCount = TotalCPTCount + 1
        txtTotalRecord = TotalCPTCount
        Loop
    CPT.Close
    Set CPT = Nothing
End If
End Sub

'Private Sub PrintCPT()
    'RptCPTList.Show
'End Sub

'*******************************Start Race *******************************
Private Sub SaveRecType()
Dim RecordSaved
Dim RecType As New adodb.Recordset
Dim SaveRecType As New adodb.Recordset
Dim strsql As String
      
    If TxtRecTypeCode = "" Then
        MsgBox "Please Enter Rec. Type Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtRecTypeCode.SetFocus
    ElseIf TxtRecTypeDesc = "" Then
        MsgBox "Please Enter Rec. Type Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtRecTypeDesc.SetFocus
    
    Else
        strsql = "Select * From ReceiptType Where RecTypeCode = '" & Trim(TxtRecTypeCode) & "'"
          
          RecType.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
          
            If RecType.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                RecType.Close
                TxtRecTypeCode.SetFocus
                TxtRecTypeCode.SelStart = 0
                TxtRecTypeCode.SelLength = Len(TxtRecTypeCode)
                RecType.Close
                Set RecType = Nothing
                Exit Sub
            Else
                RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                    If EditFlag = False Then
                        SaveRecType.Open "ReceiptType", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                        SaveRecType.AddNew
                    Else
                        strsql = "Select * From ReceiptType Where RecTypeCode = '" & Trim(TxtRecTypeCode) & "'"
                        SaveRecType.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                    End If
            
                    With SaveRecType
                        !RecTypeCode = ChkStr(TxtRecTypeCode)
                        !RecTypeDescription = ChkStr(TxtRecTypeDesc)
                        !RecTypeLastUpdateUser = Logon
                        !RecTypeLastUpdateDate = Date
                    End With
                    SaveRecType.Update
                    Call ClearForm(Me)
                    Call RecTypeListing(RACCriteria)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
                    SaveRecType.Close
                    Set SaveRecType = Nothing
                End If
                EditFlag = False
            End If
    End If

End Sub

Private Sub DeleteRecType()
Dim DeleteRecType As New adodb.Recordset
Dim rsCheckRec As New adodb.Recordset
Dim strsql As String
Dim DeleteRecord

EditFlag = False
If TxtRecTypeCode = "" Then
    MsgBox "Please Enter Rec. Type Code", vbOKOnly + vbCritical, DialogBoxTitle
ElseIf lstRecTypeList.listitems.Count = 0 Then
    MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
    DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
            strsql = "Delete From ReceiptType Where RecTypeCode = '" & Trim(TxtRecTypeCode) & "'"
            DeleteRecType.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
            CmdSave.Enabled = True
            Call ClearForm(Me)
            Call RecTypeListing
            Call DisableCmdButton(Me)
            Call EntriesEnable(True)
            TxtRecTypeCode.SetFocus
        End If
End If
End Sub

Private Sub SearchRecType()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(TxtRecTypeCode)) Then
strAppend = strAppend + " And RecTypeCode = '" & Trim(TxtRecTypeCode) & "'"
End If

If Len(Trim(TxtRecTypeDesc)) Then
strAppend = strAppend + " And RectypeDescription like '%" & ChkStr(TxtRecTypeDesc) & "%'"
End If

RecTypeCriteria = strAppend
Call RecTypeListing(strAppend)
Screen.MousePointer = Default
End Sub

Private Sub RecTypeListing(Optional strAppend As String)
Dim RecType As New adodb.Recordset
Dim RecTypeMaster As ListItem
Dim Sql As String
Dim TotalRecTypeCount As String

TotalRecTypeCount = 0
lstRecTypeList.listitems.Clear


    Sql = "SELECT RecTypeCode, RecTypeDescription From ReceiptType where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
    
    RecType.Open Sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    
    If RecType.EOF = True Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until RecType.EOF
        Set RecTypeMaster = lstRecTypeList.listitems.Add(, , RecType!RecTypeCode)
            RecTypeMaster.SubItems(1) = RecType!RecTypeDescription
            RecType.MoveNext
        TotalRecTypeCount = TotalRecTypeCount + 1
        txtTotalRecord = TotalRecTypeCount
        Loop
    RecType.Close
    Set RecType = Nothing
End If
End Sub

Private Sub lstRecTypeList_Click()
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstRecTypeList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    strsql = " SELECT RecTypeCode, RecTypeDescription " & _
             " From ReceiptType Where RecTypeCode = '" & lstRecTypeList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    TxtRecTypeCode = rst3!RecTypeCode
    TxtRecTypeDesc = rst3!RecTypeDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
   ' medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End If
End Sub

Private Sub lstRecTypeList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstRecTypeList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    strsql = " SELECT RecTypeCode, RecTypeDescription " & _
             " From ReceiptType Where RecTypeCode = '" & lstRecTypeList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    TxtRecTypeCode = rst3!RecTypeCode
    TxtRecTypeDesc = rst3!RecTypeDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
   ' medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    
End If
End Sub

Private Sub lstRecTypeList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New adodb.Recordset
Dim strsql As String

If lstRecTypeList.listitems.Count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    strsql = " SELECT RecTypeCode, RecTypeDescription " & _
             " From ReceiptType Where RecTypeCode = '" & lstRecTypeList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    TxtRecTypeCode = rst3!RecTypeCode
    TxtRecTypeDesc = rst3!RecTypeDescription
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
    
   ' medixmasterconnMaint.Close
 '   Set medixmasterconnMaint = Nothing
    
End If
End Sub

'Private Sub PrintRecType()
    'RptRecTypeList.Show
'End Sub
'*******************************End Rec. Type *******************************



