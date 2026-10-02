VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUser 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8805
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "User Maintenance"
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
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3285
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "User"
      TabPicture(0)   =   "frmUser.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstUsersList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Access Right"
      TabPicture(1)   =   "frmUser.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Department"
      TabPicture(2)   =   "frmUser.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstDPTList"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame1(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame2(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Group"
      TabPicture(3)   =   "frmUser.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lstGRPList"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame1(1)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame2(1)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame4 
         Height          =   6495
         Left            =   -74760
         TabIndex        =   55
         Top             =   360
         Width           =   8535
         Begin VB.CommandButton cmdCheck9 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7560
            TabIndex        =   107
            Top             =   3840
            Width           =   780
         End
         Begin VB.CommandButton cmdCheck8 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7440
            TabIndex        =   118
            Top             =   1785
            Width           =   780
         End
         Begin VB.CommandButton cmdCheck7 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7440
            TabIndex        =   105
            Top             =   930
            Width           =   780
         End
         Begin VB.CommandButton cmdCheck6 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7440
            TabIndex        =   106
            Top             =   345
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Verify ENDT File"
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
            Index           =   48
            Left            =   5820
            TabIndex        =   122
            Top             =   3480
            Width           =   2625
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Import ENDT File"
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
            Index           =   49
            Left            =   5820
            TabIndex        =   121
            Top             =   3720
            Width           =   2625
         End
         Begin VB.CommandButton cmdCheckAccept 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6210
            TabIndex        =   120
            Top             =   5865
            Width           =   915
         End
         Begin VB.CommandButton cmdCheckAll 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6210
            TabIndex        =   119
            Top             =   5505
            Width           =   915
         End
         Begin VB.CheckBox ACS 
            Caption         =   "User Entry"
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
            Index           =   46
            Left            =   5820
            TabIndex        =   117
            Top             =   4635
            Width           =   1530
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Error Listing"
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
            Index           =   45
            Left            =   5820
            TabIndex        =   116
            Top             =   4410
            Width           =   2625
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Export Membership No."
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
            Index           =   44
            Left            =   5820
            TabIndex        =   115
            Top             =   3180
            Width           =   2625
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Verify ASCII File"
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
            Index           =   43
            Left            =   5820
            TabIndex        =   114
            Top             =   2955
            Width           =   2625
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Import Data for New Member"
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
            Index           =   42
            Left            =   5820
            TabIndex        =   113
            Top             =   2730
            Width           =   2625
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Export To Card"
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
            Index           =   41
            Left            =   5820
            TabIndex        =   112
            Top             =   2505
            Width           =   2625
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Gen. First Authorisation Code"
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
            Index           =   40
            Left            =   5820
            TabIndex        =   111
            Top             =   1515
            Width           =   2625
         End
         Begin VB.CheckBox ACS 
            Caption         =   "By Payor"
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
            Index           =   39
            Left            =   5820
            TabIndex        =   110
            Top             =   930
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Reports"
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
            Index           =   38
            Left            =   5820
            TabIndex        =   109
            Top             =   345
            Width           =   2265
         End
         Begin VB.CommandButton cmdCheck10 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7560
            TabIndex        =   108
            Top             =   4800
            Width           =   780
         End
         Begin VB.CommandButton cmdCheck5 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4800
            TabIndex        =   81
            Top             =   720
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Group"
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
            Index           =   37
            Left            =   3360
            TabIndex        =   103
            Top             =   5070
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Department"
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
            Index           =   36
            Left            =   3360
            TabIndex        =   102
            Top             =   4845
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Premium"
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
            Index           =   35
            Left            =   3360
            TabIndex        =   101
            Top             =   4620
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Language"
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
            Index           =   34
            Left            =   3360
            TabIndex        =   100
            Top             =   4395
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "City"
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
            Index           =   33
            Left            =   3360
            TabIndex        =   99
            Top             =   4170
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "State"
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
            Index           =   32
            Left            =   3360
            TabIndex        =   98
            Top             =   3945
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Sequence Number"
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
            Index           =   31
            Left            =   3360
            TabIndex        =   97
            Top             =   3720
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Beneficiary"
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
            Index           =   30
            Left            =   3360
            TabIndex        =   96
            Top             =   3495
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Service Provider Charge"
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
            Index           =   29
            Left            =   3360
            TabIndex        =   95
            Top             =   3270
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "MCO"
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
            Index           =   28
            Left            =   3360
            TabIndex        =   94
            Top             =   3045
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Health Plan"
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
            Index           =   27
            Left            =   3360
            TabIndex        =   93
            Top             =   2820
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Race"
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
            Index           =   24
            Left            =   3360
            TabIndex        =   92
            Top             =   2145
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Plans"
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
            Index           =   23
            Left            =   3360
            TabIndex        =   91
            Top             =   1920
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
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
            Height          =   285
            Index           =   22
            Left            =   3360
            TabIndex        =   90
            Top             =   1695
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Insured Type"
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
            Index           =   21
            Left            =   3360
            TabIndex        =   89
            Top             =   1470
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Hospital"
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
            Index           =   20
            Left            =   3360
            TabIndex        =   88
            Top             =   1245
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Nationality"
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
            Index           =   19
            Left            =   3360
            TabIndex        =   87
            Top             =   1020
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Age Band"
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
            Index           =   17
            Left            =   3360
            TabIndex        =   85
            Top             =   570
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Company Information"
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
            Index           =   16
            Left            =   3360
            TabIndex        =   84
            Top             =   345
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Service Provider"
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
            Index           =   26
            Left            =   3360
            TabIndex        =   83
            Top             =   2595
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Relationship"
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
            Index           =   25
            Left            =   3360
            TabIndex        =   82
            Top             =   2370
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Reports"
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
            Index           =   51
            Left            =   3360
            TabIndex        =   80
            Top             =   5310
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Medical Report"
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
            Index           =   50
            Left            =   165
            TabIndex        =   75
            Top             =   2445
            Width           =   3075
         End
         Begin VB.CommandButton cmdCheck4 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2370
            TabIndex        =   74
            Top             =   4575
            Width           =   780
         End
         Begin VB.CommandButton cmdCheck3 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2370
            TabIndex        =   73
            Top             =   3045
            Width           =   780
         End
         Begin VB.CommandButton cmdCheck2 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2415
            TabIndex        =   72
            Top             =   1545
            Width           =   780
         End
         Begin VB.CommandButton cmdCheck1 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            TabIndex        =   71
            Top             =   480
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Authorisation Code Inquiry"
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
            Index           =   15
            Left            =   120
            TabIndex        =   70
            Top             =   5280
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Diagnosis Inquiry"
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
            Index           =   14
            Left            =   120
            TabIndex        =   69
            Top             =   5025
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Case Inquiry"
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
            Index           =   13
            Left            =   120
            TabIndex        =   68
            Top             =   4800
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Membership Inquiry"
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
            Index           =   12
            Left            =   120
            TabIndex        =   67
            Top             =   4575
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claims Delete"
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
            Index           =   11
            Left            =   120
            TabIndex        =   66
            Top             =   3945
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claims Closing"
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
            Index           =   10
            Left            =   120
            TabIndex        =   65
            Top             =   3720
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claims Adjustment"
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
            Index           =   9
            Left            =   120
            TabIndex        =   64
            Top             =   3495
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claims Inquiry"
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
            Index           =   8
            Left            =   120
            TabIndex        =   63
            Top             =   3270
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claims Entry"
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
            Index           =   7
            Left            =   120
            TabIndex        =   62
            Top             =   3045
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Generate Second Authorisation Code"
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
            Index           =   6
            Left            =   165
            TabIndex        =   61
            Top             =   2220
            Width           =   3075
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Case Adjustment"
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
            Index           =   5
            Left            =   165
            TabIndex        =   60
            Top             =   1995
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Patient's Discharge"
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
            Index           =   4
            Left            =   165
            TabIndex        =   59
            Top             =   1770
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Case Entry"
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
            Index           =   3
            Left            =   165
            TabIndex        =   58
            Top             =   1545
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Membership Adjustment"
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
            Index           =   2
            Left            =   165
            TabIndex        =   57
            Top             =   705
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Membership Registration"
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
            Index           =   1
            Left            =   165
            TabIndex        =   56
            Top             =   480
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
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
            Height          =   285
            Index           =   18
            Left            =   3360
            TabIndex        =   86
            Top             =   795
            Width           =   2265
         End
         Begin VB.Label Label21 
            Caption         =   "Check All Accept Maintenance"
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
            Left            =   3960
            TabIndex        =   11
            Top             =   5865
            Width           =   2220
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Check All"
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
            Left            =   5220
            TabIndex        =   13
            Top             =   5595
            Width           =   870
         End
         Begin VB.Label Label19 
            BackColor       =   &H000000FF&
            Caption         =   " Administration"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5820
            TabIndex        =   26
            Top             =   4185
            Width           =   2670
         End
         Begin VB.Label Label18 
            BackColor       =   &H000000FF&
            Caption         =   " Data Maintenance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5880
            TabIndex        =   27
            Top             =   2280
            Width           =   2430
         End
         Begin VB.Label Label17 
            BackColor       =   &H000000FF&
            Caption         =   " Service Provider"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5880
            TabIndex        =   28
            Top             =   1290
            Width           =   2550
         End
         Begin VB.Label Label16 
            BackColor       =   &H000000FF&
            Caption         =   " EIS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5880
            TabIndex        =   124
            Top             =   705
            Width           =   2550
         End
         Begin VB.Label Label15 
            BackColor       =   &H000000FF&
            Caption         =   " Reports"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5880
            TabIndex        =   123
            Top             =   120
            Width           =   2535
         End
         Begin VB.Label Label14 
            BackColor       =   &H000000FF&
            Caption         =   " Maintenance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3360
            TabIndex        =   104
            Top             =   120
            Width           =   2355
         End
         Begin VB.Label Label7 
            BackColor       =   &H000000FF&
            Caption         =   " Inquiry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   79
            Top             =   4350
            Width           =   3030
         End
         Begin VB.Label Label12 
            BackColor       =   &H000000FF&
            Caption         =   " Claims"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   78
            Top             =   2820
            Width           =   3030
         End
         Begin VB.Label Label11 
            BackColor       =   &H000000FF&
            Caption         =   " Admission"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   165
            TabIndex        =   77
            Top             =   1320
            Width           =   3030
         End
         Begin VB.Label Label9 
            BackColor       =   &H000000FF&
            Caption         =   " Membership"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   3030
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Index           =   0
         Left            =   -74715
         TabIndex        =   47
         Top             =   1470
         Width           =   6990
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
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
            Index           =   0
            Left            =   5940
            TabIndex        =   53
            Top             =   270
            Width           =   960
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Ok"
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
            Index           =   0
            Left            =   4995
            TabIndex        =   52
            Top             =   270
            Width           =   960
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
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
            Height          =   315
            Index           =   0
            Left            =   1980
            TabIndex        =   51
            Top             =   285
            Width           =   960
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
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
            Height          =   315
            Index           =   0
            Left            =   1035
            TabIndex        =   50
            Top             =   285
            Width           =   960
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
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
            Height          =   315
            Index           =   0
            Left            =   75
            TabIndex        =   49
            Top             =   285
            Width           =   960
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
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
            Index           =   0
            Left            =   3240
            TabIndex        =   48
            Top             =   270
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   0
         Left            =   -74760
         TabIndex        =   42
         Top             =   480
         Width           =   7005
         Begin VB.TextBox txtDPTCode 
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
            Left            =   1395
            MaxLength       =   3
            TabIndex        =   44
            Top             =   270
            Width           =   840
         End
         Begin VB.TextBox txtDPTDescription 
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
            Left            =   1395
            MaxLength       =   50
            TabIndex        =   43
            Top             =   585
            Width           =   3585
         End
         Begin VB.Label Label13 
            Caption         =   "Department Code"
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
            Left            =   90
            TabIndex        =   46
            Top             =   270
            Width           =   1500
         End
         Begin VB.Label Label2 
            Caption         =   "Description"
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
            Left            =   90
            TabIndex        =   45
            Top             =   585
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   1560
         Width           =   6990
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
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
            Index           =   1
            Left            =   3240
            TabIndex        =   40
            Top             =   270
            Width           =   960
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
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
            Height          =   315
            Index           =   1
            Left            =   75
            TabIndex        =   39
            Top             =   285
            Width           =   960
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
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
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   38
            Top             =   285
            Width           =   960
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
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
            Height          =   315
            Index           =   1
            Left            =   2070
            TabIndex        =   37
            Top             =   285
            Width           =   960
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Ok"
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
            Index           =   1
            Left            =   4950
            TabIndex        =   36
            Top             =   270
            Width           =   960
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
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
            Index           =   1
            Left            =   5940
            TabIndex        =   35
            Top             =   270
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   7005
         Begin VB.TextBox txtGRPDescription 
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
            Left            =   1035
            MaxLength       =   255
            TabIndex        =   31
            Top             =   540
            Width           =   2190
         End
         Begin VB.TextBox txtGRPCode 
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
            Left            =   1035
            MaxLength       =   4
            TabIndex        =   30
            Top             =   225
            Width           =   840
         End
         Begin VB.Label Label4 
            Caption         =   "Description"
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
            Left            =   135
            TabIndex        =   33
            Top             =   540
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Group Code"
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
            Left            =   135
            TabIndex        =   32
            Top             =   225
            Width           =   1050
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Index           =   2
         Left            =   -74760
         TabIndex        =   10
         Top             =   360
         Width           =   7125
         Begin VB.TextBox txtUSRDateCreated 
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
            IMEMode         =   3  'DISABLE
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1560
         End
         Begin VB.TextBox txtUSRCode 
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
            IMEMode         =   3  'DISABLE
            Left            =   1485
            MaxLength       =   10
            TabIndex        =   16
            Top             =   315
            Width           =   1005
         End
         Begin VB.ComboBox txtUSRStatus 
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
            Left            =   4800
            TabIndex        =   15
            Text            =   "A"
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtUSRPassword 
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
            IMEMode         =   3  'DISABLE
            Left            =   1485
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   14
            Top             =   1080
            Width           =   1920
         End
         Begin VB.TextBox txtUSRName 
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
            Left            =   1440
            MaxLength       =   200
            TabIndex        =   12
            Top             =   720
            Width           =   4035
         End
         Begin VB.Label Label10 
            Caption         =   "Status"
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
            Left            =   3720
            TabIndex        =   24
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label Label8 
            Caption         =   "Date Created"
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
            Left            =   3720
            TabIndex        =   23
            Top             =   1125
            Width           =   1185
         End
         Begin VB.Label Label2 
            Caption         =   "Staff Name"
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
            Left            =   225
            TabIndex        =   22
            Top             =   765
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "User ID"
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
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Group"
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
            Left            =   225
            TabIndex        =   20
            Top             =   1800
            Width           =   1230
         End
         Begin VB.Label Label5 
            Caption         =   "Department"
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
            Left            =   225
            TabIndex        =   19
            Top             =   1485
            Width           =   1185
         End
         Begin VB.Label Label6 
            Caption         =   "Password"
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
            Left            =   225
            TabIndex        =   18
            Top             =   1125
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         Height          =   675
         Index           =   2
         Left            =   -74760
         TabIndex        =   3
         Top             =   2760
         Width           =   7005
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
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
            Index           =   2
            Left            =   3150
            TabIndex        =   9
            Top             =   225
            Width           =   960
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
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
            Height          =   315
            Index           =   2
            Left            =   90
            TabIndex        =   8
            Top             =   225
            Width           =   960
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
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
            Height          =   315
            Index           =   2
            Left            =   1050
            TabIndex        =   7
            Top             =   225
            Width           =   960
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
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
            Height          =   315
            Index           =   2
            Left            =   2025
            TabIndex        =   6
            Top             =   225
            Width           =   960
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Ok"
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
            Index           =   2
            Left            =   5880
            TabIndex        =   5
            Top             =   225
            Width           =   960
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
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
            Index           =   2
            Left            =   7065
            TabIndex        =   4
            Top             =   225
            Width           =   960
         End
      End
      Begin MSComctlLib.ListView lstUsersList 
         Height          =   2640
         Left            =   -74760
         TabIndex        =   25
         Top             =   3600
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   4657
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Department"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstGRPList 
         Height          =   1800
         Left            =   240
         TabIndex        =   41
         Top             =   2400
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   3175
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
            Text            =   "Group Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7937
         EndProperty
      End
      Begin MSComctlLib.ListView lstDPTList 
         Height          =   2325
         Left            =   -74760
         TabIndex        =   54
         Top             =   2640
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   4101
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
            Text            =   "Department Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7937
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label18_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub lstDPTList_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lstUsersList_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub SSTab1_DblClick()

End Sub
