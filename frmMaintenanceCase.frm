VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenanceCase 
   Caption         =   "Case Maintenance"
   ClientHeight    =   7815
   ClientLeft      =   2160
   ClientTop       =   1980
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   536.029
   ScaleMode       =   0  'User
   ScaleWidth      =   933.429
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   735
      Left            =   360
      TabIndex        =   29
      Top             =   6360
      Width           =   10995
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5160
         TabIndex        =   82
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   7320
         TabIndex        =   25
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   4320
         TabIndex        =   24
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9000
         TabIndex        =   27
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   8160
         TabIndex        =   26
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   22
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   21
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3480
         TabIndex        =   23
         Top             =   240
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7035
      Left            =   120
      TabIndex        =   28
      Top             =   240
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   12409
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
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
      TabCaption(0)   =   "ICD Selection"
      TabPicture(0)   =   "frmMaintenanceCase.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label30"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label29"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label28"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label27"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lstICDCodeList"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lstOriginalICDList"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdAdd2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CmdRemove"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "ICD Maintenance"
      TabPicture(1)   =   "frmMaintenanceCase.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstICDList"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label34"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "ICD Category"
      TabPicture(2)   =   "frmMaintenanceCase.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstICDCATList"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Repudiation Code"
      TabPicture(3)   =   "frmMaintenanceCase.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lstRCList"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label7"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Repudiation Category"
      TabPicture(4)   =   "frmMaintenanceCase.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lstRepCatList"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label10"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Nature of Claim"
      TabPicture(5)   =   "frmMaintenanceCase.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame5"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "lstClaimNatureList"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label18"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Treatment Modality"
      TabPicture(6)   =   "frmMaintenanceCase.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame4"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "lstModalityList"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label14"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Hospitalisation Type"
      TabPicture(7)   =   "frmMaintenanceCase.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame7"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "lstPayTypeList"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Label26"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).ControlCount=   3
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   -74760
         TabIndex        =   76
         Top             =   960
         Width           =   10965
         Begin VB.TextBox TxtPayTypeCode 
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
            MaxLength       =   2
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox TxtPayTypeDesc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   2760
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   1080
            Width           =   8025
         End
         Begin VB.TextBox TxtPayTypeName 
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
            MaxLength       =   100
            TabIndex        =   18
            Top             =   720
            Width           =   5655
         End
         Begin VB.Label Label25 
            Caption         =   "Hospitalisation Type Desc.   *"
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
            TabIndex        =   79
            Top             =   1080
            Width           =   2595
         End
         Begin VB.Label Label24 
            Caption         =   "Hospitalisation Type Code   *"
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
            TabIndex        =   78
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label23 
            Caption         =   "Hospitalisation Type Name  *"
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
            TabIndex        =   77
            Top             =   720
            Width           =   2655
         End
      End
      Begin VB.CommandButton CmdRemove 
         Caption         =   "&Unselect"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         TabIndex        =   70
         Top             =   3360
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd2 
         Caption         =   "&Select"
         Height          =   315
         Left            =   5400
         TabIndex        =   69
         Top             =   2760
         Width           =   840
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   -74760
         TabIndex        =   63
         Top             =   960
         Width           =   10965
         Begin VB.TextBox TxtModalityCode 
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
            MaxLength       =   2
            TabIndex        =   14
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox TxtModalityDesc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   2160
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   1080
            Width           =   8625
         End
         Begin VB.TextBox TxtModalityName 
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
            MaxLength       =   50
            TabIndex        =   15
            Top             =   720
            Width           =   4575
         End
         Begin VB.Label Label12 
            Caption         =   "Modality Description *"
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
            TabIndex        =   66
            Top             =   1080
            Width           =   1995
         End
         Begin VB.Label Label13 
            Caption         =   "Modality Code            *"
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
            TabIndex        =   65
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label15 
            Caption         =   "Modality Name           *"
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
            TabIndex        =   64
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   -74760
         TabIndex        =   58
         Top             =   960
         Width           =   10965
         Begin VB.TextBox TxtClaimNatureName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   2160
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   720
            Width           =   8505
         End
         Begin VB.TextBox TxtClaimNatureCode 
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
            TabIndex        =   12
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Claim Nature Code   *"
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
            TabIndex        =   60
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label17 
            Caption         =   "Claim Nature Name  *"
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
            TabIndex        =   59
            Top             =   720
            Width           =   2235
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74760
         TabIndex        =   52
         Top             =   960
         Width           =   10965
         Begin VB.TextBox TxtRepCatDescription 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   2880
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1080
            Width           =   7905
         End
         Begin VB.TextBox TxtRepCatCode 
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
            Left            =   2880
            MaxLength       =   4
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox TxtRepCatName 
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
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   10
            Top             =   720
            Width           =   5175
         End
         Begin VB.Label Label8 
            Caption         =   "Repudiation Cat. Code              *"
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
            TabIndex        =   55
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label9 
            Caption         =   "Repudiation Cat. Name             *"
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
            TabIndex        =   54
            Top             =   720
            Width           =   2715
         End
         Begin VB.Label Label11 
            Caption         =   "Repudiation Cat. Description  *"
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
            TabIndex        =   53
            Top             =   1080
            Width           =   2715
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
         Height          =   1995
         Left            =   -74760
         TabIndex        =   47
         Top             =   960
         Width           =   10965
         Begin VB.ComboBox CmbCATCode 
            Height          =   315
            Left            =   2640
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TxtRCCode 
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
            Left            =   2640
            MaxLength       =   10
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtRCDescription 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   2640
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   1080
            Width           =   8145
         End
         Begin VB.Label Label31 
            Caption         =   "Repudiation Cat. Code     *    "
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
            TabIndex        =   71
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "Repudiation Description *"
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
            TabIndex        =   49
            Top             =   1080
            Width           =   2235
         End
         Begin VB.Label Label6 
            Caption         =   "Repudiation Code             *"
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
            TabIndex        =   48
            Top             =   720
            Width           =   2295
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
         Height          =   1635
         Left            =   -74760
         TabIndex        =   42
         Top             =   960
         Width           =   10965
         Begin VB.TextBox TxtICDCATDescription 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   2160
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   720
            Width           =   8505
         End
         Begin VB.TextBox TxtICDCATCode 
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
            MaxLength       =   10
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "ICD Category Code   *"
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
            TabIndex        =   44
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "ICD Description         *"
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
            TabIndex        =   43
            Top             =   720
            Width           =   1875
         End
      End
      Begin VB.Frame Frame10 
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
         TabIndex        =   37
         Top             =   960
         Width           =   10965
         Begin VB.TextBox TxtOtherDesc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   2280
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1440
            Width           =   8505
         End
         Begin VB.ComboBox CboICDCAT 
            Height          =   315
            Left            =   5640
            TabIndex        =   1
            Top             =   360
            Width           =   5175
         End
         Begin VB.TextBox TxtICDDescription 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   2280
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   720
            Width           =   8505
         End
         Begin VB.TextBox TxtICDCode 
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
            MaxLength       =   5
            TabIndex        =   0
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label38 
            Caption         =   "ICD Other Description *"
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
            TabIndex        =   75
            Top             =   1440
            Width           =   1995
         End
         Begin VB.Label Label37 
            Caption         =   "ICD Cat. Code         *"
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
            Left            =   3960
            TabIndex        =   74
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label32 
            Caption         =   "ICD Code                        *"
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
            TabIndex        =   39
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label33 
            Caption         =   "ICD Description            *"
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
            TabIndex        =   38
            Top             =   720
            Width           =   2115
         End
      End
      Begin MSComctlLib.ListView lstOriginalICDList 
         Height          =   4815
         Left            =   360
         TabIndex        =   31
         Top             =   1320
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
            Text            =   "ICD Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ICD Description"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstICDCodeList 
         Height          =   4815
         Left            =   6720
         TabIndex        =   32
         Top             =   1320
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
            Text            =   "ICD Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ICD Description"
            Object.Width           =   97009
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   97009
         EndProperty
      End
      Begin MSComctlLib.ListView lstICDList 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   40
         Top             =   3480
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   4683
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
            Text            =   "ICD Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ICD Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ICD Other Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ICD Category Code"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lstICDCATList 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   45
         Top             =   2880
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   5741
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
            Text            =   "ICD Category Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ICD Catergory Description"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstRCList 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   50
         Top             =   3240
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   5106
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
            Text            =   "RC Cat Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "RC Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RC Description"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstRepCatList 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   56
         Top             =   3360
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   4895
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
            Text            =   "RC CAT Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "RC Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RC Description"
            Object.Width           =   12347
         EndProperty
      End
      Begin MSComctlLib.ListView lstClaimNatureList 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   61
         Top             =   2880
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   5741
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
            Text            =   "Claim Nature Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Claim Nature Name"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstModalityList 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   67
         Top             =   3360
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   4895
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
            Text            =   "Modality Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Modality Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Modality Description"
            Object.Width           =   12347
         EndProperty
      End
      Begin MSComctlLib.ListView lstPayTypeList 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   80
         Top             =   3360
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   4895
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
            Text            =   "Hospitalisation Type Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Hospitalisation Type Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hospitalisation Type Description"
            Object.Width           =   12347
         EndProperty
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "HOSPITALISATION TYPE LIST"
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
         TabIndex        =   81
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "MODALITY CODE LIST"
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
         TabIndex        =   68
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "CLAIM NATURE CODE LIST"
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
         TabIndex        =   62
         Top             =   2640
         Width           =   2235
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "REPUDIATION CODE LIST"
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
         TabIndex        =   57
         Top             =   3120
         Width           =   2115
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "REPUDIATION CODE LIST"
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
         TabIndex        =   51
         Top             =   3000
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ICD CATEGORY CODE LIST"
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
         TabIndex        =   46
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "ICD CODE LIST"
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
         TabIndex        =   41
         Top             =   3240
         Width           =   1275
      End
      Begin VB.Label Label27 
         Caption         =   "Orginal ICD Code"
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
         Left            =   360
         TabIndex        =   36
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label28 
         Caption         =   "ICD Code Applied"
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
         Left            =   6720
         TabIndex        =   35
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label29 
         Caption         =   "<======"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   34
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "======>"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   33
         Top             =   2280
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Case Information"
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
      Height          =   219
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   73
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label36 
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
      Left            =   5760
      TabIndex        =   72
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label35 
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
      TabIndex        =   30
      Top             =   7320
      Width           =   4695
   End
End
Attribute VB_Name = "frmMaintenanceCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim bExit As Boolean
Public ICDCriteria As String
Public ICDCATCriteria As String
Public RCCriteria As String
Public RCCATCriteria As String
Public MODALITYCriteria As String
Public CLAIMNATURECriteria As String
'Public PAYMODECriteria As String
Public PAYTYPECriteria As String

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

'***************************Start EntriesEnable**********************************
Private Sub EntriesEnable(status As Boolean)
        
'    TxtCasExCode.Enabled = status
'    TxtCasExDescription.Enabled = status
    'TxtUDCode.Enabled = status
    'TxtUDName.Enabled = status
    'TxtUDDescription.Enabled = status
    CmbCATCode.Enabled = status
    TxtICDCode.Enabled = status
    TxtICDDescription.Enabled = status
    TxtICDCATCode.Enabled = status
    CboICDCAT.Enabled = status
    TxtICDCATDescription.Enabled = status
    TxtOtherDesc.Enabled = status
    TxtRCCode.Enabled = status
    TxtRCDescription.Enabled = status
    TxtRepCatCode.Enabled = status
    TxtRepCatName.Enabled = status
    TxtRepCatDescription.Enabled = status
    TxtModalityCode.Enabled = status
    TxtModalityName.Enabled = status
    TxtModalityDesc.Enabled = status
    TxtClaimNatureCode.Enabled = status
    TxtClaimNatureName.Enabled = status
    'TxtPayModeCode.Enabled = status
    'TxtPayModeName.Enabled = status
    'TxtPayModeDesc.Enabled = status
    TxtPayTypeCode.Enabled = status
    TxtPayTypeName.Enabled = status
    TxtPayTypeDesc.Enabled = status
    

End Sub
'***************************End EntriesEnable**********************************

Private Sub RCComboListing()
Dim RC As New ADODB.Recordset
Dim cmd As New ADODB.Command
        
        
    ' Set up a command object for the stored procedure.
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchRepCodeCategory"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    RC.CursorLocation = adUseClient
    RC.CursorType = adOpenForwardOnly
    RC.CacheSize = 100
        
    ' Create a recordset by executing the command.
    Set RC = cmd.Execute
    CmbCATCode.Clear
    
    Do Until RC.EOF
       CmbCATCode.AddItem RC!RepCatCode
       RC.MoveNext
    Loop
        
    RC.Close
    Set RC = Nothing
       
End Sub

Private Sub ICDCATCombolisting()
Dim ICDCAT As New ADODB.Recordset
Dim cmd As New ADODB.Command

        
    ' Set up a command object for the stored procedure.
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchICDCategory"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    ICDCAT.CursorLocation = adUseClient
    ICDCAT.CursorType = adOpenForwardOnly
    ICDCAT.CacheSize = 100
        
    ' Create a recordset by executing the command.
    Set ICDCAT = cmd.Execute
    
    
    'Screen.MousePointer = vbHourglass
    'ICDCAT.Open "Select ICDCATCode ,  ICDCATDescription from ICDCategory", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    CboICDCAT.Clear
    
    Do Until ICDCAT.EOF
         CboICDCAT.AddItem ICDCAT!ICDCATCode & " - " & ICDCAT!ICDCATDescription
         ICDCAT.MoveNext
    Loop
    ICDCAT.Close
    Set ICDCAT = Nothing
    'Screen.MousePointer = Default

        
End Sub

Private Sub CboICDCAT_Change()
    If InStr(CboICDCAT.Text, "null") Then
       CboICDCAT.Enabled = False
    End If
End Sub

Private Sub CboICDCAT_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboICDCAT.Text, CboICDCAT.SelStart) + Chr(KeyAscii) + Mid(CboICDCAT.Text, CboICDCAT.SelStart + CboICDCAT.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboICDCAT.ListCount - 1
 
        If NewText = LCase(Left(CboICDCAT.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboICDCAT.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboICDCAT.Text = ValidValue
                CboICDCAT.SelStart = 0
                CboICDCAT.SelLength = Len(ValidValue)
            End If
        End If
End Sub


Private Sub CmbCATCode_Change()
    If InStr(CmbCATCode.Text, "null") Then
       CmbCATCode.Enabled = False
    End If
End Sub

Private Sub CmbCATCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CmbCATCode.Text, CmbCATCode.SelStart) + Chr(KeyAscii) + Mid(CmbCATCode.Text, CmbCATCode.SelStart + CmbCATCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CmbCATCode.ListCount - 1
 
        If NewText = LCase(Left(CmbCATCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CmbCATCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               CmbCATCode.Text = ValidValue
               CmbCATCode.SelStart = 0
               CmbCATCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstICDList.listitems.Clear
    lstICDCATList.listitems.Clear
    lstRCList.listitems.Clear
    lstRepCatList.listitems.Clear
    lstClaimNatureList.listitems.Clear
    lstModalityList.listitems.Clear
    lstPayTypeList.listitems.Clear
    txtTotalRecord = 0
    cmdAdd.Enabled = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub CmdPrint_Click()
  
    EditFlag = False
    Call ClearForm(Me)
    'Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    'cmdSave.Enabled = false
    
    Select Case SSTab1.Tab
        Case 1
            If lstICDList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstICDList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 2
            If lstICDCATList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstICDCATList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 3
            If lstRCList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstRCList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 4
            If lstRepCatList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstRepCatList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 5
            If lstClaimNatureList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstClaimNatureList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 6
            If lstModalityList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstModalityList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 7
            If lstPayTypeList.listitems.Count <> 0 Then
                Call ExportListViewtoExcel(lstPayTypeList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
    End Select
    
End Sub

Private Sub CmdRefresh_Click()

Call ClearForm(Me)
Call EntriesEnable(True)

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

Select Case SSTab1.Tab

    'Case 0
        'Call PayorListing
    Case 1
        ICDCATCombolisting
    'Case 2
        'Call HospitalListing
    Case 3
        Call RCComboListing
    'Case 4
        'txtINScode.Refresh
        'txtHLTCode.Refresh
        'Call MCOComboListing
    'Case 5
        'Call ServiceProviderListing
    'Case 6
       'Call PRVChargesComboListing
       'Call PRVListing
    'Case 7
       'Call CASEListing
    'Case 8
       'Call ICDListing
    'Case 9
        'Call WorkListing
End Select

'medixmasterconn.Close
'Set medixmasterconn = Nothing

End Sub


'***************************Start Form**********************************


Private Sub Form_Load()
    cmdDelete.Enabled = True
    Me.Width = SSTab1.Width + 500
    ICDCriteria = ""
    ICDCATCriteria = ""
    RCCriteria = ""
    RCCATCriteria = ""
    MODALITYCriteria = ""
    CLAIMNATURECriteria = ""
    PAYTYPECriteria = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Call EnterToTab(KeyAscii)
End Sub
'***************************End Form**********************************

'Private Sub Form_Resize()
    'Dim i As Integer
'End Sub

Private Sub lstClaimNatureList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
        
    If lstClaimNatureList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtClaimNatureCode = lstClaimNatureList.SelectedItem.Text
        TxtClaimNatureName = lstClaimNatureList.SelectedItem.SubItems(1)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstClaimNatureList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
    
        'strsql = "Select ClaimNatureCode, ClaimNatureName " & _
        '"from ClaimNature Where ClaimNatureCode = '" & lstClaimNatureList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        'TxtClaimNatureCode = Trim(rst3!ClaimNatureCode)
        'TxtClaimNatureName = Trim(rst3!ClaimNatureName)
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstClaimNatureList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstClaimNatureList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtClaimNatureCode = lstClaimNatureList.SelectedItem.Text
        TxtClaimNatureName = lstClaimNatureList.SelectedItem.SubItems(1)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstClaimNatureList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
    
        'strsql = "Select ClaimNatureCode, ClaimNatureName " & _
        '"from ClaimNature Where ClaimNatureCode = '" & lstClaimNatureList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        'TxtClaimNatureCode = Trim(rst3!ClaimNatureCode)
        'TxtClaimNatureName = Trim(rst3!ClaimNatureName)
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstICDCATList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstICDCATList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtICDCATCode = lstICDCATList.SelectedItem.Text
        TxtICDCATDescription = lstICDCATList.SelectedItem.SubItems(1)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
        
        
        'If lstICDCATList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
      
        'strsql = "Select ICDCATCode, ICDCATDescription " & _
        '"from ICDCategory Where ICDCATCode = '" & lstICDCATList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtICDCATCode = Trim(rst3!ICDCATCode)
        'TxtICDCATDescription = Trim(rst3!ICDCATDescription)
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstICDCATList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstICDCATList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtICDCATCode = lstICDCATList.SelectedItem.Text
        TxtICDCATDescription = lstICDCATList.SelectedItem.SubItems(1)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
        'If lstICDCATList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
      
        'strsql = "Select ICDCATCode, ICDCATDescription " & _
        '"from ICDCategory Where ICDCATCode = '" & lstICDCATList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtICDCATCode = Trim(rst3!ICDCATCode)
        'TxtICDCATDescription = Trim(rst3!ICDCATDescription)
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstICDList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstICDList.SortKey = ColumnHeader.Index - 1
    lstICDList.Sorted = True
        If lstICDList.SortOrder = lvwAscending Then
            lstICDList.SortOrder = lvwDescending
        Else
            lstICDList.SortOrder = lvwAscending
        End If
End Sub

Private Sub lstICDList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
     If lstICDList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtICDCode = lstICDList.SelectedItem.Text
        TxtICDDescription = lstICDList.SelectedItem.SubItems(1)
        TxtOtherDesc = lstICDList.SelectedItem.SubItems(2)
        CboICDCAT = lstICDList.SelectedItem.SubItems(3)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstICDList.listitems.count <> 0 Then
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
    
        'strsql = "Select ICDCode, ICDCATCode, ICDDescription, ICDOtherDesc " & _
        '"from ICD Where ICDCode = '" & lstICDList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtICDCode = Trim(rst3!ICDCode)
        'If Len(Trim(Mid(rst3!ICDCATCode, 1, 7))) Then
           'CboICDCAT = Trim(Mid(rst3!ICDCATCode, 1, 7))
        'End If
        'TxtICDDescription = Trim(rst3!ICDDescription)
        'If Len(Trim(rst3!ICDOtherDesc)) Then
            'TxtOtherDesc = Trim(rst3!ICDOtherDesc)
        'End If
            
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstICDList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstICDList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtICDCode = lstICDList.SelectedItem.Text
        TxtICDDescription = lstICDList.SelectedItem.SubItems(1)
        TxtOtherDesc = lstICDList.SelectedItem.SubItems(2)
        CboICDCAT = lstICDList.SelectedItem.SubItems(3)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
               
    'If lstICDList.listitems.count <> 0 Then
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
    
        'strsql = "Select ICDCode, ICDCATCode, ICDDescription, ICDOtherDesc " & _
        '"from ICD Where ICDCode = '" & lstICDList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtICDCode = Trim(rst3!ICDCode)
        'If Len(Trim(Mid(rst3!ICDCATCode, 1, 7))) Then
           'CboICDCAT = Trim(Mid(rst3!ICDCATCode, 1, 7))
        'End If
        'TxtICDDescription = Trim(rst3!ICDDescription)
        'If Len(Trim(rst3!ICDOtherDesc)) Then
            'TxtOtherDesc = Trim(rst3!ICDOtherDesc)
        'End If
            
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstModalityList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstModalityList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtModalityCode = lstModalityList.SelectedItem.Text
        TxtModalityName = lstModalityList.SelectedItem.SubItems(1)
        TxtModalityDesc = lstModalityList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    
    'If lstModalityList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
            
        'strsql = "Select ModalityCode, ModalityName, ModalityDesc " & _
        '"from TreatmentModality Where ModalityCode = '" & lstModalityList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtModalityCode = Trim(rst3!ModalityCode)
        'TxtModalityName = Trim(rst3!ModalityName)
        'If Len(Trim(rst3!ModalityDesc)) Then
           'TxtModalityDesc = Trim(rst3!ModalityDesc)
        'End If
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstModalityList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstModalityList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtModalityCode = lstModalityList.SelectedItem.Text
        TxtModalityName = lstModalityList.SelectedItem.SubItems(1)
        TxtModalityDesc = lstModalityList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstModalityList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
            
        'strsql = "Select ModalityCode, ModalityName, ModalityDesc " & _
        '"from TreatmentModality Where ModalityCode = '" & lstModalityList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtModalityCode = Trim(rst3!ModalityCode)
        'TxtModalityName = Trim(rst3!ModalityName)
        'If Len(Trim(rst3!ModalityDesc)) Then
           'TxtModalityDesc = Trim(rst3!ModalityDesc)
        'End If
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstPayTypeList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstPayTypeList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtPayTypeCode = lstPayTypeList.SelectedItem.Text
        TxtPayTypeName = lstPayTypeList.SelectedItem.SubItems(1)
        TxtPayTypeDesc = lstPayTypeList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
        
    'If lstPayTypeList.listitems.count <> 0 Then
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
            
        'strsql = "Select PayTypeCode, PayTypeName, PayTypeDesc " & _
        '"from PayType Where PayTypeCode = '" & lstPayTypeList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtPayTypeCode = Trim(rst3!PayTypeCode)
        'TxtPayTypeName = Trim(rst3!PayTypeName)
        'TxtPayTypeDesc = Trim(rst3!PayTypeDesc)
            
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstPayTypeList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstPayTypeList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtPayTypeCode = lstPayTypeList.SelectedItem.Text
        TxtPayTypeName = lstPayTypeList.SelectedItem.SubItems(1)
        TxtPayTypeDesc = lstPayTypeList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
        
    'If lstPayTypeList.listitems.count <> 0 Then
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
            
        'strsql = "Select PayTypeCode, PayTypeName, PayTypeDesc " & _
        '"from PayType Where PayTypeCode = '" & lstPayTypeList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtPayTypeCode = Trim(rst3!PayTypeCode)
        'TxtPayTypeName = Trim(rst3!PayTypeName)
        'TxtPayTypeDesc = Trim(rst3!PayTypeDesc)
            
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstRCList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstRCList.SortKey = ColumnHeader.Index - 1
    lstRCList.Sorted = True
        If lstRCList.SortOrder = lvwAscending Then
            lstRCList.SortOrder = lvwDescending
        Else
            lstRCList.SortOrder = lvwAscending
        End If

End Sub

Private Sub lstRCList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstRCList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        CmbCATCode = lstRCList.SelectedItem.Text
        TxtRCCode = lstRCList.SelectedItem.SubItems(1)
        TxtRCDescription = lstRCList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    
    'If lstRCList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
            
        'strsql = "Select RepCatCode, RepCode, RepDescription " & _
        '"from RepCode Where RepCode = '" & Trim(lstRCList.SelectedItem.Text) & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'CmbCATCode = lstRCList.SelectedItem.Text
        'TxtRCCode = lstRCList.SelectedItem.SubItems(1)
        'TxtRCDescription = lstRCList.SelectedItem.SubItems(2)
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstRCList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstRCList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        CmbCATCode = lstRCList.SelectedItem.Text
        TxtRCCode = lstRCList.SelectedItem.SubItems(1)
        TxtRCDescription = lstRCList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstRCList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
            
        'strsql = "Select RepCatCode, RepCode, RepDescription " & _
        '"from RepCode Where RepCode = '" & Trim(lstRCList.SelectedItem.Text) & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'CmbCATCode = lstRCList.SelectedItem.Text
        'TxtRCCode = lstRCList.SelectedItem.SubItems(1)
        'TxtRCDescription = lstRCList.SelectedItem.SubItems(2)
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstRepCatList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstRepCatList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtRepCatCode = lstRepCatList.SelectedItem.Text
        TxtRepCatName = lstRepCatList.SelectedItem.SubItems(1)
        TxtRepCatDescription = lstRepCatList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    
    'If lstRepCatList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
                    
        'strsql = "Select RepCatCode, RepCatName, RepCatDescription " & _
        '"from RepCodeCategory Where RepCatCode = '" & lstRepCatList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtRepCatCode = Trim(rst3!RepCatCode)
        'TxtRepCatName = Trim(rst3!RepCatName)
        'TxtRepCatDescription = Trim(rst3!RepCatDescription)
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstRepCatList_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstRepCatList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtRepCatCode = lstRepCatList.SelectedItem.Text
        TxtRepCatName = lstRepCatList.SelectedItem.SubItems(1)
        TxtRepCatDescription = lstRepCatList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstRepCatList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
                    
        'strsql = "Select RepCatCode, RepCatName, RepCatDescription " & _
        '"from RepCodeCategory Where RepCatCode = '" & lstRepCatList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtRepCatCode = Trim(rst3!RepCatCode)
        'TxtRepCatName = Trim(rst3!RepCatName)
        'TxtRepCatDescription = Trim(rst3!RepCatDescription)
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

'*****************************Start Tab Clicking**********************************
Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
    Case 0 'ICD List
         txtTotalRecord = 0
         cmdAdd.Enabled = False
         cmdDelete.Enabled = False
         cmdEdit.Enabled = False
         cmdSearch.Enabled = False
         CmdRefresh.Enabled = False
         CmdClear.Enabled = False
         cmdSave.Enabled = False
         CmdPrint.Enabled = False
    Case 1 'ICD
         lstICDList.listitems.Clear
         txtTotalRecord = 0
         Call ICDCATCombolisting
         cmdAdd.Enabled = True
         cmdSearch.Enabled = True
         CmdRefresh.Enabled = True
         CmdClear.Enabled = True
         cmdSave.Enabled = True
         CmdPrint.Enabled = True
         Call ClearForm(Me)
    Case 2 'ICD Category
         lstICDCATList.listitems.Clear
         txtTotalRecord = 0
         cmdAdd.Enabled = True
         cmdSearch.Enabled = True
         CmdRefresh.Enabled = False
         CmdClear.Enabled = True
         cmdSave.Enabled = True
         CmdPrint.Enabled = True
         Call ClearForm(Me)
    Case 3 'RC
         lstRCList.listitems.Clear
         txtTotalRecord = 0
         cmdAdd.Enabled = True
         cmdSearch.Enabled = True
         CmdRefresh.Enabled = True
         CmdClear.Enabled = True
         CmdPrint.Enabled = True
         cmdSave.Enabled = True
         Call ClearForm(Me)
         Call RCComboListing
    Case 4 'RC Category
         lstRepCatList.listitems.Clear
         txtTotalRecord = 0
         cmdAdd.Enabled = True
         cmdSearch.Enabled = True
         CmdRefresh.Enabled = False
         CmdClear.Enabled = True
         cmdSave.Enabled = True
         CmdPrint.Enabled = True
         Call ClearForm(Me)
    Case 5 'Nature of Claim
         lstClaimNatureList.listitems.Clear
         txtTotalRecord = 0
         cmdAdd.Enabled = True
         cmdSearch.Enabled = True
         CmdRefresh.Enabled = False
         CmdClear.Enabled = True
         cmdSave.Enabled = True
         CmdPrint.Enabled = True
         Call ClearForm(Me)
    Case 6 'Treatment Modality
         lstModalityList.listitems.Clear
         txtTotalRecord = 0
         cmdAdd.Enabled = True
         cmdSearch.Enabled = True
         CmdRefresh.Enabled = False
         CmdClear.Enabled = True
         cmdSave.Enabled = True
         CmdPrint.Enabled = True
         Call ClearForm(Me)
    'Case 7 'Payment Mode
         'lstPayModeList.listitems.Clear
         'txtTotalRecord = 0
         'cmdAdd.Enabled = True
         'cmdSearch.Enabled = True
         'CmdRefresh.Enabled = False
         'cmdSave.Enabled = True
         'CmdPrint.Enabled = True
         'Call ClearForm(Me)
    Case 7 'Claim Type
         lstPayTypeList.listitems.Clear
         txtTotalRecord = 0
         cmdAdd.Enabled = True
         cmdSearch.Enabled = True
         CmdRefresh.Enabled = False
         CmdClear.Enabled = True
         cmdSave.Enabled = True
         CmdPrint.Enabled = True
         Call ClearForm(Me)
    End Select
    Screen.MousePointer = Default
End Sub
'*****************************End Tab Clicking**********************************

'********************Start Add, Delete, Edit, Save, Search, Cancel buttons******************
Private Sub cmdAdd_Click()
EditFlag = False
Call ClearForm(Me)
Call DisableCmdButton(Me)
Call EntriesEnable(True)
cmdSave.Enabled = True

    Select Case SSTab1.Tab
        'Case 0
         '   Call AddRecord
          '  Call ICDListView2
        Case 1
            TxtICDCode.SetFocus
        Case 2
            TxtICDCATCode.SetFocus
        Case 3
            CmbCATCode.SetFocus
        Case 4
            TxtRepCatCode.SetFocus
        Case 5
            TxtClaimNatureCode.SetFocus
        Case 6
            TxtModalityCode.SetFocus
        Case 7
            'TxtPayModeCode.SetFocus
        'Case 8
            TxtPayTypeCode.SetFocus
    End Select

End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass
    'If SSTab1.Tab = 0 Then
        'Call RemoveRecord
        'lstICDCodeList.listitems.Clear
        'Call ListDestination
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If SSTab1.Tab = 1 Then
        Call DeleteICD
    ElseIf SSTab1.Tab = 2 Then
        Call DeleteICDCAT
    ElseIf SSTab1.Tab = 3 Then
        Call DeleteRC
    ElseIf SSTab1.Tab = 4 Then
        Call DeleteRCCAT
    ElseIf SSTab1.Tab = 5 Then
        Call DeleteClaimNature
    ElseIf SSTab1.Tab = 6 Then
        Call DeleteModality
    ElseIf SSTab1.Tab = 7 Then
        'Call DeletePayMode
    'ElseIf SSTab1.Tab = 8 Then
        Call DeletePayType
    End If
    
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
cmdAdd.Enabled = True
Screen.MousePointer = Default
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    
    'If TxtCasExCode <> "" Then
     '  TxtCasExCode.Enabled = False
      ' TxtCasExDescription.SetFocus
    'End If
        
    'If TxtUDCode <> "" Then
       'TxtUDCode.Enabled = False
       'TxtUDName.SetFocus
       'TxtUDDescription.Enabled = True
    'End If
    
    If TxtICDCode <> "" Then
       TxtICDCode.Enabled = False
       TxtICDDescription.SetFocus
    End If
    
    If TxtICDCATCode <> "" Then
       TxtICDCATCode.Enabled = False
       TxtICDCATDescription.SetFocus
    End If
    
    If CmbCATCode <> "" Then
       CmbCATCode.Enabled = False
       TxtRCCode.Enabled = False
       TxtRCDescription.SetFocus
    End If
    
    If TxtRepCatCode <> "" Then
       TxtRepCatCode.Enabled = False
       TxtRepCatName.SetFocus
       TxtICDDescription.Enabled = True
    End If
    
    If TxtModalityCode <> "" Then
       TxtModalityCode.Enabled = False
       TxtModalityName.SetFocus
       TxtModalityDesc.Enabled = True
    End If
    
    If TxtClaimNatureCode <> "" Then
       TxtClaimNatureCode.Enabled = False
       TxtClaimNatureName.SetFocus
    End If
    
    'If TxtPayModeCode <> "" Then
       'TxtPayModeCode.Enabled = False
       'TxtPayModeName.SetFocus
       'TxtPayModeDesc.Enabled = True
    'End If
    
    If TxtPayTypeCode <> "" Then
       TxtPayTypeCode.Enabled = False
       TxtPayTypeName.SetFocus
       TxtPayTypeDesc.Enabled = True
    End If
    
End Sub

Private Sub cmdSave_Click()
   Screen.MousePointer = vbHourglass
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Select Case SSTab1.Tab
        'Case 0
            'Call SaveUD
        'Case 1
            'Call SaveCASE
        Case 1
            Call SaveICD
        Case 2
            Call SaveICDCAT
        Case 3
            Call SaveRC
        Case 4
            Call SaveRCCAT
        Case 5
            Call SaveClaimNature
        Case 6
            Call SaveModality
        Case 7
            'Call SavePayMode
        'Case 8
            Call SavePayType
    End Select
    cmdAdd.Enabled = True
    cmdEdit.Enabled = False
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
    Screen.MousePointer = Default
End Sub

Private Sub cmdSearch_Click()

  ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   cmdSearch.Enabled = False
   Screen.MousePointer = vbHourglass
   Select Case SSTab1.Tab
        Case 1
            Call SearchICD
            TxtICDCode.Enabled = True
            CboICDCAT.Enabled = True
            TxtICDDescription.Enabled = True
            TxtOtherDesc.Enabled = True
            TxtICDCode.SetFocus
        Case 2
            Call SearchICDCAT
            TxtICDCATCode.Enabled = True
            TxtICDCATDescription.Enabled = True
            TxtICDCATCode.SetFocus
        Case 3
            Call SearchRC
            CmbCATCode.Enabled = True
            TxtRCCode.Enabled = True
            TxtRCDescription.Enabled = True
            CmbCATCode.SetFocus
        Case 4
            Call SearchRCCAT
            TxtRepCatCode.Enabled = True
            TxtRepCatName.Enabled = True
            TxtRepCatDescription.Enabled = True
            TxtRepCatCode.SetFocus
        Case 5
            Call SearchClaimNature
            TxtClaimNatureCode.Enabled = True
            TxtClaimNatureName.Enabled = True
            TxtClaimNatureCode.SetFocus
        Case 6
            Call SearchModality
            TxtModalityCode.Enabled = True
            TxtModalityName.Enabled = True
            TxtModalityDesc.Enabled = True
            TxtModalityCode.SetFocus
        Case 7
            Call SearchPayType
            TxtPayTypeCode.Enabled = True
            TxtPayTypeName.Enabled = True
            TxtPayTypeDesc.Enabled = True
            TxtPayTypeCode.SetFocus
        
    End Select
    Screen.MousePointer = Default
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    cmdAdd.Enabled = True
    cmdSearch.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'********************End Add, Delete, Edit, Save, Search, Cancel buttons******************


'***************************Start Case Exclusion**********************************

'Private Sub SaveCASE()
'On Error Resume Next
'Dim RecordSaved
'Dim CasEx As New ADODB.Recordset
'Dim SaveCasEx As New ADODB.Recordset
'Dim strsql As String
    
    
 '   If TxtCasExCode = "" Then
  '      MsgBox "Please Enter Case Exclusion Code", vbOKOnly + vbCritical, DialogBoxTitle
   '     TxtCasExCode.SetFocus
    'ElseIf TxtCasExDescription = "" Then
     '   MsgBox "Please Enter Case Exclusion Description", vbOKOnly + vbCritical, DialogBoxTitle
      '  TxtCasExDescription.SetFocus
    'Else
     '   strsql = "Select * From CASExclusion Where CASExclusionCode = '" & Trim(TxtCasExCode) & "'"
      '  CasEx.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
       '     If CasEx.recordCount > 0 And EditFlag = False Then
        '    MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
         '   CasEx.Close
          '  TxtCasExCode.SetFocus
           ' TxtCasExCode.SelStart = 0
            'TxtCasExCode.SelLength = Len(TxtCasExCode)
            'Set CasEx = Nothing
            'Exit Sub
            'Else
            'RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            'If RecordSaved = vbYes Then
            'If EditFlag = False Then
      
             'SaveCasEx.Open "CASExclusion", medixmasterconn, adOpenKeyset, adLockOptimistic
             'SaveCasEx.AddNew
             'Else
              '  strsql = "Select * From CASExclusion Where CASExclusionCode = '" & Trim(TxtCasExCode) & "'"
               ' SaveCasEx.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            'End If
            
             '   With SaveCasEx
              '      !CASExclusionCode = TxtCasExCode
               '     !CASExclusionDescription = TxtCasExDescription
                '    !CASExclusionLastAdjustmentUser = Logon
                 '   !CASExclusionLastAdjustmentDate = Date
                    
                'End With
                'SaveCasEx.Update
                'Call ClearForm(Me)
                'Call CASEListing
                'If EditFlag = False Then
                    'MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                'Else
                 '   MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                'End If
      
            'SaveCasEx.Close
            'Set SaveCasEx = Nothing
        'End If
        'EditFlag = False
    'End If
'End If
'End Sub

'Private Sub DeleteCASE()
'Dim DeleteCASE As New ADODB.Recordset
'Dim strsql As String
'Dim DeleteRecord

'EditFlag = False

'If TxtCasExCode = "" Then
 '  MsgBox "Please Enter Case Exclusion Code", vbOKOnly + vbCritical, DialogBoxTitle
  ' TxtCasExCode.SetFocus
'ElseIf lstCaseExList.ListItems.Count = 0 Then
 '  MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
'Else
 '  DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
  '      If DeleteRecord = vbYes Then
           'strsql = " Delete Form CASExclusion Where CASExclusionCode = '" & ChkStr(lstCaseExList.SelectedItem.Text) & "'"
   '        strsql = "Delete From CASExclusion Where CASExclusionCode = '" & TxtCasExCode & "'"
    '       DeleteCASE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
     '      MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
      '     Call ClearForm(Me)
       '    Call CASEListing
        '   Call setCmdButton(Me, False)
         '  Call EntriesEnable(True)
          ' TxtCasExCode.SetFocus
          ' cmdAdd.Enabled = True
        'End If
'End If
'Exit Sub
'End Sub

'Private Sub SearchCASE()
'Dim strAppend As String
   
 '   strAppend = ""
    
  '  If Len(Trim(TxtCasExCode)) Then
   '     strAppend = strAppend + " And CASExclusionCode like '%" & Mid(ChkStr(TxtCasExCode), 1, 2) & "%'"
    'End If
    'If Len(Trim(TxtCasExDescription)) Then
     '   strAppend = strAppend + " And CASExclusionDescription like '%" & Mid(ChkStr(TxtCasExDescription), 1, 2) & "%'"
    'End If
    
    'CASECriteria = strAppend
    'Call CASEListing(strAppend)
'End Sub

'Private Sub lstCASEExList_Click()
 '   On Error Resume Next
  '  Dim rst3 As New ADODB.Recordset
   ' Dim strsql As String
    
    'Call EnableCmdButton(Me)
    'Call ClearForm(Me)
    'strsql = "Select CASExclusionCode, CASExclusionDescription " & _
    '"from CASExclusion Where CASExclusionCode = '" & lstCaseExList.SelectedItem.Text & "'"
    'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'TxtCasExCode = Trim(rst3!CASExclusionCode)
    'TxtCasExDescription = Trim(rst3!CASExclusionDescription)
    
    'Call EntriesEnable(False)
    'cmdSave.Enabled = False
    'rst3.Close
'End Sub

'Private Sub CASEListing(Optional strAppend As String)
 '   Dim rst2 As New ADODB.Recordset
  '  Dim CASEMaster As ListItem
   ' Dim sql As String
     
    'lstCaseExList.ListItems.Clear
     '   sql = " SELECT CASExclusionCode, CASExclusionDescription " & _
      '  " from CASExclusion where 0=0 "
    
    'If Len(Trim(strAppend)) Then
     '   sql = sql + strAppend
    'End If

    'rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'If rst2.recordCount = 0 Then
     '   MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    'Else
     '   Do Until rst2.EOF
      '      Set CASEMaster = lstCaseExList.ListItems.Add(, , rst2!CASExclusionCode)
       '         CASEMaster.SubItems(1) = rst2!CASExclusionDescription
                'rst2.MoveNext
        'Loop
    'rst2.Close
    'End If
'End Sub
'***************************End Case Exclusion**********************************

'***************************Start UD Exclusion**********************************
'Private Sub SaveUD()
'On Error Resume Next
'Dim RecordSaved
'Dim UD As New ADODB.Recordset
'Dim SaveUD As New ADODB.Recordset
'Dim strsql As String
    
    
    'If TxtUDCode = "" Then
        'MsgBox "Please Enter UD Code", vbOKOnly + vbCritical, DialogBoxTitle
        'TxtUDCode.SetFocus
    'ElseIf TxtUDName = "" Then
        'MsgBox "Please Enter UD Name", vbOKOnly + vbCritical, DialogBoxTitle
        'TxtUDName.SetFocus
    'ElseIf TxtUDDescription = "" Then
        'MsgBox "Please Enter UD Description", vbOKOnly + vbCritical, DialogBoxTitle
        'TxtUDDescription.SetFocus
    'Else
        'strsql = "Select * From UD Where UDCode = '" & Trim(TxtUDCode) & "'"
        'UD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            'If UD.recordCount > 0 And EditFlag = False Then
            'MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            'UD.Close
            'TxtUDCode.SetFocus
            'TxtUDCode.SelStart = 0
            'TxtUDCode.SelLength = Len(TxtUDCode)
            
            'Set UD = Nothing
            'Exit Sub
            'Else
            'RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            'If RecordSaved = vbYes Then
            'If EditFlag = False Then
      
             'SaveUD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             'SaveUD.AddNew
             'Else
                'strsql = "Select * From UD Where UDCode = '" & Trim(TxtUDCode) & "'"
                'SaveUD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            'End If
            
                'With SaveUD
                    '!UDCode = TxtUDCode
                    '!UDName = TxtUDName
                    '!UDDescription = TxtUDDescription
                    '!UDLastUpdateUser = Logon
                    '!UDLastUpdateDate = Date
                    
                'End With
                'SaveUD.Update
                'Call ClearForm(Me)
                'Call UDListing
                'If EditFlag = False Then
                    'MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                'Else
                    'MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                'End If
      
            'SaveUD.Close
            'Set SaveUD = Nothing
        'End If
        'EditFlag = False
    'End If
'End If
'End Sub


'Private Sub DeleteUD()
'Dim DeleteUD As New ADODB.Recordset
'Dim rsCheckCASUD As New ADODB.Recordset
'Dim strsql, sqlCASUD As String
'Dim DeleteRecord

'EditFlag = False
'If TxtUDCode = "" Then
   'MsgBox "Please Enter UD Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtUDCode.SetFocus
'ElseIf lstUDList.ListItems.Count = 0 Then
   'MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
'Else
   'sqlCASUD = "Select CASNumber from CAS_UD where UDCode = '" & ChkStr(lstUDList.SelectedItem.Text) & "'"
   
   
   'rsCheckCASUD.MaxRecords = 1
   'rsCheckCASUD.Open sqlCASUD, medixmasterconn, adOpenKeyset, adLockOptimistic
   
     
   'If rsCheckCASUD.recordCount Then
        'MsgBox "This record cannot be deleted because it is used in CAS Table!", vbOKOnly + vbCritical, "Error Message"
   'Else
   'DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     'If DeleteRecord = vbYes And rsCheckCASUD.recordCount = 0 Then
           'strsql = "Delete From UD Where UDCode = '" & ChkStr(lstUDList.SelectedItem.Text) & "'"
           'DeleteUD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           'MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           'cmdSave.Enabled = True
           'Call ClearForm(Me)
           'Call UDListing
           'Call DisableCmdButton(Me)
           'Call EntriesEnable(True)
           'TxtUDCode.SetFocus
         'End If
     'End If
'End If
'Exit Sub

'End Sub

'Private Sub SearchUD()
    'Dim strAppend As String
   
    'strAppend = ""
    
    'If Len(Trim(TxtUDCode)) Then
        'strAppend = strAppend + " And UDCode LIKE '%" & TxtUDCode & "%'"
    'End If
    'If Len(Trim(TxtUDName)) Then
        'strAppend = strAppend + " And UDName LIKE '%" & TxtUDName & "%'"
    'End If
    'If Len(Trim(TxtUDDescription)) Then
        'strAppend = strAppend + " And UDDescription LIKE '%" & TxtUDDescription & "%'"
    'End If
    
    'UDCriteria = strAppend
    'Call UDListing(strAppend)
    
'End Sub

'Private Sub lstUDList_Click()
    'On Error Resume Next
    'Dim rst3 As New ADODB.Recordset
    'Dim strsql As String
    
    'Call EnableCmdButton(Me)
    'Call ClearForm(Me)
    'strsql = "Select UDCode, UDName, UDDescription " & _
    '"from UD Where UDCode = '" & lstUDList.SelectedItem.Text & "'"
    'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'TxtUDCode = Trim(rst3!UDCode)
    'TxtUDName = Trim(rst3!UDName)
    'TxtUDDescription = Trim(rst3!UDDescription)
       
    'Call EntriesEnable(False)
    'cmdSave.Enabled = False
    'rst3.Close
'End Sub

'Private Sub UDListing(Optional strAppend As String)
    'Dim rst2 As New ADODB.Recordset
    'Dim UDMaster As ListItem
    'Dim sql As String
     
    'lstUDList.ListItems.Clear
        'sql = "SELECT UDCode, UDName, UDDescription From UD where 0=0"
    
    'If Len(Trim(strAppend)) Then
        'sql = sql + strAppend
    'End If

    'rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'If rst2.recordCount = 0 Then
        'MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    'Else
        'Do Until rst2.EOF
            'Set UDMaster = lstUDList.ListItems.Add(, , rst2!UDCode)
                'UDMaster.SubItems(1) = Trim(rst2!UDName)
                'UDMaster.SubItems(2) = Trim(rst2!UDDescription)
                'rst2.MoveNext
        'Loop
    'rst2.Close
    'End If
'End Sub
'***************************End UD Exclusion**********************************

'***************************Start ICD Code**********************************
Private Sub SaveICD()
'On Error Resume Next
Dim RecordSaved
Dim ICD As New ADODB.Recordset
Dim SaveICD As New ADODB.Recordset
Dim strsql As String
    
        
    If TxtICDCode = "" Then
        MsgBox "Please Enter ICD Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtICDCode.SetFocus
    ElseIf CboICDCAT = "" Then
        MsgBox "Please Enter ICD Category Code", vbOKOnly + vbCritical, DialogBoxTitle
        CboICDCAT.SetFocus
    ElseIf TxtICDDescription = "" Then
        MsgBox "Please Enter ICD Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtICDDescription.SetFocus
    ElseIf TxtOtherDesc = "" Then
        MsgBox "Please Enter ICD Other Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtOtherDesc.SetFocus
    Else
        strsql = "Select * From ICD Where ICDCode = '" & Trim(TxtICDCode) & "'"
        ICD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If ICD.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                ICD.Close
                TxtICDCode.SetFocus
                TxtICDCode.SelStart = 0
                TxtICDCode.SelLength = Len(TxtICDCode)
                Set ICD = Nothing
            'Exit Sub
            Else
                RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                    If EditFlag = False Then
      
                        SaveICD.Open "ICD", medixmasterconn, adOpenKeyset, adLockOptimistic
                        SaveICD.AddNew
                    Else
                        strsql = "Select * From ICD Where ICDCode = '" & Trim(TxtICDCode) & "'"
                        SaveICD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    End If
            
                    With SaveICD
                        !ICDCode = ChkStr(TxtICDCode)
                        !ICDCATCode = Trim(Mid(CboICDCAT, 1, 7))
                        !ICDDescription = ChkStr(TxtICDDescription)
                        !ICDOtherDesc = ChkStr(TxtOtherDesc)
                        !ICDLastUpdateUser = Logon
                        !ICDLastUpdateDate = Date
                    End With
                    SaveICD.Update
                    Call ClearForm(Me)
                    Call ICDListing(ICDCriteria)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
      
                    SaveICD.Close
                    Set SaveICD = Nothing
                End If
                EditFlag = False
            End If
    End If
    
End Sub

Private Sub DeleteICD()
Dim rsDeleteICD As New ADODB.Recordset
Dim rsCheckDIAFinal As New ADODB.Recordset
Dim rsCheckDIAFirst As New ADODB.Recordset
Dim strsql, sqlDIAFinal, sqlDIAFirst As String
Dim DeleteRecord
EditFlag = False

If TxtICDCode = "" Then
   MsgBox "Please Enter ICD Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtICDCode.SetFocus
ElseIf lstICDList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else

'sqlICDCAT = "Select ICDCATCode from ICD where ICDCode = '" & ChkStr(lstICDList.SelectedItem.Text) & "'"
sqlDIAFinal = "Select CASNumber from DIAFinal where ICDCode1 = '" & ChkStr(lstICDList.SelectedItem.Text) & "'"
sqlDIAFirst = "Select CASNumber from DIAFirst where ICDCode1 = '" & ChkStr(lstICDList.SelectedItem.Text) & "'"

'rsCheckICDCAT.MaxRecords = 1
'rsCheckICDCAT.Open sqlICDCAT, medixmasterconn, adOpenKeyset, adLockOptimistic
rsCheckDIAFinal.MaxRecords = 1
rsCheckDIAFinal.Open sqlDIAFinal, medixmasterconn, adOpenKeyset, adLockOptimistic

rsCheckDIAFirst.MaxRecords = 1
rsCheckDIAFirst.Open sqlDIAFirst, medixmasterconn, adOpenKeyset, adLockOptimistic


'If rsCheckICDCAT.recordCount Then
    'MsgBox "This record cannot be deleted because it is used in ICD Table!", vbOKOnly + vbCritical, "Error Message"
If rsCheckDIAFinal.recordCount Then
    MsgBox "This record cannot be deleted because it is used in DIAFinal Table!", vbOKOnly + vbCritical, "Error Message"
ElseIf rsCheckDIAFirst.recordCount Then
    MsgBox "This record cannot be deleted because it is used in DIAFirst Table!", vbOKOnly + vbCritical, "Error Message"
Else
    DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes And (rsCheckDIAFinal.recordCount = 0 Or rsCheckDIAFirst.recordCount = 0) Then
        strsql = "Delete from ICD where ICDcode = '" & ChkStr(lstICDList.SelectedItem.Text) & "'"
        rsDeleteICD.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        Call ClearForm(Me)
        Call ICDListing(ICDCriteria)
        Call setCmdButton(Me, False)
        Call EntriesEnable(True)
        TxtICDCode.SetFocus
        cmdAdd.Enabled = True
    End If
  End If
End If

'Exit Sub

End Sub

Private Sub SearchICD()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(TxtICDCode)) Then
        strAppend = strAppend + " And ICDCode like '%" & ChkStr(TxtICDCode) & "%'"
    End If
    
    If Len(Trim(CboICDCAT)) Then
        strAppend = strAppend + " And ICDCATCode like '" & RemStr(Mid(CboICDCAT, 1, IIf(InStr(1, CboICDCAT, "-") = 0, 4, InStr(1, CboICDCAT, "-") - 1))) & "%'"
    End If
    
    If Len(Trim(TxtICDDescription)) Then
        strAppend = strAppend + " And ICDDescription like '%" & ChkStr(TxtICDDescription) & "%'"
    End If
    
    If Len(Trim(TxtOtherDesc)) Then
        strAppend = strAppend + " And ICDOtherDesc like '%" & ChkStr(TxtOtherDesc) & "%'"
    End If
       
    ICDCriteria = strAppend
    Call ICDListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstICDList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
           
    If lstICDList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtICDCode = lstICDList.SelectedItem.Text
        TxtICDDescription = lstICDList.SelectedItem.SubItems(1)
        TxtOtherDesc = lstICDList.SelectedItem.SubItems(2)
        CboICDCAT = lstICDList.SelectedItem.SubItems(3)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstICDList.listitems.count <> 0 Then
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
    
        'strsql = "Select ICDCode, ICDCATCode, ICDDescription, ICDOtherDesc " & _
        '"from ICD Where ICDCode = '" & lstICDList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtICDCode = Trim(rst3!ICDCode)
        'If Len(Trim(Mid(rst3!ICDCATCode, 1, 7))) Then
           'CboICDCAT = Trim(Mid(rst3!ICDCATCode, 1, 7))
        'End If
        'TxtICDDescription = Trim(rst3!ICDDescription)
        'If Len(Trim(rst3!ICDOtherDesc)) Then
            'TxtOtherDesc = Trim(rst3!ICDOtherDesc)
        'End If
            
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub ICDListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim ICDMaster As ListItem
    Dim sql As String
    Dim TotalICDCount As Integer
     
    TotalICDCount = 0
    lstICDList.listitems.Clear
        sql = " SELECT ICDCode, ICDCATCode, ICDDescription, ICDOtherDesc from ICD where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ICDMaster = lstICDList.listitems.Add(, , rst2!ICDCode)
                ICDMaster.SubItems(1) = Trim(rst2!ICDDescription)
                If Len(Trim(rst2!ICDOtherDesc)) Then
                   ICDMaster.SubItems(2) = Trim(rst2!ICDOtherDesc)
                End If
                If Len(Trim(rst2!ICDCATCode)) Then
                   ICDMaster.SubItems(3) = Trim(rst2!ICDCATCode)
                End If
                rst2.MoveNext
        TotalICDCount = TotalICDCount + 1
        txtTotalRecord = TotalICDCount
        Loop
    rst2.Close
    End If
End Sub

'Private Sub PrintICD()
    'RptICDList.Show
'End Sub
'***************************End ICD Code**********************************

'***************************Start ICD Catergory Code**********************************
Private Sub SaveICDCAT()
'On Error Resume Next
Dim RecordSaved
Dim ICDCAT As New ADODB.Recordset
Dim SaveICDCAT As New ADODB.Recordset
Dim strsql As String
    
    
    If TxtICDCATCode = "" Then
        MsgBox "Please Enter ICD Category Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtICDCATCode.SetFocus
    ElseIf TxtICDCATDescription = "" Then
        MsgBox "Please Enter ICD Category Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtICDCATDescription.SetFocus
    Else
        strsql = "Select * From ICDCategory Where ICDCATCode = '" & Trim(TxtICDCATCode) & "'"
        ICDCAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If ICDCAT.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            ICDCAT.Close
            TxtICDCATCode.SetFocus
            TxtICDCATCode.SelStart = 0
            TxtICDCATCode.SelLength = Len(TxtICDCATCode)
            Set ICDCAT = Nothing
            'Exit Sub
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
      
                    SaveICDCAT.Open "ICDCategory", medixmasterconn, adOpenKeyset, adLockOptimistic
                    SaveICDCAT.AddNew
                Else
                    strsql = "Select * From ICDCategory Where ICDCATCode = '" & Trim(TxtICDCATCode) & "'"
                    SaveICDCAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                End If
            
                With SaveICDCAT
                    !ICDCATCode = ChkStr(TxtICDCATCode)
                    !ICDCATDescription = ChkStr(TxtICDCATDescription)
                    !ICDCATLastUpdateUser = Logon
                    !ICDCATLastUpdateDate = Date
                End With
                SaveICDCAT.Update
                Call ClearForm(Me)
                Call ICDCATListing(ICDCATCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
      
                SaveICDCAT.Close
                Set SaveICDCAT = Nothing
            End If
            EditFlag = False
        End If
    End If


End Sub

Private Sub DeleteICDCAT()
Dim rsDeleteICDCAT As New ADODB.Recordset
Dim rsCheckICD As New ADODB.Recordset
Dim strsql, sqlICD As String
Dim DeleteRecord
EditFlag = False


If TxtICDCATCode = "" Then
   MsgBox "Please Enter ICD Category Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtICDCATCode.SetFocus
ElseIf lstICDCATList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else

sqlICD = "Select ICDCatCode from ICD where ICDCode = '" & ChkStr(lstICDCATList.SelectedItem.Text) & "'"

rsCheckICD.MaxRecords = 1
rsCheckICD.Open sqlICD, medixmasterconn, adOpenKeyset, adLockOptimistic


If rsCheckICD.recordCount Then
    MsgBox "This record cannot be deleted because it is used in DiaFinal Table!", vbOKOnly + vbCritical, "Error Message"
Else
    DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes And rsCheckICD.recordCount = 0 Then
        strsql = "Delete from ICDCategory where ICDCatCode = '" & ChkStr(lstICDCATList.SelectedItem.Text) & "'"
        rsDeleteICDCAT.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        Call ClearForm(Me)
        Call ICDCATListing(ICDCATCriteria)
        Call setCmdButton(Me, False)
        Call EntriesEnable(True)
        TxtICDCode.SetFocus
        cmdAdd.Enabled = True
    End If
  End If
End If
'Exit Sub

End Sub

Private Sub SearchICDCAT()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(TxtICDCATCode)) Then
        strAppend = strAppend + " And ICDCATCode like '%" & ChkStr(TxtICDCATCode) & "%'"
    End If
    If Len(Trim(TxtICDCATDescription)) Then
        strAppend = strAppend + " And ICDCATDescription like '%" & ChkStr(TxtICDCATDescription) & "%'"
    End If
       
    ICDCATCriteria = strAppend
    Call ICDCATListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstICDCATList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstICDCATList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtICDCATCode = lstICDCATList.SelectedItem.Text
        TxtICDCATDescription = lstICDCATList.SelectedItem.SubItems(1)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
        
        'If lstICDCATList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
      
        'strsql = "Select ICDCATCode, ICDCATDescription " & _
        '"from ICDCategory Where ICDCATCode = '" & lstICDCATList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtICDCATCode = Trim(rst3!ICDCATCode)
        'TxtICDCATDescription = Trim(rst3!ICDCATDescription)
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstICDCATList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstICDCATList.SortKey = ColumnHeader.Index - 1
    lstICDCATList.Sorted = True
    If lstICDCATList.SortOrder = lvwAscending Then
        lstICDCATList.SortOrder = lvwDescending
    Else
        lstICDCATList.SortOrder = lvwAscending
    End If
End Sub

Private Sub ICDCATListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim ICDCATMaster As ListItem
    Dim sql As String
    Dim TotalICDCATCount As Integer
     
    TotalICDCATCount = 0
    lstICDCATList.listitems.Clear
        sql = " SELECT ICDCATCode, ICDCATDescription from ICDCategory where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ICDCATMaster = lstICDCATList.listitems.Add(, , rst2!ICDCATCode)
                ICDCATMaster.SubItems(1) = Trim(rst2!ICDCATDescription)
                rst2.MoveNext
        TotalICDCATCount = TotalICDCATCount + 1
        txtTotalRecord = TotalICDCATCount
        Loop
    rst2.Close
    End If
End Sub

'Private Sub PrintICDCAT()
    'RptICDCATList.Show
'End Sub
'***************************End ICD Catergory Code**********************************


'***************************Start Repudiation Code**********************************
Private Sub SaveRC()
On Error Resume Next
Dim RecordSaved
Dim RC As New ADODB.Recordset
Dim SaveRC As New ADODB.Recordset
Dim strsql As String
    
        
    If TxtRCCode = "" Then
        MsgBox "Please Enter Repudiation Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtRCCode.SetFocus
    ElseIf TxtRCDescription = "" Then
        MsgBox "Please Enter Repudiation Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtRCDescription.SetFocus
    Else
        strsql = "Select * From RepCode Where RepCode = '" & Trim(TxtRCCode) & "'"
        RC.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If RC.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                RC.Close
                TxtRCCode.SetFocus
                TxtRCCode.SelStart = 0
                TxtRCCode.SelLength = Len(TxtRCCode)
                Set RC = Nothing
            'Exit Sub
            Else
                RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                    If EditFlag = False Then
      
                        SaveRC.Open "RepCode", medixmasterconn, adOpenKeyset, adLockOptimistic
                        SaveRC.AddNew
                    Else
                        strsql = "Select * From RepCode Where RepCode = '" & Trim(TxtRCCode) & "'"
                        SaveRC.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    End If
            
                    With SaveRC
                        !REPCode = ChkStr(TxtRCCode)
                        !RepCatCode = ChkStr(CmbCATCode)
                        !RepDescription = LCase(TxtRCDescription)
                        !RepLastUpdateUser = Logon
                        !RepLastUpdateDate = Date
                    End With
                    SaveRC.Update
                    Call ClearForm(Me)
                    Call RCListing(RCCriteria)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
      
                    SaveRC.Close
                    Set SaveRC = Nothing
                End If
                EditFlag = False
        End If
    End If
    
        
End Sub

Private Sub DeleteRC()
Dim rsDeleteRC As New ADODB.Recordset
Dim rsCheckCAS As New ADODB.Recordset
Dim strsql, sqlCAS As String
Dim DeleteRecord
EditFlag = False


If TxtRCCode = "" Then
   MsgBox "Please Enter Repudiation Code", vbOKOnly + vbCritical, DialogBoxTitle

ElseIf lstRCList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else

sqlCAS = "Select CASNumber from CAS where CASRepCode = '" & Trim(RemStr(TxtRCCode)) & "'"
'ChkStr(lstRCList.SelectedItem.Text) & "'"

rsCheckCAS.MaxRecords = 1
rsCheckCAS.Open sqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic

If rsCheckCAS.recordCount Then
    MsgBox "This record cannot be deleted because it is used in CAS Table!", vbOKOnly + vbCritical, "Error Message"
Else
    DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes And rsCheckCAS.recordCount = 0 Then
        strsql = "Delete from RepCode where RepCode = '" & ChkStr(lstRCList.SelectedItem.SubItems(1)) & "'"
        rsDeleteRC.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        Call ClearForm(Me)
        Call RCListing(RCCriteria)
        Call setCmdButton(Me, False)
        Call EntriesEnable(True)
        TxtRCCode.SetFocus
        cmdAdd.Enabled = True
    End If
  End If
End If

'Exit Sub

End Sub

Private Sub SearchRC()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(CmbCATCode)) Then
        strAppend = strAppend + " AND RepCatCode like '" & RemStr(Mid(CmbCATCode, 1, IIf(InStr(1, CmbCATCode, "-") = 0, 4, InStr(1, CmbCATCode, "-") - 1))) & "%'"
        'strAppend = strAppend + " And RepCatCode = '" & ChkStr(CmbCATCode) & "'"
    End If
    If Len(Trim(TxtRCCode)) Then
        strAppend = strAppend + " And RepCode like '%" & ChkStr(TxtRCCode) & "%'"
    End If
    If Len(Trim(TxtRCDescription)) Then
        strAppend = strAppend + " And RepDescription like '%" & ChkStr(TxtRCDescription) & "%'"
    End If
    
    RCCriteria = strAppend
    Call RCListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstRCList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstRCList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        CmbCATCode = lstRCList.SelectedItem.Text
        TxtRCCode = lstRCList.SelectedItem.SubItems(1)
        TxtRCDescription = lstRCList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstRCList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
            
        'strsql = "Select RepCatCode, RepCode, RepDescription " & _
        '"from RepCode Where RepCode = '" & Trim(lstRCList.SelectedItem.Text) & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'CmbCATCode = lstRCList.SelectedItem.Text
        'TxtRCCode = lstRCList.SelectedItem.SubItems(1)
        'TxtRCDescription = lstRCList.SelectedItem.SubItems(2)
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub RCListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim RCMaster As ListItem
    Dim sql As String
    Dim TotalRCCount As Integer
     
    TotalRCCount = 0
    lstRCList.listitems.Clear
    
    sql = "SELECT * from RepCode where 0=0"
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set RCMaster = lstRCList.listitems.Add(, , rst2!RepCatCode)
                RCMaster.SubItems(1) = Trim(rst2!REPCode)
                RCMaster.SubItems(2) = Trim(rst2!RepDescription)
                rst2.MoveNext
                TotalRCCount = TotalRCCount + 1
                txtTotalRecord = TotalRCCount
        Loop
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

'Private Sub PrintRepCode()
    'RptRepCodeList.Show
'End Sub
'***************************End Repudiation Code**********************************

'***************************Start Repudiation Catergory Code**********************************
Private Sub SaveRCCAT()
On Error Resume Next
Dim RecordSaved
Dim RCCAT As New ADODB.Recordset
Dim SaveRCCAT As New ADODB.Recordset
Dim strsql As String
    
        
    If TxtRepCatCode = "" Then
        MsgBox "Please Enter Repudiation Catergory Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtRepCatCode.SetFocus
    ElseIf TxtRepCatName = "" Then
        MsgBox "Please Enter Repudiation Catergory Name", vbOKOnly + vbCritical, DialogBoxTitle
        TxtRepCatName.SetFocus
    ElseIf TxtRepCatDescription = "" Then
        MsgBox "Please Enter Repudiation Catergory Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtRepCatDescription.SetFocus
    Else
        strsql = "Select * From RepCodeCategory Where RepCatCode = '" & Trim(TxtRepCatCode) & "'"
        RCCAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If RCCAT.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            RCCAT.Close
            TxtRepCatCode.SetFocus
            TxtRepCatCode.SelStart = 0
            TxtRepCatCode.SelLength = Len(TxtRepCatCode)
            Set RCCAT = Nothing
            'Exit Sub
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                    If EditFlag = False Then
      
                        SaveRCCAT.Open "RepCodeCategory", medixmasterconn, adOpenKeyset, adLockOptimistic
                        SaveRCCAT.AddNew
                    Else
                        strsql = "Select * From RepCodeCategory Where RepCatCode = '" & Trim(TxtRepCatCode) & "'"
                        SaveRCCAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    End If
            
                    With SaveRCCAT
                        !RepCatCode = ChkStr(TxtRepCatCode)
                        !RepCatName = ChkStr(TxtRepCatName)
                        !RepCatDescription = LCase(TxtRepCatDescription)
                        !RepCatLastUpdateUser = Logon
                        !RepCatLastUpdateDate = Date
                    End With
                    SaveRCCAT.Update
                    Call ClearForm(Me)
                    Call RCCATListing(RCCATCriteria)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
      
                    SaveRCCAT.Close
                    Set SaveRCCAT = Nothing
                End If
                EditFlag = False
        End If
    End If
    
        
End Sub

Private Sub DeleteRCCAT()
Dim rsDeleteRCCAT As New ADODB.Recordset
Dim rsCheckRepCode As New ADODB.Recordset
Dim strsql, sqlRepCode As String
Dim DeleteRecord
EditFlag = False

If TxtRepCatCode = "" Then
   MsgBox "Please Enter Repudiation Category Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtRepCatCode.SetFocus
ElseIf lstRepCatList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else

sqlRepCode = "Select RepCatCode from RepCode where RepCode = '" & ChkStr(lstRepCatList.SelectedItem.Text) & "'"

rsCheckRepCode.MaxRecords = 1
rsCheckRepCode.Open sqlRepCode, medixmasterconn, adOpenKeyset, adLockOptimistic


If rsCheckRepCode.recordCount Then
    MsgBox "This record cannot be deleted because it is used in RepCode Table!", vbOKOnly + vbCritical, "Error Message"
ElseIf rsCheckRepCode.recordCount Then
    MsgBox "This record cannot be deleted because it is used in DiaFirst Table!", vbOKOnly + vbCritical, "Error Message"
Else
    DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes And rsCheckRepCode.recordCount = 0 Then
        strsql = "Delete from RepCodeCategory where RepCatCode = '" & ChkStr(lstRepCatList.SelectedItem.Text) & "'"
        rsDeleteRCCAT.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        Call ClearForm(Me)
        Call RCCATListing(RCCATCriteria)
        Call setCmdButton(Me, False)
        Call EntriesEnable(True)
        TxtRepCatCode.SetFocus
        cmdAdd.Enabled = True
    End If
  End If
End If
'Exit Sub

End Sub

Private Sub SearchRCCAT()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(TxtRepCatCode)) Then
        strAppend = strAppend + " And RepCatCode like '%" & ChkStr(TxtRepCatCode) & "%'"
    End If
    If Len(Trim(TxtRepCatName)) Then
        strAppend = strAppend + " And RepCatName like '%" & ChkStr(TxtRepCatName) & "%'"
    End If
    If Len(Trim(TxtRepCatDescription)) Then
        strAppend = strAppend + " And RepCatDescription like '%" & ChkStr(TxtRepCatDescription) & "%'"
    End If
    
    RCCATCriteria = strAppend
    Call RCCATListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstRepCatList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstRepCatList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtRepCatCode = lstRepCatList.SelectedItem.Text
        TxtRepCatName = lstRepCatList.SelectedItem.SubItems(1)
        TxtRepCatDescription = lstRepCatList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstRepCatList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
                    
        'strsql = "Select RepCatCode, RepCatName, RepCatDescription " & _
        '"from RepCodeCategory Where RepCatCode = '" & lstRepCatList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtRepCatCode = Trim(rst3!RepCatCode)
        'TxtRepCatName = Trim(rst3!RepCatName)
        'TxtRepCatDescription = Trim(rst3!RepCatDescription)
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstRepCatList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstRepCatList.SortKey = ColumnHeader.Index - 1
    lstRepCatList.Sorted = True
    If lstRepCatList.SortOrder = lvwAscending Then
        lstRepCatList.SortOrder = lvwDescending
    Else
        lstRepCatList.SortOrder = lvwAscending
    End If
End Sub

Private Sub RCCATListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim RCCATMaster As ListItem
    Dim sql As String
    Dim TotalRCCATCount As Integer
     
    TotalRCCATCount = 0
    lstRepCatList.listitems.Clear
        
        sql = " SELECT RepCatCode, RepCatName, RepCatDescription " & _
        "from RepCodeCategory where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set RCCATMaster = lstRepCatList.listitems.Add(, , rst2!RepCatCode)
                If Trim(rst2!RepCatName) <> "" Then
                    RCCATMaster.SubItems(1) = Trim(rst2!RepCatName)
                End If
                If Trim(rst2!RepCatDescription) <> "" Then
                    RCCATMaster.SubItems(2) = Trim(rst2!RepCatDescription)
                End If
                rst2.MoveNext
        TotalRCCATCount = TotalRCCATCount + 1
        txtTotalRecord = TotalRCCATCount
        Loop
    rst2.Close
    End If
End Sub

'Private Sub PrintRepCat()
    'RptRepCatList.Show
'End Sub
'***************************End Repudiation Catergory Code**********************************

'***************************Start Treatment Modality **********************************
Private Sub SaveModality()
On Error Resume Next
Dim RecordSaved
Dim Modality As New ADODB.Recordset
Dim SaveModality As New ADODB.Recordset
Dim strsql As String
    
        
    If TxtModalityCode = "" Then
        MsgBox "Please Enter Modality Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtModalityCode.SetFocus
    ElseIf TxtModalityName = "" Then
        MsgBox "Please Enter Modality Name", vbOKOnly + vbCritical, DialogBoxTitle
        TxtModalityName.SetFocus
    ElseIf TxtModalityDesc = "" Then
        MsgBox "Please Enter Modality Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtModalityDesc.SetFocus
    Else
        strsql = "Select * From TreatmentModality Where ModalityCode = '" & Trim(TxtModalityCode) & "'"
        Modality.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If Modality.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                Modality.Close
                TxtModalityCode.SetFocus
                TxtModalityCode.SelStart = 0
                TxtModalityCode.SelLength = Len(TxtModalityCode)
                Set Modality = Nothing
            'Exit Sub
            Else
                RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                    If EditFlag = False Then
      
                        SaveModality.Open "TreatmentModality", medixmasterconn, adOpenKeyset, adLockOptimistic
                        SaveModality.AddNew
                    Else
                        strsql = "Select * From TreatmentModality Where ModalityCode = '" & Trim(TxtModalityCode) & "'"
                        SaveModality.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    End If
            
                    With SaveModality
                        !ModalityCode = ChkStr(TxtModalityCode)
                        !ModalityName = ChkStr(TxtModalityName)
                        !ModalityDesc = ChkStr(TxtModalityDesc)
                        !ModalityLastUpdateUser = Logon
                        !ModalityLastUpdateDate = Date
                    End With
                    SaveModality.Update
                    Call ClearForm(Me)
                    Call ModalityListing(MODALITYCriteria)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
      
                    SaveModality.Close
                    Set SaveModality = Nothing
                End If
                EditFlag = False
            End If
    End If
    
        
End Sub

Private Sub DeleteModality()
Dim rsDeleteModality As New ADODB.Recordset
Dim rsCheckCAS As New ADODB.Recordset
Dim strsql, sqlCAS As String
Dim DeleteRecord
EditFlag = False


If TxtModalityCode = "" Then
   MsgBox "Please Enter Modality Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtModalityCode.SetFocus
ElseIf lstModalityList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else

sqlCAS = "Select CASNumber from CAS where ModalityCode = '" & ChkStr(lstModalityList.SelectedItem.Text) & "'"

rsCheckCAS.MaxRecords = 1
rsCheckCAS.Open sqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic

If rsCheckCAS.recordCount Then
    MsgBox "This record cannot be deleted because it is used in CAS Table!", vbOKOnly + vbCritical, "Error Message"
Else
    DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes And rsCheckCAS.recordCount = 0 Then
        strsql = "Delete from TreatmentModality where ModalityCode = '" & ChkStr(lstModalityList.SelectedItem.Text) & "'"
        rsDeleteModality.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        Call ClearForm(Me)
        Call ModalityListing  '(ICDCriteria)
        Call setCmdButton(Me, False)
        Call EntriesEnable(True)
        TxtModalityCode.SetFocus
        cmdAdd.Enabled = True
    End If
  End If
End If
'Exit Sub

End Sub

Private Sub SearchModality()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(TxtModalityCode)) Then
        strAppend = strAppend + " And ModalityCode like '%" & ChkStr(TxtModalityCode) & "%'"
    End If
    If Len(Trim(TxtModalityName)) Then
        strAppend = strAppend + " And ModalityName like '%" & ChkStr(TxtModalityName) & "%'"
    End If
    If Len(Trim(TxtModalityDesc)) Then
        strAppend = strAppend + " And ModalityDesc like '%" & ChkStr(TxtModalityDesc) & "%'"
    End If
    
    MODALITYCriteria = strAppend
    Call ModalityListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstModalityList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstModalityList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtModalityCode = lstModalityList.SelectedItem.Text
        TxtModalityName = lstModalityList.SelectedItem.SubItems(1)
        TxtModalityDesc = lstModalityList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstModalityList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
            
        'strsql = "Select ModalityCode, ModalityName, ModalityDesc " & _
        '"from TreatmentModality Where ModalityCode = '" & lstModalityList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtModalityCode = Trim(rst3!ModalityCode)
        'TxtModalityName = Trim(rst3!ModalityName)
        'If Len(Trim(rst3!ModalityDesc)) Then
           'TxtModalityDesc = Trim(rst3!ModalityDesc)
        'End If
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstModalityList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstModalityList.SortKey = ColumnHeader.Index - 1
    lstModalityList.Sorted = True
    If lstModalityList.SortOrder = lvwAscending Then
        lstModalityList.SortOrder = lvwDescending
    Else
        lstModalityList.SortOrder = lvwAscending
    End If
End Sub

Private Sub ModalityListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim ModalityMaster As ListItem
    Dim sql As String
    Dim TotalModalityCount As Integer
    
    TotalModalityCount = 0
    lstModalityList.listitems.Clear
        sql = " SELECT ModalityCode, ModalityName, ModalityDesc " & _
        "from TreatmentModality where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ModalityMaster = lstModalityList.listitems.Add(, , rst2!ModalityCode)
                ModalityMaster.SubItems(1) = Trim(rst2!ModalityName)
                If Len(Trim(rst2!ModalityDesc)) Then
                    ModalityMaster.SubItems(2) = Trim(rst2!ModalityDesc)
                End If
                rst2.MoveNext
        TotalModalityCount = TotalModalityCount + 1
        txtTotalRecord = TotalModalityCount
        Loop
    rst2.Close
    End If
End Sub

'Private Sub PrintTreatment()
    'RptTreatmentList.Show
'End Sub
'***************************End Treatment Modality **********************************

'***************************Start Claim Nature**********************************
Private Sub SaveClaimNature()
'On Error Resume Next
Dim RecordSaved
Dim ClaimNature As New ADODB.Recordset
Dim SaveClaimNature As New ADODB.Recordset
Dim strsql As String
    
    
    If TxtClaimNatureCode = "" Then
        MsgBox "Please Enter Claim Nature Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtClaimNatureCode.SetFocus
    ElseIf TxtClaimNatureName = "" Then
        MsgBox "Please Enter Claim Nature Name", vbOKOnly + vbCritical, DialogBoxTitle
        TxtClaimNatureName.SetFocus
    Else
        strsql = "Select * From ClaimNature Where ClaimNatureCode = '" & Trim(TxtClaimNatureCode) & "'"
        ClaimNature.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If ClaimNature.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                ClaimNature.Close
                TxtClaimNatureCode.SetFocus
                TxtClaimNatureCode.SelStart = 0
                TxtClaimNatureCode.SelLength = Len(TxtClaimNatureCode)
                Set ClaimNature = Nothing
            'Exit Sub
            Else
                RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                    If EditFlag = False Then
      
                        SaveClaimNature.Open "ClaimNature", medixmasterconn, adOpenKeyset, adLockOptimistic
                        SaveClaimNature.AddNew
                    Else
                        strsql = "Select * From ClaimNature Where ClaimNatureCode = '" & Trim(TxtClaimNatureCode) & "'"
                        SaveClaimNature.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    End If
            
                    With SaveClaimNature
                        !ClaimNatureCode = ChkStr(TxtClaimNatureCode)
                        !ClaimNatureName = ChkStr(TxtClaimNatureName)
                        !ClaimNatureLastUpdateUser = Logon
                        !ClaimNatureLastUpdateDate = Date
                    End With
                    SaveClaimNature.Update
                    Call ClearForm(Me)
                    Call ClaimNatureListing(CLAIMNATURECriteria)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
      
                    SaveClaimNature.Close
                    Set SaveClaimNature = Nothing
                End If
                EditFlag = False
            End If
    End If

End Sub

Private Sub DeleteClaimNature()
Dim rsDeleteClaimNature As New ADODB.Recordset
Dim rsCheckCAS As New ADODB.Recordset
Dim strsql, sqlCAS As String
Dim DeleteRecord

EditFlag = False

If TxtClaimNatureCode = "" Then
   MsgBox "Please Enter Claim Nature Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtClaimNatureCode.SetFocus
ElseIf lstClaimNatureList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else

sqlCAS = "Select CASNumber from CAS where ClaimNatureCode = '" & ChkStr(lstClaimNatureList.SelectedItem.Text) & "'"

rsCheckCAS.MaxRecords = 1
rsCheckCAS.Open sqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic


If rsCheckCAS.recordCount Then
    MsgBox "This record cannot be deleted because it is used in CAS Table!", vbOKOnly + vbCritical, "Error Message"
Else
    DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes And rsCheckCAS.recordCount = 0 Then
        strsql = "Delete from ClaimNature where ClaimNaturecode = '" & ChkStr(lstClaimNatureList.SelectedItem.Text) & "'"
        rsDeleteClaimNature.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        Call ClearForm(Me)
        Call ClaimNatureListing '(ClaimNatureCriteria)
        Call setCmdButton(Me, False)
        Call EntriesEnable(True)
        TxtClaimNatureCode.SetFocus
        cmdAdd.Enabled = True
    End If
  End If
End If
'Exit Sub

End Sub

Private Sub SearchClaimNature()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(TxtClaimNatureCode)) Then
        strAppend = strAppend + " And ClaimNatureCode like '%" & ChkStr(TxtClaimNatureCode) & "%'"
    End If
    If Len(Trim(TxtClaimNatureName)) Then
        strAppend = strAppend + " And ClaimNatureName like '%" & ChkStr(TxtClaimNatureName) & "%'"
    End If
       
    CLAIMNATURECriteria = strAppend
    Call ClaimNatureListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstClaimNatureList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
        
    If lstClaimNatureList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtClaimNatureCode = lstClaimNatureList.SelectedItem.Text
        TxtClaimNatureName = lstClaimNatureList.SelectedItem.SubItems(1)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstClaimNatureList.listitems.count <> 0 Then
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
    
        'strsql = "Select ClaimNatureCode, ClaimNatureName " & _
        '"from ClaimNature Where ClaimNatureCode = '" & lstClaimNatureList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        'TxtClaimNatureCode = Trim(rst3!ClaimNatureCode)
        'TxtClaimNatureName = Trim(rst3!ClaimNatureName)
        
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstClaimNatureList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstClaimNatureList.SortKey = ColumnHeader.Index - 1
    lstClaimNatureList.Sorted = True
    If lstClaimNatureList.SortOrder = lvwAscending Then
        lstClaimNatureList.SortOrder = lvwDescending
    Else
        lstClaimNatureList.SortOrder = lvwAscending
    End If
End Sub

Private Sub ClaimNatureListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim ClaimNatureMaster As ListItem
    Dim sql As String
    Dim TotalClaimCount As Integer
    
    TotalClaimCount = 0
    lstClaimNatureList.listitems.Clear
        sql = " SELECT ClaimNatureCode, ClaimNatureName from ClaimNature where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set ClaimNatureMaster = lstClaimNatureList.listitems.Add(, , rst2!ClaimNatureCode)
                ClaimNatureMaster.SubItems(1) = Trim(rst2!ClaimNatureName)
                rst2.MoveNext
        TotalClaimCount = TotalClaimCount + 1
        txtTotalRecord = TotalClaimCount
        Loop
    rst2.Close
    End If
End Sub

'Private Sub PrintNature()
    'RptNatureList.Show
'End Sub
'***************************End Claim Nature**********************************

'***************************Start Payment Mode**********************************
'Private Sub SavePayMode()
'Dim RecordSaved
'Dim PayMode As New ADODB.Recordset
'Dim SavePayMode As New ADODB.Recordset
'Dim strsql As String
    
    
    'If TxtPayModeCode = "" Then
        'MsgBox "Please Enter Payment Mode Code", vbOKOnly + vbCritical, DialogBoxTitle
        'TxtPayModeCode.SetFocus
    'ElseIf TxtPayModeName = "" Then
        'MsgBox "Please Enter Payment Mode Name", vbOKOnly + vbCritical, DialogBoxTitle
        'TxtPayModeName.SetFocus
    'ElseIf TxtPayModeDesc = "" Then
        'MsgBox "Please Enter Payment Mode Description", vbOKOnly + vbCritical, DialogBoxTitle
        'TxtPayModeDesc.SetFocus
    'Else
        'strsql = "Select * From PayMode Where PayModeCode = '" & Trim(TxtPayModeCode) & "'"
        'PayMode.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            'If PayMode.recordCount > 0 And EditFlag = False Then
            'MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            'PayMode.Close
            'TxtPayModeCode.SetFocus
            'TxtPayModeCode.SelStart = 0
            'TxtPayModeCode.SelLength = Len(TxtPayModeCode)
            'Set PayMode = Nothing
            'Exit Sub
            'Else
            'RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            'If RecordSaved = vbYes Then
            'If EditFlag = False Then
      
             'SavePayMode.Open "PayMode", medixmasterconn, adOpenKeyset, adLockOptimistic
             'SavePayMode.AddNew
             'Else
                'strsql = "Select * From PayMode Where PayModeCode = '" & Trim(TxtPayModeCode) & "'"
                'SavePayMode.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            'End If
            
                'With SavePayMode
                    '!PayModeCode = TxtPayModeCode
                    '!PayModeName = TxtPayModeName
                    '!PayModeDesc = TxtPayModeDesc
                    '!PayModeLastUpdateUser = Logon
                    '!PayModeLastUpdateDate = Date
                'End With
                'SavePayMode.Update
                'Call ClearForm(Me)
                'Call PayModeListing(PAYMODECriteria)
                'If EditFlag = False Then
                    'MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                'Else
                    'MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                'End If
      
            'SavePayMode.Close
            'Set SavePayMode = Nothing
        'End If
        'EditFlag = False
    'End If
'End If
'End Sub

'Private Sub DeletePayMode()
'Dim rsDeletePayMode As New ADODB.Recordset
'Dim rsCheckCAS As New ADODB.Recordset
'Dim strsql, sqlCAS As String
'Dim DeleteRecord
'EditFlag = False

'If TxtPayModeCode = "" Then
   'MsgBox "Please Enter Payment Mode Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtPayModeCode.SetFocus
'ElseIf lstPayModeList.listitems.Count = 0 Then
   'MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
'Else

'sqlCAS = "Select CASNumber from CAS where PayModeCode = '" & ChkStr(lstPayModeList.SelectedItem.Text) & "'"

'rsCheckCAS.MaxRecords = 1
'rsCheckCAS.Open sqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic


'If rsCheckCAS.recordCount Then
    'MsgBox "This record cannot be deleted because it is used in CAS Table!", vbOKOnly + vbCritical, "Error Message"
'Else
    'DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    'If DeleteRecord = vbYes And rsCheckCAS.recordCount = 0 Then
        'strsql = "Delete from PayMode where PayModeCode = '" & ChkStr(lstPayModeList.SelectedItem.Text) & "'"
        'rsDeletePayMode.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic
        'MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        'Call ClearForm(Me)
        'Call PayModeListing '(ClaimNatureCriteria)
        'Call setCmdButton(Me, False)
        'Call EntriesEnable(True)
        'TxtPayModeCode.SetFocus
        'cmdAdd.Enabled = True
    'End If
  'End If
'End If
'Exit Sub

'End Sub

'Private Sub SearchPayMode()
'Dim strAppend As String
   
    'strAppend = ""
    
    'If Len(Trim(TxtPayModeCode)) Then
        'strAppend = strAppend + " And PayModeCode like '%" & ChkStr(TxtPayModeCode) & "%'"
    'End If
    'If Len(Trim(TxtPayModeName)) Then
        'strAppend = strAppend + " And PayModeName like '%" & ChkStr(TxtPayModeName) & "%'"
    'End If
    
    'If Len(Trim(TxtPayModeDesc)) Then
        'strAppend = strAppend + " And PayModeDesc like '%" & ChkStr(TxtPayModeDesc) & "%'"
    'End If
       
    'PAYMODECriteria = strAppend
    'Call PayModeListing(strAppend)
'End Sub

'Private Sub lstPayModeList_Click()
    'On Error Resume Next
    'Dim rst3 As New ADODB.Recordset
    'Dim strsql As String
    
    'Call EnableCmdButton(Me)
    'Call ClearForm(Me)
    'strsql = "Select PayModeCode, PayModeName, PayMOdeDesc " & _
    '"from PayMode Where PayModeCode = '" & lstPayModeList.SelectedItem.Text & "'"
    'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'TxtPayModeCode = Trim(rst3!PayModeCode)
    'TxtPayModeName = Trim(rst3!PayModeName)
    'TxtPayModeDesc = Trim(rst3!PayModeDesc)
        
    'Call EntriesEnable(False)
    'cmdSave.Enabled = False
    'rst3.Close
'End Sub

'Private Sub lstPayModeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'lstPayModeList.SortKey = ColumnHeader.Index - 1
    'lstPayModeList.Sorted = True
    'If lstPayModeList.SortOrder = lvwAscending Then
        'lstPayModeList.SortOrder = lvwDescending
    'Else
        'lstPayModeList.SortOrder = lvwAscending
    'End If
'End Sub

'Private Sub PayModeListing(Optional strAppend As String)
    'Dim rst2 As New ADODB.Recordset
    'Dim PayModeMaster As ListItem
    'Dim sql As String
    'Dim TotalPayModeCount As Integer
     
    'TotalPayModeCount = 0
    'lstPayModeList.listitems.Clear
        'sql = " SELECT PayModeCode, PayModeName, PayModeDesc from PayMode where 0=0 "
    
    'If Len(Trim(strAppend)) Then
        'sql = sql + strAppend
    'End If

    'rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'If rst2.recordCount = 0 Then
        'MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    'Else
        'Do Until rst2.EOF
            'Set PayModeMaster = lstPayModeList.listitems.Add(, , rst2!PayModeCode)
                'PayModeMaster.SubItems(1) = Trim(rst2!PayModeName)
                'If Len(Trim(rst2!PayModeDesc)) Then
                    'PayModeMaster.SubItems(2) = Trim(rst2!PayModeDesc)
                'End If
                'rst2.MoveNext
        'TotalPayModeCount = TotalPayModeCount + 1
        'txtTotalRecord = TotalPayModeCount
        'Loop
    'rst2.Close
    'End If
'End Sub

'Private Sub PrintPayMode()
    'RptPayModeList.show
'End Sub
'***************************End Payment Mode**********************************

'***************************Start Payment Type**********************************
Private Sub SavePayType()
'On Error Resume Next
Dim RecordSaved
Dim PayType As New ADODB.Recordset
Dim SavePayType As New ADODB.Recordset
Dim strsql As String
    
        
    If TxtPayTypeCode = "" Then
        MsgBox "Please Enter Payment Type Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtPayTypeCode.SetFocus
    ElseIf TxtPayTypeName = "" Then
        MsgBox "Please Enter Payment Type Name", vbOKOnly + vbCritical, DialogBoxTitle
        TxtPayTypeName.SetFocus
    ElseIf TxtPayTypeDesc = "" Then
        MsgBox "Please Enter Payment Type Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtPayTypeDesc.SetFocus
    Else
        strsql = "Select * From PayType Where PayTypeCode = '" & Trim(TxtPayTypeCode) & "'"
        PayType.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If PayType.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                PayType.Close
                TxtPayTypeCode.SetFocus
                TxtPayTypeCode.SelStart = 0
                TxtPayTypeCode.SelLength = Len(TxtPayTypeCode)
                Set PayType = Nothing
            'Exit Sub
            Else
                RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
                If RecordSaved = vbYes Then
                    If EditFlag = False Then
      
                        SavePayType.Open "PayType", medixmasterconn, adOpenKeyset, adLockOptimistic
                        SavePayType.AddNew
                    Else
                        strsql = "Select * From PayType Where PayTypeCode = '" & Trim(TxtPayTypeCode) & "'"
                        SavePayType.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    End If
            
                    With SavePayType
                        !PayTypeCode = ChkStr(TxtPayTypeCode)
                        !PayTypeName = ChkStr(TxtPayTypeName)
                        !PayTypeDesc = ChkStr(TxtPayTypeDesc)
                        !PayTypeLastUpdateUser = Logon
                        !PayTypeLastUpdateDate = Date
                    End With
                    SavePayType.Update
                    Call ClearForm(Me)
                    Call PayTypeListing(PAYTYPECriteria)
                    If EditFlag = False Then
                        MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                    Else
                        MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
      
                    SavePayType.Close
                    Set SavePayType = Nothing
                End If
                EditFlag = False
            End If
    End If
    
   
End Sub

Private Sub DeletePayType()
Dim rsDeletePayType As New ADODB.Recordset
Dim rsCheckCAS As New ADODB.Recordset
Dim strsql, sqlCAS As String
Dim DeleteRecord

EditFlag = False

If TxtPayTypeCode = "" Then
   MsgBox "Please Enter Payment Type Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtPayTypeCode.SetFocus
ElseIf lstPayTypeList.listitems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else

sqlCAS = "Select CASNumber from CAS where PayTypeCode = '" & ChkStr(lstPayTypeList.SelectedItem.Text) & "'"

rsCheckCAS.MaxRecords = 1
rsCheckCAS.Open sqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic


If rsCheckCAS.recordCount Then
    MsgBox "This record cannot be deleted because it is used in CAS Table!", vbOKOnly + vbCritical, "Error Message"
Else
    DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
    If DeleteRecord = vbYes And rsCheckCAS.recordCount = 0 Then
        strsql = "Delete from PayType where PayTypeCode = '" & ChkStr(lstPayTypeList.SelectedItem.Text) & "'"
        rsDeletePayType.Open strsql, medixmasterconn, adOpenDynamic, adLockOptimistic
        MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
        Call ClearForm(Me)
        Call PayTypeListing '(PAYTYPECriteria)
        Call setCmdButton(Me, False)
        Call EntriesEnable(True)
        TxtPayTypeCode.SetFocus
        cmdAdd.Enabled = True
    End If
  End If
End If
'Exit Sub


End Sub

Private Sub SearchPayType()
Dim strAppend As String
   
    Screen.MousePointer = vbHourglass
    strAppend = ""
    
    If Len(Trim(TxtPayTypeCode)) Then
        strAppend = strAppend + " And PayTypeCode like '%" & Mid(ChkStr(TxtPayTypeCode), 1, 2) & "%'"
    End If
    If Len(Trim(TxtPayTypeName)) Then
        strAppend = strAppend + " And PayTypeName like '%" & ChkStr(TxtPayTypeName) & "%'"
    End If
    If Len(Trim(TxtPayTypeDesc)) Then
        strAppend = strAppend + " And PayTypeDesc like '%" & ChkStr(TxtPayTypeDesc) & "%'"
    End If
       
    PAYTYPECriteria = strAppend
    Call PayTypeListing(strAppend)
    Screen.MousePointer = Default
End Sub

Private Sub lstPayTypeList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    
    If lstPayTypeList.listitems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtPayTypeCode = lstPayTypeList.SelectedItem.Text
        TxtPayTypeName = lstPayTypeList.SelectedItem.SubItems(1)
        TxtPayTypeDesc = lstPayTypeList.SelectedItem.SubItems(2)
        Call EntriesEnable(False)
        cmdSave.Enabled = False
    End If
    
    'If lstPayTypeList.listitems.count <> 0 Then
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
            
        'strsql = "Select PayTypeCode, PayTypeName, PayTypeDesc " & _
        '"from PayType Where PayTypeCode = '" & lstPayTypeList.SelectedItem.Text & "'"
        'rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'TxtPayTypeCode = Trim(rst3!PayTypeCode)
        'TxtPayTypeName = Trim(rst3!PayTypeName)
        'TxtPayTypeDesc = Trim(rst3!PayTypeDesc)
            
        'Call EntriesEnable(False)
        'cmdSave.Enabled = False
        'rst3.Close
        'Set rst3 = Nothing
    'End If
End Sub

Private Sub lstPayTypeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstPayTypeList.SortKey = ColumnHeader.Index - 1
    lstPayTypeList.Sorted = True
    If lstPayTypeList.SortOrder = lvwAscending Then
        lstPayTypeList.SortOrder = lvwDescending
    Else
        lstPayTypeList.SortOrder = lvwAscending
    End If
End Sub

Private Sub PayTypeListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim PayTypeMaster As ListItem
    Dim sql As String
    Dim TotalPayTypeCount As Integer
     
    TotalPayTypeCount = 0
    lstPayTypeList.listitems.Clear
        sql = " SELECT PayTypeCode, PayTypeName, PayTypeDesc from PayType where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set PayTypeMaster = lstPayTypeList.listitems.Add(, , rst2!PayTypeCode)
                PayTypeMaster.SubItems(1) = Trim(rst2!PayTypeName)
                If Len(Trim(rst2!PayTypeDesc)) Then
                    PayTypeMaster.SubItems(2) = Trim(rst2!PayTypeDesc)
                End If
                rst2.MoveNext
        TotalPayTypeCount = TotalPayTypeCount + 1
        txtTotalRecord = TotalPayTypeCount
        Loop
    rst2.Close
    End If
End Sub

'Private Sub PrintPayType()
    'RptPayTypeList.Show
'End Sub
'***************************End Payment Mode**********************************

'***************************Start Call ListOriginal & ListDestination**********************************

Private Sub CmdAdd2_Click()
    Screen.MousePointer = vbHourglass
    Call AddRecord
    Call ICDListView2
    lstICDCodeList.listitems.Clear
    lstOriginalICDList.listitems.Clear
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Call ListDestination
    Call ListOriginal
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   '
    Screen.MousePointer = Default
End Sub

Private Sub cmdRemove_Click()
    Screen.MousePointer = vbHourglass
    Call RemoveRecord
    lstICDCodeList.listitems.Clear
    
    lstOriginalICDList.listitems.Clear
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Call ListDestination
    Call ListOriginal
    
   ' medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
    Screen.MousePointer = Default
End Sub

Private Sub ListOriginal()
    Dim ListOriginal As New ADODB.Recordset
    Dim ICDOriginal As ListItem
    Dim sql As String

        sql = " select ICDCode, ICDDescription , ICDCode2 " & _
          " from View_ICD_tracking "
         
        ListOriginal.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
        Do Until ListOriginal.EOF
            Set ICDOriginal = lstOriginalICDList.listitems.Add(, , ListOriginal!ICDCode)
                ICDOriginal.SubItems(1) = Trim(ListOriginal!ICDDescription)
                If Trim(ListOriginal!ICDCode2) <> "" Then
                    ICDOriginal.Bold = True
                Else
                    ICDOriginal.Bold = False
                End If
                
                ListOriginal.MoveNext
        Loop
                     
        ListOriginal.Close
        Set ListOriginal = Nothing
    
End Sub

Private Sub ListDestination(Optional header As String)
Dim sql1 As String
Dim ListDestination As New ADODB.Recordset
Dim ICDTest As ListItem
    
        
    sql1 = " select ICDCode, ICDDescription, " & _
           " ICDLastUpdateUser, ICDLastUpdateDate" & _
           " from ICDOriginal "

    ListDestination.Open sql1, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
  
    Do Until ListDestination.EOF
        Set ICDTest = lstICDCodeList.listitems.Add(, , ListDestination!ICDCode)
            ICDTest.SubItems(1) = Trim(ListDestination!ICDDescription)
            ListDestination.MoveNext
    Loop
        
    ListDestination.Close
    Set ListDestination = Nothing
        
End Sub

Private Sub lstICDCodeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    lstICDCodeList.SortKey = ColumnHeader.Index - 1
    lstICDCodeList.Sorted = True
        If lstICDCodeList.SortOrder = lvwAscending Then
            lstICDCodeList.SortOrder = lvwDescending
        Else
            lstICDCodeList.SortOrder = lvwAscending
        End If
End Sub

Private Sub lstOriginalICDList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    lstOriginalICDList.SortKey = ColumnHeader.Index - 1
    lstOriginalICDList.Sorted = True
        If lstOriginalICDList.SortOrder = lvwAscending Then
            lstOriginalICDList.SortOrder = lvwDescending
        Else
            lstOriginalICDList.SortOrder = lvwAscending
        End If
End Sub

Private Sub RemoveRecord()
Dim ii As Integer
Dim D_tmpICDCode, D_tmpICDDescription As String
Dim D_ttmpICDCode, D_ttmpICDDescription As String
Dim D_ICDCodeArray() As String
Dim D_ICDDescriptionArray() As String
Dim cnt As Integer

cnt = 0
 
    If lstICDCodeList.SelectedItem Is Nothing Then Exit Sub
       For ii = 1 To lstICDCodeList.listitems.Count
        
           If lstICDCodeList.listitems(ii).Selected = True Then
             Set lstICDCodeList.SelectedItem = lstICDCodeList.listitems(ii)
                D_tmpICDCode = lstICDCodeList.SelectedItem
                D_tmpICDDescription = lstICDCodeList.SelectedItem.SubItems(1)
            
                    If D_tmpICDCode = D_ttmpICDCode And _
                        D_tmpICDDescription = D_ttmpICDDescription Then
                
                    Else
                         D_ttmpICDCode = D_ttmpICDCode
                         D_ttmpICDDescription = D_ttmpICDDescription
                                 
                         cnt = cnt + 1
                         ReDim Preserve D_ICDCodeArray(cnt)
                         ReDim Preserve D_ICDDescriptionArray(cnt)
            
                         D_ICDCodeArray(cnt) = D_tmpICDCode
                         D_ICDDescriptionArray(cnt) = D_tmpICDDescription
                 End If
           End If
        Next ii
    
    For ii = 1 To UBound(D_ICDCodeArray)
      RemoveICDFile D_ICDCodeArray(ii), D_ICDDescriptionArray(ii)
    Next ii

End Sub

Private Sub AddRecord()
Dim ii As Integer
Dim tmpICDCode, tmpICDDescription As String
Dim ttmpICDCode, ttmpICDDescription As String
Dim ICDCodeArray() As String
Dim ICDDescriptionArray() As String
Dim cnt As Integer

cnt = 0
    
    If lstOriginalICDList.SelectedItem Is Nothing Then Exit Sub
       For ii = 1 To lstOriginalICDList.listitems.Count
        
           If lstOriginalICDList.listitems(ii).Selected = True Then
                
             Set lstOriginalICDList.SelectedItem = lstOriginalICDList.listitems(ii)

                tmpICDCode = lstOriginalICDList.SelectedItem
                tmpICDDescription = lstOriginalICDList.SelectedItem.SubItems(1)
            
                    If tmpICDCode = ttmpICDCode And _
                        tmpICDDescription = ttmpICDDescription Then
                
                        Else
                            ttmpICDCode = ttmpICDCode
                            ttmpICDDescription = ttmpICDDescription
                                    
                            cnt = cnt + 1
                            ReDim Preserve ICDCodeArray(cnt)
                            ReDim Preserve ICDDescriptionArray(cnt)
               
                            ICDCodeArray(cnt) = tmpICDCode
                            ICDDescriptionArray(cnt) = tmpICDDescription
                    End If
           End If
        Next ii
    
    For ii = 1 To UBound(ICDCodeArray)
      UpdateICDFile ICDCodeArray(ii), ICDDescriptionArray(ii)

    Next ii
End Sub

Private Sub UpdateICDFile(tICDCode As String, tICDDescription As String)
On Error Resume Next
Dim rsUpdate As New ADODB.Recordset
Dim sql As String
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    sql = "SELECT * FROM ICDOriginal"
         
    rsUpdate.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
    rsUpdate.AddNew
        With rsUpdate
           !ICDCode = ChkStr(tICDCode)
           !ICDDescription = ChkStr(tICDDescription)
           !ICDLastUpdateUser = Logon
           !ICDLastUpdateDate = Date
        End With
        rsUpdate.Update
        rsUpdate.Close
        Set rsUpdate = Nothing
        
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing

End Sub

Private Sub RemoveICDFile(D_tICDCode As String, D_tICDDescription As String)
On Error Resume Next
Dim rsRemove As New ADODB.Recordset
Dim sql As String
  
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    sql = "SELECT * FROM ICDOriginal"
    sql = sql & " WHERE ICDCode = '" & D_tICDCode & "'"
        
    rsRemove.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
    rsRemove.Delete
    rsRemove.UpdateBatch
    rsRemove.Close
    Set rsRemove = Nothing
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End Sub

Private Sub ICDListView2()
  
 Dim X As Object
 Dim i As Integer
 Dim listitems As ListItem
 If lstOriginalICDList.SelectedItem Is Nothing Then Exit Sub
    For i = 1 To lstOriginalICDList.listitems.Count
        If lstOriginalICDList.listitems(i).Selected = True Then
           Set listitems = lstICDCodeList.listitems.Add(, , lstOriginalICDList.listitems(i).Text)
               listitems.SubItems(1) = lstOriginalICDList.listitems(i).SubItems(1)
               listitems.SubItems(2) = ""
        End If
    Next i

lstICDCodeList.listitems.Clear
Call ListDestination
End Sub
'***************************End Call ListOriginal & ListDestination**********************************
