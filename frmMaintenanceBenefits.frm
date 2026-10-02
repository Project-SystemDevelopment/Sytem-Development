VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMaintenanceBenefits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Benefits Maintenance "
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      Height          =   705
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Width           =   11655
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   5040
         TabIndex        =   16
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9960
         TabIndex        =   19
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   9120
         TabIndex        =   18
         Top             =   270
         Width           =   840
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   13
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   4200
         TabIndex        =   15
         Top             =   285
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   8280
         TabIndex        =   17
         Top             =   270
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Category"
      TabPicture(0)   =   "frmMaintenanceBenefits.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label38"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstCATList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Sub Category"
      TabPicture(1)   =   "frmMaintenanceBenefits.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(0)"
      Tab(1).Control(1)=   "lstSubCatList"
      Tab(1).Control(2)=   "Label5"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Benefit Entry"
      TabPicture(2)   =   "frmMaintenanceBenefits.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FGrid"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Benefits Listing"
      TabPicture(3)   =   "frmMaintenanceBenefits.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(1)=   "Frame1(4)"
      Tab(3).Control(2)=   "lstBNFList"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   -74640
         TabIndex        =   64
         Top             =   2400
         Width           =   11055
         Begin VB.ComboBox CboPayorCode 
            Height          =   315
            ItemData        =   "frmMaintenanceBenefits.frx":0070
            Left            =   1920
            List            =   "frmMaintenanceBenefits.frx":0072
            TabIndex        =   9
            Top             =   240
            Width           =   3615
         End
         Begin VB.ComboBox CboGrpCom 
            Height          =   315
            ItemData        =   "frmMaintenanceBenefits.frx":0074
            Left            =   7680
            List            =   "frmMaintenanceBenefits.frx":0076
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label16 
            Caption         =   "Payor Code    "
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
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Group Company  "
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
            Left            =   6000
            TabIndex        =   65
            Top             =   240
            Width           =   1815
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FGrid 
         Height          =   4575
         Left            =   -74520
         TabIndex        =   58
         Top             =   1680
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8070
         _Version        =   393216
         Rows            =   40
         Cols            =   5
         FixedCols       =   2
         AllowUserResizing=   3
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   -74640
         TabIndex        =   51
         Top             =   480
         Width           =   10575
         Begin VB.ComboBox cboTInsCode 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   60
            Top             =   600
            Width           =   3615
         End
         Begin MSMask.MaskEdBox txtTEffdate 
            Height          =   315
            Left            =   7200
            TabIndex        =   59
            Top             =   600
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboTPlnCode 
            Height          =   315
            Left            =   7200
            TabIndex        =   53
            Top             =   240
            Width           =   3105
         End
         Begin VB.ComboBox cboTHLTCode 
            Height          =   315
            Left            =   1800
            TabIndex        =   52
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Label9 
            Caption         =   "Eff Date       *"
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
            Left            =   5760
            TabIndex        =   61
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label8 
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
            Left            =   135
            TabIndex        =   56
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label10 
            Caption         =   "Plan Code     *"
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
            Left            =   5760
            TabIndex        =   55
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Insured Type *"
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
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   1485
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
         Height          =   1860
         Index           =   4
         Left            =   -74640
         TabIndex        =   42
         Top             =   480
         Width           =   11115
         Begin VB.TextBox TxtVersion 
            Height          =   315
            Left            =   8040
            MaxLength       =   3
            TabIndex        =   67
            Text            =   "  "
            Top             =   0
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtRMPerDis 
            Height          =   315
            Left            =   7680
            MaxLength       =   15
            TabIndex        =   8
            Top             =   1440
            Width           =   3225
         End
         Begin VB.TextBox TxtDay 
            Height          =   315
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "  "
            Top             =   1440
            Width           =   3615
         End
         Begin VB.TextBox txtBNFclaimable 
            Height          =   315
            Left            =   7680
            MaxLength       =   15
            TabIndex        =   6
            Top             =   1080
            Width           =   3225
         End
         Begin VB.ComboBox cboBHLTCode 
            Height          =   315
            Left            =   1920
            TabIndex        =   1
            Top             =   360
            Width           =   3615
         End
         Begin VB.ComboBox cboBPLNCode 
            Height          =   315
            Left            =   1920
            TabIndex        =   3
            Top             =   720
            Width           =   3615
         End
         Begin VB.ComboBox cboBINSCode 
            Height          =   315
            Left            =   7680
            TabIndex        =   2
            Top             =   360
            Width           =   3225
         End
         Begin VB.ComboBox cboCatCode 
            Height          =   315
            Left            =   7680
            TabIndex        =   4
            Top             =   720
            Width           =   3225
         End
         Begin MSMask.MaskEdBox txtEffdate 
            Height          =   315
            Left            =   1920
            TabIndex        =   5
            Top             =   1080
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtTempDate 
            Height          =   285
            Left            =   9120
            TabIndex        =   62
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label Label14 
            Caption         =   "RM Per Disability"
            Height          =   255
            Left            =   6000
            TabIndex        =   63
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label28 
            Caption         =   "Sub Cat Code*"
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
            Left            =   6000
            TabIndex        =   50
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "No of Days        "
            Height          =   240
            Left            =   255
            TabIndex        =   49
            Top             =   1440
            Width           =   1275
         End
         Begin VB.Label Label18 
            Caption         =   "Insured Type *"
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
            Left            =   6000
            TabIndex        =   48
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label19 
            Caption         =   "Plan Code     *"
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
            Left            =   255
            TabIndex        =   47
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label Label22 
            Caption         =   "RM / Day"
            Height          =   195
            Left            =   6000
            TabIndex        =   46
            Top             =   1080
            Width           =   1320
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Claimable"
            Height          =   330
            Left            =   -1845
            TabIndex        =   45
            Top             =   675
            Width           =   1050
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
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Eff Date       *"
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
            Left            =   255
            TabIndex        =   43
            Top             =   1080
            Width           =   1215
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
         Height          =   1620
         Index           =   0
         Left            =   -74760
         TabIndex        =   32
         Top             =   480
         Width           =   11145
         Begin VB.ComboBox cboSCatCode 
            Height          =   315
            Left            =   1680
            TabIndex        =   22
            Top             =   240
            Width           =   6855
         End
         Begin VB.TextBox TxtsubcatDescription 
            Height          =   495
            Left            =   1680
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   960
            Width           =   8775
         End
         Begin VB.TextBox TxtsubcatCode 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   23
            Text            =   "  "
            Top             =   600
            Width           =   6855
         End
         Begin VB.Label Label6 
            Caption         =   "Cat Code        *"
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
            Left            =   135
            TabIndex        =   38
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label4 
            Caption         =   "Description     *"
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
            TabIndex        =   35
            Top             =   960
            Width           =   1635
         End
         Begin VB.Label Label2 
            Caption         =   "Sub Cat Code *"
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
            Left            =   135
            TabIndex        =   34
            Top             =   600
            Width           =   1560
         End
         Begin VB.Label Label1 
            Caption         =   "Plan Claimable"
            Height          =   330
            Left            =   -1845
            TabIndex        =   33
            Top             =   675
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
         Height          =   1620
         Index           =   5
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   11025
         Begin VB.TextBox TxtCATCode 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   20
            Text            =   "  "
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox TxtCATDescription 
            Height          =   855
            Left            =   1680
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   600
            Width           =   9255
         End
         Begin VB.Label Label34 
            Caption         =   "Plan Claimable"
            Height          =   330
            Left            =   -1845
            TabIndex        =   29
            Top             =   675
            Width           =   1050
         End
         Begin VB.Label Label36 
            Caption         =   "Category Code  *"
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
            Left            =   135
            TabIndex        =   28
            Top             =   240
            Width           =   1560
         End
         Begin VB.Label Label37 
            Caption         =   "Description       *"
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
            TabIndex        =   27
            Top             =   600
            Width           =   1635
         End
      End
      Begin MSComctlLib.ListView lstCATList 
         Height          =   3765
         Left            =   240
         TabIndex        =   30
         Top             =   2400
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6641
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
            Text            =   "Category Code"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Category Description"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComctlLib.ListView lstSubCatList 
         Height          =   3765
         Left            =   -74760
         TabIndex        =   36
         Top             =   2400
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6641
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
            Text            =   "Cat Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Category Desc"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sub Cat Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   10583
         EndProperty
      End
      Begin MSComctlLib.ListView lstBNFList 
         Height          =   2805
         Left            =   -74640
         TabIndex        =   57
         Top             =   3360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4948
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sub Cat. Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sub Cat. Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Insured Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Plan Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Health Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Days/Months"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "RM/Day"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "RM Per Disability"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Payor Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "BENEFIT SUB CATERGORY LIST"
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
         TabIndex        =   37
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "BENEFIT CATERGORY LIST"
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
         TabIndex        =   31
         Top             =   2160
         Width           =   2295
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Benefits Maintenance"
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
      TabIndex        =   68
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label13 
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
      TabIndex        =   41
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label txtTotalRecord 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   40
      Top             =   7680
      Width           =   1455
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
      Left            =   240
      TabIndex        =   39
      Top             =   7680
      Width           =   4695
   End
End
Attribute VB_Name = "frmMaintenanceBenefits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim bExit As Boolean
Public BNFCriteria As String
Public CATCriteria As String
Public subcatCriteria As String
Public TBNFCriteria As String
Dim S As Boolean
Dim recordInBNF As Integer
Private m_blnDirection(1 To 11) As Boolean
Dim FCol, Frow As Integer

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

Private Sub FGrid_Click()
    Dim temp, tempCol, tempRow, tempTitle As String
    Dim days, value As Double
    tempCol = FGrid.col
    tempRow = FGrid.row
    FGrid.row = 0
    tempTitle = FGrid.Text
    FGrid.row = tempRow
    FGrid.col = 1
    temp = InputBox(tempTitle, FGrid.Text)
    FGrid.col = tempCol
    If temp <> "" Then
        FGrid = temp
        If tempCol = 3 Then
            FGrid.col = 2
            days = FGrid.Text
            If IsNumeric(days) = True Then
                FGrid.col = 3
                value = FGrid.Text
                If IsNumeric(value) = True Then
                    FGrid.col = 4
                    FGrid = days * value
                End If
            End If
        End If
    End If
End Sub

'**************************Start Form******************************
'Private Sub Form_Activate()
    
    'SSTab1.Tab = 0
'End Sub

Private Sub Form_Load()
    'Frame3.Width = Screen.Width
    CmdRefresh.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = False 'True
    Me.Width = SSTab1.Width + 500
    'Me.Height = SSTab1.Height + Frame3.Height + 600
    FGrid.ColAlignment(2) = 7
    FGrid.ColAlignment(3) = 7
    FGrid.ColAlignment(4) = 7
    BNFCriteria = ""
    CATCriteria = ""
    subcatCriteria = ""
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub
'**************************End Form******************************

'**************************Start Add, Edit, Delete, Search, Save, Exit******************************
Private Sub cmdAdd_Click()
    EditFlag = False
    'For FCol = 2 To 4
    '    FGrid.Col = FCol
    '    For Frow = 1 To 22
    '        FGrid.Row = Frow
    '        FGrid = ""
    '    Next
    'Next

    Call ClearForm(Me)
    Call EntriesEnable(True)
    CmdSave.Enabled = True

    Select Case SSTab1.Tab
        Case 0
            TxtCATCode.SetFocus
        Case 2
            cboCatCode.SetFocus
    End Select
End Sub

Private Sub cmdEdit_Click()
   EditFlag = True
    Call EntriesEnable(True)
    CmdSave.Enabled = True
    CmdAdd.Enabled = True
    CmdDelete.Enabled = False
    If cboCatCode <> "" Then
       cboCatCode.Enabled = False
       cboBINSCode.Enabled = False
       cboBHLTCode.Enabled = False
       cboBPLNCode.Enabled = False
       txtEffdate.Enabled = False
       TxtDay.Enabled = True
       txtBNFclaimable.Enabled = True
       txtRMPerDis.Enabled = True
       CmdEdit.Enabled = False
       CmdDelete.Enabled = False
    End If
    
    If TxtCATCode <> "" Then
       TxtCATCode.Enabled = False
       TxtCATDescription.Enabled = True
       TxtCATDescription.SetFocus
    End If
              
    If Trim(cboSCatCode) <> "" Then
        cboSCatCode.Enabled = False
        TxtsubcatCode.Enabled = False
       
    End If
End Sub

Private Sub cmdDelete_Click()
    Screen.MousePointer = vbHourglass

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If SSTab1.Tab = 0 Then
      Call DeleteCAT
    ElseIf SSTab1.Tab = 1 Then
      Call Deletesubcat
    ElseIf SSTab1.Tab = 3 Then
      Call DeleteBNF
    End If
    CmdAdd.Enabled = True
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    

Screen.MousePointer = Default
End Sub

Private Sub cmdSearch_Click()

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    CmdDelete.Enabled = False
    CmdSearch.Enabled = False
    CmdSave.Enabled = False
    
    Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SearchCAT
            TxtCATCode.Enabled = True
            TxtCATDescription.Enabled = True
            TxtCATCode.SetFocus
            CmdAdd.Enabled = True
        Case 1
            Call Searchsubcat
            CmdAdd.Enabled = True
           ' cboSCatCode.SetFocus
        Case 2
            
            Call searchtBNF
            'CmdAdd.Enabled = False
            'cboCatCode.SetFocus
        Case 3
            Call searchBNF
            cboBINSCode.Enabled = True
            cboBHLTCode.Enabled = True
            cboBPLNCode.Enabled = True
            CmdAdd.Enabled = True
            txtBNFclaimable.Enabled = False
          
        End Select

    Screen.MousePointer = Default
    
 '   medixmasterconn.Close
 '   Set medixmasterconn = Nothing
    CmdDelete.Enabled = True
    CmdSearch.Enabled = True
    CmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    Screen.MousePointer = vbHourglass
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   '
    Select Case SSTab1.Tab
        Case 0
            Call SaveCAT
        Case 1
            Call SaveSubCAT
        Case 2
            Call SaveBNF
        Case 3
            Call SaveBNFList
    End Select
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
    Screen.MousePointer = Default
End Sub


Private Sub CmdExit_Click()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


Private Sub CmdPrint_Click()
    
    Screen.MousePointer = vbHourglass
        
    Select Case SSTab1.Tab
        Case 0
            If lstCATList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstCATList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 1
            If lstSubCatList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstSubCatList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 3
            If lstBNFList.ListItems.Count <> 0 Then
                Call ExportListViewtoExcel(lstBNFList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
    End Select
    
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = vbDefault
    
    'Screen.MousePointer = vbHourglass
    'Select Case SSTab1.Tab
        'Case 0
            'Call PrintCAT
        'Case 1
            'Call PrintSUBCAT
        'Case 3
            'Call PrintBNF
    'End Select
    'CmdAdd.Enabled = True
    'CmdEdit.Enabled = False
    'Screen.MousePointer = Default
End Sub


Private Sub lstBNFList_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
    Dim strsql As String
    If lstBNFList.ListItems.Count > 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        cboCatCode = lstBNFList.SelectedItem.Text
        cboBINSCode = lstBNFList.SelectedItem.SubItems(2)
        cboBPLNCode = lstBNFList.SelectedItem.SubItems(3)
        cboBHLTCode = lstBNFList.SelectedItem.SubItems(4)
        TxtDay = lstBNFList.SelectedItem.SubItems(5)
        txtBNFclaimable = lstBNFList.SelectedItem.SubItems(6)
        txtRMPerDis = lstBNFList.SelectedItem.SubItems(7)
        txtEffdate = lstBNFList.SelectedItem.SubItems(8)
        'txtTempDate = lstBNFList.SelectedItem.SubItems(7)
        
        Call EntriesEnable(False)
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstBNFList_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
    Dim strsql As String
    If lstBNFList.ListItems.Count > 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        cboCatCode = lstBNFList.SelectedItem.Text
        cboBINSCode = lstBNFList.SelectedItem.SubItems(2)
        cboBPLNCode = lstBNFList.SelectedItem.SubItems(3)
        cboBHLTCode = lstBNFList.SelectedItem.SubItems(4)
        TxtDay = lstBNFList.SelectedItem.SubItems(5)
        txtBNFclaimable = lstBNFList.SelectedItem.SubItems(6)
        txtRMPerDis = lstBNFList.SelectedItem.SubItems(7)
        txtEffdate = lstBNFList.SelectedItem.SubItems(8)
        'txtTempDate = lstBNFList.SelectedItem.SubItems(7)
        
        Call EntriesEnable(False)
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
End Sub

Private Sub lstCATList_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
Dim rstCAT As New ADODB.Recordset
Dim strsql As String

    If lstCATList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtCATCode = lstCATList.SelectedItem.Text
        TxtCATDescription = lstCATList.SelectedItem.SubItems(1)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        CmdDelete.Enabled = True
    End If
    
    'If lstCATList.listitems.count Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
        
        'strsql = "Select CATCode, CATDescription " & _
        '"From BNFCategory where CATCode = '" & lstCATList.SelectedItem.Text & "'"
        'rstCAT.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        'TxtCATCode = rstCAT!CATCode
        'TxtCATDescription = rstCAT!catDescription
        
        'Call EntriesEnable(False)
        'CmdSave.Enabled = False
        'CmdDelete.Enabled = True
        'rstCAT.Close
        'Set rstCAT = Nothing
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
    'End If
End Sub

Private Sub lstCATList_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
Dim rstCAT As New ADODB.Recordset
Dim strsql As String

    If lstCATList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtCATCode = lstCATList.SelectedItem.Text
        TxtCATDescription = lstCATList.SelectedItem.SubItems(1)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        CmdDelete.Enabled = True
    End If
    
    'If lstCATList.listitems.count Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
        
        'strsql = "Select CATCode, CATDescription " & _
        '"From BNFCategory where CATCode = '" & lstCATList.SelectedItem.Text & "'"
        'rstCAT.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        'TxtCATCode = rstCAT!CATCode
        'TxtCATDescription = rstCAT!catDescription
        
        'Call EntriesEnable(False)
        'CmdSave.Enabled = False
        'CmdDelete.Enabled = True
        'rstCAT.Close
        'Set rstCAT = Nothing
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
    'End If
End Sub

Private Sub lstSubCatList_KeyDown(KeyCode As Integer, Shift As Integer)

    If lstSubCatList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        cboSCatCode = lstSubCatList.SelectedItem.Text
        TxtsubcatCode = lstSubCatList.SelectedItem.SubItems(2)
        TxtsubcatDescription = lstSubCatList.SelectedItem.SubItems(3)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        CmdDelete.Enabled = True
    End If
End Sub

Private Sub lstSubCatList_KeyUp(KeyCode As Integer, Shift As Integer)

    If lstSubCatList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        cboSCatCode = lstSubCatList.SelectedItem.Text
        TxtsubcatCode = lstSubCatList.SelectedItem.SubItems(2)
        TxtsubcatDescription = lstSubCatList.SelectedItem.SubItems(3)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        CmdDelete.Enabled = True
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
'    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    EditFlag = False
    Select Case SSTab1.Tab
    Case 0
        Call ClearForm(Me)
        CmdAdd.Enabled = True
        CmdEdit.Enabled = False
        CmdDelete.Enabled = False
        CmdRefresh.Enabled = False
        CmdSave.Enabled = True
        CmdPrint.Enabled = True
        CmdSearch.Enabled = True
        lstCATList.ListItems.Clear
        txtTotalRecord = 0
    Case 1
       If cboSCatCode.ListCount = 0 Then
            Call BNFSubcombolisting
        End If
        lstSubCatList.ListItems.Clear
        CmdRefresh.Enabled = True
        CmdAdd.Enabled = True
        CmdSearch.Enabled = True
        CmdSave.Enabled = True
        CmdPrint.Enabled = True
        txtTotalRecord = 0
    Case 2
        Call ClearForm(Me)
        cboTHLTCode.Enabled = True
        cboTInsCode.Enabled = True
        cboTPlnCode.Enabled = True
        txtTEffdate.Enabled = True
        CmdEdit.Enabled = False
        CmdDelete.Enabled = False
        CmdRefresh.Enabled = False
        CmdSearch.Enabled = True
        CmdAdd.Enabled = True
        CmdPrint.Enabled = False
        CmdSave.Enabled = True
        If cboTInsCode.ListCount = 0 Then
            Call BNFcombolisting
        End If
        txtTotalRecord = 0
    Case 3
        CmdAdd.Enabled = True
        CmdEdit.Enabled = False
        CmdDelete.Enabled = False
        CmdPrint.Enabled = True
        CmdRefresh.Enabled = True
        CmdSearch.Enabled = True
        CmdSave.Enabled = True
        If cboBINSCode.ListCount = 0 Then
            Call BNFCombolisting2
        End If
        lstBNFList.ListItems.Clear
        txtTotalRecord = 0
    End Select
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing

End Sub


Private Sub CmdRefresh_Click()
Screen.MousePointer = vbHourglass
Call ClearForm(Me)
Call EntriesEnable(True)

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

Select Case SSTab1.Tab
    Case 1
        Call BNFSubcombolisting
    Case 2
        Call BNFcombolisting
    Case 3
        Call BNFCombolisting2
End Select

'medixmasterconn.Close
'Set medixmasterconn = Nothing

Screen.MousePointer = Default
End Sub

Private Sub CmdClear_Click()

    'For FCol = 2 To 4
    '    FGrid.Col = FCol
    '    For Frow = 1 To 22
    '        FGrid.Row = Frow
    '        FGrid = ""
    '    Next
    'Next
    Call ClearGrid
    Call ClearForm(Me)
    lstCATList.ListItems.Clear
    lstSubCatList.ListItems.Clear
    lstBNFList.ListItems.Clear
    txtTotalRecord = 0
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
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
'**************************END EntriesEnable******************************


'############################### Start BNF (Benefit) ####Sunny#########################

Private Sub BNFcombolisting()
Dim HLT As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
Dim subCat As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cnt As Integer
FGrid.row = 0
FGrid.ColWidth(1) = 3500
FGrid.ColWidth(2) = 2000
FGrid.ColWidth(3) = 2000
FGrid.ColWidth(4) = 2000
FGrid.col = 0
FGrid = "Sub Cat Code"
FGrid.col = 1
FGrid = "Sub Cat"
FGrid.col = 2
FGrid = "Days/Months"
FGrid.col = 3
FGrid = "RM (Days/Months)"
FGrid.col = 4
FGrid = "RM Per Disability"

Screen.MousePointer = vbHourglass


' Set up a command object for the stored procedure.
Set cmd.ActiveConnection = medixmasterconn
cmd.CommandText = "SearchHLTType"
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 15
HLT.CursorLocation = adUseClient
HLT.CursorType = adOpenForwardOnly
HLT.CacheSize = 100

Set cmd2.ActiveConnection = medixmasterconn
cmd2.CommandText = "SearchINS"
cmd2.CommandType = adCmdStoredProc
cmd2.CommandTimeout = 15
INS.CursorLocation = adUseClient
INS.CursorType = adOpenForwardOnly
INS.CacheSize = 100

Set cmd3.ActiveConnection = medixmasterconn
cmd3.CommandText = "SearchPlan"
cmd3.CommandType = adCmdStoredProc
cmd3.CommandTimeout = 15
PLN.CursorLocation = adUseClient
PLN.CursorType = adOpenForwardOnly
PLN.CacheSize = 100

' Create a recordset by executing the command.
Set HLT = cmd.Execute
Set INS = cmd2.Execute
Set PLN = cmd3.Execute


subCat.Open "select BNFSubCode , BNFSubNAme from BNFSUBCAT", medixmasterconn, adOpenKeyset, adLockOptimistic
    

cboTHLTCode.Clear
cboTPlnCode.Clear
cboTInsCode.Clear

Do Until HLT.EOF
     cboTHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
     HLT.MoveNext
Loop
      
Do Until INS.EOF
     cboTInsCode.AddItem INS!INSCode & " - " & INS!INSDescription
     INS.MoveNext
Loop
      
Do Until PLN.EOF
     cboTPlnCode.AddItem PLN!PLNCode '& "-" & PLN!PLNDescription
     PLN.MoveNext
Loop

cnt = 1
FGrid.Rows = subCat.recordCount + 1
recordInBNF = subCat.recordCount + 1
Do Until subCat.EOF
     cboCatCode.AddItem subCat!BNFSubCode & " -- " & subCat!BNFSubName
     FGrid.col = 0
     FGrid.row = cnt
     FGrid.Text = subCat!BNFSubCode

     FGrid.col = 1
     FGrid.row = cnt
     FGrid.Text = subCat!BNFSubName
     cnt = cnt + 1
     subCat.MoveNext
Loop
    HLT.Close
    Set HLT = Nothing
    INS.Close
    Set INS = Nothing
    PLN.Close
    Set PLN = Nothing
    subCat.Close
    Set subCat = Nothing
    
Screen.MousePointer = Default


End Sub

Private Sub BNFCombolisting2()
Dim HLT As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
Dim subCat As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim GRPCOM As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim cmd5 As New ADODB.Command
Dim cmd6 As New ADODB.Command

Screen.MousePointer = vbHourglass


' Set up a command object for the stored procedure.
Set cmd.ActiveConnection = medixmasterconn
cmd.CommandText = "SearchHLTType"
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 15
HLT.CursorLocation = adUseClient
HLT.CursorType = adOpenForwardOnly
HLT.CacheSize = 100

Set cmd2.ActiveConnection = medixmasterconn
cmd2.CommandText = "SearchINS"
cmd2.CommandType = adCmdStoredProc
cmd2.CommandTimeout = 15
INS.CursorLocation = adUseClient
INS.CursorType = adOpenForwardOnly
INS.CacheSize = 100

Set cmd3.ActiveConnection = medixmasterconn
cmd3.CommandText = "SearchPlan"
cmd3.CommandType = adCmdStoredProc
cmd3.CommandTimeout = 15
PLN.CursorLocation = adUseClient
PLN.CursorType = adOpenForwardOnly
PLN.CacheSize = 100

Set cmd4.ActiveConnection = medixmasterconn
cmd4.CommandText = "SearchBNFSubCat"
cmd4.CommandType = adCmdStoredProc
cmd4.CommandTimeout = 15
subCat.CursorLocation = adUseClient
subCat.CursorType = adOpenForwardOnly
subCat.CacheSize = 100

Set cmd5.ActiveConnection = medixmasterconn
cmd5.CommandText = "SearchPayor"
cmd5.CommandType = adCmdStoredProc
cmd5.CommandTimeout = 15
PAY.CursorLocation = adUseClient
PAY.CursorType = adOpenForwardOnly
PAY.CacheSize = 100

Set cmd6.ActiveConnection = medixmasterconn
cmd6.CommandText = "SearchGrpCompany"
cmd6.CommandType = adCmdStoredProc
cmd6.CommandTimeout = 15
GRPCOM.CursorLocation = adUseClient
GRPCOM.CursorType = adOpenForwardOnly
GRPCOM.CacheSize = 100

' Create a recordset by executing the command.
Set HLT = cmd.Execute
Set INS = cmd2.Execute
Set PLN = cmd3.Execute
Set subCat = cmd4.Execute
Set PAY = cmd5.Execute
Set GRPCOM = cmd6.Execute


cboBHLTCode.Clear
cboBPLNCode.Clear
cboCatCode.Clear
cboBINSCode.Clear
CboPayorCode.Clear
CboGrpCom.Clear

Do Until HLT.EOF
     cboBHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
     HLT.MoveNext
Loop
HLT.Close
Set HLT = Nothing
  
Do Until INS.EOF
     cboBINSCode.AddItem INS!INSCode & " - " & INS!INSDescription
     INS.MoveNext
Loop
INS.Close
Set INS = Nothing
      
Do Until PLN.EOF
     cboBPLNCode.AddItem PLN!PLNCode '& "-" & PLN!PLNDescription
     PLN.MoveNext
Loop
PLN.Close
Set PLN = Nothing

Do Until PAY.EOF
     CboPayorCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
     PAY.MoveNext
Loop
PAY.Close
Set PAY = Nothing

Do Until GRPCOM.EOF
     If Len(GRPCOM!GrpCompany) Then
        CboGrpCom.AddItem GRPCOM!GrpCompany
     End If
     GRPCOM.MoveNext
Loop
GRPCOM.Close
Set GRPCOM = Nothing

Do Until subCat.EOF
     cboCatCode.AddItem subCat!BNFSubCode & " -- " & subCat!BNFSubName
     subCat.MoveNext
Loop
subCat.Close
Set subCat = Nothing


Screen.MousePointer = Default


End Sub

Private Sub BNFListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim BNF As ListItem
Dim sql As String
Dim TotalBNFCount As String
   
    TotalBNFCount = "0"
    lstBNFList.ListItems.Clear
    
    sql = " select BNFCode, PLNCode, HLTCode, INSCode, BNFSubName, " & _
      " BNFClaimAbleAmount, BNFNoOfDays , BNFRMPerDis, BNFEffDate, " & _
      " BNFSubName, PAYCode, GRPCompany " & _
      " From VIEW_BNF_PLN_SubCat where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        
        Do Until rst2.EOF
            
            Set BNF = lstBNFList.ListItems.Add(, , rst2!BNFCode)
                BNF.SubItems(1) = rst2!BNFSubName
                'BNF.SubItems(2) = rst2!INSCode
                If Len(Trim(rst2!INSCode)) Then
                    BNF.SubItems(2) = rst2!INSCode
                End If
                BNF.SubItems(3) = rst2!PLNCode
                BNF.SubItems(4) = rst2!HLTCode
                
                If IsNull(rst2!BNFNoOFDays) = False Then
                    BNF.SubItems(5) = rst2!BNFNoOFDays
                End If
                If IsNull(rst2!BNFClaimAbleAmount) = False Then
                    BNF.SubItems(6) = rst2!BNFClaimAbleAmount 'FormatCurrency()
                End If
                If IsNull(rst2!BNFRMperDis) = False Then
                    BNF.SubItems(7) = rst2!BNFRMperDis
                End If
                
                If IsNull(rst2!BNFEffDate) = False Then
                    BNF.SubItems(8) = rst2!BNFEffDate
                End If
                
                If IsNull(rst2!PAYCode) = False Then
                    BNF.SubItems(9) = rst2!PAYCode
                End If
                
                If IsNull(rst2!GrpCompany) = False Then
                    BNF.SubItems(10) = rst2!GrpCompany
                End If
                
                rst2.MoveNext
        TotalBNFCount = TotalBNFCount + 1
        txtTotalRecord = TotalBNFCount
        Loop
    
    rst2.Close
    Set rst2 = Nothing
    End If
End Sub

Private Sub lstBNFlist_click()
'On Error Resume Next
    Dim strsql As String
    If lstBNFList.ListItems.Count > 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        cboCatCode = lstBNFList.SelectedItem.Text
        cboBINSCode = lstBNFList.SelectedItem.SubItems(2)
        cboBPLNCode = lstBNFList.SelectedItem.SubItems(3)
        cboBHLTCode = lstBNFList.SelectedItem.SubItems(4)
        TxtDay = lstBNFList.SelectedItem.SubItems(5)
        txtBNFclaimable = lstBNFList.SelectedItem.SubItems(6)
        txtRMPerDis = lstBNFList.SelectedItem.SubItems(7)
        If lstBNFList.SelectedItem.SubItems(8) <> "__/__/____" Then
            txtEffdate = lstBNFList.SelectedItem.SubItems(8)
        End If
        'txtTempDate = lstBNFList.SelectedItem.SubItems(7)
        
        Call EntriesEnable(False)
        CmdDelete.Enabled = True
        CmdSave.Enabled = False
    End If
End Sub
 
Private Sub lstBNFList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstBNFList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstBNFList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstBNFList, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstBNFList, ColumnHeader.Index, ldtString, m_blnDirection(11)
    End Select
End Sub
''''''''''''''''''''save, edit delete''''''''''''''''''''''''''''
Private Sub SaveBNF()
On Error GoTo UpdateErr
Dim BNF As New ADODB.Recordset
Dim SaveBNF As New ADODB.Recordset
Dim RecordSaved
Dim strsql As String
Dim currentRecord As Integer
Dim beginTrans As Boolean


If Len(Trim(cboTHLTCode)) = 0 Then
 MsgBox "Please Enter Health Type !", vbOKOnly + vbCritical, DialogBoxTitle
    cboTHLTCode.SetFocus
ElseIf Len(Trim(cboTInsCode)) = 0 Then
    MsgBox "Please Enter Insured Type!", vbOKOnly + vbCritical, DialogBoxTitle
    cboTInsCode.SetFocus
ElseIf Len(Trim(cboTPlnCode)) = 0 Then
    MsgBox "Please Enter Plan Code!", vbOKOnly + vbCritical, DialogBoxTitle
    cboTPlnCode.SetFocus
ElseIf txtTEffdate = "__/__/____" Or IsDate(txtTEffdate) = False Then
    MsgBox "Please Enter (Valid )effectice Date!", vbOKOnly + vbCritical, DialogBoxTitle
    txtTEffdate.SetFocus
Else
  strsql = "Select BNFCode From BNF where PLNCode ='" & RemStr(cboTPlnCode) & "' and INSCode ='" & Left(cboTInsCode, 1) & "'" & _
            " AND BNFEffDate ='" & Format(txtTEffdate, "mm/dd/yyyy") & "'"
  BNF.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
  
  If BNF.recordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
        BNF.Close
        cboTHLTCode.SetFocus
        cboTHLTCode.SelStart = 0
        cboTHLTCode.SelLength = Len(cboTHLTCode)
        Set BNF = Nothing
        'Exit Sub
  Else
      RecordSaved = MsgBox("Do you want to save this record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
'            If EditFlag = False Then
 '               SaveBNF.Open "select * from BNF", medixmasterconn, adOpenKeyset, adLockOptimistic
  '              SaveBNF.AddNew
  '          Else
   '             strsql = "Select * From BNF Where " & _
    '                    " AND PLNCode ='" & RemStr(cboTPlnCode) & "' and INSCode ='" & Left(cboTInsCode, 1) & "'" & _
     '                   " AND BNFEffDate ='" & Format(txtTEffdate, "mm/dd/yyyy") & "'"
      '          SaveBNF.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic '
   '         End If
'            medixmasterconn.beginTrans
 '           beginTrans = True
            
            '
             For currentRecord = 1 To recordInBNF - 1 Step 1
               
                With SaveBNF
                FGrid.row = currentRecord
                FGrid.col = 4
                If Trim(FGrid.Text) <> "" Then
                    If EditFlag = False Then
                       SaveBNF.Open "select * from BNF", medixmasterconn, adOpenKeyset, adLockOptimistic
                       SaveBNF.AddNew
                    End If
                    !INSCode = Trim(Left(ChkStr(cboTInsCode.Text), 1))
                    !PLNCode = Trim(Left(ChkStr(cboTPlnCode.Text), 4))
                    !HLTCode = (ChkStr(Mid(cboTHLTCode.Text, 1, 1)))
                    !BNFEffDate = txtTEffdate
                    !BNFLastUpdateUser = Logon
                    !BNFLastUpdateDate = Date
                    FGrid.col = 0
                    !BNFCode = FGrid.Text
                    FGrid.col = 2
                    If Trim(FGrid.Text) <> "" Then
                        !BNFNoOFDays = FGrid.Text
                    End If
                    FGrid.col = 3
                    If Trim(FGrid.Text) <> "" Then
                        !BNFClaimAbleAmount = FGrid.Text
                    End If
                    FGrid.col = 4
                    !BNFRMperDis = FGrid.Text
                    SaveBNF.Update
                    SaveBNF.Close
                    Set SaveBNF = Nothing
                End If
                End With
            Next
            'Call ClearForm(Me)
            'Call BNFListing(BNFCriteria)
            
            If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            End If
       End If
       EditFlag = False
   End If
End If

  '  If beginTrans = True Then
   '     medixmasterconn.CommitTrans
    'End If
    
Exit Sub
UpdateErr:
    'If beginTrans = True Then
     '   medixmasterconn.RollbackTrans
    'End If
  MsgBox Err.Description
End Sub

Private Sub searchBNF()
  Dim strAppend As String
   
    strAppend = ""
    
    If Len(Trim(cboCatCode)) Then
        strAppend = strAppend + " And BNFCode = '" & Trim(Mid(cboCatCode, 1, IIf(InStr(1, cboCatCode, "--") = 0, 10, InStr(1, cboCatCode, "--") - 1))) & "'"
    End If
    If Len(Trim(cboBHLTCode)) Then
        strAppend = strAppend + " And HLTCode = '" & Mid(cboBHLTCode, 1, 1) & "'"
    End If
    If Len(Trim(cboBINSCode)) Then
        strAppend = strAppend + " And INSCode = '" & Mid(cboBINSCode, 1, 1) & "'"
    End If
    If Len(Trim(cboBPLNCode)) Then
        strAppend = strAppend + " And PLNCode = '" & Trim(Mid(cboBPLNCode, 1, 4)) & "'"
    End If
    
    If Len(Trim(txtEffdate)) > 0 And IsDate(txtEffdate) = True Then
        strAppend = strAppend + " And BNFEffDate = '" & Format(txtEffdate, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(CboPayorCode)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayorCode, 1, IIf(InStr(1, CboPayorCode, "-") = 0, 4, InStr(1, CboPayorCode, "-") - 1))) & "'"
    End If

    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend + " AND GRPCompany like '" & RemStr(Mid(CboGrpCom, 1, IIf(InStr(1, CboGrpCom, "-") = 0, 4, InStr(1, CboGrpCom, "-") - 1))) & "%'"
    End If
    
    BNFCriteria = strAppend
    Call BNFListing(strAppend)
    
End Sub

Private Sub DeleteBNF()
Dim rsDeleteBNF As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

EditFlag = False

If Trim(cboCatCode) = "" Then
    MsgBox "Please Enter Benefit Sub Category Code", vbOKOnly + vbCritical, DialogBoxTitle
    'cboCatCode.SetFocus
ElseIf lstBNFList.ListItems.Count = 0 Then
    MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
    sqlMBM = "Select BNFCode From VIEW_BNF_MBM_CodeChecking"
    sqlMBM = sqlMBM & " Where BNFCode ='" & Trim(cboCatCode) & "'"
   
    rsCheckMBM.MaxRecords = 1
    rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic

    If rsCheckMBM.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Member table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
            If DeleteRecord = vbYes And rsCheckMBM.recordCount = 0 Then
                strsql = "Delete From BNF Where BNFCode = '" & ChkStr(lstBNFList.SelectedItem.Text) & "'"
                rsDeleteBNF.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
                CmdSave.Enabled = True
                Call ClearForm(Me)
                Call BNFListing
                Call DisableCmdButton(Me)
                Call EntriesEnable(True)
                cboCatCode.SetFocus
            Else
                CmdDelete.Enabled = False
                CmdEdit.Enabled = False
            End If
    End If
End If

Exit Sub
   
End Sub
    
'Private Sub PrintBNF()
    'RptBNFList.Show
'End Sub

Private Sub cboBHLTCode_Click()
Dim sql As String
Dim PLN As New ADODB.Recordset
Dim strsql As String
Dim PLNCode As String

If Mid(cboBHLTCode, 1, 1) = "S" Then
    S = True
    If S = True Then
    
    
    'PLN.Open "PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
    'strsql = "SELECT PRMCode, PRMAmount  From PRM "
    'strsql = strsql & "WHERE (INSCode = '" & InsureType & "') AND (HLTCode = '" & HealthValue & "') "
    'strsql = strsql & "AND (PLNCode = '" & PlanValue & "') AND (AGECode = '" & AgeValue & "') "
    'PLNCode = Mid(PLN!PLNCode, 1, 1)
    
    'PLN.Close
    strsql = "Select PLNCode, HLTCode From PLN "
    strsql = strsql & "Where (PLNCode = '" & PLNCode & "') "
    strsql = strsql & " and  (HLTCode = '" & Mid(Trim(cboBHLTCode), 1, 1) & "')"
    PLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
'  strsql = "Select AGECode, HLTCode From AGE "
'        strsql = strsql & "where (AGECode = '" & TxtAgeCode & "')  and (HLTCode = '" & Trim(CblHealth) & "')"
'        AGE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until PLN.EOF
    'txtPLNcode.AddItem PLN!PLNCode
    cboBPLNCode = PLN!PLNCode
    PLN.MoveNext
    Loop
    
    'ElseIf PLN!HLTCode = "N" Then
    'Do Until PLN.EOF
    'txtPLNcode = PLN!PLNCode
    'PLN.MoveNext
    'Loop

    End If
    End If
End Sub

Private Sub cboBHLTCode_KeyPress(KeyAscii As Integer)
'If secondeditflag = False Then

   Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboBHLTCode.Text, cboBHLTCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboBHLTCode.Text, cboBHLTCode.SelStart + cboBHLTCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboBHLTCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboBHLTCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboBHLTCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboBHLTCode.Text = ValidValue
              cboBHLTCode.SelStart = 0
              cboBHLTCode.SelLength = Len(ValidValue)
            End If
            
'End If

End Sub

Private Sub cboBINSCode_KeyPress(KeyAscii As Integer)
'If secondeditflag = False Then

Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboBINSCode.Text, cboBINSCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboBINSCode.Text, cboBINSCode.SelStart + cboBINSCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboBINSCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboBINSCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboBINSCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboBINSCode.Text = ValidValue
              cboBINSCode.SelStart = 0
              cboBINSCode.SelLength = Len(ValidValue)
              End If
'End If
End Sub

Private Sub cboBPLNCode_KeyPress(KeyAscii As Integer)
'If secondeditflag = False Then

Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboBPLNCode.Text, cboBPLNCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboBPLNCode.Text, cboBPLNCode.SelStart + cboBPLNCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboBPLNCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboBPLNCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboBPLNCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboBPLNCode.Text = ValidValue
              cboBPLNCode.SelStart = 0
              cboBPLNCode.SelLength = Len(ValidValue)
              End If
            
'End If

End Sub

Private Sub cbocatcode_KeyPress(KeyAscii As Integer)
Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboCatCode.Text, cboCatCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboCatCode.Text, cboCatCode.SelStart + cboCatCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboCatCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboCatCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboCatCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboCatCode.Text = ValidValue
              cboCatCode.SelStart = 0
              cboCatCode.SelLength = Len(ValidValue)
              End If
End Sub


'############################### Start BNF category (Benefit) ####Sunny#########################
Private Sub SaveCAT()
Dim RecordSaved
Dim CAT As New ADODB.Recordset
Dim SaveCAT As New ADODB.Recordset
Dim strsql As String
    
        
    If TxtCATCode = "" Then
        MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtCATCode.SetFocus
    ElseIf TxtCATDescription = "" Then
        MsgBox "Please Enter Category Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtCATDescription.SetFocus
    Else
        strsql = "Select * From BNFCategory where CATCode = '" & Trim(TxtCATCode) & "'"
        CAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If CAT.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            CAT.Close
            TxtCATCode.SetFocus
            TxtCATCode.SelStart = 0
            TxtCATCode.SelLength = Len(TxtCATCode)
            Set CAT = Nothing
            'Exit Sub
            
        Else
            RecordSaved = MsgBox("Do you want to save this record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                   SaveCAT.Open "BNFCategory", medixmasterconn, adOpenKeyset, adLockOptimistic
                   SaveCAT.AddNew
                Else
                   strsql = "Select * From BNFCategory Where CATCode = '" & ChkStr(TxtCATCode) & "'"
                   SaveCAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                End If
                With SaveCAT
                     !CATCode = ChkStr(TxtCATCode)
                     !catDescription = ChkStr(TxtCATDescription)
                     !CATLastUpdateUser = Logon
                     !CATLastUpdateDate = Date
                End With
                SaveCAT.Update
                Call ClearForm(Me)
                Call CATListing(CATCriteria)
                'lstCATList.listitems.clear
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        
                SaveCAT.Close
                Set SaveCAT = Nothing
            End If
            EditFlag = False
      End If
  End If
  
End Sub

Private Sub DeleteCAT()
Dim rsDeleteCAT As New ADODB.Recordset
Dim rsCheckBNF As New ADODB.Recordset
Dim strsql, sqlBNF As String
Dim DeleteRecord
    
EditFlag = False
       
If TxtCATCode = "" Then
   MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtCATCode.SetFocus
ElseIf lstCATList.ListItems.Count = 0 Then
   MsgBox "Please search th record", vbOKOnly + vbCritical, DialogBoxTitle
Else
    sqlBNF = "Select BNFCATCode From BNFSubCAt where bnfCATCode = '" & Trim(lstCATList.SelectedItem.Text) & "'"
    
    rsCheckBNF.MaxRecords = 1
    rsCheckBNF.Open sqlBNF, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rsCheckBNF.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Benefit table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes And rsCheckBNF.recordCount = 0 Then
             strsql = "Delete From BNFCategory Where CATCode = '" & ChkStr(lstCATList.SelectedItem.Text) & "'"
             rsDeleteCAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call CATListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             TxtCATCode.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
End If

Exit Sub

End Sub

Private Sub SearchCAT()
    
    Dim strAppend As String
   
    strAppend = ""
    
    If Len(Trim(TxtCATCode)) Then
        strAppend = strAppend + " And CATCode like '" & ChkStr(TxtCATCode) & "%'"
    End If
   
    If Len(Trim(TxtCATDescription)) Then
        strAppend = strAppend + " And CATDescription like '%" & RemStr(TxtCATDescription) & "%'"
    End If
    
    CATCriteria = strAppend
    Call CATListing(strAppend)
    
End Sub

Private Sub CATListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim CAT As ListItem
Dim sql As String
Dim TotalCATCount As Integer
   
    TotalCATCount = 0
    lstCATList.ListItems.Clear
    
    sql = " select CATCode, CATDescription " & _
      " From BNFCategory where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found!", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            With rst2
            Set CAT = lstCATList.ListItems.Add(, , rst2!CATCode)
                CAT.SubItems(1) = rst2!catDescription
                rst2.MoveNext
            End With
        TotalCATCount = TotalCATCount + 1
        txtTotalRecord = TotalCATCount
        Loop
    End If
    
    rst2.Close
    Set rst2 = Nothing
    
End Sub

Private Sub lstCATList_Click()
'On Error Resume Next
    
    If lstCATList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
        TxtCATCode = lstCATList.SelectedItem.Text
        TxtCATDescription = lstCATList.SelectedItem.SubItems(1)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        CmdDelete.Enabled = True
    End If
    
    'If lstCATList.listitems.count Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'Dim rstCAT As New ADODB.Recordset
        'Dim strsql As String
        
        'Call EnableCmdButton(Me)
        'Call ClearForm(Me)
        
        'strsql = "Select CATCode, CATDescription " & _
        '"From BNFCategory where CATCode = '" & lstCATList.SelectedItem.Text & "'"
        'rstCAT.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        
        'TxtCATCode = rstCAT!CATCode
        'TxtCATDescription = rstCAT!catDescription
        
        'Call EntriesEnable(False)
        'CmdSave.Enabled = False
        'CmdDelete.Enabled = True
        'rstCAT.Close
        'Set rstCAT = Nothing
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
    'End If
End Sub

Private Sub lstCATList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCATList.SortKey = ColumnHeader.Index - 1
    lstCATList.Sorted = True
    If lstCATList.SortOrder = lvwAscending Then
        lstCATList.SortOrder = lvwDescending
    Else
        lstCATList.SortOrder = lvwAscending
    End If
End Sub

'Private Sub PrintCAT()
    'RptBNFCatList.Show
'End Sub

'############################### Start BNF Sub category (Benefit) ####   Eric  #########################

Private Sub BNFSubcombolisting()
Dim CAT As New ADODB.Recordset
Dim cmd As New ADODB.Command

    Screen.MousePointer = vbHourglass
    
       
    ' Set up a command object for the stored procedure.
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchBNFCategory"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    CAT.CursorLocation = adUseClient
    CAT.CursorType = adOpenForwardOnly
    CAT.CacheSize = 100
    
    ' Create a recordset by executing the command.
    Set CAT = cmd.Execute
        
    cboSCatCode.Clear
    
    Do Until CAT.EOF
         cboSCatCode.AddItem CAT!CATCode & " * " & CAT!catDescription
         CAT.MoveNext
    Loop
    CAT.Close
    Set CAT = Nothing
    
        
    Screen.MousePointer = Default
End Sub

Private Sub SaveSubCAT()
Dim RecordSaved
Dim subCat As New ADODB.Recordset
Dim SaveSubCAT  As New ADODB.Recordset
Dim strsql As String
    
        
    If cboSCatCode = "" Then
        MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboSCatCode.SetFocus
    ElseIf TxtsubcatCode = "" Then
        MsgBox "Please Enter Sub Category Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtsubcatCode.SetFocus
    ElseIf TxtsubcatDescription = "" Then
        MsgBox "Please Enter Sub Category Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtsubcatDescription.SetFocus
    Else
        strsql = "Select BNFSubCode  From BNFsubcat where BNFsubCode = '" & Trim(TxtsubcatCode) & "'"
        subCat.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If subCat.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"

            TxtsubcatCode.SetFocus
            TxtsubcatCode.SelStart = 0
            TxtsubcatCode.SelLength = Len(TxtsubcatCode)
            subCat.Close
            Set subCat = Nothing
            'Exit Sub
        Else
            RecordSaved = MsgBox("Do you want to save this record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                   SaveSubCAT.Open "select * from BNFsubcat ", medixmasterconn, adOpenKeyset, adLockOptimistic
                   SaveSubCAT.AddNew
                Else
                   strsql = "Select * From BNFsubcat  Where BNFsubCode = '" & ChkStr(TxtsubcatCode) & "'"
                   SaveSubCAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                End If
                With SaveSubCAT
                     !BNFCatCode = Trim(Mid(cboSCatCode, 1, IIf(InStr(1, cboSCatCode, "*") = 0, 10, InStr(1, cboSCatCode, "*") - 1)))
                     !BNFSubCode = TxtsubcatCode
                     !BNFSubName = TxtsubcatDescription
                     !BNFLastUpdateUser = Logon
                     !BNFLastUpdateDate = Date
                End With
                SaveSubCAT.Update
                Call ClearForm(Me)
                Call subcatListing(subcatCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        
                SaveSubCAT.Close
                Set SaveSubCAT = Nothing
            End If
            EditFlag = False
      End If
  End If
  
End Sub

Private Sub Deletesubcat()
Dim rsDeletesubcat As New ADODB.Recordset
Dim rsCheckBNF As New ADODB.Recordset
Dim strsql, sqlBNF As String
Dim DeleteRecord
    

EditFlag = False
       
If TxtsubcatCode = "" Then
    MsgBox "Please Enter subcategory Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtsubcatCode.SetFocus
ElseIf lstSubCatList.ListItems.Count = 0 Then
    MsgBox "Please select a record to be deleted! ", vbOKOnly + vbCritical, DialogBoxTitle
Else
    sqlBNF = "Select BNFCode From BNF where BNFCode = '" & Trim(lstSubCatList.SelectedItem.SubItems(2)) & "'"
    
    rsCheckBNF.MaxRecords = 1
    rsCheckBNF.Open sqlBNF, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rsCheckBNF.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Benefit table!", vbCritical, "Error in deleting record..."
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes And rsCheckBNF.recordCount = 0 Then
             strsql = "Delete From BNFsubcat Where BNFsubcode = '" & RemStr(lstSubCatList.SelectedItem.SubItems(2)) & "'"
             rsDeletesubcat.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call subcatListing(subcatCriteria)
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             TxtsubcatCode.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
End If

Exit Sub

End Sub

Private Sub Searchsubcat()
    Screen.MousePointer = vbHourglass
    Dim strAppend As String
   
    strAppend = ""
    
    If Len(Trim(cboSCatCode)) Then
        strAppend = strAppend + " And BNFCATCode like '" & RemStr(Mid(cboSCatCode, 1, IIf(InStr(1, cboSCatCode, "*") = 0, 4, InStr(1, cboSCatCode, "*") - 1))) & "%'"
    End If
    
    If Len(Trim(TxtsubcatCode)) Then
        strAppend = strAppend + " And BNFsubCode like '" & RemStr(TxtsubcatCode) & "%'"
    End If
   
    If Len(Trim(TxtsubcatDescription)) Then
        strAppend = strAppend + " And BNFSUbName like '%" & RemStr(TxtsubcatDescription) & "%'"
    End If
    
    subcatCriteria = strAppend
    Call subcatListing(strAppend)
    Screen.MousePointer = Default
End Sub


Private Sub subcatListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim subCat As ListItem
Dim sql As String
Dim TotalSubCatCount As Integer
   
    TotalSubCatCount = 0
    lstSubCatList.ListItems.Clear
    
    sql = " select BNFCatCode, BNFsubCode, BNFsubName , CATDescription" & _
      " From View_BNF_CAT_subcat  where 0=0 "
     
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        
        Do Until rst2.EOF
            With rst2
            Set subCat = lstSubCatList.ListItems.Add(, , rst2!BNFCatCode)
                subCat.SubItems(1) = rst2!catDescription
                subCat.SubItems(2) = rst2!BNFSubCode
                subCat.SubItems(3) = rst2!BNFSubName
                rst2.MoveNext
            End With
        TotalSubCatCount = TotalSubCatCount + 1
        txtTotalRecord = TotalSubCatCount
        Loop
    End If
    
    rst2.Close
    Set rst2 = Nothing
    
End Sub

'Private Sub PrintSUBCAT()
    'RptBNFSubList.Show
'End Sub

Private Sub lstsubcatList_Click()

    If lstSubCatList.ListItems.Count <> 0 Then
        Call EnableCmdButton(Me)
        cboSCatCode = lstSubCatList.SelectedItem.Text
        TxtsubcatCode = lstSubCatList.SelectedItem.SubItems(2)
        TxtsubcatDescription = lstSubCatList.SelectedItem.SubItems(3)
        Call EntriesEnable(False)
        CmdSave.Enabled = False
        CmdDelete.Enabled = True
    End If
End Sub

Private Sub lstsubcatList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstSubCatList.SortKey = ColumnHeader.Index - 1
    lstSubCatList.Sorted = True
    If lstSubCatList.SortOrder = lvwAscending Then
        lstSubCatList.SortOrder = lvwDescending
    Else
        lstSubCatList.SortOrder = lvwAscending
    End If
End Sub

Private Sub cboSCatCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboSCatCode.Text, cboSCatCode.SelStart) + Chr(KeyAscii) + Mid(cboSCatCode.Text, cboSCatCode.SelStart + cboSCatCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboSCatCode.ListCount - 1
 
        If NewText = LCase(Left(cboSCatCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboSCatCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboSCatCode.Text = ValidValue
                cboSCatCode.SelStart = 0
                cboSCatCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

' # ***********************************  BNF Table

Private Sub searchtBNF()
  Dim strAppend As String
   
    strAppend = ""
    
    'If Len(Trim(cboTcatCode)) Then
     '   strAppend = strAppend + " And BNFCode = '" & Trim(Mid(cboTcatCode, 1, IIf(InStr(1, cboTcatCode, "--") = 0, 10, InStr(1, cboTcatCode, "--") - 1))) & "'"
    'End If
    If Len(Trim(cboTHLTCode)) = 0 Then
        MsgBox "Please key in Health Type !", vbCritical
    ElseIf Len(Trim(cboTInsCode)) = 0 Then
        MsgBox "Please key in Insured Type !", vbCritical
    ElseIf Len(Trim(cboTPlnCode)) = 0 Then
        MsgBox "Please key in Plan Code!", vbCritical
    ElseIf txtTEffdate = "__/__/____" Then
        MsgBox "Please key in Effective Date!", vbCritical
    ElseIf IsDate(txtTEffdate) = False Then
        MsgBox "Please key in VALID Effective Date!", vbCritical
    Else
        strAppend = strAppend + " And PLNCode = '" & Trim(Mid(cboTPlnCode, 1, 4)) & "'"
        strAppend = strAppend + " And INSCode = '" & Mid(cboTInsCode, 1, 1) & "'"
        strAppend = strAppend + " And HLTCode = '" & Mid(cboTHLTCode, 1, 1) & "'"
        strAppend = strAppend + " And BNFEffDate = '" & Format(txtTEffdate, "mm/dd/yyyy") & "'"
        TBNFCriteria = strAppend
        Call tBNFlisting
    End If
End Sub

Private Sub tBNFlisting()
Dim BNF As New ADODB.Recordset
Dim strsql As String
Dim counter As Integer

    strsql = "Select BNFCode , BNFNoOFDays, BNFClaimableAmount , BNFRMPerDis  , BNFEffDate from BNF where 0=0 " & TBNFCriteria
    strsql = strsql & " order by BNFCOde "
    BNF.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Call ClearGrid
    If BNF.recordCount Then
    Do Until BNF.EOF
         With BNF
            For counter = 1 To recordInBNF - 1 Step 1
             FGrid.row = counter
             FGrid.col = 0
                If FGrid.Text = !BNFCode Then
                    FGrid.col = 2
                    If !BNFNoOFDays <> "" Then
                        FGrid.Text = !BNFNoOFDays
                    End If
                    FGrid.col = 3
                    If !BNFClaimAbleAmount <> "" Then
                        
                        FGrid.Text = !BNFClaimAbleAmount
                    End If
                    FGrid.col = 4
                    If !BNFRMperDis <> "" Then
                        
                        FGrid.Text = !BNFRMperDis
                    End If
                End If
            Next
            .MoveNext
         End With
    Loop
    End If
    BNF.Close
    Set BNF = Nothing
End Sub

''  Key Press  *******************************************************

Private Sub cbotHLTcode_KeyPress(KeyAscii As Integer)
'If secondeditflag = False Then

   Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboTHLTCode.Text, cboTHLTCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboTHLTCode.Text, cboTHLTCode.SelStart + cboTHLTCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboTHLTCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboTHLTCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboTHLTCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboTHLTCode.Text = ValidValue
              cboTHLTCode.SelStart = 0
              cboTHLTCode.SelLength = Len(ValidValue)
            End If
            
'End If

End Sub

Private Sub cbotinscode_KeyPress(KeyAscii As Integer)
'If secondeditflag = False Then

Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboTInsCode.Text, cboTInsCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboTInsCode.Text, cboTInsCode.SelStart + cboTInsCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboTInsCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboTInsCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboTInsCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboTInsCode.Text = ValidValue
              cboTInsCode.SelStart = 0
              cboTInsCode.SelLength = Len(ValidValue)
              End If
'End If
End Sub

Private Sub cbotplncode_KeyPress(KeyAscii As Integer)
'If secondeditflag = False Then

Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboTPlnCode.Text, cboTPlnCode.SelStart) & _
                Chr(KeyAscii) + Mid(cboTPlnCode.Text, cboTPlnCode.SelStart + cboTPlnCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboTPlnCode.ListCount - 1
            
            If NewText = ChkStr(Left(cboTPlnCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboTPlnCode.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboTPlnCode.Text = ValidValue
              cboTPlnCode.SelStart = 0
              cboTPlnCode.SelLength = Len(ValidValue)
              End If
            
'End If

End Sub

'Private Sub cbotCatCode_KeyPress(KeyAscii As Integer)
'Dim NewText As String
 '   Dim ValidCount As Integer
  '  Dim ValidValue As String
   ' Dim i As Integer
    
'    If KeyAscii > 32 And KeyAscii < 127 Then
 '       NewText = ChkStr(Left(cboTcatCode.Text, cboTcatCode.SelStart) & _
  '              Chr(KeyAscii) + Mid(cboTcatCode.Text, cboTcatCode.SelStart + cboTcatCode.SelLength + 1))
   ''
     '   ValidCount = 0
      '  ValidValue = ""
        
       ' For i = 0 To cboTcatCode.ListCount - 1
            
'            If NewText = ChkStr(Left(cboTcatCode.list(i), Len(NewText))) Then
 '                   ValidCount = ValidCount + 1
  '                  ValidValue = cboTcatCode.list(i)
   '         End If
    '    Next
     '       If ValidCount <= 1 Then KeyAscii = 0
      '      End If
       '     If ValidCount = 1 Then
        '      cboTcatCode.Text = ValidValue
         '     cboTcatCode.SelStart = 0
          '    cboTcatCode.SelLength = Len(ValidValue)
           '   End If
'End Sub


Private Sub SaveBNFList()
'On Error GoTo UpdateErr
Dim BNFList As New ADODB.Recordset
Dim SaveBNFList As New ADODB.Recordset
Dim RecordSaved
Dim strsql As String
Dim GenerateBNFVersion1 As String


If Len(Trim(cboCatCode)) = 0 Then
    MsgBox "Please Enter Benefit sub category Code", vbOKOnly + vbCritical, DialogBoxTitle
    cboCatCode.SetFocus
ElseIf Len(Trim(cboBHLTCode)) = 0 Then
 MsgBox "Please Enter Health Type !", vbOKOnly + vbCritical, DialogBoxTitle
    cboBHLTCode.SetFocus
ElseIf Len(Trim(cboBINSCode)) = 0 Then
    MsgBox "Please Enter Insured Type!", vbOKOnly + vbCritical, DialogBoxTitle
    cboBINSCode.SetFocus
ElseIf Len(Trim(cboBPLNCode)) = 0 Then
    MsgBox "Please Enter Plan Code!", vbOKOnly + vbCritical, DialogBoxTitle
    cboBPLNCode.SetFocus
ElseIf Len(Trim(TxtDay)) = 0 Then
    MsgBox "Please Enter No of Day! ", vbOKOnly + vbCritical, DialogBoxTitle
    TxtDay.SetFocus
ElseIf txtBNFclaimable = "" Then
    MsgBox "Please Enter Annual Limit!", vbOKOnly + vbCritical, DialogBoxTitle
    txtBNFclaimable.SetFocus
ElseIf txtEffdate = "__/__/____" Or IsDate(txtEffdate) = False Then
    MsgBox "Please Enter (Valid )effectice Date!", vbOKOnly + vbCritical, DialogBoxTitle
    txtEffdate.SetFocus
    
Else
  If Trim(txtTempDate) = "" Then
    txtTempDate = txtEffdate
  End If
  strsql = " Select BNFCode From BNF where HLTCode ='" & Left(RemStr(cboBHLTCode), 1) & "'" & _
            " AND PLNCode ='" & RemStr(cboBPLNCode) & "'" & _
           " AND INSCode ='" & Left(cboBINSCode, 1) & "'" & _
           " AND BNFEffDate ='" & Format(txtTempDate, "mm/dd/yyyy") & "'"
  BNFList.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
  
  If BNFList.recordCount > 0 And EditFlag = False Then
        MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
        BNFList.Close
        cboBHLTCode.SetFocus
        cboBHLTCode.SelStart = 0
        cboBHLTCode.SelLength = Len(cboBHLTCode)
        Set BNFList = Nothing
        'Exit Sub
  Else
      RecordSaved = MsgBox("Do you want to save this record?", vbYesNo + vbExclamation)

        If RecordSaved = vbYes Then
            If EditFlag = False Then
            'MembershipNo = Trim(txtPAYCode) & Trim(CboPlnCode) & Trim(cboINSCode) & Mid(BusinessYear, 3, 2) & GenerateMembershipNo(Trim(txtPAYCode), Trim(CboPlnCode), Trim(cboINSCode))

                SaveBNFList.Open "BNF", medixmasterconn, adOpenKeyset, adLockOptimistic
                SaveBNFList.AddNew
                GenerateBNFVersion1 = GenerateBNFVersion(Trim(Mid(cboCatCode, 1, 3)), Trim(Mid(cboBHLTCode, 1, 1)), Trim(Mid(cboBINSCode, 1, 1)), Trim(Mid(cboBPLNCode, 1, 2)), Trim(CDate(txtEffdate)))

            Else
                strsql = "Select * From BNF Where BNFCode = '" & ChkStr(cboCatCode) & "' " & _
                        " AND PLNCode ='" & RemStr(cboBPLNCode) & "' and INSCode ='" & Left(cboBINSCode, 1) & "'" & _
                        " AND BNFEffDate ='" & Format(txtTempDate, "mm/dd/yyyy") & "'"
                SaveBNFList.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic '
            End If
          
            With SaveBNFList
                !BNFCode = ChkStr(Mid(cboCatCode, 1, IIf(InStr(1, cboCatCode, "--") = 0, 10, InStr(1, cboCatCode, "--") - 1)))
                !INSCode = Left(ChkStr(cboBINSCode.Text), 1)
                !PLNCode = Trim(Left(ChkStr(cboBPLNCode.Text), 4))
                !HLTCode = (ChkStr(Mid(cboBHLTCode.Text, 1, 1)))
                !BNFNoOFDays = Left(ChkStr(TxtDay.Text), 3)
                !BNFClaimAbleAmount = ChkStr(txtBNFclaimable)
                !BNFEffDate = txtEffdate
                !BNFEffDateVersion = GenerateBNFVersion1
                !BNFLastUpdateUser = Logon
                !BNFLastUpdateDate = Date
            End With
            SaveBNFList.Update
            Call ClearForm(Me)
            Call BNFListing(BNFCriteria)
         
            If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            End If
   
            SaveBNFList.Close
            Set SaveBNFList = Nothing
       End If
       EditFlag = False
  
   End If
End If

End Sub


Private Sub txtEffDate_LostFocus()
    If IsDate(txtEffdate) Then
    Else
        If txtEffdate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtEffdate.SetFocus
        End If
    End If
End Sub

'Private Sub txtEffdate_Change()
    'txtTempEffDate = txtEffDate
'End Sub

'Private Sub txtEffdate_LostFocus()
    'txtTempEffDate = txtEffDate
'End Sub

Private Sub ClearGrid()
    For FCol = 2 To 4
        FGrid.col = FCol
        For Frow = 1 To 22
            FGrid.row = Frow
            FGrid = ""
        Next
    Next
End Sub

