VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserPermissionMaintenance 
   Caption         =   "Security Maintenance"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7680
      Left            =   240
      TabIndex        =   201
      Top             =   240
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   13547
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Users"
      TabPicture(0)   =   "frmUserPermissionMaintenance.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label23"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstUsersList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtSearch"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Access Control (HIS)"
      TabPicture(1)   =   "frmUserPermissionMaintenance.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Access Control (CIS)"
      TabPicture(2)   =   "frmUserPermissionMaintenance.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   202
         Top             =   3480
         Width           =   840
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
         Height          =   2280
         Left            =   480
         TabIndex        =   204
         Top             =   360
         Width           =   10245
         Begin VB.TextBox txtUSRLGCode 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   1
            Top             =   315
            Width           =   1005
         End
         Begin VB.ComboBox txtAuthStatus 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6600
            TabIndex        =   8
            Text            =   "N"
            Top             =   1800
            Width           =   2055
         End
         Begin VB.ComboBox TxtSRVCode 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1845
            TabIndex        =   5
            Top             =   1800
            Width           =   3090
         End
         Begin VB.TextBox txtUSRDateCreated 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   6600
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2040
         End
         Begin VB.TextBox txtUSRCode 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1845
            MaxLength       =   10
            TabIndex        =   0
            Top             =   315
            Width           =   1005
         End
         Begin VB.ComboBox txtUSRStatus 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6600
            TabIndex        =   7
            Text            =   "A"
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox txtUSRPassword 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1845
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   1080
            Width           =   3120
         End
         Begin VB.TextBox txtUSRName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1845
            MaxLength       =   200
            TabIndex        =   2
            Top             =   720
            Width           =   7755
         End
         Begin VB.ComboBox txtGRPCode 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1845
            TabIndex        =   4
            Top             =   1440
            Width           =   3090
         End
         Begin VB.Label Label39 
            Caption         =   "User LG ID"
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
            Left            =   3240
            TabIndex        =   245
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label33 
            Caption         =   "Auth. Status"
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
            TabIndex        =   243
            Top             =   1800
            Width           =   1125
         End
         Begin VB.Label Label5 
            Caption         =   "Service Provider"
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
            Left            =   585
            TabIndex        =   221
            Top             =   1800
            Width           =   1230
         End
         Begin VB.Label Label10 
            Caption         =   "Status"
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
            TabIndex        =   210
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label Label8 
            Caption         =   "Date Created"
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
            TabIndex        =   209
            Top             =   1125
            Width           =   1185
         End
         Begin VB.Label Label7 
            Caption         =   "Staff Name"
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
            Left            =   585
            TabIndex        =   208
            Top             =   765
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "User ID"
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
            Left            =   600
            TabIndex        =   207
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Department"
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
            Left            =   585
            TabIndex        =   206
            Top             =   1440
            Width           =   1230
         End
         Begin VB.Label Label6 
            Caption         =   "Password"
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
            Left            =   585
            TabIndex        =   205
            Top             =   1125
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         Height          =   675
         Left            =   480
         TabIndex        =   203
         Top             =   2640
         Width           =   10245
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3600
            TabIndex        =   12
            Top             =   225
            Width           =   960
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   600
            TabIndex        =   9
            Top             =   225
            Width           =   960
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1530
            TabIndex        =   10
            Top             =   225
            Width           =   960
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2505
            TabIndex        =   11
            Top             =   225
            Width           =   960
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Sa&ve"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7155
            TabIndex        =   13
            Top             =   225
            Width           =   960
         End
         Begin VB.CommandButton cmdExit 
            Cancel          =   -1  'True
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8145
            TabIndex        =   14
            Top             =   225
            Width           =   960
         End
      End
      Begin MSComctlLib.ListView lstUsersList1 
         Height          =   3240
         Left            =   -74910
         TabIndex        =   211
         Top             =   3750
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   5715
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
      Begin TabDlg.SSTab SSTab2 
         Height          =   7005
         Left            =   -74880
         TabIndex        =   200
         Top             =   480
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   12356
         _Version        =   393216
         TabOrientation  =   2
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmUserPermissionMaintenance.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label21"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label20"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label13"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label12"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label22(0)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label11"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label24"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label29"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label30"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "ACS(8)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "cmdCheckAccept"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "cmdCheckAll"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "ACS(67)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "cmdCheck4"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "ACS(65)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "ACS(64)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "ACS(27)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "ACS(22)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "cmdCheck3"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "ACS(14)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "ACS(13)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "ACS(12)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "cmdCheck2"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "cmdCheck1"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "ACS(2)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "ACS(1)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "ACS(3)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "ACS(4)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "ACS(5)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "ACS(16)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "ACS(17)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "cmdCheck7"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "cmdCheck8"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "ACS(31)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "ACS(32)"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "ACS(41)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "ACS(40)"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "ACS(42)"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "ACS(29)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "ACS(62)"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "ACS(55)"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "ACS(61)"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "ACS(59)"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "ACS(30)"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "ACS(11)"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "ACS(56)"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "ACS(63)"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "ACS(6)"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "ACS(15)"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "ACS(9)"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "ACS(44)"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "ACS(18)"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "ACS(10)"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "ACS(21)"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "ACS(20)"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).Control(56)=   "ACS(19)"
         Tab(0).Control(56).Enabled=   0   'False
         Tab(0).Control(57)=   "ACS(7)"
         Tab(0).Control(57).Enabled=   0   'False
         Tab(0).Control(58)=   "CmdCheck9"
         Tab(0).Control(58).Enabled=   0   'False
         Tab(0).Control(59)=   "ACS(35)"
         Tab(0).Control(59).Enabled=   0   'False
         Tab(0).Control(60)=   "ACS(36)"
         Tab(0).Control(60).Enabled=   0   'False
         Tab(0).Control(61)=   "ACS(37)"
         Tab(0).Control(61).Enabled=   0   'False
         Tab(0).Control(62)=   "ACS(34)"
         Tab(0).Control(62).Enabled=   0   'False
         Tab(0).Control(63)=   "ACS(38)"
         Tab(0).Control(63).Enabled=   0   'False
         Tab(0).Control(64)=   "ACS(39)"
         Tab(0).Control(64).Enabled=   0   'False
         Tab(0).Control(65)=   "ACS(43)"
         Tab(0).Control(65).Enabled=   0   'False
         Tab(0).Control(66)=   "ACS(23)"
         Tab(0).Control(66).Enabled=   0   'False
         Tab(0).Control(67)=   "ACS(26)"
         Tab(0).Control(67).Enabled=   0   'False
         Tab(0).Control(68)=   "ACS(25)"
         Tab(0).Control(68).Enabled=   0   'False
         Tab(0).Control(69)=   "ACS(24)"
         Tab(0).Control(69).Enabled=   0   'False
         Tab(0).Control(70)=   "ACS(33)"
         Tab(0).Control(70).Enabled=   0   'False
         Tab(0).Control(71)=   "ACS(45)"
         Tab(0).Control(71).Enabled=   0   'False
         Tab(0).Control(72)=   "ACS(46)"
         Tab(0).Control(72).Enabled=   0   'False
         Tab(0).Control(73)=   "ACS(47)"
         Tab(0).Control(73).Enabled=   0   'False
         Tab(0).Control(74)=   "ACS(48)"
         Tab(0).Control(74).Enabled=   0   'False
         Tab(0).Control(75)=   "ACS(50)"
         Tab(0).Control(75).Enabled=   0   'False
         Tab(0).Control(76)=   "ACS(51)"
         Tab(0).Control(76).Enabled=   0   'False
         Tab(0).Control(77)=   "ACS(52)"
         Tab(0).Control(77).Enabled=   0   'False
         Tab(0).Control(78)=   "ACS(53)"
         Tab(0).Control(78).Enabled=   0   'False
         Tab(0).Control(79)=   "CmdUncheck"
         Tab(0).Control(79).Enabled=   0   'False
         Tab(0).Control(80)=   "ACS(66)"
         Tab(0).Control(80).Enabled=   0   'False
         Tab(0).Control(81)=   "ACS(68)"
         Tab(0).Control(81).Enabled=   0   'False
         Tab(0).Control(82)=   "ACS(54)"
         Tab(0).Control(82).Enabled=   0   'False
         Tab(0).Control(83)=   "ACS(28)"
         Tab(0).Control(83).Enabled=   0   'False
         Tab(0).Control(84)=   "ACS(49)"
         Tab(0).Control(84).Enabled=   0   'False
         Tab(0).Control(85)=   "ACS(57)"
         Tab(0).Control(85).Enabled=   0   'False
         Tab(0).Control(86)=   "ACS(60)"
         Tab(0).Control(86).Enabled=   0   'False
         Tab(0).Control(87)=   "ACS(145)"
         Tab(0).Control(87).Enabled=   0   'False
         Tab(0).Control(88)=   "ACS(146)"
         Tab(0).Control(88).Enabled=   0   'False
         Tab(0).Control(89)=   "ACS(148)"
         Tab(0).Control(89).Enabled=   0   'False
         Tab(0).Control(90)=   "ACS(149)"
         Tab(0).Control(90).Enabled=   0   'False
         Tab(0).Control(91)=   "ACS(150)"
         Tab(0).Control(91).Enabled=   0   'False
         Tab(0).Control(92)=   "ACS(152)"
         Tab(0).Control(92).Enabled=   0   'False
         Tab(0).Control(93)=   "ACS(151)"
         Tab(0).Control(93).Enabled=   0   'False
         Tab(0).Control(94)=   "ACS(154)"
         Tab(0).Control(94).Enabled=   0   'False
         Tab(0).Control(95)=   "ACS(156)"
         Tab(0).Control(95).Enabled=   0   'False
         Tab(0).Control(96)=   "ACS(155)"
         Tab(0).Control(96).Enabled=   0   'False
         Tab(0).Control(97)=   "ACS(157)"
         Tab(0).Control(97).Enabled=   0   'False
         Tab(0).Control(98)=   "ACS(58)"
         Tab(0).Control(98).Enabled=   0   'False
         Tab(0).Control(99)=   "ACS(169)"
         Tab(0).Control(99).Enabled=   0   'False
         Tab(0).Control(100)=   "ACS(170)"
         Tab(0).Control(100).Enabled=   0   'False
         Tab(0).Control(101)=   "ACS(171)"
         Tab(0).Control(101).Enabled=   0   'False
         Tab(0).Control(102)=   "ACS(172)"
         Tab(0).Control(102).Enabled=   0   'False
         Tab(0).Control(103)=   "ACS(173)"
         Tab(0).Control(103).Enabled=   0   'False
         Tab(0).Control(104)=   "ACS(174)"
         Tab(0).Control(104).Enabled=   0   'False
         Tab(0).Control(105)=   "ACS(175)"
         Tab(0).Control(105).Enabled=   0   'False
         Tab(0).Control(106)=   "ACS(176)"
         Tab(0).Control(106).Enabled=   0   'False
         Tab(0).Control(107)=   "ACS(177)"
         Tab(0).Control(107).Enabled=   0   'False
         Tab(0).Control(108)=   "ACS(178)"
         Tab(0).Control(108).Enabled=   0   'False
         Tab(0).ControlCount=   109
         TabCaption(1)   =   "Maintenance"
         TabPicture(1)   =   "frmUserPermissionMaintenance.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ACS(168)"
         Tab(1).Control(1)=   "ACS(159)"
         Tab(1).Control(2)=   "ACS(158)"
         Tab(1).Control(3)=   "ACS(147)"
         Tab(1).Control(4)=   "ACS(110)"
         Tab(1).Control(5)=   "ACS(79)"
         Tab(1).Control(6)=   "ACS(78)"
         Tab(1).Control(7)=   "ACS(112)"
         Tab(1).Control(8)=   "ACS(83)"
         Tab(1).Control(9)=   "ACS(69)"
         Tab(1).Control(10)=   "ACS(111)"
         Tab(1).Control(11)=   "ACS(86)"
         Tab(1).Control(12)=   "ACS(80)"
         Tab(1).Control(13)=   "ACS(88)"
         Tab(1).Control(14)=   "ACS(87)"
         Tab(1).Control(15)=   "ACS(85)"
         Tab(1).Control(16)=   "ACS(84)"
         Tab(1).Control(17)=   "ACS(90)"
         Tab(1).Control(18)=   "ACS(89)"
         Tab(1).Control(19)=   "ACS(108)"
         Tab(1).Control(20)=   "ACS(107)"
         Tab(1).Control(21)=   "ACS(103)"
         Tab(1).Control(22)=   "ACS(104)"
         Tab(1).Control(23)=   "ACS(105)"
         Tab(1).Control(24)=   "ACS(106)"
         Tab(1).Control(25)=   "ACS(82)"
         Tab(1).Control(26)=   "ACS(81)"
         Tab(1).Control(27)=   "ACS(75)"
         Tab(1).Control(28)=   "ACS(76)"
         Tab(1).Control(29)=   "ACS(77)"
         Tab(1).Control(30)=   "ACS(97)"
         Tab(1).Control(31)=   "ACS(109)"
         Tab(1).Control(32)=   "ACS(98)"
         Tab(1).Control(33)=   "ACS(99)"
         Tab(1).Control(34)=   "ACS(102)"
         Tab(1).Control(35)=   "ACS(101)"
         Tab(1).Control(36)=   "ACS(100)"
         Tab(1).Control(37)=   "cmdCheck5"
         Tab(1).Control(38)=   "ACS(91)"
         Tab(1).Control(39)=   "ACS(74)"
         Tab(1).Control(40)=   "ACS(70)"
         Tab(1).Control(41)=   "ACS(92)"
         Tab(1).Control(42)=   "ACS(71)"
         Tab(1).Control(43)=   "ACS(93)"
         Tab(1).Control(44)=   "ACS(73)"
         Tab(1).Control(45)=   "ACS(96)"
         Tab(1).Control(46)=   "ACS(95)"
         Tab(1).Control(47)=   "ACS(94)"
         Tab(1).Control(48)=   "ACS(72)"
         Tab(1).Control(49)=   "Label40"
         Tab(1).Control(50)=   "Label38"
         Tab(1).Control(51)=   "Label15"
         Tab(1).Control(52)=   "Label28"
         Tab(1).Control(53)=   "Label26"
         Tab(1).Control(54)=   "Label25"
         Tab(1).Control(55)=   "Label19"
         Tab(1).Control(56)=   "Label17"
         Tab(1).Control(57)=   "Label16"
         Tab(1).Control(58)=   "Label18"
         Tab(1).Control(59)=   "Label14"
         Tab(1).ControlCount=   60
         Begin VB.CheckBox ACS 
            Caption         =   "Print Payment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   178
            Left            =   420
            TabIndex        =   258
            Top             =   4320
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Sales - Invoice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   177
            Left            =   9360
            TabIndex        =   257
            Top             =   2955
            Width           =   1605
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Bordx - Payor"
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
            Index           =   176
            Left            =   9360
            TabIndex        =   256
            Top             =   2760
            Width           =   1605
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Bordx - MNRB"
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
            Index           =   175
            Left            =   9360
            TabIndex        =   255
            Top             =   2520
            Width           =   1605
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Payment - MBM"
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
            Index           =   174
            Left            =   9360
            TabIndex        =   254
            Top             =   2280
            Width           =   1605
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Payment - HOS"
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
            Index           =   173
            Left            =   9360
            TabIndex        =   253
            Top             =   2040
            Width           =   1605
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Receipt - Payor"
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
            Index           =   172
            Left            =   9360
            TabIndex        =   252
            Top             =   1800
            Width           =   1605
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Receipt - MNRB"
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
            Index           =   171
            Left            =   9360
            TabIndex        =   251
            Top             =   1560
            Width           =   1605
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Invoice - MBM"
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
            Index           =   170
            Left            =   9360
            TabIndex        =   250
            Top             =   1320
            Width           =   1605
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Invoice - HOS"
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
            Index           =   169
            Left            =   9360
            TabIndex        =   249
            Top             =   1080
            Width           =   1605
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Payment Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   58
            Left            =   7380
            TabIndex        =   85
            Top             =   1320
            Width           =   1755
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Data Transfer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   168
            Left            =   -67620
            TabIndex        =   248
            Top             =   3960
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   159
            Left            =   -67680
            TabIndex        =   150
            Top             =   5640
            Width           =   1770
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Listing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   158
            Left            =   -67680
            TabIndex        =   149
            Top             =   5400
            Width           =   1770
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Trust Account"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   157
            Left            =   7380
            TabIndex        =   93
            Top             =   3160
            Width           =   1515
         End
         Begin VB.CheckBox ACS 
            Caption         =   "DCU Report 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   155
            Left            =   7380
            TabIndex        =   79
            Top             =   5490
            Width           =   1785
         End
         Begin VB.CheckBox ACS 
            Caption         =   "DCU Report 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   156
            Left            =   7380
            TabIndex        =   80
            Top             =   5715
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Send to Claims"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   154
            Left            =   2280
            TabIndex        =   35
            Top             =   3840
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Storage Rep Case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   151
            Left            =   5640
            TabIndex        =   72
            Top             =   4635
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Storage Rep Enq"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   152
            Left            =   5640
            TabIndex        =   73
            Top             =   4815
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "MMA Fees"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   150
            Left            =   2280
            TabIndex        =   36
            Top             =   4080
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "NIAM Report"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   149
            Left            =   2160
            TabIndex        =   25
            Top             =   2040
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Invoice Manual Upd"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   148
            Left            =   420
            TabIndex        =   41
            Top             =   5520
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "MMA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   147
            Left            =   -67680
            TabIndex        =   148
            Top             =   4725
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Enq -  Pay2MBM"
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
            Index           =   146
            Left            =   7380
            TabIndex        =   92
            Top             =   2940
            Width           =   1635
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Enq -  Pay2Hos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   145
            Left            =   7380
            TabIndex        =   91
            Top             =   2740
            Width           =   1635
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Suspension"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   110
            Left            =   -67620
            TabIndex        =   145
            Top             =   2085
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Special Approval"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   79
            Left            =   -73920
            TabIndex        =   114
            Top             =   3360
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Agent"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   78
            Left            =   -73920
            TabIndex        =   113
            Top             =   3120
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Receipt fr Others"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   60
            Left            =   7380
            TabIndex        =   87
            Top             =   1740
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Update Invoice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   57
            Left            =   7380
            TabIndex        =   84
            Top             =   1050
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Storage Enquiry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   49
            Left            =   5640
            TabIndex        =   71
            Top             =   4395
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking (Jupiter)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   28
            Left            =   5640
            TabIndex        =   49
            Top             =   1200
            Width           =   1635
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Date Selection"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   54
            Left            =   5640
            TabIndex        =   78
            Top             =   5970
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking (Jupiter)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   68
            Left            =   7380
            TabIndex        =   99
            Top             =   5040
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Send to A/C"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   66
            Left            =   7380
            TabIndex        =   97
            Top             =   4560
            Width           =   2265
         End
         Begin VB.CommandButton CmdUncheck 
            Height          =   255
            Left            =   7920
            TabIndex        =   101
            Top             =   6000
            Width           =   255
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking 4 (Jup.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   53
            Left            =   5640
            TabIndex        =   77
            Top             =   5805
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking 3 (Jup.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   52
            Left            =   5640
            TabIndex        =   76
            Top             =   5565
            Width           =   1785
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking 2 (Jup.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   51
            Left            =   5640
            TabIndex        =   75
            Top             =   5325
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking 1 (Jup.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   50
            Left            =   5640
            TabIndex        =   74
            Top             =   5085
            Width           =   1785
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Storage Entry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   48
            Left            =   5640
            TabIndex        =   70
            Top             =   4200
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "CLM Tracking 4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   47
            Left            =   5640
            TabIndex        =   69
            Top             =   3960
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "CLM Tracking 3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   46
            Left            =   5640
            TabIndex        =   68
            Top             =   3720
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "CLM Tracking 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   45
            Left            =   5640
            TabIndex        =   67
            Top             =   3480
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Print Bordx Excess"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   33
            Left            =   3900
            TabIndex        =   55
            Top             =   3570
            Width           =   1905
         End
         Begin VB.CheckBox ACS 
            Caption         =   "HOS Accounting"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   24
            Left            =   3900
            TabIndex        =   45
            Top             =   1080
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "HOS (DR to MBM)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   3900
            TabIndex        =   46
            Top             =   1320
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "HOS (CR to MBM)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   26
            Left            =   3900
            TabIndex        =   47
            Top             =   1560
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "MBM Accounting"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   3900
            TabIndex        =   44
            Top             =   840
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Bordx Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   43
            Left            =   3900
            TabIndex        =   65
            Top             =   6000
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "A/C Period"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   39
            Left            =   3900
            TabIndex        =   61
            Top             =   5040
            Width           =   1785
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Send to Accounts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   38
            Left            =   3900
            TabIndex        =   60
            Top             =   4800
            Width           =   1785
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Update Bordx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   34
            Left            =   3900
            TabIndex        =   56
            Top             =   3840
            Width           =   1785
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Send to MNRB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   37
            Left            =   3900
            TabIndex        =   59
            Top             =   4560
            Width           =   1785
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Returned Fr Payor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   36
            Left            =   3900
            TabIndex        =   58
            Top             =   4320
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Send to Payor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   35
            Left            =   3900
            TabIndex        =   57
            Top             =   4080
            Width           =   1545
         End
         Begin VB.CommandButton CmdCheck9 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2655
            TabIndex        =   42
            Top             =   4785
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking (Jupiter)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   420
            TabIndex        =   21
            Top             =   2040
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Invoice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   19
            Left            =   420
            TabIndex        =   38
            Top             =   4815
            Width           =   2235
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Fee Listing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   420
            TabIndex        =   39
            Top             =   5040
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Refund Listing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   21
            Left            =   420
            TabIndex        =   40
            Top             =   5280
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Upload Bordx by IC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   2160
            TabIndex        =   24
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claims ASCII"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   2280
            TabIndex        =   34
            Top             =   3600
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "CLM Tracking 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   44
            Left            =   5640
            TabIndex        =   66
            Top             =   3240
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Generate ASCII"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   2160
            TabIndex        =   23
            Top             =   1560
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking (Jupiter)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   420
            TabIndex        =   31
            Top             =   3855
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   420
            TabIndex        =   20
            Top             =   1725
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "User Permission"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   112
            Left            =   -67620
            TabIndex        =   147
            Top             =   3720
            Width           =   1890
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Premium"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   83
            Left            =   -73920
            TabIndex        =   118
            Top             =   4680
            Width           =   1890
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Receipt Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   63
            Left            =   7380
            TabIndex        =   90
            Top             =   2510
            Width           =   1650
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Access"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   69
            Left            =   -73920
            TabIndex        =   104
            Top             =   1000
            Width           =   1890
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Payment to MBM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   56
            Left            =   7380
            TabIndex        =   83
            Top             =   840
            Width           =   1755
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Registration"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   420
            TabIndex        =   27
            Top             =   2835
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Insert Worksheet"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   30
            Left            =   3900
            TabIndex        =   52
            Top             =   2835
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Receipt fr MBM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   59
            Left            =   7380
            TabIndex        =   86
            Top             =   1500
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Receipt fr Payor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   61
            Left            =   7380
            TabIndex        =   88
            Top             =   1980
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Payment to HOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   55
            Left            =   7380
            TabIndex        =   82
            Top             =   600
            Width           =   1890
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Receipt fr MNRB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   62
            Left            =   7380
            TabIndex        =   89
            Top             =   2220
            Width           =   1635
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Main DCN Checking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   29
            Left            =   3900
            TabIndex        =   50
            Top             =   1800
            Width           =   2115
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Bordx && Worksheet"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   42
            Left            =   3900
            TabIndex        =   64
            Top             =   5760
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Bordx Control"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   40
            Left            =   3900
            TabIndex        =   62
            Top             =   5280
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Outstanding Bordx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   41
            Left            =   3900
            TabIndex        =   63
            Top             =   5520
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Print Bordx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   32
            Left            =   3900
            TabIndex        =   54
            Top             =   3345
            Width           =   1305
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Department"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   111
            Left            =   -67620
            TabIndex        =   146
            Top             =   3500
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Remove Worksheet"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   31
            Left            =   3900
            TabIndex        =   53
            Top             =   3090
            Width           =   1995
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   86
            Left            =   -73920
            TabIndex        =   121
            Top             =   5760
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Health Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   80
            Left            =   -73920
            TabIndex        =   115
            Top             =   4005
            Width           =   2265
         End
         Begin VB.CommandButton cmdCheck8 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9600
            TabIndex        =   94
            Top             =   550
            Width           =   780
         End
         Begin VB.CommandButton cmdCheck7 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9600
            TabIndex        =   100
            Top             =   4035
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "LG Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   2280
            TabIndex        =   33
            Top             =   3360
            Width           =   1305
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Audit Review"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   420
            TabIndex        =   32
            Top             =   4080
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Deletion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   420
            TabIndex        =   19
            Top             =   1500
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Leaflet && Card"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   420
            TabIndex        =   18
            Top             =   1260
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Enquiry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   420
            TabIndex        =   17
            Top             =   1050
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Benefit Entry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   88
            Left            =   -73920
            TabIndex        =   123
            Top             =   6195
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Sub Category"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   87
            Left            =   -73920
            TabIndex        =   122
            Top             =   5955
            Width           =   1665
         End
         Begin VB.CheckBox ACS 
            Caption         =   "MCO Fee"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   85
            Left            =   -73920
            TabIndex        =   120
            Top             =   5115
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Annual Limit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   84
            Left            =   -73920
            TabIndex        =   119
            Top             =   4875
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "ICD Selection"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   90
            Left            =   -70740
            TabIndex        =   125
            Top             =   1000
            Width           =   1890
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Benefit Listing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   89
            Left            =   -73920
            TabIndex        =   124
            Top             =   6480
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Policy Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   108
            Left            =   -70740
            TabIndex        =   143
            Top             =   5865
            Width           =   1905
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Endorsement Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   107
            Left            =   -70740
            TabIndex        =   142
            Top             =   5610
            Width           =   1905
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Language"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   103
            Left            =   -70740
            TabIndex        =   138
            Top             =   4695
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "City"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   104
            Left            =   -70740
            TabIndex        =   139
            Top             =   4920
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "State"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   105
            Left            =   -70740
            TabIndex        =   140
            Top             =   5130
            Width           =   1905
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Work Nature"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   106
            Left            =   -70740
            TabIndex        =   141
            Top             =   5385
            Width           =   1905
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Plan Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   82
            Left            =   -73920
            TabIndex        =   117
            Top             =   4440
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Insured Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   81
            Left            =   -73920
            TabIndex        =   116
            Top             =   4200
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Hospital Group"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   75
            Left            =   -73920
            TabIndex        =   110
            Top             =   2355
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Clinic Group"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   76
            Left            =   -73920
            TabIndex        =   111
            Top             =   2595
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Payor Group"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   77
            Left            =   -73920
            TabIndex        =   112
            Top             =   2835
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Hospitalisation Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   97
            Left            =   -70740
            TabIndex        =   132
            Top             =   2670
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Valid Date Range"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   109
            Left            =   -67620
            TabIndex        =   144
            Top             =   1000
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Age Band"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   98
            Left            =   -70740
            TabIndex        =   133
            Top             =   3525
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Occupation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   99
            Left            =   -70740
            TabIndex        =   134
            Top             =   3720
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Race"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   102
            Left            =   -70740
            TabIndex        =   137
            Top             =   4440
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Nationality"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   101
            Left            =   -70740
            TabIndex        =   136
            Top             =   4200
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Relationship"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   100
            Left            =   -70740
            TabIndex        =   135
            Top             =   3915
            Width           =   1905
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Registration"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   420
            TabIndex        =   15
            Top             =   555
            Width           =   2235
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Adjustment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   420
            TabIndex        =   16
            Top             =   795
            Width           =   2265
         End
         Begin VB.CommandButton cmdCheck1 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2655
            TabIndex        =   26
            Top             =   525
            Width           =   780
         End
         Begin VB.CommandButton cmdCheck2 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2655
            TabIndex        =   37
            Top             =   2805
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Adjustment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   420
            TabIndex        =   28
            Top             =   3090
            Width           =   1995
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Inquiry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   420
            TabIndex        =   29
            Top             =   3345
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   420
            TabIndex        =   30
            Top             =   3585
            Width           =   2265
         End
         Begin VB.CommandButton cmdCheck3 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6120
            TabIndex        =   51
            Top             =   525
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Processing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   22
            Left            =   3900
            TabIndex        =   43
            Top             =   555
            Width           =   1905
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   27
            Left            =   5640
            TabIndex        =   48
            Top             =   960
            Width           =   1185
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Registration"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   64
            Left            =   7380
            TabIndex        =   95
            Top             =   4035
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Adjustment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   65
            Left            =   7380
            TabIndex        =   96
            Top             =   4275
            Width           =   2265
         End
         Begin VB.CommandButton cmdCheck4 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6165
            TabIndex        =   81
            Top             =   2805
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   67
            Left            =   7380
            TabIndex        =   98
            Top             =   4800
            Width           =   2265
         End
         Begin VB.CommandButton cmdCheck5 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   -65700
            TabIndex        =   151
            Top             =   5760
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "ICD Maintenance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   91
            Left            =   -70740
            TabIndex        =   126
            Top             =   1200
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Service Provider"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   74
            Left            =   -73920
            TabIndex        =   109
            Top             =   2115
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Hospital Info"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   70
            Left            =   -73920
            TabIndex        =   105
            Top             =   1200
            Width           =   1785
         End
         Begin VB.CheckBox ACS 
            Caption         =   "ICD Category"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   92
            Left            =   -70740
            TabIndex        =   127
            Top             =   1485
            Width           =   1425
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Clinic Info"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   71
            Left            =   -73920
            TabIndex        =   106
            Top             =   1425
            Width           =   1785
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Repudiation Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   93
            Left            =   -70740
            TabIndex        =   128
            Top             =   1710
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Payor Branch"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   73
            Left            =   -73920
            TabIndex        =   108
            Top             =   1890
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Treatment Modality"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   96
            Left            =   -70740
            TabIndex        =   131
            Top             =   2400
            Width           =   1905
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Nature of Claim"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   95
            Left            =   -70740
            TabIndex        =   130
            Top             =   2160
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Repudiation Category"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   94
            Left            =   -70740
            TabIndex        =   129
            Top             =   1935
            Width           =   2145
         End
         Begin VB.CommandButton cmdCheckAll 
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
            Left            =   7920
            TabIndex        =   102
            Top             =   6240
            Width           =   255
         End
         Begin VB.CommandButton cmdCheckAccept 
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
            Left            =   7920
            TabIndex        =   103
            Top             =   6480
            Width           =   255
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Payor Info"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   72
            Left            =   -73920
            TabIndex        =   107
            Top             =   1695
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Upload Bordx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   2160
            TabIndex        =   22
            Top             =   1320
            Width           =   1665
         End
         Begin VB.Label Label40 
            BackColor       =   &H00FFFF80&
            Caption         =   "Print Control"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -67680
            TabIndex        =   246
            Top             =   5160
            Width           =   2670
         End
         Begin VB.Label Label38 
            BackColor       =   &H00FFFF80&
            Caption         =   " Maintenance MMA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -67680
            TabIndex        =   242
            Top             =   4440
            Width           =   2670
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFF80&
            Caption         =   " Maintenance Suspension"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -67620
            TabIndex        =   233
            Top             =   1800
            Width           =   2670
         End
         Begin VB.Label Label30 
            BackColor       =   &H008080FF&
            Caption         =   "MCO Fee"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   420
            TabIndex        =   232
            Top             =   4560
            Width           =   3030
         End
         Begin VB.Label Label29 
            Caption         =   "Uncheck All"
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
            Left            =   8280
            TabIndex        =   199
            Top             =   6000
            Width           =   1215
         End
         Begin VB.Label Label28 
            BackColor       =   &H00FFFF80&
            Caption         =   " Maintenance Security"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -67620
            TabIndex        =   230
            Top             =   3240
            Width           =   2670
         End
         Begin VB.Label Label26 
            BackColor       =   &H00FFFF80&
            Caption         =   " Maintenance Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -67620
            TabIndex        =   229
            Top             =   720
            Width           =   2670
         End
         Begin VB.Label Label25 
            BackColor       =   &H00FFFF80&
            Caption         =   " Maintenance Others"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -70740
            TabIndex        =   228
            Top             =   3240
            Width           =   2670
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FFFF80&
            Caption         =   " Maintenance Case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -70740
            TabIndex        =   227
            Top             =   720
            Width           =   2670
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFF80&
            Caption         =   " Maintenance Benefits"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -73920
            TabIndex        =   226
            Top             =   5520
            Width           =   2670
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFF80&
            Caption         =   " Maintenance Plan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -73920
            TabIndex        =   225
            Top             =   3720
            Width           =   2670
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFFF00&
            Caption         =   " Maintenance Entity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -73920
            TabIndex        =   224
            Top             =   720
            Width           =   2670
         End
         Begin VB.Label Label24 
            BackColor       =   &H008080FF&
            Caption         =   " Accounts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7380
            TabIndex        =   223
            Top             =   300
            Width           =   3015
         End
         Begin VB.Label Label11 
            BackColor       =   &H008080FF&
            Caption         =   " Case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   420
            TabIndex        =   219
            Top             =   2580
            Width           =   3030
         End
         Begin VB.Label Label22 
            BackColor       =   &H008080FF&
            Caption         =   " Worksheet"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   7380
            TabIndex        =   222
            Top             =   3780
            Width           =   3015
         End
         Begin VB.Label Label1 
            BackColor       =   &H008080FF&
            Caption         =   " Membership"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   420
            TabIndex        =   220
            Top             =   300
            Width           =   3030
         End
         Begin VB.Label Label12 
            BackColor       =   &H008080FF&
            Caption         =   " Invoice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3900
            TabIndex        =   218
            Top             =   300
            Width           =   3030
         End
         Begin VB.Label Label13 
            BackColor       =   &H008080FF&
            Caption         =   " Claim Bordx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3900
            TabIndex        =   217
            Top             =   2580
            Width           =   3030
         End
         Begin VB.Label Label14 
            BackColor       =   &H008080FF&
            Caption         =   " Maintenance Module"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -73920
            TabIndex        =   216
            Top             =   180
            Width           =   4395
         End
         Begin VB.Label Label20 
            Caption         =   "Check All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8280
            TabIndex        =   215
            Top             =   6240
            Width           =   870
         End
         Begin VB.Label Label21 
            Caption         =   "Check All Except Maintenance"
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
            Left            =   8280
            TabIndex        =   214
            Top             =   6480
            Width           =   2580
         End
      End
      Begin MSComctlLib.ListView lstUsersList 
         Height          =   3480
         Left            =   480
         TabIndex        =   231
         Top             =   3840
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   6138
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
         NumItems        =   7
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
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "User LG ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Department"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Service Provider"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Authorised Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   234
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   11668
         _Version        =   393216
         TabOrientation  =   2
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmUserPermissionMaintenance.frx":008C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label36"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label35"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label34"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label32"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label31"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label27"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label37"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "ACS(134)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "ACS(140)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "ACS(138)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "ACS(139)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "cmdCheck15"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "ACS(141)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "ACS(143)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "ACS(136)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "cmdCheck14"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "ACS(135)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "ACS(133)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "ACS(142)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "ACS(130)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "ACS(129)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "ACS(131)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "ACS(128)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "ACS(127)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "ACS(126)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "ACS(125)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "ACS(137)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "cmdCheck13"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "cmdCheck11"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "ACS(123)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "cmdCheck10"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "ACS(116)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "ACS(115)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "ACS(117)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "ACS(120)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "ACS(121)"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "ACS(122)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "ACS(124)"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "cmdCheck16"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "ACS(114)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "ACS(113)"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "ACS(118)"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "ACS(119)"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "cmdCheck17"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "ACS(144)"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "ACS(132)"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "ACS(153)"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "ACS(160)"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "ACS(161)"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "ACS(162)"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "ACS(163)"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "ACS(164)"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "ACS(165)"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "ACS(167)"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "ACS(166)"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "ACS(179)"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).ControlCount=   56
         Begin VB.CheckBox ACS 
            Caption         =   "Invoice  After Bordx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   179
            Left            =   4140
            TabIndex        =   259
            Top             =   3960
            Width           =   2235
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claim Tracking (Jupiter)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   166
            Left            =   7800
            TabIndex        =   247
            Top             =   8760
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking (Jupiter)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   167
            Left            =   4140
            TabIndex        =   187
            Top             =   6200
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   165
            Left            =   4140
            TabIndex        =   186
            Top             =   5940
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claim Tracking 2 (Jupiter)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   164
            Left            =   4140
            TabIndex        =   178
            Top             =   3675
            Width           =   2625
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claim Tracking (Jupiter)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   163
            Left            =   4140
            TabIndex        =   177
            Top             =   3480
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking (Jupiter)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   162
            Left            =   420
            TabIndex        =   163
            Top             =   3600
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   161
            Left            =   420
            TabIndex        =   162
            Top             =   3360
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking (Jupiter)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   160
            Left            =   420
            TabIndex        =   154
            Top             =   960
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "To Member - Batch"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   153
            Left            =   4140
            TabIndex        =   183
            Top             =   5160
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claim Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   132
            Left            =   4140
            TabIndex        =   175
            Top             =   3000
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Clinic"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   144
            Left            =   7680
            TabIndex        =   197
            Top             =   2775
            Width           =   2235
         End
         Begin VB.CommandButton cmdCheck17 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9915
            TabIndex        =   198
            Top             =   2745
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Adjustment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   119
            Left            =   420
            TabIndex        =   161
            Top             =   3120
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Member"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   118
            Left            =   420
            TabIndex        =   160
            Top             =   2880
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Enquiry"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   113
            Left            =   420
            TabIndex        =   152
            Top             =   495
            Width           =   2235
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   114
            Left            =   420
            TabIndex        =   153
            Top             =   720
            Width           =   2265
         End
         Begin VB.CommandButton cmdCheck16 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2655
            TabIndex        =   155
            Top             =   465
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Print Bordx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   124
            Left            =   4140
            TabIndex        =   167
            Top             =   960
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Insert Worksheet"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   122
            Left            =   4140
            TabIndex        =   165
            Top             =   480
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "To Clinic - Batch"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   121
            Left            =   4140
            TabIndex        =   182
            Top             =   4920
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Claim Tracking 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   120
            Left            =   4140
            TabIndex        =   176
            Top             =   3200
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Clinic"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   117
            Left            =   420
            TabIndex        =   159
            Top             =   2640
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "MCO Invoice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   115
            Left            =   420
            TabIndex        =   156
            Top             =   1560
            Width           =   2235
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Fee Listing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   116
            Left            =   420
            TabIndex        =   157
            Top             =   1815
            Width           =   2265
         End
         Begin VB.CommandButton cmdCheck10 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2655
            TabIndex        =   158
            Top             =   1545
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Remove Worksheet"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   123
            Left            =   4140
            TabIndex        =   166
            Top             =   765
            Width           =   2145
         End
         Begin VB.CommandButton cmdCheck11 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2655
            TabIndex        =   164
            Top             =   2520
            Width           =   780
         End
         Begin VB.CommandButton cmdCheck13 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6375
            TabIndex        =   179
            Top             =   465
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Entity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   137
            Left            =   7680
            TabIndex        =   189
            Top             =   480
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Send to Payor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   125
            Left            =   4140
            TabIndex        =   168
            Top             =   1200
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Returned Fr Payor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   126
            Left            =   4140
            TabIndex        =   169
            Top             =   1500
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Send to Accounts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   127
            Left            =   4140
            TabIndex        =   170
            Top             =   1750
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "A/C Period"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   128
            Left            =   4140
            TabIndex        =   171
            Top             =   2040
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Bordx && Worksheet"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   131
            Left            =   4140
            TabIndex        =   174
            Top             =   2760
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Bordx Control"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   129
            Left            =   4140
            TabIndex        =   172
            Top             =   2280
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Outstanding Bordx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   130
            Left            =   4140
            TabIndex        =   173
            Top             =   2520
            Width           =   2265
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Benefit - Clinic"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   142
            Left            =   7680
            TabIndex        =   194
            Top             =   1680
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "To Clinic"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   133
            Left            =   4140
            TabIndex        =   180
            Top             =   4440
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Fr Member"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   135
            Left            =   4140
            TabIndex        =   184
            Top             =   5445
            Width           =   2145
         End
         Begin VB.CommandButton cmdCheck14 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6360
            TabIndex        =   188
            Top             =   4425
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Fr Payor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   136
            Left            =   4140
            TabIndex        =   185
            Top             =   5640
            Width           =   1905
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Check Item"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   143
            Left            =   7680
            TabIndex        =   195
            Top             =   1920
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Benefit - Member"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   141
            Left            =   7680
            TabIndex        =   193
            Top             =   1440
            Width           =   2265
         End
         Begin VB.CommandButton cmdCheck15 
            Caption         =   "Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9915
            TabIndex        =   196
            Top             =   465
            Width           =   780
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Case"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   139
            Left            =   7680
            TabIndex        =   191
            Top             =   1005
            Width           =   2145
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Plan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   138
            Left            =   7680
            TabIndex        =   190
            Top             =   720
            Width           =   2025
         End
         Begin VB.CheckBox ACS 
            Caption         =   "Calculation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   140
            Left            =   7680
            TabIndex        =   192
            Top             =   1260
            Width           =   1545
         End
         Begin VB.CheckBox ACS 
            Caption         =   "To Member"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   134
            Left            =   4140
            TabIndex        =   181
            Top             =   4680
            Width           =   1905
         End
         Begin VB.Label Label37 
            BackColor       =   &H008080FF&
            Caption         =   "Tracking"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7680
            TabIndex        =   241
            Top             =   2520
            Width           =   3030
         End
         Begin VB.Label Label27 
            BackColor       =   &H008080FF&
            Caption         =   " Membership"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   420
            TabIndex        =   240
            Top             =   240
            Width           =   3030
         End
         Begin VB.Label Label31 
            BackColor       =   &H008080FF&
            Caption         =   " MCO Fee"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   420
            TabIndex        =   239
            Top             =   1320
            Width           =   3030
         End
         Begin VB.Label Label32 
            BackColor       =   &H008080FF&
            Caption         =   " Invoice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   420
            TabIndex        =   238
            Top             =   2280
            Width           =   3030
         End
         Begin VB.Label Label34 
            BackColor       =   &H008080FF&
            Caption         =   " Claim Bordx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4140
            TabIndex        =   237
            Top             =   240
            Width           =   3030
         End
         Begin VB.Label Label35 
            BackColor       =   &H008080FF&
            Caption         =   " Accounts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4140
            TabIndex        =   236
            Top             =   4200
            Width           =   3030
         End
         Begin VB.Label Label36 
            BackColor       =   &H008080FF&
            Caption         =   " Maintenance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7680
            TabIndex        =   235
            Top             =   240
            Width           =   3030
         End
      End
      Begin VB.Label Label23 
         Caption         =   "Search For Staff Id"
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
         Left            =   1920
         TabIndex        =   213
         Top             =   3525
         Width           =   1770
      End
      Begin VB.Label Label9 
         Caption         =   "Search For Staff Id"
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
         Left            =   -74100
         TabIndex        =   212
         Top             =   3435
         Width           =   1770
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Security Maintenance - Access Control"
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
      TabIndex        =   244
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmUserPermissionMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim tmpEncryptPassword As String

Private Sub cmdCheck16_Click()
    ACS(113).Value = 1
    ACS(114).Value = 1
    ACS(160).Value = 1
End Sub

Private Sub cmdCheck17_Click()
    ACS(144).Value = 1
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

Private Sub cmdCheck1_Click()
Dim AA As Integer
    For AA = 1 To 10
        ACS(AA).Value = 1
    Next AA
    ACS(149).Value = 1
End Sub

Private Sub cmdCheck2_Click()
Dim AA As Integer
    For AA = 11 To 18
        ACS(AA).Value = 1
    Next AA
    ACS(150).Value = 1
    ACS(154).Value = 1
End Sub

Private Sub cmdCheck3_Click()
Dim AA As Integer
    For AA = 22 To 29
        ACS(AA).Value = 1
    Next AA
End Sub

Private Sub cmdCheck4_Click()
Dim AA As Integer
    For AA = 30 To 54
        ACS(AA).Value = 1
    Next AA
    ACS(151).Value = 1
    ACS(152).Value = 1
    ACS(155).Value = 1
    ACS(156).Value = 1
End Sub

Private Sub cmdCheck5_Click()
Dim AA As Integer
    For AA = 69 To 112
        ACS(AA).Value = 1
    Next AA
    ACS(147).Value = 1
    ACS(158).Value = 1
    ACS(159).Value = 1
End Sub

Private Sub cmdCheck7_Click()
Dim AA As Integer
    For AA = 64 To 68
        ACS(AA).Value = 1
    Next AA
End Sub

Private Sub cmdCheck8_Click()
Dim AA As Integer
    For AA = 55 To 63
        ACS(AA).Value = 1
    Next AA
    ACS(145).Value = 1
    ACS(146).Value = 1
    ACS(157).Value = 1
    For AA = 169 To 177
        ACS(AA).Value = 1
    Next AA
End Sub

Private Sub CmdCheck9_Click()
    ACS(19).Value = 1
    ACS(20).Value = 1
    ACS(21).Value = 1
    ACS(148).Value = 1
End Sub

Private Sub cmdCheck10_Click()
    ACS(115).Value = 1
    ACS(116).Value = 1
End Sub

Private Sub cmdCheck11_Click()
    ACS(117).Value = 1
    ACS(118).Value = 1
    ACS(119).Value = 1
    ACS(161).Value = 1
    ACS(162).Value = 1
End Sub

Private Sub cmdCheck13_Click()
Dim AA As Integer
    ACS(120).Value = 1
    For AA = 122 To 132
        ACS(AA).Value = 1
    Next AA
    ACS(163).Value = 1
    ACS(164).Value = 1
End Sub

Private Sub cmdCheck14_Click()
    ACS(133).Value = 1
    ACS(134).Value = 1
    ACS(135).Value = 1
    ACS(136).Value = 1
    ACS(121).Value = 1
    ACS(153).Value = 1
    ACS(165).Value = 1
    ACS(167).Value = 1
End Sub

Private Sub cmdCheck15_Click()
Dim AA As Integer
    For AA = 137 To 143
        ACS(AA).Value = 1
    Next AA
End Sub

Private Sub cmdCheckAccept_Click()
'On Error Resume Next
Dim CheckLoop As Integer
Dim AA As Integer
    For CheckLoop = 1 To ACS.Count '- 1
        ACS(CheckLoop).Value = 1
    Next CheckLoop

    For AA = 69 To 107
        ACS(AA).Value = 1
    Next AA
    For AA = 109 To 112
        ACS(AA).Value = 1
    Next AA
    For AA = 137 To 143
        ACS(AA).Value = 1
    Next AA
    ACS(147).Value = 0
End Sub

Private Sub CmdUncheck_Click()
    Dim CheckLoop As Integer
    For CheckLoop = 1 To ACS.Count '- 1
        ACS(CheckLoop).Value = 0
    Next CheckLoop

End Sub

Private Sub cmdCheckALL_Click()
'On Error Resume Next
Dim CheckLoop As Integer
        
    For CheckLoop = 1 To ACS.Count '- 1
        ACS(CheckLoop).Value = 1
    Next CheckLoop
       
End Sub

Private Sub cmdSave_Click()
    'On Error Resume Next
    Dim RecordSaved
    Dim USR As New ADODB.Recordset
    Dim StrSql As String
    Dim ListCheckBox As Integer
    Dim LoopCount As Integer
    Dim AccessString As String
    
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If txtUSRCode = "" Then
        MsgBox "Please Enter User Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtUSRCode.SetFocus
    ElseIf txtUSRName = "" Then
        MsgBox "Please Enter User Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtUSRName.SetFocus
    ElseIf txtUSRPassword = "" Then
        MsgBox "Please Enter Password", vbOKOnly + vbCritical, DialogBoxTitle
        txtUSRPassword.SetFocus
    ElseIf txtGRPCode = "" Then
        MsgBox "Please Enter Department", vbOKOnly + vbCritical, DialogBoxTitle
        txtGRPCode.SetFocus
    ElseIf TxtSRVCode = "" Then
        MsgBox "Please Enter Service Provider", vbOKOnly + vbCritical, DialogBoxTitle
        TxtSRVCode.SetFocus
    ElseIf txtUSRStatus = "" Then
        MsgBox "Please Enter Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtAuthStatus = "" Then
        MsgBox "Please Enter Authorised Status", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(Len(txtUSRPassword)) < 7 Then
        MsgBox "Password minimum length is 8 characters", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        StrSql = "Select USRCode From USR "
        StrSql = StrSql & "where (USRCode = '" & Trim(txtUSRCode) & "')"
        USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
        If USR.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                USR.Close
                'txtUSRCode.SetFocus
                'txtUSRCode.SelStart = 0
                'txtUSRCode.SelLength = Len(txtUSRCode)
                Set USR = Nothing
                Exit Sub
        Else
        USR.Close
        RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            If EditFlag = False Then
                USR.Open "USR", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                USR.AddNew
            Else
                StrSql = "Select * From USR Where USRCode = '" & txtUSRCode & "'"
                USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            End If
                
                ListCheckBox = ACS.Count
                For LoopCount = 1 To ListCheckBox '- 1
                    AccessString = AccessString + CStr(ACS(LoopCount))
                Next LoopCount
                
                Call Crypt(txtUSRPassword)
                
                With USR
                    !USRCode = ChkStr(txtUSRCode)
                    !USRLGCode = ChkStr(txtUSRLGCode)
                    !USRName = ChkStr(txtUSRName)
                    !USRPassword = tmpEncryptPassword
                    !USRDateCreated = txtUSRDateCreated
                    !USRStatus = txtUSRStatus
                    !AuthorisedStatus = txtAuthStatus
                    !GrpCode = Trim(Mid(txtGRPCode, 1, 3))
                    !SRVCode = Trim(Mid(TxtSRVCode, 1, 2))
                    !USRAccess = AccessString
                End With
                USR.Update
            Call ClearForm(Me)
            If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            End If
            'txtUSRCode.SetFocus
            USR.Close
            'Call USRListing
            cmdAdd.Enabled = True
        End If
        EditFlag = False
    End If
   End If
   
 '  medixmasterconnMaint.Close
 '  Set medixmasterconnMaint = Nothing
   
End Sub

Private Sub cmdSearch_Click()
Dim USR As New ADODB.Recordset
Dim StrSql As String
Dim USRMaster As ListItem

 '   medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If txtSearch = "" Then
        lstUsersList.ListItems.Clear
        StrSql = "Select * From USR "
        USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
            Do Until USR.EOF
                Set USRMaster = lstUsersList.ListItems.Add(, , USR!USRCode)
                    USRMaster.SubItems(1) = USR!USRName
                    USRMaster.SubItems(2) = USR!GrpCode
                    USRMaster.SubItems(3) = USR!SRVCode
                    USRMaster.SubItems(4) = USR!USRStatus
                    USR.MoveNext
            Loop
        USR.Close
        
    Else
        lstUsersList.ListItems.Clear
        StrSql = "Select * From USR where USRCode LIKE '%" & txtSearch & "%'"
        USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
            Do Until USR.EOF
                Set USRMaster = lstUsersList.ListItems.Add(, , USR!USRCode)
                    USRMaster.SubItems(1) = USR!USRName
                    USRMaster.SubItems(2) = USR!GrpCode
                    USRMaster.SubItems(3) = USR!SRVCode
                    USRMaster.SubItems(4) = USR!USRStatus
                    USR.MoveNext
            Loop
        USR.Close
    End If
    
'    medixmasterconnMaint.Close
 '   Set medixmasterconnMaint = Nothing
 End Sub

Private Sub Form_Load()
    'Frame3.Width = Screen.Width
    txtUSRDateCreated = Date
    cmdAdd.Enabled = True
    'Call DisableCmdButton(Me)
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Call ComboListing
    Call USRListing
    
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
    frmUserPermissionMaintenance.SSTab1.Tab = 0
        
End Sub

Private Sub CmdAdd_Click()
    EditFlag = False
    Call ClearForm(Me)
    Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    txtUSRDateCreated = Date
    txtUSRStatus = "A"
    txtAuthStatus = "N"
    txtUSRCode.SetFocus
    cmdSave.Enabled = True
    
End Sub

Private Sub cmdDelete_Click()
Dim rst4 As New ADODB.Recordset
Dim UserArchieve As New ADODB.Recordset
Dim StrSql As String
Dim RecordSaved
    
    EditFlag = False
    cmdDelete.Enabled = False
    cmdSearch.Enabled = False
    cmdSave.Enabled = False
   ' medixmasterconnA.Open "File Name=" & BOSPath & "Conn\mediconnectA.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If txtUSRCode = "" Then
        MsgBox "Please Enter User Code ...", vbOKOnly + vbInformation, DialogBoxTitle
    Else
      RecordSaved = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            
            UserArchieve.Open "USRArchive", medixmasterconnA, adOpenKeyset, adLockOptimistic
                
                With UserArchieve
                    .AddNew
                    !USRCode = Trim(txtUSRCode)
                    !USRName = Trim(txtUSRName)
                    !USRLastUpdateDate = Date
                    !USRLastUpdateUser = Logon
                    .Update
                End With
            
            UserArchieve.Close
            Set UserArchieve = Nothing
            
            StrSql = "Delete From USR Where USRCode = '" & txtUSRCode & "'"
            rst4.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
            cmdSave.Enabled = True
            Call ClearForm(Me)
            'Call USRListing
            Call DisableCmdButton(Me)
            Call EntriesEnable(True)
        
            txtUSRCode.SetFocus
            cmdAdd.Enabled = True
        End If
    End If
    
   ' medixmasterconnA.Close
   ' Set medixmasterconnA = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    cmdDelete.Enabled = True
    cmdSearch.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub CmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    cmdSave.Enabled = True
    txtUSRCode.SetFocus
    txtUSRCode.Enabled = False
    
End Sub


Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
        'bExit = True
        Unload Me
'    End If
End Sub


Private Sub USRListing()
On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim UserMaster As ListItem
    
    lstUsersList.ListItems.Clear
    rst2.Open "USR", medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        Do Until rst2.EOF
            Set UserMaster = lstUsersList.ListItems.Add(, , rst2!USRCode)
                UserMaster.SubItems(1) = Trim(rst2!USRName)
                UserMaster.SubItems(2) = Trim(rst2!USRLGCode)
                UserMaster.SubItems(3) = rst2!GrpCode
                UserMaster.SubItems(4) = rst2!SRVCode
                UserMaster.SubItems(5) = rst2!USRStatus
                UserMaster.SubItems(6) = rst2!AuthorisedStatus
                rst2.MoveNext
        Loop
    rst2.Close
End Sub

Private Sub lstUsersList_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   lstUsersList.SortKey = ColumnHeader.Index - 1
   lstUsersList.Sorted = True
End Sub

Private Sub EntriesEnable(status As Boolean)
        On Error Resume Next
        Dim listloop As Integer
        txtUSRCode.Enabled = status
        txtUSRLGCode.Enabled = status
        txtUSRName.Enabled = status
        txtUSRPassword.Enabled = status
        txtGRPCode.Enabled = status
        TxtSRVCode.Enabled = status
        txtAuthStatus.Enabled = status
        txtUSRStatus.Enabled = status
        cmdCheck1.Enabled = status
        cmdCheck2.Enabled = status
        cmdCheck3.Enabled = status
        cmdCheck4.Enabled = status
        cmdCheck5.Enabled = status
        'cmdCheck6.Enabled = status
        cmdCheck7.Enabled = status
        cmdCheck8.Enabled = status
        CmdUncheck.Enabled = status
        cmdCheckAll.Enabled = status
        cmdCheckAccept.Enabled = status
        
    
    For listloop = 0 To ACS.Count
        ACS(listloop).Enabled = status
    Next listloop
    
End Sub

Private Sub ComboListing()
    Dim GRP As New ADODB.Recordset
    Dim SRV As New ADODB.Recordset
  
    GRP.Open "GRP", medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
    SRV.Open "SRV", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
    Do Until GRP.EOF
        txtGRPCode.AddItem GRP!GrpCode & " - " & Trim(GRP!GRPDescription)
        GRP.MoveNext
    Loop
    
    Do Until SRV.EOF
        TxtSRVCode.AddItem SRV!SRVCode & " - " & Trim(SRV!SRVname)
        SRV.MoveNext
    Loop
    
    txtUSRStatus.AddItem "A"
    txtUSRStatus.AddItem "S"
    
    txtAuthStatus.AddItem "Y"
    txtAuthStatus.AddItem "N"
    
    GRP.Close
    SRV.Close
End Sub

Private Sub lstUsersList_Click()
'On Error Resume Next
Dim StrSql As String
Dim USR As New ADODB.Recordset
Dim listloop As Integer
    
    
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    StrSql = "SELECT * "
    StrSql = StrSql & "FROM view_USR_GRP where USRCode='" & Trim(lstUsersList.SelectedItem.Text) & "'"
    USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    Do While Not USR.EOF
        With USR
            txtUSRCode = !USRCode
            If Len(!USRLGCode) Then
                txtUSRLGCode = !USRLGCode
            Else
                txtUSRLGCode = ""
            End If
            txtUSRName = !USRName
            txtUSRPassword = !USRPassword
            txtUSRDateCreated = !USRDateCreated
            txtUSRStatus = !USRStatus
            txtAuthStatus = !AuthorisedStatus
            txtGRPCode = !GrpCode & " - " & Trim(!GRPDescription)
            TxtSRVCode = !SRVCode & " - " & Trim(!SRVname)
        End With
        
        For listloop = 1 To Len(USR!USRAccess)
            If Mid(USR!USRAccess, listloop, 1) = 1 Then
                'ACS(listloop - 1).value = 1
                ACS(listloop).Value = 1
            Else
                'ACS(listloop - 1).value = 0
                ACS(listloop).Value = 0
            End If
        Next listloop
    USR.MoveNext
    Loop
    Call EnableCmdButton(Me)
    Call EntriesEnable(False)
    
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub lstUsersList_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
Dim StrSql As String
Dim USR As New ADODB.Recordset
Dim listloop As Integer
    
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    StrSql = "SELECT *"
    StrSql = StrSql & "FROM view_USR_GRP where USRCode='" & Trim(lstUsersList.SelectedItem.Text) & "'"
    USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        With USR
            txtUSRCode = !USRCode
            If Len(!USRLGCode) Then
                txtUSRLGCode = !USRLGCode
            Else
                txtUSRLGCode = ""
            End If
            txtUSRName = !USRName
            txtUSRPassword = !USRPassword
            txtUSRDateCreated = !USRDateCreated
            txtUSRStatus = !USRStatus
            txtAuthStatus = !AuthorisedStatus
            txtGRPCode = !GrpCode & " - " & Trim(!GRPDescription)
            TxtSRVCode = !SRVCode & " - " & Trim(!SRVname)
        End With
        
    For listloop = 1 To Len(USR!USRAccess)
        If Mid(USR!USRAccess, listloop, 1) = 1 Then
            'ACS(listloop - 1).value = 1
            ACS(listloop).Value = 1
        Else
            'ACS(listloop - 1).value = 0
            ACS(listloop).Value = 0
        End If
    Next listloop
    
    Call EnableCmdButton(Me)
    Call EntriesEnable(False)
    
   ' medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub lstUsersList_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
Dim StrSql As String
Dim USR As New ADODB.Recordset
Dim listloop As Integer
    
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    StrSql = "SELECT *"
    StrSql = StrSql & "FROM view_USR_GRP where USRCode='" & Trim(lstUsersList.SelectedItem.Text) & "'"
    USR.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        With USR
            txtUSRCode = !USRCode
            If Len(!USRLGCode) Then
                txtUSRLGCode = !USRLGCode
            Else
                txtUSRLGCode = ""
            End If
            txtUSRName = !USRName
            txtUSRPassword = !USRPassword
            txtUSRDateCreated = !USRDateCreated
            txtUSRStatus = !USRStatus
            txtAuthStatus = !AuthorisedStatus
            txtGRPCode = !GrpCode & " - " & Trim(!GRPDescription)
            TxtSRVCode = !SRVCode & " - " & Trim(!SRVname)
        End With
        
    For listloop = 1 To Len(USR!USRAccess)
        If Mid(USR!USRAccess, listloop, 1) = 1 Then
            'ACS(listloop - 1).value = 1
            ACS(listloop).Value = 1
        Else
            'ACS(listloop - 1).value = 0
            ACS(listloop).Value = 0
        End If
    Next listloop
    
    Call EnableCmdButton(Me)
    Call EntriesEnable(False)
    
 '   medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub txtGRPCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtGRPCode.Text, txtGRPCode.SelStart) + Chr(KeyAscii) + Mid(txtGRPCode.Text, txtGRPCode.SelStart + txtGRPCode.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To txtGRPCode.ListCount - 1
        
        If NewText = LCase(Left(txtGRPCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = txtGRPCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
            If ValidCount = 1 Then
                txtGRPCode.Text = ValidValue
                txtGRPCode.SelStart = 0
                txtGRPCode.SelLength = Len(ValidValue)
            End If
        
        End If
End Sub
'End Sub

Private Sub txtSRVCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(TxtSRVCode.Text, TxtSRVCode.SelStart) + Chr(KeyAscii) + Mid(TxtSRVCode.Text, TxtSRVCode.SelStart + TxtSRVCode.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To TxtSRVCode.ListCount - 1
        
        If NewText = LCase(Left(TxtSRVCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = TxtSRVCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
            If ValidCount = 1 Then
                TxtSRVCode.Text = ValidValue
                TxtSRVCode.SelStart = 0
                TxtSRVCode.SelLength = Len(ValidValue)
            End If
        
        End If
End Sub

Private Sub txtUSRStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtUSRStatus.Text, txtUSRStatus.SelStart) + Chr(KeyAscii) + Mid(txtUSRStatus.Text, txtUSRStatus.SelStart + txtUSRStatus.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To txtUSRStatus.ListCount - 1
        
        If NewText = LCase(Left(txtUSRStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = txtUSRStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
            If ValidCount = 1 Then
                txtUSRStatus.Text = ValidValue
                txtUSRStatus.SelStart = 0
                txtUSRStatus.SelLength = Len(ValidValue)
            End If
        
        End If

End Sub

Private Sub txtAuthStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtAuthStatus.Text, txtAuthStatus.SelStart) + Chr(KeyAscii) + Mid(txtAuthStatus.Text, txtAuthStatus.SelStart + txtAuthStatus.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To txtUSRStatus.ListCount - 1
        
        If NewText = LCase(Left(txtAuthStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = txtAuthStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
            If ValidCount = 1 Then
                txtAuthStatus.Text = ValidValue
                txtAuthStatus.SelStart = 0
                txtAuthStatus.SelLength = Len(ValidValue)
            End If
        
        End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
End Sub

Public Function Crypt(strOriginal) As String
Dim valString$, x%

    For x = 1 To Len(strOriginal)
        'If encrypting = True Then
            valString = Asc(Mid(strOriginal, x, 1)) + (x ^ 2)
                Do While valString > 255
                    valString = valString - 255
                Loop
        'Else
        '    valString = Asc(Mid(strOriginal, x, 1)) - (x ^ 2)
        '        Do While valString <= 0
        '            valString = valString + 255
        '        Loop
        'End If
        Crypt = Crypt & Chr(valString)
    Next
    tmpEncryptPassword = Crypt
End Function

