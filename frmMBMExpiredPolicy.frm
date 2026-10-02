VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMExpiredPolicy 
   Caption         =   "Member Expired Policy"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   9225
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   60
      TabIndex        =   34
      Top             =   2040
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Special Note"
      TabPicture(0)   =   "frmMBMExpiredPolicy.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstSpecialNoteList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Repudiation"
      TabPicture(1)   =   "frmMBMExpiredPolicy.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label65(58)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblPatient"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label65(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label65(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame18"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lstRepList"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CmdDiag"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton CmdDiag 
         Caption         =   ">"
         Height          =   255
         Left            =   -73320
         TabIndex        =   49
         Top             =   480
         Width           =   225
      End
      Begin MSComctlLib.ListView lstSpecialNoteList 
         Height          =   1935
         Left            =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Patient Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Hospital"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Diagnosis"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DOA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOD"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Bill Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Index"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "MemberNo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
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
         Height          =   1155
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   8955
         Begin VB.ComboBox cboPatient 
            Height          =   315
            Left            =   885
            TabIndex        =   10
            Top             =   240
            Width           =   3675
         End
         Begin VB.ComboBox cboHOSList 
            Height          =   315
            Left            =   885
            TabIndex        =   11
            Top             =   510
            Width           =   3675
         End
         Begin MSMask.MaskEdBox txtDOA 
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
            Left            =   5760
            TabIndex        =   13
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDOD 
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
            Left            =   5760
            TabIndex        =   14
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtBillAmt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5760
            MaxLength       =   8
            TabIndex        =   15
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtDiag 
            Height          =   285
            Left            =   885
            MaxLength       =   100
            TabIndex        =   12
            Top             =   785
            Width           =   3675
         End
         Begin VB.Label Label73 
            Caption         =   "Diagnosis"
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
            TabIndex        =   41
            Top             =   800
            Width           =   660
         End
         Begin VB.Label Label31 
            Caption         =   "Patient"
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
            TabIndex        =   40
            Top             =   255
            Width           =   1020
         End
         Begin VB.Label Label64 
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
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   535
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "DOA"
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
            Left            =   4800
            TabIndex        =   38
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "DOD"
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
            Left            =   4800
            TabIndex        =   37
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Bill Amount"
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
            Left            =   4800
            TabIndex        =   36
            Top             =   720
            Width           =   1275
         End
      End
      Begin MSComctlLib.ListView lstRepList 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   48
         Top             =   2160
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MBM Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Patient Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rep 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rep 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Rep 3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Index"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame18 
         Height          =   1185
         Left            =   -74880
         TabIndex        =   44
         Top             =   720
         Width           =   8895
         Begin VB.ComboBox cboREPCode1 
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
            Left            =   90
            Style           =   1  'Simple Combo
            TabIndex        =   45
            Top             =   240
            Width           =   8715
         End
         Begin VB.ComboBox cboREPCode2 
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
            Left            =   90
            Style           =   1  'Simple Combo
            TabIndex        =   46
            Top             =   525
            Width           =   8715
         End
         Begin VB.ComboBox cboREPCode3 
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
            Left            =   90
            Style           =   1  'Simple Combo
            TabIndex        =   47
            Top             =   795
            Width           =   8715
         End
      End
      Begin VB.Label Label65 
         Caption         =   "Please double click to reprint repudiation letter"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   57
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label Label65 
         Caption         =   "Patient Name :"
         Height          =   255
         Index           =   0
         Left            =   -69960
         TabIndex        =   53
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblPatient 
         Height          =   255
         Left            =   -68880
         TabIndex        =   52
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label65 
         Caption         =   "Repudiation Resons"
         Height          =   255
         Index           =   58
         Left            =   -74880
         TabIndex        =   50
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   60
      TabIndex        =   22
      Top             =   240
      Width           =   9135
      Begin VB.CommandButton cmdSearchMBM 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   23
         Top             =   225
         Width           =   675
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1000
         MaxLength       =   25
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Height          =   285
         Left            =   3480
         TabIndex        =   1
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label76 
         Caption         =   "Mem Num."
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
         TabIndex        =   24
         Top             =   250
         Width           =   1275
      End
   End
   Begin VB.Frame Frame10 
      Height          =   1185
      Left            =   60
      TabIndex        =   18
      Top             =   840
      Width           =   9135
      Begin VB.TextBox txtPOLStatus 
         Height          =   330
         Left            =   5880
         MaxLength       =   50
         TabIndex        =   5
         Top             =   187
         Width           =   1035
      End
      Begin VB.TextBox txtMBMName 
         Height          =   330
         Left            =   1000
         MaxLength       =   50
         TabIndex        =   2
         Top             =   187
         Width           =   3675
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   330
         Left            =   1000
         MaxLength       =   30
         TabIndex        =   3
         Top             =   480
         Width           =   3675
      End
      Begin VB.TextBox txtMBMPolNo 
         Height          =   330
         Left            =   1000
         MaxLength       =   25
         TabIndex        =   4
         Top             =   760
         Width           =   3675
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
         Left            =   7800
         TabIndex        =   8
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
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
         Left            =   7800
         TabIndex        =   9
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtDataStatus 
         Height          =   330
         Left            =   5880
         MaxLength       =   50
         TabIndex        =   6
         Top             =   480
         Width           =   1035
      End
      Begin VB.TextBox txtEXPStatus 
         Height          =   330
         Left            =   5880
         MaxLength       =   50
         TabIndex        =   7
         Top             =   760
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Exp Date"
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
         Left            =   7080
         TabIndex        =   29
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Eff Date"
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
         Left            =   7080
         TabIndex        =   28
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label25 
         Caption         =   "Pol Status"
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
         Left            =   4800
         TabIndex        =   27
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label90 
         Caption         =   "Data Status"
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
         Left            =   4800
         TabIndex        =   26
         Top             =   495
         Width           =   915
      End
      Begin VB.Label Label103 
         Caption         =   "Exp. Status"
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
         Left            =   4800
         TabIndex        =   25
         Top             =   795
         Width           =   915
      End
      Begin VB.Label Label64 
         Caption         =   "IC Num"
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
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   500
         Width           =   615
      End
      Begin VB.Label Label31 
         Caption         =   "Mem Name"
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
         Left            =   120
         TabIndex        =   20
         Top             =   200
         Width           =   1020
      End
      Begin VB.Label Label73 
         Caption         =   "Policy Num"
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
         Left            =   120
         TabIndex        =   19
         Top             =   800
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   30
      Top             =   5640
      Width           =   9135
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   5475
         TabIndex        =   16
         Top             =   200
         Width           =   900
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   6360
         TabIndex        =   43
         Top             =   200
         Width           =   900
      End
      Begin VB.CommandButton cmdRepudiate 
         Caption         =   "Rep&udiate"
         Height          =   315
         Left            =   7260
         TabIndex        =   33
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   315
         Left            =   6360
         TabIndex        =   32
         Top             =   200
         Width           =   900
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8160
         TabIndex        =   31
         Top             =   200
         Width           =   900
      End
      Begin VB.Label lblRepPatient 
         Caption         =   "Label2"
         Height          =   255
         Left            =   1680
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblRepIndex 
         Caption         =   "Label2"
         Height          =   255
         Left            =   1680
         TabIndex        =   55
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblMBMNumber 
         Caption         =   "MBMNumber"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMBMIndex 
         Caption         =   "Index"
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Member Expired Policy"
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
      TabIndex        =   17
      Top             =   0
      Width           =   11445
   End
End
Attribute VB_Name = "frmMBMExpiredPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditStatus As Boolean
Dim TmpPLNCode As String
Dim objword As Word.Application
Dim objdoc As Word.Document

Private Sub cmdAdd_Click()
    cmdEdit.Visible = False
    cmdUpdate.Visible = True
    cmdUpdate.Caption = "&Save"
    Frame1.Enabled = True
    EditStatus = False
End Sub

Private Sub CmdDiag_Click()
    FrmSearchRepudiation.Source = 3
    FrmSearchRepudiation.Show vbModal
End Sub

Private Sub cmdEdit_Click()
    cmdEdit.Visible = False
    cmdUpdate.Visible = True
    cmdUpdate.Caption = "&Update"
    Frame1.Enabled = True
    EditStatus = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRepudiate_Click()
    If lblPatient = "" Then
        MsgBox "Please Select Patient", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboREPCode1 = "" Then
        MsgBox "Please Enter Repudiation Reason", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        SaveRepudiation
    End If
End Sub

Private Sub cmdSearchMBM_Click()
    Call ClearEntries
    frmSearchMBM.Source = 7
    frmSearchMBM.Show
End Sub

Public Sub cmdShow_Click()
Dim tempMBMNumber As String
    Screen.MousePointer = vbHourglass
    If txtMBMNumber = "" Then
        MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
    Else
        cmdEdit.Visible = True
        cmdUpdate.Visible = False
            Call FindMBM
    End If
    Screen.MousePointer = Default

End Sub

Private Sub cmdUpdate_Click()
    cmdEdit.Visible = True
    cmdUpdate.Visible = False
    Frame1.Enabled = False
    Call SaveSpecialNote
End Sub

Private Sub Form_Load()
    Call ComboList
End Sub

Private Sub lstRepList_Click()
    Call GetRepudiation
End Sub

Private Sub lstRepList_DblClick()
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call GetRepudiation
        Call funcword
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub lstRepList_KeyDown(KeyCode As Integer, Shift As Integer)
    Call GetRepudiation
End Sub

Private Sub lstRepList_KeyUp(KeyCode As Integer, Shift As Integer)
    Call GetRepudiation
End Sub

Private Sub lstSpecialNoteList_Click()
    Call GetSpecialNote
End Sub

Private Sub lstSpecialNoteList_KeyDown(KeyCode As Integer, Shift As Integer)
    Call GetSpecialNote
End Sub

Private Sub lstSpecialNoteList_KeyUp(KeyCode As Integer, Shift As Integer)
    Call GetSpecialNote
End Sub

Private Function FindMBM()
Dim XrefSql1 As String
Dim XrefSql2 As String
Dim XrefDB As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpREcCnt As Integer
Dim tempMBMNumber  As String
    XrefSql1 = ""
    XrefSql2 = ""
    cboPatient.Clear
    tempMBMNumber = txtMBMNumber
    Call ClearEntries
    txtMBMNumber = tempMBMNumber

    If Len(txtMBMNumber) = 8 Or Len(txtMBMNumber) = 9 Then
        frmSearchMBM05.Source = 7
        frmSearchMBM05.Show
        'tmpREcCnt = frmSearchMBM05.SearchRecCnt
        If frmSearchMBM05.SearchRecCnt < 1 Then
            Unload frmSearchMBM05
        ElseIf frmSearchMBM05.SearchRecCnt = 1 Then
            txtMBMNumber = frmSearchMBM05.tmpMBMNo
            Unload frmSearchMBM05
            tmpREcCnt = 2

        End If
    Else
        tmpREcCnt = "2"
    End If
    
    If tmpREcCnt > 1 Then
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
        XrefSql1 = "Select MBMPolicyNo,MBMIcBcPp,MBMName,MBMPayorEffDate,MBMPayorEXPDate,MBMDataStatus,MBMStatus,INSCode,PLNCode "
        XrefSql2 = " And MBMNumber = '" & Trim(txtMBMNumber) & "'"
    
        Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "SearchMBMCrossRef"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("SelField", adVarChar, adParamInput, 2000, XrefSql1)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, XrefSql2)
        cmd.Parameters.Append Prm2
        XrefDB.CursorLocation = adUseClient
        XrefDB.CursorType = adOpenForwardOnly
        XrefDB.CacheSize = 100
        Set XrefDB = cmd.Execute
        
        If XrefDB.EOF = False Then
            With XrefDB
                txtMBMPolNo = !MBMPolicyNo
                txtMBMIcBcPP = !MBMIcBcPp
                txtMBMName = !MBMName
                txtMBMPayorEffDate = !MBMPayorEffDate
                txtMBMPayorExpDate = !MBMPayorEXPDate
                txtDataStatus = !MBMDataStatus
                txtPOLStatus = !MBMStatus
                txtEXPStatus = IIf(CDate(Date) > CDate(!MBMPayorEXPDate), "Y", "N")
                cboPatient.AddItem !MBMName
                TmpPLNCode = !PLNCode
                If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                    Call showMBMCov
                End If
            End With
            Call SpecialNoteList
            Call RepudiationList

            If Trim(txtEXPStatus) = "N" Then
                MsgBox "This Policy is not expired!", vbCritical
                Call EnableEntries(False)
            Else
                Call EnableEntries(True)
            End If
        Else
            MsgBox "No record found!", vbCritical
        End If
        XrefDB.Close
        Set XrefDB = Nothing
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
    End If
End Function

Private Sub showMBMCov()
    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
                
    'sqlstr = "select MBMCNumber, MBMCCoverID, MBMCName," & _
        " MBMCICBCPPNo , MBMCDOB, MBMCSex, MBMCAdultChild, MBMCAgeband, " & _
        " MBMCRELCode,   " & _
        " MBMCStatus , MBMCAnnualLimit,  " & _
        " MBMCStatus " & _
        " from  MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    'MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "14"
    strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Registration"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        MBMCov.CursorLocation = adUseClient
        MBMCov.CursorType = adOpenForwardOnly
        MBMCov.CacheSize = 100
        Set MBMCov = cmd.Execute
    
    If MBMCov.recordCount Then
        Do While Not MBMCov.EOF
            cboPatient.AddItem MBMCov!MBMCName
            MBMCov.MoveNext
        Loop
    End If
    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub ClearSpecialNote()
    cboPatient = ""
    cboHOSList = ""
    txtDiag = ""
    txtDOA.mask = ""
    txtDOA.Text = ""
    txtDOA.mask = "##/##/####"
    txtDOD.mask = ""
    txtDOD.Text = ""
    txtDOD.mask = "##/##/####"
    txtBillAmt = ""
End Sub

Private Sub ClearEntries()
   On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ComboBox Then
          '  If ctl.Name = "cboPayMode" Then
          '      cboPayMode.ListIndex = 0
            'If ctl.Name = "CboPayType" Then
            '    CboPayType.ListIndex = 1
            'Else
                ctl.Text = ""
            'End If
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        End If
    Next ctl
    lblMBMNumber = ""
End Sub

Private Sub GetSpecialNote()
    Frame1.Enabled = False
    If lstSpecialNoteList.ListItems.Count > 0 Then
        With lstSpecialNoteList.SelectedItem
            cboPatient = .Text
            cboHOSList = .SubItems(1)
            txtDiag = .SubItems(2)
            txtDOA = .SubItems(3)
            txtDOD = .SubItems(4)
            txtBillAmt = .SubItems(5)
            lblMBMIndex = .SubItems(6)
            lblPatient = cboPatient
        End With
    End If
End Sub

Private Sub SaveSpecialNote()
Dim rst As New ADODB.Recordset
Dim strsqlrst As String
Dim strsql As String
Dim strsqlPLN As String
Dim tmpSpecialNotes As String
Dim tmpName As String
Dim tmpDOA As String
Dim tmpDOD As String
Dim tmpdate As String

    If Trim(cboPatient) = "" Then
        MsgBox "Please select Patient Name", vbOKOnly + vbInformation, DialogBoxTitle
        Exit Sub
    ElseIf Trim(cboHOSList) = "" Then
        MsgBox "Please select Hospital Name", vbOKOnly + vbInformation, DialogBoxTitle
        Exit Sub
    ElseIf Trim(txtDiag) = "" Then
        MsgBox "Please enter diagnosis", vbOKOnly + vbInformation, DialogBoxTitle
        Exit Sub
    ElseIf txtDOA = "__/__/____" Or IsDate(txtDOA) = False Then
        MsgBox "Please select Admission Date", vbOKOnly + vbInformation, DialogBoxTitle
        Exit Sub
    ElseIf txtDOD = "__/__/____" Or IsDate(txtDOD) = False Then
        MsgBox "Please enter Discharge Date", vbOKOnly + vbInformation, DialogBoxTitle
        Exit Sub
    ElseIf Trim(txtBillAmt) = "" Then
        MsgBox "Please enter Bill Amount", vbOKOnly + vbInformation, DialogBoxTitle
        Exit Sub
    Else
    tmpSpecialNotes = ""
    strsqlPLN = ""
    strsql = ""
    tmpName = Replace(cboPatient, "'", "''")
    
    tmpDOA = Format(txtDOA, "YYYY/MM/DD")
    tmpDOD = Format(txtDOD, "YYYY/MM/DD")
    tmpdate = Format(Date, "YYYY/MM/DD")
    
   'tmpSpecialNotes = ""
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If EditStatus = True Then
    
        strsql = " update MBMExpSpecialNote set " & _
                 " MBMNumber = '" & Trim(txtMBMNumber) & "', " & _
                 " MBMPatientName ='" & tmpName & "', " & _
                 " MBMHOSName = '" & cboHOSList & "', " & _
                 " MBMDiag = '" & txtDiag & "', " & _
                 " MBMDOA = '" & tmpDOA & "', " & _
                 " MBMDOD = '" & tmpDOD & "', " & _
                 " MBMBillAmt = " & txtBillAmt & ", " & _
                 " RegDate = '" & tmpdate & "'," & _
                 " RegUser = '" & Logon & "', " & _
                 " LastUpdateDate = '" & tmpdate & "', " & _
                 " LastUpdateUser = '" & Logon & "'" & _
                 " Where MBMIndex = '" & Trim(lblMBMIndex) & "'"
    Else
    strsql = ""
        strsql = "insert into MBMExpSpecialNote (MBMNumber, MBMPatientName, MBMHOSName,MBMDiag,MBMDOA,MBMDOD, MBMBillAmt,RegDate,RegUser,LastUpdateDate, LastUpdateUser) " & _
                "Values('" & txtMBMNumber & "', '" & tmpName & "', '" & cboHOSList & "', '" & txtDiag & "', '" & tmpDOA & "', '" & tmpDOD & "', " & txtBillAmt & ", '" & tmpdate & "', '" & Logon & "', '" & tmpdate & "', '" & Logon & "')"
    End If
        medixmasterconn.Execute strsql
   
        tmpSpecialNotes = " *Start* " & "Expired Policy" & " - " & "(" & "Patient : " & tmpName & " ; " & "Hopital : " & cboHOSList & " ; " & "Diagnois : " & txtDiag & _
                          " ; DOA : " & txtDOA & " ; " & "DOD : " & txtDOD & " ; " & "Bill Amt : " & txtBillAmt & ")" & " *End* "
        
        strsql = "Select MBMNumber,MBMRemarks from MBMOthersTwo where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        rst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            'If EditStatus = False Then
                tmpSpecialNotes = rst!MBMRemarks & tmpSpecialNotes
                rst!MBMRemarks = tmpSpecialNotes
                rst.Update
            'End If
        rst.Close
        Set rst = Nothing
        medixmasterconn.Close
        Set medixmasterconn = Nothing
    
        If EditStatus = True Then
            MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
        Else
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
        End If
        Call FindMBM
    End If
    
End Sub

Private Sub ComboList()
Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String
    cboHOSList.Clear
    tmpSelField = "SELECT HOSCode as tmpCode, HOSName as tmpName "
    tmpTable = "HOS"
    tmpSelCriteria = "0=0"
    Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, cboHOSList)
End Sub

Private Sub LoadComboListing(tmpSelField As String, tmpSelCriteria As String, tmpTableName As String, tmpComboBox As ComboBox)
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    tmpComboBox.Clear
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelField)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, tmpSelCriteria)
    cmd.Parameters.Append Prm3
    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    
    Do Until TmpRec.EOF
        With TmpRec
        If tmpTableName = "HOS" Then
            tmpComboBox.AddItem !tmpName
        End If
    TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Function SpecialNoteList()
Dim strsql As String
Dim rst As New ADODB.Recordset
Dim Record As ListItem
        
        lstSpecialNoteList.ListItems.Clear
        strsql = "Select MBMIndex,MBMNumber,MBMPatientName,MBMHOSName,MBMDiag,MBMDOA,MBMDOD,MBMBillAmt from MBMEXPSpecialNOte where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        rst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If rst.recordCount > 0 Then
            Do While Not rst.EOF
                With rst
                    Set Record = lstSpecialNoteList.ListItems.Add(, , !MBMPatientName)
                        Record.SubItems(1) = !MBMHOSName
                        Record.SubItems(2) = !MBMDiag
                        Record.SubItems(3) = !MBMDOA
                        Record.SubItems(4) = !MBMDOD
                        Record.SubItems(5) = !MBMBillAmt
                        Record.SubItems(6) = !MBMIndex
                        Record.SubItems(7) = !MBMNumber
                        lblMBMNumber = !MBMNumber
                End With
            rst.MoveNext
            Loop
        End If
        rst.Close
        Set rst = Nothing
End Function

Private Sub EnableEntries(status As Boolean)
    cmdAdd.Enabled = status
    cmdUpdate.Enabled = status
    cmdEdit.Enabled = status
    cmdRepudiate.Enabled = status
    cmdExit.Enabled = True
End Sub

Private Sub txtDOA_LostFocus()
    If IsDate(txtDOA) Then
        If txtDOA <> "__/__/____" Then
            If CDate(txtDOA) > CDate(Date) Then
                MsgBox "Admission Date after System Date ! ", vbCritical
                txtDOD.SetFocus
                Exit Sub
            End If
            
            If CDate(txtDOA) < CDate(txtMBMPayorExpDate) Then
                MsgBox "Admission Date before Expiry Date! ", vbCritical
                txtDOA.SetFocus
            End If
        End If

    Else
        If txtDOA <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOA.SetFocus
        End If
    End If

End Sub


Private Sub txtDOD_LostFocus()
    If IsDate(txtDOD) Then
        If txtDOD <> "__/__/____" Then
            If CDate(txtDOD) > CDate(Date) Then
                MsgBox "Dsicharge Date after System Date ! ", vbCritical
                txtDOD.SetFocus
                Exit Sub
            End If
            
            If CDate(txtDOD) < CDate(txtMBMPayorExpDate) Then
                MsgBox "Discharge Date before Expiry Date! ", vbCritical
                txtDOD.SetFocus
                Exit Sub
            End If
            
            If CDate(txtDOD) < CDate(txtDOA) Then
                MsgBox "Discharge Date before Admission Date! ", vbCritical
                txtDOD.SetFocus
                Exit Sub
            End If

        End If

    Else
        If txtDOA <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOA.SetFocus
        End If
    End If
End Sub

Private Sub SaveRepudiation()
Dim rst As New ADODB.Recordset
Dim strsql As String
Dim Rep1, Rep2, Rep3 As String
Dim tmpName As String
Dim tmpdate As String

    strsql = ""
    tmpName = Replace(lblPatient, "'", "''")
    Rep1 = IIf(Len(cboREPCode1), cboREPCode1, "")
    Rep2 = IIf(Len(cboREPCode2), cboREPCode2, "")
    Rep3 = IIf(Len(cboREPCode3), cboREPCode3, "")
    tmpdate = Format(Date, "YYYY/MM/DD")
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
strsql = ""
    strsql = "insert into MBMExpRepudiation (MBMNumber, MBMPatientName, RepReason1,RepReason2,RepReason3,RegDate,RegUser) " & _
            "Values('" & txtMBMNumber & "', '" & tmpName & "', '" & Rep1 & "', '" & Rep2 & "', '" & Rep2 & "', '" & tmpdate & "', '" & Logon & "')"
    medixmasterconn.Execute strsql
   
    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
    
    
    Call funcword
    Call RepudiationList
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    

End Sub

Private Sub funcword()
On Error GoTo openerror

        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                    Set objdoc = Nothing
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetterMBM.DOT\")
        ElseIf Not objdoc Is Nothing Then
              If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
                vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              ElseIf vbCancel Then
                Exit Sub
              End If
                Set objdoc = Nothing
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "REPLetterMBM.DOT")
        End If
        bookmark ("REPLetterMBM.DOT")
        Me.Show
        Exit Sub
openerror:
    Me.Show
    If Err.Number = "5356" Then
        MsgBox Err.Description & "Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Function bookmark(Optional strAppend As String) As Integer
Dim MBM As New ADODB.Recordset
Dim strsqlMBM As String
Dim tmpMBMName As String
Dim tmpMBMAddress1 As String
Dim tmpMBMAddress2 As String
Dim tmpMBMAddress3 As String
Dim tmpCTYCode As String
Dim tmpMBMPostCode As String
Dim tmpSTACode As String
Dim tmpPAYCode As String
Dim tmpPAYName As String
Dim cmd As New ADODB.Command
Dim cmd1 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim cmd5 As New ADODB.Command
Dim cmd6 As New ADODB.Command
Dim cmd7 As New ADODB.Command
Dim cmd8 As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim strsqlREP As String
Dim REP As New ADODB.Recordset
Dim REPCat As New ADODB.Recordset
Dim strsqlREPCat As String
Dim PAY As New ADODB.Recordset
Dim strsqlPAY As String

        strsqlMBM = "Select MBMName,MBMAddress1,MBMAddress2,MBMAddress3,CTYCode,MBMPostCode,STACode,PAYCode from MBM where MBMNumber = '" & Trim(lblMBMNumber) & "'"
        MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBM.recordCount Then
            With MBM
                tmpMBMName = IIf(Len(!MBMName), !MBMName, "")
                tmpMBMAddress1 = IIf(Len(!MBMAddress1), !MBMAddress1, "")
                tmpMBMAddress2 = IIf(Len(!MBMAddress2), !MBMAddress2, "")
                tmpMBMAddress3 = IIf(Len(!MBMAddress3), !MBMAddress3, "")
                tmpCTYCode = IIf(Len(!CTYCode), !CTYCode, "")
                tmpMBMPostCode = IIf(Len(!MBMPostCode), !MBMPostCode, "")
                tmpSTACode = IIf(Len(!STACode), !STACode, "")
                tmpPAYCode = IIf(Len(!PAYCode), !PAYCode, "")
            End With
        End If
        MBM.Close
        Set MBM = Nothing
        
        strsqlPAY = "Select PAYCode,PAYCompanyName from PAY where PAYCode = '" & Trim(tmpPAYCode) & "'"
        PAY.Open strsqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
        If PAY.recordCount Then
            With PAY
                tmpPAYName = IIf(Len(!PAYCompanyName), !PAYCompanyName, "")
            End With
        End If
        PAY.Close
        Set PAY = Nothing
        
        bookmark = False
        
        objdoc.Bookmarks("PrincipalName").Range.Text = tmpMBMName
        objdoc.Bookmarks("Add1").Range.Text = tmpMBMAddress1
        objdoc.Bookmarks("Add2").Range.Text = tmpMBMAddress2 & tmpMBMAddress3
        objdoc.Bookmarks("PostCode").Range.Text = tmpMBMPostCode & "" & tmpCTYCode
        objdoc.Bookmarks("State").Range.Text = tmpSTACode
        
        If txtDiag <> "" Then
            objdoc.Bookmarks("Reasons").Range.Text = txtDiag
        End If

        If lblMBMNumber <> "" Then
            objdoc.Bookmarks("mbmnumber").Range.Text = lblMBMNumber
        End If
        
        If txtMBMPolNo <> "" Then
            objdoc.Bookmarks("policynumber").Range.Text = txtMBMPolNo
        End If
        
        If txtMBMName <> "" Then
            objdoc.Bookmarks("InsuredName").Range.Text = txtMBMName
        End If

        If Trim(lblPatient) <> "" Then
          objdoc.Bookmarks("patientname").Range.Text = Trim(lblPatient)
        End If
                
                
        objdoc.Bookmarks("hospital").Range.Text = Trim(lstSpecialNoteList.SelectedItem.SubItems(1))

        If txtMBMPayorExpDate <> "" Then
            objdoc.Bookmarks("ExpDate").Range.Text = txtMBMPayorExpDate
        End If
        
        objdoc.Bookmarks("DOA").Range.Text = Trim(lstSpecialNoteList.SelectedItem.SubItems(3))
                
        If Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - 1)) <> "RC6" Then
            If Len(Trim(cboREPCode1)) Then
                'strsqlREP = "Select * from REPCode where RepCode = '" & Trim(Mid(cboREPCode1, 1, 7)) & "'"
                'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "20"
                StrAppend1 = " AND RepCode = '" & Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - 1)) & "'"
                Set cmd3.ActiveConnection = medixmasterconn
                cmd3.CommandText = "Case_Adjustment"
                cmd3.CommandType = adCmdStoredProc
                cmd3.CommandTimeout = 300
                Set Prm1 = cmd3.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd3.Parameters.Append Prm1
                Set Prm2 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd3.Parameters.Append Prm2
                REP.CursorLocation = adUseClient
                REP.CursorType = adOpenForwardOnly
                REP.CacheSize = 100
                Set REP = cmd3.Execute
        
                'strsqlREPCat = "Select * from REPCodeCategory where REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                'REPCat.Open strsqlREPCat, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "21"
                StrAppend1 = " AND REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                Set cmd4.ActiveConnection = medixmasterconn
                cmd4.CommandText = "Case_Adjustment"
                cmd4.CommandType = adCmdStoredProc
                cmd4.CommandTimeout = 300
                Set Prm1 = cmd4.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd4.Parameters.Append Prm1
                Set Prm2 = cmd4.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd4.Parameters.Append Prm2
                REPCat.CursorLocation = adUseClient
                REPCat.CursorType = adOpenForwardOnly
                REPCat.CacheSize = 100
                Set REPCat = cmd4.Execute
                
                If Len(REPCat!RepCatName) Then
                    objdoc.Bookmarks("REPCatNamePol1").Range.Text = Trim(REPCat!RepCatName)
                    objdoc.Bookmarks("REPCatName1").Range.Text = Trim(REPCat!RepCatName)
                End If
                
                If Len(REPCat!REPCatDescription) Then
                    objdoc.Bookmarks("REPCat1Desc").Range.Text = Trim(REPCat!REPCatDescription)
                End If
                
                objdoc.Bookmarks("REPReason1").Range.Text = Trim(REP!RepDescription)
        
                REP.Close
                REPCat.Close
                Set REP = Nothing
                Set REPCat = Nothing
            End If
            
            If Len(Trim(cboREPCode2)) Then
                'strsqlREP = "Select * from REPCode where RepCode = '" & Trim(Mid(cboREPCode2, 1, 7)) & "'"
                'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "20"
                StrAppend1 = " AND RepCode = '" & Trim(Mid(cboREPCode2, 1, InStr(1, cboREPCode2, "-") - 1)) & "'"
                Set cmd5.ActiveConnection = medixmasterconn
                cmd5.CommandText = "Case_Adjustment"
                cmd5.CommandType = adCmdStoredProc
                cmd5.CommandTimeout = 300
                Set Prm1 = cmd5.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd5.Parameters.Append Prm1
                Set Prm2 = cmd5.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd5.Parameters.Append Prm2
                REP.CursorLocation = adUseClient
                REP.CursorType = adOpenForwardOnly
                REP.CacheSize = 100
                Set REP = cmd5.Execute
                
                'strsqlREPCat = "Select * from REPCodeCategory where REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                'REPCat.Open strsqlREPCat, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "21"
                StrAppend1 = " AND REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                Set cmd6.ActiveConnection = medixmasterconn
                cmd6.CommandText = "Case_Adjustment"
                cmd6.CommandType = adCmdStoredProc
                cmd6.CommandTimeout = 300
                Set Prm1 = cmd6.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd6.Parameters.Append Prm1
                Set Prm2 = cmd6.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd6.Parameters.Append Prm2
                REPCat.CursorLocation = adUseClient
                REPCat.CursorType = adOpenForwardOnly
                REPCat.CacheSize = 100
                Set REPCat = cmd6.Execute
                
                If Not REPCat.EOF Then
                    objdoc.Bookmarks("REPCatNamePol2").Range.Text = ", " & Trim(REPCat!RepCatName)
                    objdoc.Bookmarks("REPCatName2").Range.Text = Trim(REPCat!RepCatName)
                End If
                
                If Len(REPCat!REPCatDescription) Then
                    objdoc.Bookmarks("REPCat2Desc").Range.Text = Trim(REPCat!REPCatDescription)
                End If
    
                objdoc.Bookmarks("REPReason2").Range.Text = Trim(REP!RepDescription)
                
                REP.Close
                REPCat.Close
                Set REP = Nothing
                Set REPCat = Nothing
            End If
            
            If Len(Trim(cboREPCode3)) Then
                'strsqlREP = "Select * from REPCode where RepCode = '" & Trim(Mid(cboREPCode3, 1, 7)) & "'"
                'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "20"
                StrAppend1 = " AND RepCode = '" & Trim(Mid(cboREPCode3, 1, InStr(1, cboREPCode3, "-") - 1)) & "'"
                Set cmd7.ActiveConnection = medixmasterconn
                cmd7.CommandText = "Case_Adjustment"
                cmd7.CommandType = adCmdStoredProc
                cmd7.CommandTimeout = 300
                Set Prm1 = cmd7.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd7.Parameters.Append Prm1
                Set Prm2 = cmd7.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd7.Parameters.Append Prm2
                REP.CursorLocation = adUseClient
                REP.CursorType = adOpenForwardOnly
                REP.CacheSize = 100
                Set REP = cmd7.Execute
                
                'strsqlREPCat = "Select * from REPCodeCategory where REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                'REPCat.Open strsqlREPCat, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "21"
                StrAppend1 = " AND REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                Set cmd8.ActiveConnection = medixmasterconn
                cmd8.CommandText = "Case_Adjustment"
                cmd8.CommandType = adCmdStoredProc
                cmd8.CommandTimeout = 300
                Set Prm1 = cmd8.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd8.Parameters.Append Prm1
                Set Prm2 = cmd8.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd8.Parameters.Append Prm2
                REPCat.CursorLocation = adUseClient
                REPCat.CursorType = adOpenForwardOnly
                REPCat.CacheSize = 100
                Set REPCat = cmd8.Execute
                
                If Not REP.EOF Then
                    objdoc.Bookmarks("REPCatNamePol3").Range.Text = ", " & Trim(REPCat!RepCatName)
                    objdoc.Bookmarks("REPCatName3").Range.Text = Trim(REPCat!RepCatName)
                End If
                
                If Len(REPCat!REPCatDescription) Then
                    objdoc.Bookmarks("REPCat3Desc").Range.Text = REPCat!REPCatDescription
                End If
    
                objdoc.Bookmarks("REPReason3").Range.Text = Trim(REP!RepDescription)
                    
                REP.Close
                REPCat.Close
                Set REP = Nothing
                Set REPCat = Nothing
            End If
        End If
        If tmpPAYName <> "" Then
            objdoc.Bookmarks("Payor").Range.Text = tmpPAYName
        End If
       
        If tmpPAYName <> "" Then
            objdoc.Bookmarks("Payor2").Range.Text = tmpPAYName
        End If
        
        'If Len(PAY!PAYAdmissionContact) Then
        '    objdoc.Bookmarks("Pic").Range.Text = Trim(PAY!PAYAdmissionContact)
        'End If
        bookmark = True

End Function

Private Function RepudiationList()
Dim strsql As String
Dim rst As New ADODB.Recordset
Dim Record As ListItem
        
        lstRepList.ListItems.Clear
        strsql = "Select MBMIndex,MBMNumber,MBMPatientName,RepReason1,RepReason2,RepReason3 from MBMExpRepudiation where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        rst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If rst.recordCount > 0 Then
            Do While Not rst.EOF
                With rst
                    Set Record = lstRepList.ListItems.Add(, , !MBMNumber)
                        Record.SubItems(1) = !MBMPatientName
                        Record.SubItems(2) = !RepReason1
                        Record.SubItems(3) = !RepReason2
                        Record.SubItems(4) = !RepReason3
                        Record.SubItems(5) = !MBMIndex
                End With
            rst.MoveNext
            Loop
        End If
        rst.Close
        Set rst = Nothing
End Function

Private Sub GetRepudiation()
    Frame1.Enabled = False
    If lstRepList.ListItems.Count > 0 Then
        With lstRepList.SelectedItem
            'cboPatient = .Text
            lblPatient = .SubItems(1)
            cboREPCode1 = .SubItems(2)
            cboREPCode2 = .SubItems(3)
            cboREPCode3 = .SubItems(4)
            lblRepIndex = .SubItems(5)
        End With
    End If
End Sub
