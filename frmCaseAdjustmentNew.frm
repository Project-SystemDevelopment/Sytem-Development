VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseAdjustmentNew 
   Caption         =   "frmCaseAdjustment"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   12075
   Begin VB.Frame Frame20 
      Height          =   495
      Index           =   10
      Left            =   120
      TabIndex        =   99
      Top             =   240
      Width           =   11895
      Begin VB.ComboBox cboPending 
         Height          =   315
         Left            =   10560
         TabIndex        =   308
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox cboClaimability 
         Height          =   315
         Left            =   8040
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   300
         Left            =   3890
         TabIndex        =   3
         Top             =   130
         Width           =   690
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Prev"
         Height          =   300
         Left            =   3200
         TabIndex        =   2
         Top             =   130
         Width           =   690
      End
      Begin VB.ComboBox cboStatus1 
         Height          =   315
         Left            =   6280
         TabIndex        =   5
         Top             =   140
         Width           =   855
      End
      Begin VB.TextBox txtCaseNo 
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
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         Top             =   160
         Width           =   1515
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4560
         TabIndex        =   4
         Top             =   130
         Width           =   690
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Height          =   300
         Left            =   2640
         TabIndex        =   1
         Top             =   130
         Width           =   570
      End
      Begin VB.Label Label65 
         Caption         =   "Pending"
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
         Index           =   34
         Left            =   9960
         TabIndex        =   309
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Claimability"
         Height          =   255
         Index           =   108
         Left            =   7200
         TabIndex        =   102
         Top             =   165
         Width           =   855
      End
      Begin VB.Label Label65 
         Caption         =   "Case Status"
         Height          =   255
         Index           =   107
         Left            =   5400
         TabIndex        =   101
         Top             =   165
         Width           =   855
      End
      Begin VB.Label Label65 
         Caption         =   "Cas Number"
         Height          =   255
         Index           =   37
         Left            =   120
         TabIndex        =   100
         Top             =   180
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7065
      Left            =   120
      TabIndex        =   114
      Top             =   1440
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12462
      _Version        =   393216
      Tabs            =   13
      Tab             =   8
      TabsPerRow      =   13
      TabHeight       =   529
      TabCaption(0)   =   "General Info"
      TabPicture(0)   =   "frmCaseAdjustmentNew.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstMember"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Case Details"
      TabPicture(1)   =   "frmCaseAdjustmentNew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame20(4)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Case Others"
      TabPicture(2)   =   "frmCaseAdjustmentNew.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Take Over/Others"
      TabPicture(3)   =   "frmCaseAdjustmentNew.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame20(3)"
      Tab(3).Control(1)=   "Frame20(18)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Benefit"
      TabPicture(4)   =   "frmCaseAdjustmentNew.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstBenefits"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Audit Review"
      TabPicture(5)   =   "frmCaseAdjustmentNew.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label65(110)"
      Tab(5).Control(1)=   "Label65(111)"
      Tab(5).Control(2)=   "cboSuppExc"
      Tab(5).Control(3)=   "txtPrincipalExc"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Case Mgmt Notes"
      TabPicture(6)   =   "frmCaseAdjustmentNew.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label65(98)"
      Tab(6).Control(1)=   "Label65(97)"
      Tab(6).Control(2)=   "Label65(96)"
      Tab(6).Control(3)=   "lblPage"
      Tab(6).Control(4)=   "lblPage1"
      Tab(6).Control(5)=   "lblPage2"
      Tab(6).Control(6)=   "lblPage3"
      Tab(6).Control(7)=   "lblPage4"
      Tab(6).Control(8)=   "lblPage5"
      Tab(6).Control(9)=   "txtCASRemark4"
      Tab(6).Control(10)=   "txtCASRemark5"
      Tab(6).Control(11)=   "txtCASRemark3"
      Tab(6).Control(12)=   "txtCASRemarkTwo"
      Tab(6).Control(13)=   "txtCASRemark"
      Tab(6).Control(14)=   "Frame20(17)"
      Tab(6).Control(15)=   "Frame20(16)"
      Tab(6).ControlCount=   16
      TabCaption(7)   =   "Mgmt 1"
      TabPicture(7)   =   "frmCaseAdjustmentNew.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label1(11)"
      Tab(7).Control(1)=   "Label1(2)"
      Tab(7).Control(2)=   "txtSpecialRemark"
      Tab(7).Control(3)=   "txtInvRemark"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   "Mgmt2"
      TabPicture(8)   =   "frmCaseAdjustmentNew.frx":00E0
      Tab(8).ControlEnabled=   -1  'True
      Tab(8).Control(0)=   "Label1(9)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Label1(6)"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "Label1(5)"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).Control(3)=   "Label1(1)"
      Tab(8).Control(3).Enabled=   0   'False
      Tab(8).Control(4)=   "Label1(12)"
      Tab(8).Control(4).Enabled=   0   'False
      Tab(8).Control(5)=   "lstNote1(5)"
      Tab(8).Control(5).Enabled=   0   'False
      Tab(8).Control(6)=   "lstNote1(3)"
      Tab(8).Control(6).Enabled=   0   'False
      Tab(8).Control(7)=   "lstNote1(8)"
      Tab(8).Control(7).Enabled=   0   'False
      Tab(8).Control(8)=   "lstINVListing"
      Tab(8).Control(8).Enabled=   0   'False
      Tab(8).Control(9)=   "lstNote1(2)"
      Tab(8).Control(9).Enabled=   0   'False
      Tab(8).ControlCount=   10
      TabCaption(9)   =   "Mgmt3"
      TabPicture(9)   =   "frmCaseAdjustmentNew.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label1(4)"
      Tab(9).Control(1)=   "Label1(3)"
      Tab(9).Control(2)=   "Label1(10)"
      Tab(9).Control(3)=   "Label1(7)"
      Tab(9).Control(4)=   "lstLGListing1"
      Tab(9).Control(5)=   "lstNote1(7)"
      Tab(9).Control(6)=   "lstNote1(1)"
      Tab(9).Control(7)=   "lstNote1(0)"
      Tab(9).Control(8)=   "lstLGListing"
      Tab(9).Control(9)=   "cmdDocReqAdm"
      Tab(9).Control(10)=   "cmdClearDoc"
      Tab(9).ControlCount=   11
      TabCaption(10)  =   "WS"
      TabPicture(10)  =   "frmCaseAdjustmentNew.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lstWorksheetListing"
      Tab(10).Control(1)=   "Frame20(12)"
      Tab(10).ControlCount=   2
      TabCaption(11)  =   "Forms"
      TabPicture(11)  =   "frmCaseAdjustmentNew.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "CmnDlg"
      Tab(11).Control(1)=   "cmdDelGL"
      Tab(11).Control(2)=   "List3"
      Tab(11).Control(3)=   "Frame20(21)"
      Tab(11).Control(4)=   "Frame20(19)"
      Tab(11).Control(5)=   "List1"
      Tab(11).ControlCount=   6
      TabCaption(12)  =   "Policy"
      TabPicture(12)  =   "frmCaseAdjustmentNew.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "lstPolicyWording"
      Tab(12).ControlCount=   1
      Begin VB.ListBox lstPolicyWording 
         Height          =   5130
         ItemData        =   "frmCaseAdjustmentNew.frx":016C
         Left            =   -74880
         List            =   "frmCaseAdjustmentNew.frx":016E
         TabIndex        =   430
         ToolTipText     =   "right click to randomize contents"
         Top             =   480
         Width           =   4455
      End
      Begin VB.ListBox List1 
         Height          =   2010
         ItemData        =   "frmCaseAdjustmentNew.frx":0170
         Left            =   -74880
         List            =   "frmCaseAdjustmentNew.frx":0172
         TabIndex        =   429
         ToolTipText     =   "right click to randomize contents"
         Top             =   540
         Width           =   4575
      End
      Begin VB.Frame Frame20 
         Height          =   1035
         Index           =   19
         Left            =   -70080
         TabIndex        =   422
         Top             =   420
         Width           =   6645
         Begin VB.CommandButton CmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Left            =   5760
            TabIndex        =   426
            Top             =   630
            Width           =   735
         End
         Begin VB.TextBox TxtDest 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   425
            Top             =   600
            Width           =   4575
         End
         Begin VB.TextBox TxtSrc 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   424
            Top             =   240
            Width           =   4575
         End
         Begin VB.CommandButton CmdBrowse 
            Caption         =   "Browse"
            Height          =   255
            Left            =   5760
            TabIndex        =   423
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Destination :"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   428
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Source :"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   427
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Internal Documents"
         Height          =   2655
         Index           =   21
         Left            =   -74880
         TabIndex        =   420
         Top             =   2820
         Width           =   4695
         Begin VB.ListBox List2 
            Height          =   2310
            ItemData        =   "frmCaseAdjustmentNew.frx":0174
            Left            =   120
            List            =   "frmCaseAdjustmentNew.frx":0176
            Style           =   1  'Checkbox
            TabIndex        =   421
            ToolTipText     =   "right click to randomize contents"
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.ListBox List3 
         Height          =   2310
         ItemData        =   "frmCaseAdjustmentNew.frx":0178
         Left            =   -69960
         List            =   "frmCaseAdjustmentNew.frx":017A
         Style           =   1  'Checkbox
         TabIndex        =   419
         ToolTipText     =   "right click to randomize contents"
         Top             =   3060
         Width           =   4455
      End
      Begin VB.CommandButton cmdDelGL 
         Caption         =   "Delete GL"
         Height          =   375
         Left            =   -65280
         TabIndex        =   418
         Top             =   4980
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdClearDoc 
         Caption         =   "Clear Doc Record"
         Height          =   405
         Left            =   -64440
         TabIndex        =   406
         Top             =   6060
         Width           =   1215
      End
      Begin VB.TextBox txtInvRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   -74880
         MaxLength       =   1500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   403
         Top             =   4920
         Width           =   11655
      End
      Begin VB.TextBox txtSpecialRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4020
         Left            =   -74880
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   398
         Top             =   600
         Width           =   11655
      End
      Begin VB.CommandButton cmdDocReqAdm 
         Caption         =   "Delete"
         Height          =   285
         Left            =   -65400
         TabIndex        =   390
         Top             =   6060
         Width           =   855
      End
      Begin VB.Frame Frame20 
         Height          =   615
         Index           =   12
         Left            =   -74880
         TabIndex        =   375
         Top             =   5340
         Width           =   11415
         Begin VB.ComboBox cboBordxRevStatus 
            Height          =   315
            Left            =   10440
            TabIndex        =   379
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cboREVBordxInd 
            Height          =   315
            Left            =   960
            TabIndex        =   378
            Top             =   675
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtCLMAppCode 
            Height          =   285
            Left            =   7080
            MaxLength       =   15
            TabIndex        =   377
            Top             =   240
            Width           =   2055
         End
         Begin VB.ComboBox cboAppStatus 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5040
            TabIndex        =   376
            Top             =   200
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtAppDate 
            Height          =   300
            Left            =   2880
            TabIndex        =   380
            Top             =   195
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label65 
            Caption         =   "Reverse Status"
            Height          =   255
            Index           =   69
            Left            =   9240
            TabIndex        =   388
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "Rev Bordx"
            Height          =   255
            Index           =   119
            Left            =   120
            TabIndex        =   387
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "App Code"
            Height          =   255
            Index           =   118
            Left            =   6360
            TabIndex        =   386
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblClaimNo 
            Height          =   255
            Left            =   840
            TabIndex        =   385
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "Claim No"
            Height          =   255
            Index           =   115
            Left            =   120
            TabIndex        =   384
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "App Date"
            Height          =   255
            Index           =   113
            Left            =   2160
            TabIndex        =   383
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label65 
            Caption         =   "App Status"
            Height          =   255
            Index           =   112
            Left            =   4200
            TabIndex        =   382
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblBordxAmt 
            Caption         =   "Label1"
            Height          =   135
            Left            =   1200
            TabIndex        =   381
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame20 
         Height          =   615
         Index           =   18
         Left            =   -74880
         TabIndex        =   167
         Top             =   420
         Width           =   11415
         Begin VB.ComboBox cboIntExtTO 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5520
            TabIndex        =   93
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtLinkCASNo 
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
            Left            =   10080
            MaxLength       =   15
            TabIndex        =   95
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtCASTOPlan 
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
            Left            =   8040
            MaxLength       =   4
            TabIndex        =   94
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtCASTakeOverAmt 
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
            TabIndex        =   92
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cboCASTakeOveInd 
            Height          =   315
            Left            =   840
            TabIndex        =   91
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "Link Case No"
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
            Index           =   92
            Left            =   9120
            TabIndex        =   172
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "T/O Plan"
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
            Index           =   66
            Left            =   7320
            TabIndex        =   171
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label65 
            Caption         =   "T/O Amt"
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
            Index           =   61
            Left            =   2160
            TabIndex        =   170
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label65 
            Caption         =   "T/O Ind"
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
            Index           =   60
            Left            =   120
            TabIndex        =   169
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label65 
            Caption         =   "Int/Ext TakeOver"
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
            Index           =   109
            Left            =   4200
            TabIndex        =   168
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtPrincipalExc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         Left            =   -74880
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   166
         Top             =   660
         Width           =   11415
      End
      Begin VB.ComboBox cboSuppExc 
         Height          =   315
         Left            =   -73680
         TabIndex        =   165
         Top             =   2700
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame Frame20 
         Caption         =   "Frame12"
         Height          =   4935
         Index           =   16
         Left            =   -67560
         TabIndex        =   164
         Top             =   960
         Width           =   15
      End
      Begin VB.Frame Frame20 
         Caption         =   "Frame11"
         Height          =   4935
         Index           =   17
         Left            =   -73560
         TabIndex        =   163
         Top             =   960
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Height          =   5445
         Left            =   -74880
         TabIndex        =   158
         Top             =   480
         Width           =   11415
         Begin VB.Frame FrameNS13RateMin 
            Enabled         =   0   'False
            Height          =   2055
            Index           =   5
            Left            =   8760
            TabIndex        =   349
            Top             =   120
            Width           =   975
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   29
               Left            =   60
               TabIndex        =   355
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   28
               Left            =   60
               TabIndex        =   354
               Top             =   1440
               Width           =   855
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   27
               Left            =   60
               TabIndex        =   353
               Top             =   940
               Width           =   855
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   26
               Left            =   60
               TabIndex        =   352
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   25
               Left            =   60
               TabIndex        =   351
               Top             =   680
               Width           =   855
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   24
               Left            =   60
               TabIndex        =   350
               Top             =   400
               Width           =   855
            End
            Begin VB.Label Label90 
               Alignment       =   2  'Center
               Caption         =   "Anaes 13th New"
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
               Index           =   4
               Left            =   120
               TabIndex        =   356
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame FrameNS13RateMin 
            Enabled         =   0   'False
            Height          =   2055
            Index           =   4
            Left            =   4320
            TabIndex        =   341
            Top             =   120
            Width           =   1095
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   23
               Left            =   60
               TabIndex        =   347
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   22
               Left            =   60
               TabIndex        =   346
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   21
               Left            =   60
               TabIndex        =   345
               Top             =   940
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   20
               Left            =   60
               TabIndex        =   344
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   19
               Left            =   60
               TabIndex        =   343
               Top             =   680
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   18
               Left            =   60
               TabIndex        =   342
               Top             =   400
               Width           =   975
            End
            Begin VB.Label Label90 
               Alignment       =   2  'Center
               Caption         =   "13th New"
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
               Index           =   3
               Left            =   120
               TabIndex        =   348
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame FrameNS13RateMin 
            Enabled         =   0   'False
            Height          =   2055
            Index           =   3
            Left            =   5400
            TabIndex        =   333
            Top             =   120
            Width           =   1095
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   17
               Left            =   60
               TabIndex        =   339
               Top             =   400
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   16
               Left            =   60
               TabIndex        =   338
               Top             =   680
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   15
               Left            =   60
               TabIndex        =   337
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   14
               Left            =   60
               TabIndex        =   336
               Top             =   940
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   13
               Left            =   60
               TabIndex        =   335
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   12
               Left            =   60
               TabIndex        =   334
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label90 
               Alignment       =   2  'Center
               Caption         =   "13th New (GST)"
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
               Index           =   2
               Left            =   240
               TabIndex        =   340
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame FrameNS13RateMin 
            Enabled         =   0   'False
            Height          =   2055
            Index           =   2
            Left            =   4440
            TabIndex        =   324
            Top             =   2160
            Width           =   1095
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   11
               Left            =   60
               TabIndex        =   330
               Top             =   400
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   10
               Left            =   60
               TabIndex        =   329
               Top             =   680
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   9
               Left            =   60
               TabIndex        =   328
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   8
               Left            =   60
               TabIndex        =   327
               Top             =   940
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   7
               Left            =   60
               TabIndex        =   326
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   6
               Left            =   60
               TabIndex        =   325
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label90 
               Alignment       =   2  'Center
               Caption         =   "13th New"
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
               Index           =   1
               Left            =   120
               TabIndex        =   332
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame FrameNS13RateMin 
            Enabled         =   0   'False
            Height          =   2055
            Index           =   1
            Left            =   5400
            TabIndex        =   317
            Top             =   2160
            Width           =   1095
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   5
               Left            =   60
               TabIndex        =   323
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   4
               Left            =   60
               TabIndex        =   322
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   3
               Left            =   60
               TabIndex        =   321
               Top             =   940
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   2
               Left            =   60
               TabIndex        =   320
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   1
               Left            =   60
               TabIndex        =   319
               Top             =   680
               Width           =   975
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Index           =   0
               Left            =   60
               TabIndex        =   318
               Top             =   400
               Width           =   975
            End
            Begin VB.Label Label90 
               Alignment       =   2  'Center
               Caption         =   "13th New (GST)"
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
               Index           =   0
               Left            =   240
               TabIndex        =   331
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame Frame13Rate 
            Enabled         =   0   'False
            Height          =   2025
            Left            =   9480
            TabIndex        =   294
            Top             =   5400
            Width           =   1095
            Begin VB.TextBox txtAnae13Rate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   300
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtAnae13Rate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   299
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox txtAnae13Rate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   298
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtAnae13Rate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   297
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox txtAnae13Rate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   296
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox txtAnae13Rate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   295
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Anaes 13th Sch"
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
               Index           =   6
               Left            =   120
               TabIndex        =   301
               Top             =   105
               Width           =   735
            End
         End
         Begin VB.Frame FrameS13Rate 
            Enabled         =   0   'False
            Height          =   2055
            Left            =   6480
            TabIndex        =   285
            Top             =   120
            Width           =   1095
            Begin VB.TextBox txtS13Rate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   293
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox txtS13Rate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   290
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox txtS13Rate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   289
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox txtS13Rate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   288
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtS13Rate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   287
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox txtS13Rate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   286
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "13th Old"
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
               Index           =   1
               Left            =   120
               TabIndex        =   292
               Top             =   105
               Width           =   735
            End
         End
         Begin VB.Frame FrameS13MinRate 
            Enabled         =   0   'False
            Height          =   2055
            Left            =   7320
            TabIndex        =   278
            Top             =   5400
            Width           =   735
            Begin VB.TextBox txtS13MinRate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   284
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtS13MinRate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   283
               Top             =   1440
               Width           =   615
            End
            Begin VB.TextBox txtS13MinRate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   282
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox txtS13MinRate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   281
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtS13MinRate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   280
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox txtS13MinRate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   279
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "13th Old (Min)"
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
               Index           =   0
               Left            =   0
               TabIndex        =   291
               Top             =   120
               Width           =   735
            End
         End
         Begin VB.Frame FrameNS13Rate 
            Enabled         =   0   'False
            Height          =   2055
            Left            =   6480
            TabIndex        =   270
            Top             =   2160
            Width           =   1215
            Begin VB.TextBox txtNS13Rate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   276
               Top             =   400
               Width           =   1095
            End
            Begin VB.TextBox txtNS13Rate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   275
               Top             =   660
               Width           =   1095
            End
            Begin VB.TextBox txtNS13Rate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   274
               Top             =   900
               Width           =   1095
            End
            Begin VB.TextBox txtNS13Rate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   273
               Top             =   1140
               Width           =   1095
            End
            Begin VB.TextBox txtNS13Rate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   272
               Top             =   1410
               Width           =   1095
            End
            Begin VB.TextBox txtNS13Rate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   271
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label90 
               Alignment       =   2  'Center
               Caption         =   "13th Old"
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
               Index           =   10
               Left            =   240
               TabIndex        =   277
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame FrameNS13RateMin 
            Enabled         =   0   'False
            Height          =   2055
            Index           =   0
            Left            =   9720
            TabIndex        =   262
            Top             =   5400
            Width           =   1095
            Begin VB.TextBox txtNS13MinRate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   268
               Top             =   400
               Width           =   975
            End
            Begin VB.TextBox txtNS13MinRate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   267
               Top             =   660
               Width           =   975
            End
            Begin VB.TextBox txtNS13MinRate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   266
               Top             =   900
               Width           =   975
            End
            Begin VB.TextBox txtNS13MinRate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   265
               Top             =   1140
               Width           =   975
            End
            Begin VB.TextBox txtNS13MinRate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   264
               Top             =   1410
               Width           =   975
            End
            Begin VB.TextBox txtNS13MinRate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   60
               TabIndex        =   263
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label90 
               Alignment       =   2  'Center
               Caption         =   "13th Sch (Min)"
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
               Index           =   12
               Left            =   0
               TabIndex        =   269
               Top             =   120
               Width           =   735
            End
         End
         Begin VB.Frame Frame20 
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
            Height          =   2040
            Index           =   28
            Left            =   9240
            TabIndex        =   254
            Top             =   5400
            Width           =   1215
            Begin VB.TextBox txtNSMMARate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   80
               TabIndex        =   261
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox txtNSMMARate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   80
               TabIndex        =   260
               Top             =   1410
               Width           =   1095
            End
            Begin VB.TextBox txtNSMMARate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   80
               TabIndex        =   259
               Top             =   1140
               Width           =   1095
            End
            Begin VB.TextBox txtNSMMARate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   80
               TabIndex        =   258
               Top             =   900
               Width           =   1095
            End
            Begin VB.TextBox txtNSMMARate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   80
               TabIndex        =   257
               Top             =   660
               Width           =   1095
            End
            Begin VB.TextBox txtNSMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   80
               TabIndex        =   256
               Top             =   400
               Width           =   1095
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "MMA"
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
               TabIndex        =   255
               Top             =   120
               Width           =   615
            End
         End
         Begin VB.Frame Frame20 
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
            Height          =   2025
            Index           =   13
            Left            =   7680
            TabIndex        =   231
            Top             =   120
            Width           =   1095
            Begin VB.TextBox txtAnaesActFees6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   237
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox txtAnaesActFees1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   236
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtAnaesActFees2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   235
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox txtAnaesActFees3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   234
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtAnaesActFees4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   233
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox txtAnaesActFees5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   232
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Anaes Actual"
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
               Index           =   4
               Left            =   120
               TabIndex        =   250
               Top             =   105
               Width           =   735
            End
         End
         Begin VB.Frame FrameAnaesMMA 
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
            Height          =   2025
            Left            =   8160
            TabIndex        =   217
            Top             =   5400
            Width           =   1095
            Begin VB.TextBox txtAnaesMMARate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   223
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   222
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   221
               Top             =   720
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   220
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   219
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   218
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Anaes MMA"
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
               Index           =   5
               Left            =   240
               TabIndex        =   251
               Top             =   120
               Width           =   615
            End
         End
         Begin VB.Frame Frame20 
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
            Height          =   2025
            Index           =   20
            Left            =   10200
            TabIndex        =   210
            Top             =   5400
            Width           =   1215
            Begin VB.TextBox txtSMMARate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   45
               TabIndex        =   216
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox txtSMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   45
               TabIndex        =   215
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox txtSMMARate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   45
               TabIndex        =   214
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtSMMARate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   45
               TabIndex        =   213
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox txtSMMARate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   45
               TabIndex        =   212
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox txtSMMARate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   45
               TabIndex        =   211
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "MMA"
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
               Index           =   2
               Left            =   240
               TabIndex        =   248
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "OPCS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1785
            Index           =   14
            Left            =   2280
            TabIndex        =   209
            Top             =   360
            Width           =   1035
            Begin VB.CommandButton cmdOPCS 
               Caption         =   ">"
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
               Index           =   1
               Left            =   720
               TabIndex        =   239
               Top             =   0
               Width           =   255
            End
            Begin VB.TextBox txtProcCode1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   61
               Top             =   240
               Width           =   900
            End
            Begin VB.TextBox txtProcCode2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   63
               Top             =   480
               Width           =   900
            End
            Begin VB.TextBox txtProcCode3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   65
               Top             =   720
               Width           =   900
            End
            Begin VB.TextBox txtProcCode4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   67
               Top             =   960
               Width           =   900
            End
            Begin VB.TextBox txtProcCode5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   69
               Top             =   1200
               Width           =   900
            End
            Begin VB.TextBox txtProcCode6 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   71
               Top             =   1440
               Width           =   900
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Surgical Procedure"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1785
            Index           =   99
            Left            =   120
            TabIndex        =   208
            Top             =   360
            Width           =   2145
            Begin VB.TextBox txtProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   60
               Top             =   240
               Width           =   2055
            End
            Begin VB.TextBox txtProcedure2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   62
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtProcedure3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   64
               Top             =   720
               Width           =   2055
            End
            Begin VB.TextBox txtProcedure4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   66
               Top             =   960
               Width           =   2055
            End
            Begin VB.TextBox txtProcedure5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   68
               Top             =   1200
               Width           =   2055
            End
            Begin VB.TextBox txtProcedure6 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   70
               Top             =   1440
               Width           =   2055
            End
         End
         Begin VB.Frame Frame19 
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
            Height          =   2040
            Index           =   1
            Left            =   3360
            TabIndex        =   201
            Top             =   2160
            Width           =   1185
            Begin VB.TextBox txtNSActFee1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   207
               Top             =   405
               Width           =   1020
            End
            Begin VB.TextBox txtNSActFee2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   206
               Top             =   645
               Width           =   1020
            End
            Begin VB.TextBox txtNSActFee3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   205
               Top             =   885
               Width           =   1020
            End
            Begin VB.TextBox txtNSActFee4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   204
               Top             =   1140
               Width           =   1020
            End
            Begin VB.TextBox txtNSActFee5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   203
               Top             =   1410
               Width           =   1020
            End
            Begin VB.TextBox txtNSActFee6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   202
               Top             =   1680
               Width           =   1020
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Actual"
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
               Index           =   8
               Left            =   240
               TabIndex        =   252
               Top             =   120
               Width           =   615
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "OPCS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1920
            Index           =   22
            Left            =   2280
            TabIndex        =   199
            Top             =   2280
            Width           =   1035
            Begin VB.TextBox txtNSProcCode1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   80
               MaxLength       =   6
               TabIndex        =   73
               Top             =   300
               Width           =   885
            End
            Begin VB.CommandButton cmdOPCS 
               Caption         =   ">"
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
               Index           =   0
               Left            =   720
               TabIndex        =   200
               Top             =   0
               Width           =   255
            End
            Begin VB.TextBox txtNSProcCode2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   80
               MaxLength       =   6
               TabIndex        =   74
               Top             =   555
               Width           =   885
            End
            Begin VB.TextBox txtNSProcCode3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   80
               MaxLength       =   6
               TabIndex        =   75
               Top             =   795
               Width           =   885
            End
            Begin VB.TextBox txtNSProcCode4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   80
               MaxLength       =   6
               TabIndex        =   76
               Top             =   1050
               Width           =   885
            End
            Begin VB.TextBox txtNSProcCode5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   80
               MaxLength       =   6
               TabIndex        =   77
               Top             =   1305
               Width           =   885
            End
            Begin VB.TextBox txtNSProcCode6 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   80
               MaxLength       =   6
               TabIndex        =   78
               Top             =   1560
               Width           =   885
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Non- Surgical Procedure"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1920
            Index           =   98
            Left            =   120
            TabIndex        =   198
            Top             =   2280
            Width           =   2145
            Begin VB.TextBox txtNSProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   5
               Left            =   40
               MaxLength       =   200
               TabIndex        =   373
               Top             =   1560
               Width           =   2055
            End
            Begin VB.TextBox txtNSProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   4
               Left            =   40
               MaxLength       =   200
               TabIndex        =   372
               Top             =   1320
               Width           =   2055
            End
            Begin VB.TextBox txtNSProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   3
               Left            =   40
               MaxLength       =   200
               TabIndex        =   371
               Top             =   1080
               Width           =   2055
            End
            Begin VB.TextBox txtNSProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   40
               MaxLength       =   200
               TabIndex        =   370
               Top             =   820
               Width           =   2055
            End
            Begin VB.TextBox txtNSProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   40
               MaxLength       =   200
               TabIndex        =   369
               Top             =   560
               Width           =   2055
            End
            Begin VB.TextBox txtNSProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   40
               MaxLength       =   200
               TabIndex        =   72
               Top             =   285
               Width           =   2055
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Medical Management"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1185
            Index           =   1
            Left            =   7800
            TabIndex        =   160
            Top             =   2160
            Width           =   3495
            Begin VB.TextBox txtMEDSURDone1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   85
               Top             =   285
               Width           =   3375
            End
            Begin VB.TextBox txtMEDSURDone2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   86
               Top             =   540
               Width           =   3375
            End
            Begin VB.TextBox txtMEDSURDone3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   150
               TabIndex        =   87
               Top             =   795
               Width           =   3375
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Att Doc"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   2
            Left            =   8040
            TabIndex        =   159
            Top             =   3360
            Width           =   3255
            Begin VB.CommandButton cmdLoadDoc 
               Caption         =   "Load Doc"
               Height          =   255
               Left            =   840
               TabIndex        =   364
               Top             =   1560
               Width           =   1095
            End
            Begin VB.ComboBox txtATTDoc 
               Height          =   315
               Index           =   4
               Left            =   0
               TabIndex        =   363
               Top             =   1200
               Width           =   3135
            End
            Begin VB.ComboBox txtATTDoc 
               Height          =   315
               Index           =   3
               Left            =   0
               TabIndex        =   362
               Top             =   960
               Width           =   3135
            End
            Begin VB.ComboBox txtATTDoc 
               Height          =   315
               Index           =   2
               Left            =   0
               TabIndex        =   361
               Top             =   720
               Width           =   3135
            End
            Begin VB.ComboBox txtATTDoc 
               Height          =   315
               Index           =   1
               Left            =   0
               TabIndex        =   360
               Top             =   480
               Width           =   3135
            End
            Begin VB.ComboBox txtATTDoc 
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   359
               Top             =   240
               Width           =   3135
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Anaes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1115
            Index           =   24
            Left            =   4320
            TabIndex        =   197
            Top             =   4200
            Width           =   3615
            Begin VB.TextBox txtAnaesDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   120
               MaxLength       =   30
               TabIndex        =   90
               Top             =   750
               Width           =   3405
            End
            Begin VB.TextBox txtAnaesDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   120
               MaxLength       =   30
               TabIndex        =   89
               Top             =   495
               Width           =   3405
            End
            Begin VB.TextBox txtAnaesDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   120
               MaxLength       =   30
               TabIndex        =   88
               Top             =   240
               Width           =   3405
            End
         End
         Begin VB.Frame Frame19 
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
            Height          =   2025
            Index           =   0
            Left            =   3360
            TabIndex        =   224
            Top             =   120
            Width           =   945
            Begin VB.TextBox txtSActFee6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   230
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox txtSActFee1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   226
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox txtSActFee2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   228
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtSActFee3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   229
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox txtSActFee4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   227
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtSActFee5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               TabIndex        =   225
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               Caption         =   "Actual"
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
               Left            =   240
               TabIndex        =   249
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Investigation"
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
            Index           =   0
            Left            =   120
            TabIndex        =   162
            Top             =   4200
            Width           =   4095
            Begin VB.CommandButton cmdInvList 
               Caption         =   ">"
               Height          =   195
               Left            =   3720
               TabIndex        =   161
               Top             =   0
               Width           =   255
            End
            Begin VB.TextBox txtPROCINVDone1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   79
               Top             =   285
               Width           =   3015
            End
            Begin VB.TextBox txtPROCINVDone2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   81
               Top             =   540
               Width           =   3015
            End
            Begin VB.TextBox txtPROCINVDone3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   83
               Top             =   795
               Width           =   3015
            End
            Begin VB.TextBox txtINVCode1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               MaxLength       =   150
               TabIndex        =   80
               Top             =   285
               Width           =   855
            End
            Begin VB.TextBox txtINVCode2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               MaxLength       =   150
               TabIndex        =   82
               Top             =   540
               Width           =   855
            End
            Begin VB.TextBox txtINVCode3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               MaxLength       =   150
               TabIndex        =   84
               Top             =   795
               Width           =   855
            End
         End
      End
      Begin VB.Frame Frame20 
         Height          =   5685
         Index           =   4
         Left            =   -74880
         TabIndex        =   117
         Top             =   360
         Width           =   11415
         Begin VB.CheckBox ChkDOD 
            Caption         =   "Check1"
            Height          =   195
            Left            =   11115
            TabIndex        =   138
            Top             =   700
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox ChkDOA 
            Caption         =   "Check1"
            Height          =   255
            Left            =   11115
            TabIndex        =   137
            Top             =   420
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.Frame Frame20 
            Height          =   1365
            Index           =   5
            Left            =   5640
            TabIndex        =   132
            Top             =   2640
            Width           =   5655
            Begin VB.CommandButton CmdDiagFinal 
               Caption         =   ">"
               Height          =   255
               Index           =   2
               Left            =   5280
               TabIndex        =   133
               Top             =   180
               Width           =   225
            End
            Begin VB.TextBox cboFDICDCode1 
               Height          =   315
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   46
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtFinalDiag1 
               Height          =   315
               Index           =   0
               Left            =   90
               MaxLength       =   150
               TabIndex        =   44
               Top             =   480
               Width           =   3870
            End
            Begin VB.TextBox cboFDICDCode2 
               Height          =   315
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   48
               Top             =   735
               Width           =   615
            End
            Begin VB.TextBox cboFDICDCode3 
               Height          =   315
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   50
               Top             =   1005
               Width           =   615
            End
            Begin MSMask.MaskEdBox txtFDDOS1 
               Height          =   315
               Left            =   3960
               TabIndex        =   45
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtFDDOS2 
               Height          =   315
               Left            =   3960
               TabIndex        =   47
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtFDDOS3 
               Height          =   315
               Left            =   3960
               TabIndex        =   49
               Top             =   1005
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtFinalDiag1 
               Height          =   315
               Index           =   1
               Left            =   90
               MaxLength       =   150
               TabIndex        =   401
               Top             =   770
               Width           =   3870
            End
            Begin VB.TextBox txtFinalDiag1 
               Height          =   315
               Index           =   2
               Left            =   90
               MaxLength       =   150
               TabIndex        =   402
               Top             =   1020
               Width           =   3870
            End
            Begin VB.Label Label65 
               Caption         =   "ICD"
               Height          =   255
               Index           =   57
               Left            =   4920
               TabIndex        =   136
               Top             =   180
               Width           =   615
            End
            Begin VB.Label Label65 
               Caption         =   "Date Onset"
               Height          =   255
               Index           =   56
               Left            =   3960
               TabIndex        =   135
               Top             =   180
               Width           =   855
            End
            Begin VB.Label Label65 
               Caption         =   "Final Diag."
               Height          =   255
               Index           =   55
               Left            =   120
               TabIndex        =   134
               Top             =   180
               Width           =   975
            End
         End
         Begin VB.Frame Frame20 
            Height          =   1365
            Index           =   7
            Left            =   120
            TabIndex        =   127
            Top             =   2640
            Width           =   5415
            Begin VB.CommandButton CmdDiagPre 
               Caption         =   ">"
               Height          =   255
               Index           =   0
               Left            =   5040
               TabIndex        =   128
               Top             =   180
               Width           =   225
            End
            Begin VB.TextBox cboICDCode 
               Height          =   315
               Index           =   0
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   37
               Top             =   465
               Width           =   615
            End
            Begin VB.TextBox txtPrelimDiag 
               Height          =   285
               Index           =   0
               Left            =   90
               MaxLength       =   150
               TabIndex        =   35
               Top             =   465
               Width           =   3650
            End
            Begin VB.TextBox cboICDCode 
               Height          =   315
               Index           =   1
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   40
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox cboICDCode 
               Height          =   315
               Index           =   2
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   43
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtPrelimDiag 
               Height          =   285
               Index           =   1
               Left            =   90
               MaxLength       =   150
               TabIndex        =   38
               Top             =   720
               Width           =   3650
            End
            Begin VB.TextBox txtPrelimDiag 
               Height          =   285
               Index           =   2
               Left            =   90
               MaxLength       =   150
               TabIndex        =   41
               Top             =   960
               Width           =   3650
            End
            Begin MSMask.MaskEdBox txtPDDOS 
               Height          =   300
               Index           =   0
               Left            =   3720
               TabIndex        =   36
               Top             =   465
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtPDDOS 
               Height          =   300
               Index           =   1
               Left            =   3720
               TabIndex        =   39
               Top             =   720
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtPDDOS 
               Height          =   300
               Index           =   2
               Left            =   3720
               TabIndex        =   42
               Top             =   960
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label65 
               Caption         =   "ICD"
               Height          =   255
               Index           =   51
               Left            =   4680
               TabIndex        =   131
               Top             =   180
               Width           =   615
            End
            Begin VB.Label Label65 
               Caption         =   "Date Onset"
               Height          =   255
               Index           =   50
               Left            =   3720
               TabIndex        =   130
               Top             =   180
               Width           =   975
            End
            Begin VB.Label Label65 
               Caption         =   "Prelim. Diag"
               Height          =   255
               Index           =   49
               Left            =   120
               TabIndex        =   129
               Top             =   180
               Width           =   1215
            End
         End
         Begin VB.TextBox txtPatientIC 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   20
            Top             =   130
            Width           =   3495
         End
         Begin VB.ComboBox cboPatientList 
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
            Left            =   1400
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   120
            Width           =   4740
         End
         Begin VB.ComboBox txtHOSCode 
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
            Left            =   1400
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   400
            Width           =   4740
         End
         Begin VB.TextBox KIVPatientName 
            Height          =   285
            Left            =   1440
            TabIndex        =   126
            Top             =   120
            Visible         =   0   'False
            Width           =   4635
         End
         Begin VB.Frame Frame20 
            Height          =   1545
            Index           =   6
            Left            =   5640
            TabIndex        =   123
            Top             =   3960
            Width           =   5655
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
               Sorted          =   -1  'True
               TabIndex        =   368
               Top             =   1200
               Width           =   5480
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
               Sorted          =   -1  'True
               TabIndex        =   367
               Top             =   840
               Width           =   5480
            End
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
               Sorted          =   -1  'True
               TabIndex        =   366
               Top             =   480
               Width           =   5480
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Load Repudiation"
               Height          =   255
               Left            =   3360
               TabIndex        =   365
               Top             =   180
               Width           =   1575
            End
            Begin VB.CommandButton CmdDiagRepudiation 
               Caption         =   ">"
               Height          =   255
               Index           =   3
               Left            =   1680
               TabIndex        =   124
               Top             =   180
               Width           =   225
            End
            Begin VB.Label Label65 
               Caption         =   "Repudiation Resons"
               Height          =   255
               Index           =   58
               Left            =   120
               TabIndex        =   125
               Top             =   180
               Width           =   1815
            End
         End
         Begin VB.Frame Frame20 
            Height          =   1545
            Index           =   8
            Left            =   120
            TabIndex        =   118
            Top             =   3960
            Width           =   5415
            Begin VB.CommandButton CmdDiagUD 
               Caption         =   ">"
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   119
               Top             =   180
               Width           =   225
            End
            Begin VB.TextBox txtUD1 
               Height          =   285
               Index           =   0
               Left            =   100
               MaxLength       =   150
               TabIndex        =   51
               Top             =   465
               Width           =   3650
            End
            Begin VB.TextBox txtUD2 
               Height          =   285
               Left            =   100
               MaxLength       =   150
               TabIndex        =   54
               Top             =   720
               Width           =   3650
            End
            Begin VB.TextBox txtUD3 
               Height          =   285
               Left            =   100
               MaxLength       =   150
               TabIndex        =   57
               Top             =   960
               Width           =   3650
            End
            Begin VB.TextBox cboUDICDCode1 
               Height          =   315
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   53
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox cboUDICDCode2 
               Height          =   315
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   56
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox cboUDICDCode3 
               Height          =   315
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   59
               Top             =   960
               Width           =   615
            End
            Begin MSMask.MaskEdBox txtUDDOS1 
               Height          =   300
               Left            =   3720
               TabIndex        =   52
               Top             =   465
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtUDDOS2 
               Height          =   300
               Left            =   3720
               TabIndex        =   55
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtUDDOS3 
               Height          =   300
               Left            =   3720
               TabIndex        =   58
               Top             =   975
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label65 
               Caption         =   "ICD"
               Height          =   255
               Index           =   54
               Left            =   4680
               TabIndex        =   122
               Top             =   180
               Width           =   615
            End
            Begin VB.Label Label65 
               Caption         =   "Date Onset"
               Height          =   255
               Index           =   53
               Left            =   3720
               TabIndex        =   121
               Top             =   180
               Width           =   975
            End
            Begin VB.Label Label65 
               Caption         =   "Underlying Disease"
               Height          =   255
               Index           =   52
               Left            =   120
               TabIndex        =   120
               Top             =   180
               Width           =   1455
            End
         End
         Begin MSMask.MaskEdBox txtFirstCall 
            Height          =   300
            Left            =   7560
            TabIndex        =   21
            Top             =   390
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFirstTreated 
            Height          =   300
            Left            =   7560
            TabIndex        =   23
            Top             =   660
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDateAdmitted 
            Height          =   300
            Left            =   9720
            TabIndex        =   22
            Top             =   390
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDateDischarge 
            Height          =   300
            Left            =   9720
            TabIndex        =   24
            Top             =   660
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtCASLGAmt 
            Alignment       =   1  'Right Justify
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
            Left            =   9480
            TabIndex        =   26
            Top             =   4515
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtCASReserveAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   7320
            TabIndex        =   25
            Top             =   4515
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cboCondwillOcc 
            Height          =   315
            Left            =   7560
            TabIndex        =   27
            Top             =   930
            Width           =   3495
         End
         Begin MSMask.MaskEdBox txtLGTime 
            Height          =   285
            Left            =   9480
            TabIndex        =   28
            Top             =   4800
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboFLWTreatment 
            Height          =   315
            Left            =   7560
            TabIndex        =   29
            Top             =   1215
            Width           =   3495
         End
         Begin MSMask.MaskEdBox txtFollowUpDate 
            Height          =   300
            Left            =   7560
            TabIndex        =   30
            Top             =   1500
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboFolUp 
            Height          =   315
            Left            =   9720
            TabIndex        =   31
            Top             =   1500
            Width           =   1335
         End
         Begin VB.TextBox txtMBMPolYr 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7560
            TabIndex        =   32
            Top             =   1765
            Width           =   1215
         End
         Begin VB.TextBox txtPatientAge 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9720
            TabIndex        =   33
            Top             =   1785
            Width           =   500
         End
         Begin VB.ComboBox cboYearDay 
            Height          =   315
            Left            =   10200
            TabIndex        =   34
            Top             =   1785
            Width           =   855
         End
         Begin VB.TextBox txtHOSName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1400
            MaxLength       =   100
            TabIndex        =   13
            Top             =   690
            Width           =   4740
         End
         Begin VB.TextBox txtReason 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   530
            Left            =   1400
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   950
            Width           =   4740
         End
         Begin VB.TextBox txtREFBy 
            Height          =   285
            Left            =   1400
            MaxLength       =   100
            TabIndex        =   15
            Top             =   1440
            Width           =   4740
         End
         Begin VB.ComboBox cboNature 
            Height          =   315
            Left            =   1400
            TabIndex        =   16
            Top             =   1695
            Width           =   2000
         End
         Begin VB.ComboBox cboStayType 
            Height          =   315
            Left            =   4410
            TabIndex        =   18
            Top             =   1695
            Width           =   1755
         End
         Begin VB.ComboBox cboModality 
            Height          =   315
            Left            =   1400
            TabIndex        =   17
            Top             =   1965
            Width           =   2000
         End
         Begin VB.ComboBox CboPayType 
            Height          =   315
            Left            =   4410
            TabIndex        =   19
            Top             =   1965
            Width           =   1755
         End
         Begin VB.ComboBox cboPendingReason 
            Height          =   315
            Left            =   7560
            TabIndex        =   313
            Top             =   2040
            Width           =   3495
         End
         Begin VB.ComboBox cboOPChemo 
            Height          =   315
            ItemData        =   "frmCaseAdjustmentNew.frx":017C
            Left            =   1400
            List            =   "frmCaseAdjustmentNew.frx":017E
            TabIndex        =   315
            Top             =   2250
            Width           =   2000
         End
         Begin VB.TextBox txtHPNumber 
            Height          =   285
            Left            =   7560
            TabIndex        =   357
            Top             =   2300
            Width           =   3495
         End
         Begin VB.Label Label65 
            Caption         =   "Mobile Number"
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
            Index           =   35
            Left            =   6360
            TabIndex        =   358
            Top             =   2320
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "OP Chemo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   316
            Top             =   2300
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "Pending Reason"
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
            Index           =   42
            Left            =   6360
            TabIndex        =   314
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label65 
            Caption         =   "Hospital Name"
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   307
            Top             =   675
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "Patient Age"
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
            Index           =   14
            Left            =   8880
            TabIndex        =   306
            Top             =   1815
            Width           =   1095
         End
         Begin VB.Label lblFaxDocStatus 
            Caption         =   "Fax documents not in system!"
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
            Left            =   3480
            TabIndex        =   303
            Top             =   2400
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label Label65 
            Caption         =   "Policy Year"
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
            Left            =   6360
            TabIndex        =   302
            Top             =   1780
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "LG Time"
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
            Index           =   68
            Left            =   8760
            TabIndex        =   238
            Top             =   4800
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Fol-Up"
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
            Index           =   48
            Left            =   8880
            TabIndex        =   157
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "LG Amt"
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
            Index           =   32
            Left            =   8760
            TabIndex        =   156
            Top             =   4560
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "D.O.D"
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
            Index           =   31
            Left            =   9000
            TabIndex        =   155
            Top             =   680
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "D.O.A"
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
            Index           =   30
            Left            =   9000
            TabIndex        =   154
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Fol-Up Date"
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
            Index           =   29
            Left            =   6360
            TabIndex        =   153
            Top             =   1530
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Fol-Up Req"
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
            Index           =   28
            Left            =   6360
            TabIndex        =   152
            Top             =   1255
            Width           =   855
         End
         Begin VB.Label Label65 
            Caption         =   "Recur"
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
            Index           =   27
            Left            =   6360
            TabIndex        =   151
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Reserve Amt"
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
            Index           =   26
            Left            =   6240
            TabIndex        =   150
            Top             =   4560
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label65 
            Caption         =   "F. Call"
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
            Index           =   25
            Left            =   6360
            TabIndex        =   149
            Top             =   680
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "F. Call"
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
            Left            =   6360
            TabIndex        =   148
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "IC Number"
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
            Left            =   6360
            TabIndex        =   147
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Cash Type"
            Height          =   255
            Index           =   22
            Left            =   3480
            TabIndex        =   146
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label65 
            Caption         =   "Stay Type"
            Height          =   255
            Index           =   21
            Left            =   3480
            TabIndex        =   145
            Top             =   1750
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Treatment Mode"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   144
            Top             =   2000
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "Nat of Clm"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   143
            Top             =   1750
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Refer"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   142
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Reason"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   141
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Hospital"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   140
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Patient"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   139
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.TextBox txtCASRemark 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4980
         Left            =   -74640
         Locked          =   -1  'True
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   116
         Top             =   960
         Width           =   8895
      End
      Begin VB.TextBox txtCASRemarkTwo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4980
         Left            =   -74640
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   115
         Top             =   960
         Width           =   8895
      End
      Begin MSComctlLib.ListView lstBenefitDetails 
         Height          =   5115
         Left            =   -74880
         TabIndex        =   173
         Top             =   360
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   9022
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
            Text            =   "Benefits Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Benefit Description"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "No Of Days"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM per Day"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Disability"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMember 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   174
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cover ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "I/C No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Relation"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOB"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Supp Annual Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Warranty Exc"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstBenefits 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   175
         Top             =   480
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   9551
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
            Text            =   "Code"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bnf Desc"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Days"
            Object.Width           =   952
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM PDay"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Dis"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Frame Frame20 
         Height          =   585
         Index           =   3
         Left            =   -74880
         TabIndex        =   193
         Top             =   1020
         Width           =   11415
         Begin MSMask.MaskEdBox txtDateRECDoc 
            Height          =   300
            Left            =   8520
            TabIndex        =   194
            Top             =   2760
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboPayPending 
            Height          =   315
            Left            =   1320
            TabIndex        =   96
            Text            =   "N - No"
            Top             =   150
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtFirstDocDate 
            Height          =   315
            Left            =   3840
            TabIndex        =   97
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSubDocDate 
            Height          =   315
            Left            =   6720
            TabIndex        =   98
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label65 
            Caption         =   "SubSeq. Doc Date"
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
            Left            =   5280
            TabIndex        =   253
            Top             =   195
            Width           =   1455
         End
         Begin VB.Label Label65 
            Caption         =   "Pymt P'ding App"
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
            Index           =   67
            Left            =   120
            TabIndex        =   196
            Top             =   195
            Width           =   1335
         End
         Begin VB.Label Label65 
            Caption         =   "First Doc Date"
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
            Index           =   59
            Left            =   2760
            TabIndex        =   195
            Top             =   195
            Width           =   1455
         End
      End
      Begin VB.TextBox txtCASRemark3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4980
         Left            =   -74640
         Locked          =   -1  'True
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   240
         Top             =   960
         Width           =   8895
      End
      Begin VB.TextBox txtCASRemark5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4980
         Left            =   -74640
         Locked          =   -1  'True
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   247
         Top             =   960
         Width           =   8895
      End
      Begin VB.TextBox txtCASRemark4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4980
         Left            =   -74640
         Locked          =   -1  'True
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   246
         Top             =   960
         Width           =   8895
      End
      Begin MSComctlLib.ListView lstWorksheetListing 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   374
         Top             =   420
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ClaimNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Actual B4 GST"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "GST"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Actual with GST"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "App Amount"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Pay2HOS"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Pay2MBM"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "BordxNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "BordxDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "BordxAmt"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Payor App Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Payor App Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Payor App Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Revised Bordx"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Cartoon No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Cartoon Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "CR/DN Ref."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "CR/DN Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "RevStatus"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstNote1 
         Height          =   1200
         Index           =   2
         Left            =   120
         TabIndex        =   392
         Top             =   5820
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2117
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstINVListing 
         Height          =   1680
         Left            =   120
         TabIndex        =   393
         Top             =   1980
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   2963
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Claim No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "INV No"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "INV Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pay to"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Payee"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "INV Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Bordx No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Bordx Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "INV Ind"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView lstNote1 
         Height          =   1200
         Index           =   8
         Left            =   6240
         TabIndex        =   394
         Top             =   5820
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2117
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstLGListing 
         Height          =   1680
         Left            =   -74880
         TabIndex        =   400
         Top             =   660
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   2963
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LG No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "LG Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "LG Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Req Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Req Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Purpose"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "GL Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Issued By"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Remark"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Purpose"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstNote1 
         Height          =   1080
         Index           =   3
         Left            =   120
         TabIndex        =   407
         Top             =   660
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   1905
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Claim No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Excess Item"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   1940
         EndProperty
      End
      Begin MSComctlLib.ListView lstNote1 
         Height          =   1320
         Index           =   0
         Left            =   -74880
         TabIndex        =   409
         Top             =   2580
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2328
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Member Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cover ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Policy No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Exp Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Data Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Policy Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstNote1 
         Height          =   1320
         Index           =   1
         Left            =   -69240
         TabIndex        =   410
         Top             =   2580
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2328
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Member No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DOD"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Claimability"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Final Diag"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Case Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Hosp Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstNote1 
         Height          =   1320
         Index           =   7
         Left            =   -74880
         TabIndex        =   413
         Top             =   4140
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   2328
         SortKey         =   1
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Claim No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ref No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ref Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cheque No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Remark"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstNote1 
         Height          =   1605
         Index           =   5
         Left            =   120
         TabIndex        =   415
         Top             =   3900
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   2831
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Claim No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bordx No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "WS Reg Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amnt B4 GST"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "GST"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Amnt wt GST"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "App Amt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Pay2HOS"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "PaytoMBM"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "WS Ind"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Payor App Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Payor App Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Payor App Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Revised Bordx"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Carton No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Carton Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "CR/DN Ref"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "CR/DN Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComDlg.CommonDialog CmnDlg 
         Left            =   -64200
         Top             =   1620
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lstLGListing1 
         Height          =   1065
         Left            =   -68400
         TabIndex        =   389
         Top             =   4440
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1879
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LG No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "LG Date"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "LG Amt"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Cov ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Patient"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "DOA"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Hospital"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Issued By"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Remark"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "User LG Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "System LG Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "LG Purpose"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "GL Listing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   417
         Top             =   420
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Claims"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   416
         Top             =   3660
         Width           =   2025
      End
      Begin VB.Label Label1 
         Caption         =   "Receipt/Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -74880
         TabIndex        =   414
         Top             =   3900
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Membership History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   412
         Top             =   2340
         Width           =   2025
      End
      Begin VB.Label Label1 
         Caption         =   "Case History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -69240
         TabIndex        =   411
         Top             =   2340
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Excess Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   408
         Top             =   420
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Investigation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   404
         Top             =   4680
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Attenting Doctor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   397
         Top             =   5580
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Anaesthesia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6360
         TabIndex        =   396
         Top             =   5580
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Hospital Invoice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   395
         Top             =   1740
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -74880
         TabIndex        =   391
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lblPage5 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64200
         TabIndex        =   245
         Top             =   2100
         Width           =   135
      End
      Begin VB.Label lblPage4 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64440
         TabIndex        =   244
         Top             =   2100
         Width           =   135
      End
      Begin VB.Label lblPage3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64680
         TabIndex        =   243
         Top             =   2100
         Width           =   135
      End
      Begin VB.Label lblPage2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64920
         TabIndex        =   242
         Top             =   2100
         Width           =   135
      End
      Begin VB.Label lblPage1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -65160
         TabIndex        =   241
         Top             =   2100
         Width           =   135
      End
      Begin VB.Label Label65 
         Caption         =   "Supplementary"
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
         Index           =   111
         Left            =   -74880
         TabIndex        =   181
         Top             =   2700
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label65 
         Caption         =   "Audit Review"
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
         Index           =   110
         Left            =   -74760
         TabIndex        =   180
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblPage 
         Caption         =   "Page 1"
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
         Left            =   -64920
         TabIndex        =   179
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Date"
         Height          =   255
         Index           =   96
         Left            =   -74640
         TabIndex        =   178
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label65 
         Caption         =   "Remarks"
         Height          =   255
         Index           =   97
         Left            =   -73560
         TabIndex        =   177
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label65 
         Caption         =   "Person Incharge"
         Height          =   255
         Index           =   98
         Left            =   -67560
         TabIndex        =   176
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame20 
      Height          =   555
      Index           =   11
      Left            =   120
      TabIndex        =   182
      Top             =   8520
      Width           =   11925
      Begin VB.CommandButton Command2 
         Caption         =   "&Print Audit"
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
         Height          =   285
         Left            =   8280
         TabIndex        =   405
         Top             =   160
         Width           =   960
      End
      Begin VB.CommandButton cmdPrintNote 
         Caption         =   "&Mgmt Note"
         Height          =   285
         Left            =   10065
         TabIndex        =   399
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton cmdReimReq 
         Caption         =   "Reim Req"
         Height          =   285
         Left            =   5685
         TabIndex        =   312
         Top             =   165
         Width           =   855
      End
      Begin VB.CommandButton cmdGPQA 
         Caption         =   "GP/QA"
         Height          =   285
         Left            =   4965
         TabIndex        =   305
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdDocReqDis 
         Caption         =   "Doc Disch"
         Height          =   285
         Left            =   4005
         TabIndex        =   304
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton cmdTHDrugs 
         Caption         =   "C. Ref AA"
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
         Left            =   3170
         TabIndex        =   191
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdPreUnd 
         Caption         =   "C. Ref RR"
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
         TabIndex        =   190
         Top             =   160
         Width           =   900
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Prev Mgmt Note"
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
         Left            =   6600
         TabIndex        =   189
         Top             =   160
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrior 
         Caption         =   "&Notify"
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
         TabIndex        =   188
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton cmdRepudiate 
         Caption         =   "Rep&udiate"
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
         Height          =   285
         Left            =   780
         TabIndex        =   186
         Top             =   160
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton cmdLG 
         Caption         =   "&LG"
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
         Height          =   285
         Left            =   120
         TabIndex        =   185
         Top             =   160
         Width           =   675
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
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
         Height          =   285
         Left            =   11040
         TabIndex        =   184
         Top             =   160
         Width           =   600
      End
      Begin VB.CommandButton cmdUpdate1 
         Caption         =   "&Update"
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
         Height          =   285
         Left            =   9360
         TabIndex        =   183
         Top             =   160
         Width           =   720
      End
      Begin VB.CommandButton cmdSaveAdd 
         Caption         =   "Upd &Add"
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
         Height          =   285
         Left            =   960
         TabIndex        =   187
         ToolTipText     =   "Update Address"
         Top             =   160
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame20 
      Height          =   720
      Index           =   9
      Left            =   120
      TabIndex        =   103
      Top             =   680
      Width           =   11895
      Begin VB.TextBox txtMBMIcBcPP 
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
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   9
         Top             =   130
         Width           =   1575
      End
      Begin VB.TextBox txtMBMNumber 
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
         Left            =   1050
         MaxLength       =   18
         TabIndex        =   7
         Top             =   130
         Width           =   2000
      End
      Begin VB.TextBox txtPayCode 
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
         Height          =   285
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtMBMName 
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
         Height          =   285
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   8
         Top             =   360
         Width           =   2000
      End
      Begin MSMask.MaskEdBox txtMBMEffDate 
         Height          =   285
         Left            =   8760
         TabIndex        =   104
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMBMExpiryDate 
         Height          =   285
         Left            =   8760
         TabIndex        =   105
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPolEffDate 
         Height          =   285
         Left            =   6540
         TabIndex        =   310
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label65 
         Caption         =   "Pol Eff Date"
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
         Left            =   5520
         TabIndex        =   311
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label65 
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
         Index           =   47
         Left            =   8040
         TabIndex        =   113
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label65 
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
         Index           =   46
         Left            =   8040
         TabIndex        =   112
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label65 
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
         Height          =   255
         Index           =   43
         Left            =   5520
         TabIndex        =   111
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label65 
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
         Height          =   255
         Index           =   41
         Left            =   3000
         TabIndex        =   110
         Top             =   780
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label65 
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
         Height          =   255
         Index           =   40
         Left            =   3120
         TabIndex        =   109
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label65 
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
         Index           =   39
         Left            =   120
         TabIndex        =   108
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label65 
         Caption         =   "Mem Num"
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
         Index           =   38
         Left            =   120
         TabIndex        =   107
         Top             =   150
         Width           =   975
      End
      Begin VB.Label txtPolStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6540
         TabIndex        =   106
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label65 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Case Adjustment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   93
      Left            =   0
      TabIndex        =   192
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "frmCaseAdjustmentNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intLastPoz2 As Integer
Dim FirstDiag As Integer
Dim FinalDiag As Integer
Dim UD1 As Integer
Dim SaveOthers As Integer
Dim REPReasons As Integer
Dim Letter As String
Dim LGNumber As String
Dim FOLNumber As String
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim EXCNumber As String
Dim MBMNumber As String
Dim TempCASNo As String
Dim RB As String
Dim tmpLGVer As String
Dim patient As String
Dim ReprintLG As Boolean
Dim ReprintBTBG As Boolean
Dim SndLG As Boolean
Dim bTabDone As Boolean
Dim ScanCode As String
Dim HOspitalName As String
Dim PrevDOA As String
Dim PatientName As String
Dim CaseNumber As String
Dim REDAmtInd As Boolean
Dim RedAmt As Double
Dim CaseRegDate As String
Dim txtPlanCode  As String
Dim MBMCPLNCode As String
Dim sqlstrCAS As String
Dim CASSearch As New ADODB.Recordset
Dim Check As Boolean
Dim TotalFinalDiag As String
Dim TotalProc As String
Dim LGNo As String
Dim LGVer As Integer
Dim PreIllDate As String
Dim PreIllCAT As String
Dim PreIllOthers As String
Dim PreIllInfBy As String
Dim UndIllDate As String
Dim UndIllCAT As String
Dim UndIllOthers As String
Dim UndIllInfBy As String
Dim UpdClaimability As Boolean
Public Document, Personnel, Attention, AdmStatus As String
Public Underlying As Boolean
Public CASUnderlying As String
Dim MMASurgFees As Double
Dim MMAAnaeFees As Double
Dim MMASurgFees13th As Double
Dim MMAAnaeFees13th As Double
Dim MMASurgFees13thMin As Double
Dim MMARateCal As Integer
Dim tmpCLMNumber As String
Dim tmpCLMversion As String
Dim tmpPLNPerVisit As String
Dim tmpHltCode As String
Dim tmpINSCode As String
Dim tmpPlnCode As String
Dim tmpRBType As String
Dim tmpMBMTermination As Boolean
Dim tmpMBMGroupCompany As String
Dim tmpLGTime As String
Dim tmpSysTime As String
Dim tmpRepStatus As String
Dim PictureFileName As String
Dim tmpMBMPymtMode As String
Private WithEvents cFileCopy As clsFileCopy
Attribute cFileCopy.VB_VarHelpID = -1
Dim tmpRemarks As String
Dim tmpRemarks2 As String
Dim tmpPolHistory As String
Dim tmpCaseHistory As String
Dim tmpFinalDiagnosis As String
Dim tmpDiagnosis As String
Dim tmpPatientDOB As String
Dim LGRemark As String
Dim tmpProductName As String
Dim txtAdd1, txtAdd2, txtAdd3 As String
Dim txtPostCode, txtCity, txtState As String
Dim tmpChkAnnBalLmt As Double
Dim tmpChkLTBalLmt As Double
Dim LGStatus As Boolean
Dim tmpPolicyWording As String
Dim tmpPrincipalDOB As String
Dim tmpPAYLGDesc As String
Dim tmptxtMBMPolNo As String
Dim tmpPLNLGPrivate As String
Dim tmpPLNLGGov As String
Dim tmpRel As String
Dim tmpMBMEmpNo As String
Dim HOSType As String
Dim tmpCashType As String
Dim tmpRMAmt As String
Dim tmpICUAmt As String
Dim ICU As String
Dim tmpFolDays As String
Dim HOSArr As Collection
Dim tmpHospChk As String
Dim tmpAgentCode As String
Dim term5 As String
Dim term6 As String
Dim term7 As String
Dim term8 As String
Dim tmpMBMGLIssuance As String
Public CASGLPurpose As String
Public CASGLPurposeRcd As Boolean
Dim strAppend As String
Dim tmpDeleteGL As String
Dim tmpDelGLWeb As String
Dim tmpETQIndGL As String
Dim strsqlCASLG As String
Dim tmpStatusTime As String
Dim tmpStatusDate As String
Dim tmpMBMTelNoM As String
Dim tmpCASLGSMS As String
Dim tmpCASPendingSMS As String
Dim tmpCASRejectSMS As String
Dim tmpCASCancelSMS As String
Dim tmpCasPendingMail As String
Dim tmpCasRejectMail As String
Dim tmpInitialLG As String
Dim tmpFinalLG As String
Dim tmpMsgprom As String
Dim tmpMBMEmail As String
Dim TmpExclusion As String
Dim tmpMBMStatus As String
Dim tmpMBMAppName As String
Dim tmpCASDocDate As String
Dim tmpCASDocType As String
Dim tmpCaseOpenBy As String
Dim RepType As String
Dim IsEtiqa As Boolean
Dim sFilePath As String

Public tmpAgent As String
Public repremarks As String


Private Type PointAPI
    X  As Long
    Y  As Long
End Type
Dim tmpRejectCat As String

Private mbActive                As Boolean
Private mlCurThumb              As Long
Private Const SRCCOPY           As Long = &HCC0020
Private Const STRETCH_HALFTONE  As Long = &H4&
Private Const SW_RESTORE        As Long = &H9&

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As PointAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecute1 Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const WSURL = "http://www.etiqahc.com.my/AndroidService.asmx?wsdl"




Private Sub cboAppStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboAppStatus, cboAppStatus.SelStart) + Chr(KeyAscii) + Mid(cboAppStatus, cboAppStatus.SelStart + cboAppStatus.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboAppStatus.ListCount - 1
        
        If NewText = LCase(Left(cboAppStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboAppStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboAppStatus = ValidValue
               cboAppStatus.SelStart = 0
               cboAppStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub


Private Sub cboBordxRevStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboBordxRevStatus, cboBordxRevStatus.SelStart) + Chr(KeyAscii) + Mid(cboBordxRevStatus, cboBordxRevStatus.SelStart + cboBordxRevStatus.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboBordxRevStatus.ListCount - 1
        
        If NewText = LCase(Left(cboBordxRevStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboBordxRevStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboBordxRevStatus = ValidValue
               cboBordxRevStatus.SelStart = 0
               cboBordxRevStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCASTakeOveInd_Click()
    If Mid(cboCASTakeOveInd, 1, 1) = "Y" Then
        txtCASTakeOverAmt.Enabled = True
    Else
        txtCASTakeOverAmt.Enabled = False
    End If

End Sub

Private Sub cboCASTakeOveInd_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCASTakeOveInd, cboCASTakeOveInd.SelStart) + Chr(KeyAscii) + Mid(cboCASTakeOveInd, cboCASTakeOveInd.SelStart + cboCASTakeOveInd.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboCASTakeOveInd.ListCount - 1
        
        If NewText = LCase(Left(cboCASTakeOveInd.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCASTakeOveInd.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboCASTakeOveInd = ValidValue
               cboCASTakeOveInd.SelStart = 0
               cboCASTakeOveInd.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCASTakeOveInd_LostFocus()
    If Mid(cboCASTakeOveInd, 1, 1) = "Y" Then
        txtCASTakeOverAmt.Enabled = True
    Else
        txtCASTakeOverAmt.Enabled = False
    End If
End Sub
Private Sub cboClaimability_LostFocus()
If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
      Call Command1_Click
End If
End Sub

Private Sub cboIntExtTO_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboIntExtTO, cboIntExtTO.SelStart) + Chr(KeyAscii) + Mid(cboIntExtTO, cboIntExtTO.SelStart + cboIntExtTO.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboIntExtTO.ListCount - 1
        
        If NewText = LCase(Left(cboIntExtTO.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboIntExtTO.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboIntExtTO = ValidValue
               cboIntExtTO.SelStart = 0
               cboIntExtTO.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboPatientList_GotFocus()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim strsql As String
Dim strsqlCOV As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String

   If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
    
        If Mid(cboPatientList, 1, 2) = "00" Then
            'strsql = "Select MBMICBcPp from MBMCrossReference where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppend = ""
            strAppendCase = "1"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd.Execute
            
                If MBM.EOF = True Then
                    MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    txtPatientIC = MBM!MBMIcBcPp
                    tmpRel = "EMPLOYEE"
                    tmpMBMStatus = IIf(Len(MBM!MBMStatus), MBM!MBMStatus, "")
                End If
            MBM.Close
            Set MBM = Nothing
        ElseIf Mid(cboPatientList, 1, 2) <> "00" Then
            'strsqlCOV = "Select MBMCICBCPPNo,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            'MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppend = ""

            strAppendCase = "1"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Adjustment"
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
            
                If MBMCov.EOF = False Then
                    txtPatientIC = MBMCov!MBMCICBCPPNo
                    cboSuppExc = cboPatientList
                    If Len(MBMCov!MBMCPLNCode) Then
                        MBMCPLNCode = MBMCov!MBMCPLNCode
                    End If
                    
                    If Trim(MBMCov!MBMCRELCode) = "W" Then
                        tmpRel = "Wife"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "H" Then
                        tmpRel = "Husband"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "D" Then
                        tmpRel = "Daughter"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "S" Then
                        tmpRel = "Son"
                    End If
                    tmpMBMStatus = IIf(Len(MBMCov!MBMCStatus), MBMCov!MBMCStatus, "")
                End If
            MBMCov.Close
            Set MBMCov = Nothing
        End If
    End If
    
    If Trim(tmpMBMStatus) = "A" Then
        cmdLG.Enabled = True
    Else
        cmdLG.Enabled = False
    End If
    
End Sub

Private Sub cboPayPending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayPending, cboPayPending.SelStart) + Chr(KeyAscii) + Mid(cboPayPending, cboPayPending.SelStart + cboPayPending.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboPayPending.ListCount - 1
        
        If NewText = LCase(Left(cboPayPending.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayPending.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboPayPending = ValidValue
               cboPayPending.SelStart = 0
               cboPayPending.SelLength = Len(ValidValue)
            End If
        End If

End Sub


Private Sub cboPendingReason_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPendingReason, cboPendingReason.SelStart) + Chr(KeyAscii) + Mid(cboPendingReason, cboPendingReason.SelStart + cboPendingReason.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboPendingReason.ListCount - 1
        
        If NewText = LCase(Left(cboPendingReason.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPendingReason.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboPendingReason = ValidValue
               cboPendingReason.SelStart = 0
               cboPendingReason.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboREVBordxInd_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboREVBordxInd, cboREVBordxInd.SelStart) + Chr(KeyAscii) + Mid(cboREVBordxInd, cboREVBordxInd.SelStart + cboREVBordxInd.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboREVBordxInd.ListCount - 1
        
        If NewText = LCase(Left(cboREVBordxInd.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboREVBordxInd.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboREVBordxInd = ValidValue
               cboREVBordxInd.SelStart = 0
               cboREVBordxInd.SelLength = Len(ValidValue)
            End If
        End If
End Sub




Private Sub cboYearDay_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboYearDay, cboYearDay.SelStart) + Chr(KeyAscii) + Mid(cboYearDay, cboYearDay.SelStart + cboYearDay.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboYearDay.ListCount - 1
        
        If NewText = LCase(Left(cboYearDay.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboYearDay.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboYearDay = ValidValue
               cboYearDay.SelStart = 0
               cboYearDay.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub CmdBrowse_Click()
Dim cnt As Integer
Dim tmpString As String
Dim X As Boolean

    X = False
    X = CreatePath(DocPath & "ScannedDoc\" & txtCaseNo)

'Set filter to all files
    tmpString = ""
    cnt = 0
    CmnDlg.filename = ""
    CmnDlg.CancelError = True
      On Error GoTo ErrHandler
      
      CmnDlg.InitDir = DocPath & "FaxDocuments"
      
      CmnDlg.Flags = cdlOFNHideReadOnly
      CmnDlg.Filter = "All files|*.*"
    'Show the open file window
    CmnDlg.ShowOpen
    'Put the file bath in the source text
    TxtSrc.Text = CmnDlg.filename
    
    Dim sArr() As String
    Dim fname As String
    Dim lenth As Integer
    sArr = Split(TxtSrc, "\")
    lenth = UBound(sArr) + 1
    fname = sArr(lenth - 1)
    tmpString = DocPath & "FaxDocuments\"
    cnt = Len(tmpString)
    TxtDest.Text = DocPath & "ScannedDoc\" & txtCaseNo & "\" & fname & ""
ErrHandler:
      Exit Sub


End Sub

Private Sub cmdClearDoc_Click()
Dim strsqlCASLG As String
Dim UPRS As New ADODB.Recordset
Dim RecordSaved

    RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then
    
        strsqlCASLG = ""
        strsqlCASLG = " update tblRecord set " & _
                 " IsLGissued = '1', IsWSdone = '1'" & _
                 " Where CaseNo = '" & Trim(txtCaseNo) & "'"
        UPRS.Open strsqlCASLG, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        
        MsgBox "Record Updated Successful..."
    End If

End Sub




Private Sub CmdCopy_Click()
On Error GoTo CopyErr

Dim ErrCode As Long
Dim ErrMsg As String
Dim tmpString As String
Dim cnt As Integer

   ' Screen.MousePointer = vbHourglass
    tmpString = ""
    cnt = 0
    tmpString = DocPath & "FaxDocuments\"
    cnt = Len(tmpString)

    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    ElseIf UCase(Mid(TxtSrc.Text, 1, cnt)) <> UCase(tmpString) Then
        MsgBox "Please check source file!", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else

    Set cFileCopy = New clsFileCopy
    
    'Disable the copy button
    CmdCopy.Enabled = False
    
    'The source file the you want to copy
    cFileCopy.SourceFile = TxtSrc
    'The destination file name
    cFileCopy.TargetFile = TxtDest
    
    cFileCopy.ByteSize = 4096 '4kb
    
    'Check for verify
    'cFileCopy.Verify = (chkVerify.Value = vbChecked)
        
    'Show percentage
    'LblPercent.Caption = "0"
    
    If Not cFileCopy.copyfile(ErrCode, ErrMsg) Then
        GoTo CopyErr
    End If
    
    'Hide progress bar
    'ProgressBar1.Visible = False
    
    'Skip the error message and exit sub
    'Call DestroyFile(TxtSrc)
    Call LoadDoc
    
    MsgBox "File Transfered Successful..."
    'Screen.MousePointer = Default

        GoTo Ex
    End If
CopyErr:
    If ErrCode = 0 Then
        ErrCode = Err.Number
        ErrMsg = Err.Description
    End If
    
    MsgBox "Error: " & ErrCode & " - " & ErrMsg, vbCritical, "Error"
    
Ex:
    'Close files
    'Close #nSF
    'Close #nDF
    
    Set cFileCopy = Nothing
    
    'Re-enable the copy button
    CmdCopy.Enabled = True
End Sub

Sub DestroyFile(sFileName As String)
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, Offset As Long
    'Create two buffers with a specified 'wi
    '     pe-out' characters
    'Const BLOCKSIZE = 4096
    'Block1 = String(BLOCKSIZE, "X")
    'Block2 = String(BLOCKSIZE, " ")
    'Overwrite the file contents with the wi
    '     pe-out characters
    'hFileHandle = FreeFile
    'Open sFileName For Binary As hFileHandle
    'Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1


    'For iLoop = 1 To Blocks
    '    offset = Seek(hFileHandle)
    '    Put hFileHandle, , Block1
    '    Put hFileHandle, offset, Block2
    'Next iLoop
    'Close hFileHandle
    'Now you can delete the file, which cont
    '     ains no sensitive data
    Kill sFileName
   ' Kill sFileName2
End Sub

Sub DestroyFileGL(sFileName As String, sFileName2 As String)
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, Offset As Long
    'Create two buffers with a specified 'wi
    '     pe-out' characters
    'Const BLOCKSIZE = 4096
    'Block1 = String(BLOCKSIZE, "X")
    'Block2 = String(BLOCKSIZE, " ")
    'Overwrite the file contents with the wi
    '     pe-out characters
    'hFileHandle = FreeFile
    'Open sFileName For Binary As hFileHandle
    'Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1


    'For iLoop = 1 To Blocks
    '    offset = Seek(hFileHandle)
    '    Put hFileHandle, , Block1
    '    Put hFileHandle, offset, Block2
    'Next iLoop
    'Close hFileHandle
    'Now you can delete the file, which cont
    '     ains no sensitive data
    Kill sFileName
    Kill sFileName2
End Sub

Private Sub LoadDoc()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
Call addTheFiles



'f = FreeFile
'Open App.Path & "\temp.html" For Output As #f
'    Print #f, "<html><body><div align=center><table border=0><tr><td valign=center><center><a href=" & Chr(34) & "Please Click File to view" & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">visit xeek.net</a></td></tr></table></body>"
'Close #f
'WebBrowser1.Navigate App.Path & "\temp.html"
Exit Sub
End Sub

Private Sub LoadWebDoc()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
Call addTheFilesWeb



'f = FreeFile
'Open App.Path & "\temp.html" For Output As #f
'    Print #f, "<html><body><div align=center><table border=0><tr><td valign=center><center><a href=" & Chr(34) & "Please Click File to view" & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">visit xeek.net</a></td></tr></table></body>"
'Close #f
'WebBrowser1.Navigate App.Path & "\temp.html"
Exit Sub
End Sub

Sub addTheFiles()
Dim B As String
List1.Clear

PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & ""
'\Medical Rpt"
B = Dir(PictureFileName & "\*.*")
'b = Dir(Dir1.Path & "\*.*")
If B <> "" Then
    If Right(B, 4) = ".jpg" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".tif" Then
        List1.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".gif" Then
        List1.AddItem B
    End If
    
    If Right(B, 4) = ".JPG" Then
        List1.AddItem B
    End If
    If Right(B, 4) = ".TIF" Then
        List1.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".GIF" Then
        List1.AddItem B
    End If
    
    If UCase(Right(B, 4)) = ".PDF" Then
        List1.AddItem B
    End If

    If UCase(Right(B, 4)) = ".PNG" Then
        List1.AddItem B
    End If

Else
    Exit Sub
End If
Do: DoEvents
    B = Dir
    If B <> "" Then
        If Right(B, 4) = ".jpg" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".tif" Then
            List1.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".gif" Then
            List1.AddItem B
        End If
        
        If Right(B, 4) = ".JPG" Then
            List1.AddItem B
        End If
        If Right(B, 4) = ".TIF" Then
            List1.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".GIF" Then
            List1.AddItem B
        End If
        
        If UCase(Right(B, 4)) = ".PDF" Then
            List1.AddItem B
        End If
        
        If UCase(Right(B, 4)) = ".PNG" Then
            List1.AddItem B
        End If
        
    Else
        Exit Do
    End If
Loop

'List1.Height = Me.Height - 2520 - 350
End Sub

Sub addTheFilesWeb()
Dim B As String
List3.Clear

PictureFileName = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & ""
'\Medical Rpt"
B = Dir(PictureFileName & "\*.*")
'b = Dir(Dir1.Path & "\*.*")
If B <> "" Then
    If Right(B, 4) = ".jpg" Then
        List3.AddItem B
    End If
    If Right(B, 4) = ".tif" Then
        List3.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".gif" Then
        List3.AddItem B
    End If
    
    If Right(B, 4) = ".JPG" Then
        List3.AddItem B
    End If
    If Right(B, 4) = ".TIF" Then
        List3.AddItem B
    End If
    If Right(B, 4) = ".GIF" Then
        List3.AddItem B
    End If
    If UCase(Right(B, 4)) = ".PDF" Then
        List3.AddItem B
    End If

Else
    Exit Sub
End If
Do: DoEvents
    B = Dir
    If B <> "" Then
        If Right(B, 4) = ".jpg" Then
            List3.AddItem B
        End If
        If Right(B, 4) = ".tif" Then
            List3.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".gif" Then
            List3.AddItem B
        End If
        
        If Right(B, 4) = ".JPG" Then
            List3.AddItem B
        End If
        If Right(B, 4) = ".TIF" Then
            List3.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".GIF" Then
            List3.AddItem B
        End If
        
    If UCase(Right(B, 4)) = ".PDF" Then
        List3.AddItem B
    End If
    
    Else
        Exit Do
    End If
Loop

'List1.Height = Me.Height - 2520 - 350
End Sub

Private Sub cmdDelGL_Click()
On Error GoTo CopyErr

Dim ErrCode As Long
Dim ErrMsg As String
Dim tmpString As String
Dim cnt As Integer
Dim RecordSaved
   ' Screen.MousePointer = vbHourglass
'    tmpString = ""
'    cnt = 0
'    cnt = Len(tmpString)

    'If txtCaseNo = "" Then
    '    MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    '    Exit Sub
    'ElseIf UCase(Mid(TxtSrc.Text, 1, cnt)) <> UCase(tmpString) Then
    '    MsgBox "Please check source file!", vbOKOnly + vbCritical, DialogBoxTitle
    '    Exit Sub
    'Else

    'Set cFileCopy = New clsFileCopy
    
    'Disable the copy button
    'CmdCopy.Enabled = False
    
    'The source file the you want to copy
   ' cFileCopy.SourceFile = TxtSrc
   ' 'The destination file name
   ' cFileCopy.TargetFile = TxtDest
   '
   ' cFileCopy.ByteSize = 4096 '4kb
    
    'Check for verify
    'cFileCopy.Verify = (chkVerify.Value = vbChecked)
        
    'Show percentage
    'LblPercent.Caption = "0"
    
    'If Not cFileCopy.copyfile(ErrCode, ErrMsg) Then
    '    GoTo CopyErr
    'End If
    
    'Hide progress bar
    'ProgressBar1.Visible = False
    
    'Skip the error message and exit sub
    
    RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then

        Call DestroyFileGL(tmpDeleteGL, tmpDelGLWeb)
        
        'Call LoadDoc
        cmdShow_Click
        MsgBox "File Deleted Successful..."
    End If
    'Screen.MousePointer = Default

        GoTo Ex
  '  End If
CopyErr:
    If ErrCode = 0 Then
        ErrCode = Err.Number
        ErrMsg = Err.Description
    End If
    
    MsgBox "Error: " & ErrCode & " - " & ErrMsg, vbCritical, "Error"
    
Ex:
    
    Set cFileCopy = Nothing

End Sub

Private Sub CmdDiagFinal_Click(a As Integer)
    'If a = "3" Then
    '    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
    '        FrmSearchRepudiation.Source = 1
    '        FrmSearchRepudiation.show vbModal
    '    End If
        'Else
        '    frmSearchICD.Options = 0
        '    frmSearchICD.optRep.Visible = False
        '    frmSearchICD.Source = 1
        '    frmSearchICD.show vbModal
        'End If
    'Else
        'If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Then
        '    FrmSearchRepudiation.Source = 1
        '    FrmSearchRepudiation.show vbModal
        'Else
            frmSearchICD.Options = 2
            frmSearchICD.optRep.Visible = False
            frmSearchICD.Source = 1
            frmSearchICD.Show vbModal
        'End If
    'End If
End Sub

Private Sub CmdDiagPre_Click(a As Integer)
            frmSearchICD.Options = 1
            frmSearchICD.optRep.Visible = False
            frmSearchICD.Source = 1
            frmSearchICD.Show vbModal
End Sub


Private Sub CmdDiagRepudiation_Click(a As Integer)
    If a = "3" Then
        If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Or Mid(cboClaimability, 1, 2) = "PP" Then
            FrmSearchRepudiation.Source = 1
            FrmSearchRepudiation.Paycoce = txtPayCode
            FrmSearchRepudiation.Show vbModal
        End If
    End If
End Sub

Private Sub CmdDiagUD_Click(a As Integer)

    'If a = "3" Then
    '    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
    '        FrmSearchRepudiation.Source = 1
    '        FrmSearchRepudiation.show vbModal
    '    End If
        'Else
        '    frmSearchICD.Options = 0
        '    frmSearchICD.optRep.Visible = False
        '    frmSearchICD.Source = 1
        '    frmSearchICD.show vbModal
        'End If
    'Else
        'If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Then
        '    FrmSearchRepudiation.Source = 1
        '    FrmSearchRepudiation.show vbModal
        'Else
            frmSearchICD.Options = 3
            frmSearchICD.optRep.Visible = False
            frmSearchICD.Source = 1
            frmSearchICD.Show vbModal
        'End If
    'End If
End Sub

Private Sub cboFolUp_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboFolUp, cboFolUp.SelStart) + Chr(KeyAscii) + Mid(cboFolUp, cboFolUp.SelStart + cboFolUp.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboFolUp.ListCount - 1
        
        If NewText = LCase(Left(cboFolUp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboFolUp.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboFolUp = ValidValue
               cboFolUp.SelStart = 0
               cboFolUp.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub CboPayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayType, CboPayType.SelStart) + Chr(KeyAscii) + Mid(CboPayType, CboPayType.SelStart + CboPayType.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboPayType.ListCount - 1
        
        If NewText = LCase(Left(CboPayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboPayType = ValidValue
               CboPayType.SelStart = 0
               CboPayType.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboPending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPending, cboPending.SelStart) + Chr(KeyAscii) + Mid(cboPending, cboPending.SelStart + cboPending.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboPending.ListCount - 1
        
        If NewText = LCase(Left(cboPending.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPending.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboPending = ValidValue
               cboPending.SelStart = 0
               cboPending.SelLength = Len(ValidValue)
            End If
        End If
End Sub


Private Sub CboStayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboStayType, cboStayType.SelStart) + Chr(KeyAscii) + Mid(cboStayType, cboStayType.SelStart + cboStayType.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboStayType.ListCount - 1
        
        If NewText = LCase(Left(cboStayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboStayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboStayType = ValidValue
               cboStayType.SelStart = 0
               cboStayType.SelLength = Len(ValidValue)
            End If
        End If
End Sub





Private Sub cmdDocReqAdm_Click()
Dim str As String
Dim GLDel As New ADODB.Recordset
Dim RecordSaved

    RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then

        str = "Delete from CASLG Where CaseNo =  '" & Mid(lstLGListing.SelectedItem.Text, 2, 10) & "' and LGNo =  '" & Mid(lstLGListing.SelectedItem.Text, 13, 1) & "'"
        GLDel.Open str, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        'Call LoadDoc
        cmdShow_Click
        MsgBox "GL Deleted Successful..."
    End If
End Sub

Private Sub cmdDocReqDis_Click()
Dim RecordSaved
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If Trim(txtMBMNumber) = "" Then
        MsgBox "Please Enter Membership Information", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
    ElseIf Trim(cboPatientList) = "" Then
        MsgBox "Please Enter Patient Information", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
    ElseIf Trim(txtHOSCode) = "" Then
        MsgBox "Please Enter Hospital Information", vbOKOnly + vbCritical, DialogBoxTitle
        txtHOSCode.SetFocus
    ElseIf txtDateAdmitted = "__/__/____" Then
        MsgBox "Please Enter Admission Date", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
    ElseIf cboPatientList = "" Then
        MsgBox "Please Enter Patient's ID", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
    
            Letter = "D"
            Call funcword
        
            Dim UPRS As New ADODB.Recordset
            tmpStatusTime = Format$(Now, "dd/mm/yyyy hh:mm")
            'tmpStatusDate = (Date, "dd/mm/yyyy")
            strsqlCASLG = ""
            strsqlCASLG = " update tblRecord set " & _
                     " IsRejected = '0', IsLGissued = '0', CaseProgress = 'Pending document from hospital'" & _
                     " Where CaseNo = '" & Trim(txtCaseNo) & "'"
            UPRS.Open strsqlCASLG, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        End If
    End If
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End Sub




Private Sub cmdGPQA_Click()
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

        Letter = "I"
        Call funcword
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub cmdInvList_Click()
    frmSearchInv.Source = 1
    frmSearchInv.Show vbModal
End Sub

Private Sub cmdLoadDoc_Click()
    Call bindDocs
End Sub

Private Sub cmdOPCS_Click(Index As Integer)
    If Index = 1 Then
        frmSearchCPT.Source = 2
    Else
        frmSearchCPT.Source = 3
    End If
        frmSearchCPT.Show vbModal
End Sub
Private Sub cmdPreUnd_Click()
    Dim UPRS As New ADODB.Recordset
    Dim RecordSaved

    RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    
    If RecordSaved = vbYes Then
        
        tmpStatusTime = Format$(Now, "dd/mm/yyyy hh:mm")
        strsqlCASLG = ""
        strsqlCASLG = " update tblRecord set " & _
                 " IsRejected = '1', IsLGissued = '0', CaseProgress = 'Case declined for referral doctor'" & _
                 " Where CaseNo = '" & Trim(txtCaseNo) & "'"
        UPRS.Open strsqlCASLG, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        
        MsgBox "Cross Referral Doctor has been rejected", vbOKOnly + vbInformation, DialogBoxTitle
    End If
End Sub

Private Sub cmdPrint_Click()
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If txtCaseNo <> "" Then
        Letter = "M"
        Call FuncWordRemark
    End If
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
End Sub

Private Sub cmdPrintNote_Click()
    Screen.MousePointer = vbHourglass
    If txtCaseNo <> "" Then
        Call funcwordMgmt
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = Default
End Sub

Private Sub cmdPrior_Click()
Dim Days As String
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If IsDate(txtMBMExpiryDate) Then
        'Days = CDate(txtMBMExpiryDate) - Date
        'If Days <= 45 And Days >= 0 Then
            Letter = "P"
            Call FuncPriorLetter
        'End If
        If CDate(txtMBMExpiryDate) < CDate(CaseRegDate) Then
            Letter = "E"
            Call FuncPriorLetter
        End If
    End If
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    cmdShow.Enabled = True
    'cmdLG.Enabled = True
    cmdRepudiate.Enabled = True
    cmdSaveAdd.Enabled = True
    cmdPrior.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdUpdate1.Enabled = True
End Sub

Private Sub cmdReimReq_Click()
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

        Letter = "S"
        Call funcword
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing

End Sub

Private Sub cmdTHDrugs_Click()
    Dim UPRS As New ADODB.Recordset
    Dim RecordSaved

    RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    
    If RecordSaved = vbYes Then
        
        tmpStatusTime = Format$(Now, "dd/mm/yyyy hh:mm")
        strsqlCASLG = ""
        strsqlCASLG = " update tblRecord set " & _
                 " IsRejected = '0', IsLGissued = '1', CaseProgress = 'Case approved for referral doctor'" & _
                 " Where CaseNo = '" & Trim(txtCaseNo) & "'"
        UPRS.Open strsqlCASLG, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        
        MsgBox "Cross Referral Doctor has been accepted", vbOKOnly + vbInformation, DialogBoxTitle
    
    End If
    
End Sub

Private Sub Command1_Click()
Dim PAYCode As String
Dim PLNCode As String

Dim PVRec As New ADODB.Recordset
Dim sqlQry As String

Dim PLNCount As Integer
PLNCount = 0


If Len(txtPayCode) Then

    sqlQry = ""
    sqlQry = "Select Count(*) as RecCnt From Tbl_PaycodeDisclouser where PayCode ='" + txtPayCode + "'"
    PVRec.Open sqlQry, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
    Do While Not PVRec.EOF
        PLNCount = CInt(IIf(Len(PVRec!Reccnt), PVRec!Reccnt, 0))
        PVRec.MoveNext
    Loop
    PVRec.Close
    Set PVRec = Nothing
    
    
    If PLNCount > 0 Then
        PAYCode = txtPayCode
        PLNCode = txtPlanCode
        cboREPCode1.Clear
        cboREPCode2.Clear
        cboREPCode3.Clear
        sqlQry = ""
        If PLNCount = 1 Then
            sqlQry = "EXEC  GetDisclouser '" & Trim(PAYCode) & "'"
        Else
            sqlQry = "EXEC  GetDisclouserByPLN '" & Trim(PAYCode) & "','" & Trim(PLNCode) & "'"
        End If
        PVRec.Open sqlQry, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
        Do While Not PVRec.EOF
            cboREPCode1.AddItem (PVRec!DisClouser & "||" & PVRec!REPCatCode)
            cboREPCode2.AddItem (PVRec!DisClouser & "||" & PVRec!REPCatCode)
            cboREPCode3.AddItem (PVRec!DisClouser & "||" & PVRec!REPCatCode)
            PVRec.MoveNext
        Loop
        PVRec.Close
        Set PVRec = Nothing
        If PLNCount = 1 Then
            sqlQry = "EXEC  GetSubDisclouser '" & Trim(PAYCode) & "'"
        Else
            sqlQry = "EXEC  GetSubDisclouserByPLN '" & Trim(PAYCode) & "','" & Trim(PLNCode) & "'"
        End If
        PVRec.Open sqlQry, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
        Do While Not PVRec.EOF
            cboREPCode1.AddItem (PVRec!SubDisclouser & "||" & PVRec!REPCatCode)
            cboREPCode2.AddItem (PVRec!SubDisclouser & "||" & PVRec!REPCatCode)
            cboREPCode3.AddItem (PVRec!SubDisclouser & "||" & PVRec!REPCatCode)
            PVRec.MoveNext
        Loop
        PVRec.Close
        Set PVRec = Nothing
        
        If PLNCount = 1 Then
            sqlQry = "EXEC  GetNonDisclouser '" & Trim(PAYCode) & "'"
        Else
            sqlQry = "EXEC  GetNonDisclouserByPLN '" & Trim(PAYCode) & "','" & Trim(PLNCode) & "'"
        End If
        PVRec.Open sqlQry, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
        Do While Not PVRec.EOF
            cboREPCode1.AddItem (PVRec!NonDisclouser & "||" & PVRec!REPCatCode)
            cboREPCode2.AddItem (PVRec!NonDisclouser & "||" & PVRec!REPCatCode)
            cboREPCode3.AddItem (PVRec!NonDisclouser & "||" & PVRec!REPCatCode)
            PVRec.MoveNext
        Loop
        PVRec.Close
        Set PVRec = Nothing
    End If

End If
End Sub

Private Sub Command2_Click()
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If txtPrincipalExc <> "" Then
        Letter = "M"
        Call FuncWordAuditreview
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
Private Sub CboClaimability_Change()
    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
        If cboREPCode1 = "" Then
            MsgBox "Please Select Repudiation reason", vbOKOnly + vbCritical, DialogBoxTitle
            cmdRepudiate.Enabled = False
            Call Command1_Click
            SSTab1.Tab = 1
        Else
            cmdRepudiate.Enabled = True
        End If
        cmdLG.Enabled = False
    Else
        cmdRepudiate.Enabled = False
    End If

End Sub

Private Sub cboClaimability_Click()
           
    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
        If cboREPCode1 = "" Then
            MsgBox "Please Select Repudiation reason", vbOKOnly + vbCritical, DialogBoxTitle
            cmdRepudiate.Enabled = False
            'Frame18.Enabled = True
            SSTab1.Tab = 1
        Else
            cmdRepudiate.Enabled = True
            'Frame18.Enabled = True
        End If
        cmdLG.Enabled = False
    Else
        cmdRepudiate.Enabled = False
        'Frame18.Enabled = False
        'cmdLG.Enabled = True

    End If
End Sub

Private Sub CboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboClaimability, cboClaimability.SelStart) + Chr(KeyAscii) + Mid(cboClaimability, cboClaimability.SelStart + cboClaimability.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboClaimability.ListCount - 1
        
        If NewText = LCase(Left(cboClaimability.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboClaimability.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboClaimability = ValidValue
               cboClaimability.SelStart = 0
               cboClaimability.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboCondwillOcc_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCondwillOcc, cboCondwillOcc.SelStart) + Chr(KeyAscii) + Mid(cboCondwillOcc, cboCondwillOcc.SelStart + cboCondwillOcc.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboCondwillOcc.ListCount - 1
        
        If NewText = LCase(Left(cboCondwillOcc.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCondwillOcc.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboCondwillOcc = ValidValue
               cboCondwillOcc.SelStart = 0
               cboCondwillOcc.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboCondwillOcc_LostFocus()
    If Len(cboCondwillOcc) > 7 Then
        MsgBox "Condition Will Occ cannot exceed 7 characters!", vbCritical
        cboCondwillOcc.SetFocus
    End If

End Sub

Private Sub cboFLWTreatment_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboFLWTreatment, cboFLWTreatment.SelStart) + Chr(KeyAscii) + Mid(cboFLWTreatment, cboFLWTreatment.SelStart + cboFLWTreatment.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboFLWTreatment.ListCount - 1
        
        If NewText = LCase(Left(cboFLWTreatment.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboFLWTreatment.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboFLWTreatment = ValidValue
               cboFLWTreatment.SelStart = 0
               cboFLWTreatment.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboFLWTreatment_LostFocus()
    If Len(cboFLWTreatment) > 7 Then
        MsgBox "FLW Treatment cannot exceed 7 characters!", vbCritical
        cboFLWTreatment.SetFocus
    End If

End Sub

Private Sub CboModality_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboModality, cboModality.SelStart) + Chr(KeyAscii) + Mid(cboModality, cboModality.SelStart + cboModality.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboModality.ListCount - 1
        
        If NewText = LCase(Left(cboModality.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboModality.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboModality = ValidValue
               cboModality.SelStart = 0
               cboModality.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboModality_LostFocus()
    If Len(cboModality) > 30 Then
        MsgBox "Modality cannot exceed 30 character!", vbCritical
        cboModality.SetFocus
    End If
End Sub

Private Sub cboNature_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboNature, cboNature.SelStart) + Chr(KeyAscii) + Mid(cboNature, cboNature.SelStart + cboNature.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboNature.ListCount - 1
        
        If NewText = LCase(Left(cboNature.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboNature.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboNature = ValidValue
               cboNature.SelStart = 0
               cboNature.SelLength = Len(ValidValue)
            End If
        End If
    End Sub

Private Sub cboNature_LostFocus()
    If Len(cboNature) > 30 Then
        MsgBox "Claim Nature cannot exceed 30 character!", vbCritical
        cboNature.SetFocus
    End If

End Sub

Private Sub cboPatientList_Change()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim strsql As String
Dim strsqlCOV As String
Dim is_integer As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String

   If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
    
        If Mid(cboPatientList, 1, 2) = "00" Then
            'strsql = "Select MBMICBcPp from MBMCrossReference where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd.Execute
        
                If MBM.EOF = True Then
                    MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    txtPatientIC = MBM!MBMIcBcPp
                    tmpRel = "EMPLOYEE"
                    tmpMBMStatus = IIf(Len(MBM!MBMStatus), MBM!MBMStatus, "")
                End If
            MBM.Close
            Set MBM = Nothing
        ElseIf Mid(cboPatientList, 1, 2) <> "00" Then
            'strsqlCOV = "Select MBMCICBCPPNo,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            'MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppend = ""

            strAppendCase = "1"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Adjustment"
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
            
                If MBMCov.EOF = False Then
                    txtPatientIC = MBMCov!MBMCICBCPPNo
                    cboSuppExc = cboPatientList
                    If Len(MBMCov!MBMCPLNCode) Then
                        MBMCPLNCode = MBMCov!MBMCPLNCode
                    End If
                    If Trim(MBMCov!MBMCRELCode) = "W" Then
                        tmpRel = "Wife"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "H" Then
                        tmpRel = "Husband"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "D" Then
                        tmpRel = "Daughter"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "S" Then
                        tmpRel = "Son"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "C" Then
                        tmpRel = "Child"
                    End If
                    tmpMBMStatus = IIf(Len(MBMCov!MBMCStatus), MBMCov!MBMCStatus, "")
                End If
            MBMCov.Close
            Set MBMCov = Nothing
        End If
    End If
 
    If Trim(tmpMBMStatus) = "A" Then
        cmdLG.Enabled = True
    Else
        cmdLG.Enabled = False
    End If
End Sub

Private Sub cboPatientList_Click()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim strsql As String
Dim strsqlCOV As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String

   If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
    
        If Mid(cboPatientList, 1, 2) = "00" Then
            'strsql = "Select MBMICBcPp from MBMCrossReference where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppend = ""

            strAppendCase = "1"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd.Execute
            
                If MBM.EOF = True Then
                    MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    txtPatientIC = MBM!MBMIcBcPp
                    tmpRel = "EMPLOYEE"
                    tmpMBMStatus = IIf(Len(MBM!MBMStatus), MBM!MBMStatus, "")
                End If
            MBM.Close
            Set MBM = Nothing
        ElseIf Mid(cboPatientList, 1, 2) <> "00" Then
            'strsqlCOV = "Select MBMCICBCPPNo,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            'MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppend = ""

            strAppendCase = "1"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Adjustment"
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
            
                If MBMCov.EOF = False Then
                    txtPatientIC = MBMCov!MBMCICBCPPNo
                    cboSuppExc = cboPatientList
                    If Len(MBMCov!MBMCPLNCode) Then
                        MBMCPLNCode = MBMCov!MBMCPLNCode
                    End If
                    
                    If Trim(MBMCov!MBMCRELCode) = "W" Then
                        tmpRel = "Wife"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "H" Then
                        tmpRel = "Husband"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "D" Then
                        tmpRel = "Daughter"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "S" Then
                        tmpRel = "Son"
                    End If
                    
                    tmpMBMStatus = IIf(Len(MBMCov!MBMCStatus), MBMCov!MBMCStatus, "")
                End If
            MBMCov.Close
            Set MBMCov = Nothing
        End If
    End If
    
    If Trim(tmpMBMStatus) = "A" Then
        cmdLG.Enabled = True
    Else
        cmdLG.Enabled = False
    End If
    
End Sub

Private Sub cboPatientList_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPatientList, cboPatientList.SelStart) + Chr(KeyAscii) + Mid(cboPatientList, cboPatientList.SelStart + cboPatientList.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboPatientList.ListCount - 1
        
        If NewText = LCase(Left(cboPatientList.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPatientList.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboPatientList = ValidValue
               cboPatientList.SelStart = 0
               cboPatientList.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboPatientList_LostFocus()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim strsql As String
Dim strsqlCOV As String
Dim is_integer As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
        
    tmpMBMStatus = ""
   If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
    
        If Mid(cboPatientList, 1, 2) = "00" Then
            'strsql = "Select MBMICBcPp from MBMCrossReference where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppend = ""

            strAppendCase = "1"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd.Execute
            
                If MBM.EOF = True Then
                    MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    txtPatientIC = MBM!MBMIcBcPp
                    tmpRel = "EMPLOYEE"
                    tmpMBMStatus = IIf(Len(MBM!MBMStatus), MBM!MBMStatus, "")
                End If
            MBM.Close
            Set MBM = Nothing
        ElseIf Mid(cboPatientList, 1, 2) <> "00" Then
            'strsqlCOV = "Select MBMCICBCPPNo,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            'MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppend = ""

            strAppendCase = "1"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd.Execute
            
                If MBMCov.EOF = False Then
                    txtPatientIC = MBMCov!MBMCICBCPPNo
                    cboSuppExc = cboPatientList
                    If Len(MBMCov!MBMCPLNCode) Then
                        MBMCPLNCode = MBMCov!MBMCPLNCode
                    End If
                    
                    If Trim(MBMCov!MBMCRELCode) = "W" Then
                        tmpRel = "Wife"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "H" Then
                        tmpRel = "Husband"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "D" Then
                        tmpRel = "Daughter"
                    ElseIf Trim(MBMCov!MBMCRELCode) = "S" Then
                        tmpRel = "Son"
                    End If
                    
                    tmpMBMStatus = IIf(Len(MBMCov!MBMCStatus), MBMCov!MBMCStatus, "")
                    
                End If
            MBMCov.Close
            Set MBMCov = Nothing
        End If
    End If
    
    If Trim(tmpMBMStatus) = "A" Then
        cmdLG.Enabled = True
    Else
        cmdLG.Enabled = False
    End If
    
End Sub

Private Sub cboREPCode1_Change()
    If cboREPCode1 <> "" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    End If
End Sub

Private Sub cboREPCode1_Click()
    If cboREPCode1 <> "" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    End If
End Sub

Private Sub cboREPCode1_GotFocus()
    If cboREPCode1 <> "" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    End If
End Sub

Private Sub cboREPCode1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboREPCode1, cboREPCode1.SelStart) + Chr(KeyAscii) + Mid(cboREPCode1, cboREPCode1.SelStart + cboREPCode1.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboREPCode1.ListCount - 1
        
        If NewText = LCase(Left(cboREPCode1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboREPCode1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboREPCode1 = ValidValue
               cboREPCode1.SelStart = 0
               cboREPCode1.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboREPCode1_LostFocus()
    If cboREPCode1 <> "" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    End If
End Sub

Private Sub cboREPCode2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboREPCode2, cboREPCode2.SelStart) + Chr(KeyAscii) + Mid(cboREPCode2, cboREPCode2.SelStart + cboREPCode2.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboREPCode2.ListCount - 1
        
        If NewText = LCase(Left(cboREPCode2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboREPCode2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboREPCode2 = ValidValue
               cboREPCode2.SelStart = 0
               cboREPCode2.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboREPCode3_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboREPCode3, cboREPCode3.SelStart) + Chr(KeyAscii) + Mid(cboREPCode3, cboREPCode3.SelStart + cboREPCode3.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboREPCode3.ListCount - 1
        
        If NewText = LCase(Left(cboREPCode3.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboREPCode3.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboREPCode3 = ValidValue
               cboREPCode3.SelStart = 0
               cboREPCode3.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboSuppExc_Change()
Dim strsql As String
Dim MBMCov As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim conn As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
    

    If txtMBMNumber <> "" And IsNumeric(Mid(cboSuppExc, 1, 1)) = True Then
        'strsql = "Select MBMCnumber , MBMCCoverID , MBMCExclusion from MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Trim(Mid(cboSuppExc, 1, 2)) & "'"
        'MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockPessimistic
        strAppend = ""

        strAppendCase = "2"
        strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Trim(Mid(cboSuppExc, 1, 2)) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
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
        
        If MBMCov.EOF = False Then
            If Trim(MBMCov!MBMCExclusion) = "" Then
               ' txtSuppExc = MBMCov!MBMCExclusion
            End If
        End If
        MBMCov.Close
        Set MBMCov = Nothing
    End If
End Sub

Private Sub cboSuppExc_Click()
Dim strsql As String
Dim MBMCov As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String

    ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    If txtMBMNumber <> "" And IsNumeric(Mid(cboSuppExc, 1, 1)) Then
        'strsql = "Select MBMCnumber , MBMCCoverID , MBMCExclusion from MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = " & Trim(Mid(cboSuppExc, 1, 2)) & ""
        'MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockPessimistic
        strAppend = ""

        strAppendCase = "2"
        strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = " & Trim(Mid(cboSuppExc, 1, 2)) & ""
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
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
        
        If MBMCov.EOF = False Then
            If Trim(MBMCov!MBMCExclusion) = "" Then
             '   txtSuppExc = MBMCov!MBMCExclusion
            End If
        End If
        MBMCov.Close
        Set MBMCov = Nothing
    End If
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub cboSuppExc_GotFocus()
Dim strsql As String
Dim MBMCov As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String

    ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
     
    If txtMBMNumber <> "" And IsNumeric(Mid(cboSuppExc, 1, 1)) Then
        'strsql = "Select MBMCnumber , MBMCCoverID , MBMCExclusion from MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = " & Trim(Mid(cboSuppExc, 1, 2)) & ""
        'MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockPessimistic
        strAppend = ""

        strAppendCase = "2"
        strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = " & Trim(Mid(cboSuppExc, 1, 2)) & ""
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
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
        
        If MBMCov.EOF = False Then
            If Trim(MBMCov!MBMCExclusion) = "" Then
               ' txtSuppExc = MBMCov!MBMCExclusion
            End If
        End If
        MBMCov.Close
        Set MBMCov = Nothing
    End If
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub CmdExit_Click()
        Unload Me
        Set frmCaseAdjustmentNew = Nothing
End Sub

Public Sub cmdLG_Click()
Dim RecordSaved
Dim CASLG As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
Dim strsqlPLN As String
Dim strsqlBTBG As String
Dim startTrans
Dim ChkVer As String
Dim TempSysDate As String
Dim TempLGAmt As String
Dim strsqlMBM As String
Dim Including As String
Dim Terms As String
Dim COPay As Boolean
Dim MRI As Boolean
Dim cmd2 As New ADODB.Command
Dim MBM As New ADODB.Recordset
Dim strsqlLG As String
Dim CLMLG As New ADODB.Recordset
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim ClaimLGStatus As Boolean
Dim oOApp As Outlook.Application
Dim oOMail As Outlook.MailItem
Dim RecordSave
Dim tmpGLPurpose As String
Dim tmpGLPurposeStatus As Boolean
Dim tmpGLSMSStatus As String
Dim CASSMS As New ADODB.Recordset
Dim strsqlSMS As String
    tmpLGVer = ""
    Dim stringmessage As String
    stringmessage = ""
    frmGLMenu.GLSource = 1
    If frmGLMenu.GLForm = False Then
        ClaimLGStatus = False
            strsqlLG = ""
            CASGLPurpose = ""
            CASGLPurposeRcd = False
            strsqlLG = "Select CASENo,LGPurpose,LGNo  from CASLG " & _
            " where CASENo   ='" & Trim(txtCaseNo) & "'"
            CLMLG.Open strsqlLG, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                Do While Not CLMLG.EOF
                    CASGLPurpose = IIf(Len(CLMLG!LGPurpose), CLMLG!LGPurpose, "")
                    CASGLPurposeRcd = True
                    CLMLG.MoveNext
                Loop
            CLMLG.Close
            Set CLMLG = Nothing
        
            frmGLMenu.txtGLDate = Format(Date, "dd/mm/yyyy")
            frmGLMenu.txtLGTime = Format(Time, "hh:mm")
            frmGLMenu.txtFaxDate = Format(Date, "dd/mm/yyyy")
            frmGLMenu.txtFaxTime = Format(Time, "hh:mm")
            frmGLMenu.txtGLDate.Enabled = False
            frmGLMenu.txtLGTime.Enabled = False
            frmGLMenu.txtFaxDate.Enabled = False
            frmGLMenu.txtFaxTime.Enabled = False
            frmGLMenu.LblCasNUmber = txtCaseNo
            frmGLMenu.chkName.Caption = "Patient name correct? " & "(" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & ")"
            frmGLMenu.chkIC.Caption = "Patient IC correct? " & "(" & txtPatientIC & ")"
            frmGLMenu.chkExpDate.Caption = "Expiry Date corect? " & "(" & txtMBMExpiryDate & ")"
            strsqlLG = ""
            strsqlLG = "Select * from tblRecord where CASENo   ='" & Trim(txtCaseNo) & "'"
            CLMLG.Open strsqlLG, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                
                If CLMLG.EOF = False Then
                    If Len(CLMLG!CaseDate) Then
                        frmGLMenu.txtRequestDate.Enabled = False
                        frmGLMenu.txtRequestTime.Enabled = False
                        frmGLMenu.txtRequestDate = IIf(Len(CLMLG!CaseDate), Format(CLMLG!CaseDate, "dd/mm/yyyy"), "__/__/____")
                        frmGLMenu.txtRequestTime = IIf(Len(CLMLG!CaseTime), Mid(CLMLG!CaseTime, 1, 5), "__:__")
                    Else
                        frmGLMenu.txtRequestDate.Enabled = True
                        frmGLMenu.txtRequestTime.Enabled = True
                    End If
                Else
                    frmGLMenu.txtRequestDate.Enabled = True
                    frmGLMenu.txtRequestTime.Enabled = True
                End If
            
            CLMLG.Close
            Set CLMLG = Nothing
                                                
        If ClaimLGStatus = True Then
            MsgBox "Please Generate Final LG in Worksheet Adjustment", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            frmGLMenu.Show vbModal
        End If
    Else
        If txtPayCode = "AX" Then
            If txtPolEffDate <> txtMBMEffDate Then
                RecordSave = MsgBox("Have you checked Take Over clause?", vbYesNo + vbExclamation)
                If RecordSave = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        tmpChkAnnBalLmt = 0
        tmpChkLTBalLmt = 0
        LGStatus = False
        ClaimLGStatus = False
        frmGLMenu.LblCasNUmber = txtCaseNo
        
        If Trim(txtMBMNumber) = "" Then
            MsgBox "Please Enter Membership Information", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMNumber.SetFocus
        ElseIf Trim(cboPatientList) = "" Then
            MsgBox "Please Enter Patient Information", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 1
        ElseIf Trim(txtHOSCode) = "" Then
            MsgBox "Please Enter Hospital Information", vbOKOnly + vbCritical, DialogBoxTitle
            txtHOSCode.SetFocus
        ElseIf txtDateAdmitted = "__/__/____" Then
            MsgBox "Please Enter Admission Date", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 1
        ElseIf cboPatientList = "" Then
            MsgBox "Please Enter Patient's ID", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call CheckLimit
            If LGStatus = True Then
                RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
            Else
                RecordSaved = MsgBox("Please check LG Amount, do you want to continue?", vbYesNo + vbExclamation)
            End If
                If RecordSaved = vbYes Then
                    strsqlLG = ""
                    ClaimLGStatus = False
                    strsqlLG = "Select CASENo,LGPurpose,LGNo  from CASLG " & _
                    " where CASENo   ='" & Trim(txtCaseNo) & "' and LGPurpose = '" & frmGLMenu.cbPurpose & "'"
                    CLMLG.Open strsqlLG, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                    ClaimLGStatus = False
                        Do While Not CLMLG.EOF
                            ClaimLGStatus = True
                            tmpLGVer = CLMLG!LGNo
                            CLMLG.MoveNext
                        Loop
                    CLMLG.Close
                    Set CLMLG = Nothing
                    strsqlLG = ""
                    strsqlLG = "Select CASENo,LGPurpose,LGNo  from CASLG " & _
                    " where CASENo   ='" & Trim(txtCaseNo) & "' AND  (LGIndicator = 'O' OR LGIndicator IS NULL)"
                    CLMLG.Open strsqlLG, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                    tmpGLPurposeStatus = False
                        Do While Not CLMLG.EOF
                            tmpGLPurposeStatus = True
                            tmpGLPurpose = IIf(Len(CLMLG!LGPurpose), CLMLG!LGPurpose, "")
                            tmpLGVer = CLMLG!LGNo
                            CLMLG.MoveNext
                        Loop
                    CLMLG.Close
                    Set CLMLG = Nothing
                    
                    If frmGLMenu.cbPurpose = "Follow-up" Then
                        tmpGLPurposeStatus = False
                    End If
                    
                    strsqlLG = ""
                    If tmpGLPurposeStatus = True Then
                        If tmpGLPurpose <> frmGLMenu.cbPurpose Then
                            strsqlLG = " update CASLG set " & _
                             " LGIndicator = 'C' " & _
                             " Where CASENo = '" & Trim(txtCaseNo) & "' and LGNo = '" & Trim(tmpLGVer) & "'"
                            medixmasterconnMaint.Execute strsqlLG
                        End If
                    End If
                    
                    Letter = ""
                    Letter = "L"
                    If ClaimLGStatus = False Then
                        LGNumber = GetLGCode(ChkStr(CaseNumber))
                    Else
                        LGNumber = Letter & txtCaseNo & "-" & tmpLGVer
                    End If
                    TempSysDate = Format(GetServerTime, "yyyy/mm/dd")
                    tmpSysTime = Format(Time, "hh:mm")
                    TempLGAmt = CDbl(IIf(IsNumeric(frmGLMenu.txtCASLGAmt) = True, frmGLMenu.txtCASLGAmt, 0))
                        If Trim(txtPayCode) = "ZN" Then
                            strAppendCase = "19"
                            StrAppend1 = " AND PLNCode = '" & Trim(txtPlanCode) & "'"
                        Else
                            strAppendCase = "19"
                            StrAppend1 = " AND PLNCode = '" & Trim(txtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
                        End If
                    Set cmd2.ActiveConnection = medixmasterconn
                    cmd2.CommandText = "Case_Adjustment"
                    cmd2.CommandType = adCmdStoredProc
                    cmd2.CommandTimeout = 300
                    Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                    cmd2.Parameters.Append Prm1
                    Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                    cmd2.Parameters.Append Prm2
                    MBM.CursorLocation = adUseClient
                    MBM.CursorType = adOpenForwardOnly
                    MBM.CacheSize = 100
                    Set MBM = cmd2.Execute
                    If MBM.EOF = False Then
                        If Trim(MBM!PLNPerVisit) = "Y" Then
                            tmpPLNPerVisit = "Y"
                        Else
                            tmpPLNPerVisit = ""
                        End If
                        
                        tmpRBType = IIf(Len(MBM!PLNRBType), MBM!PLNRBType, "")
                        If Len(MBM!PLNLGType) Then
                            Letter = MBM!PLNLGType
                        End If
                        tmpPLNLGPrivate = IIf(Len(MBM!PLNLGPrivate), MBM!PLNLGPrivate, "")
                        tmpPLNLGGov = IIf(Len(MBM!PLNLGGov), MBM!PLNLGGov, "")
                        tmpETQIndGL = IIf(Len(MBM!PLNETQIndGL), MBM!PLNETQIndGL, "")
                    End If
                    MBM.Close
                    Set MBM = Nothing
                    If Trim(tmpPLNLGPrivate) <> "1" Or tmpPLNLGGov <> "1" Then
                        strsqlPLN = "EXEC Case_Adjustment '" & Trim(txtMBMNumber) & "',3"
                        PLN.Open strsqlPLN, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                        If Not PLN.EOF Then
                            Letter = "B"
                            LGNumber = "B" & Mid(LGNumber, 2, Len(LGNumber))
                        End If
                        PLN.Close
                        Set PLN = Nothing
                    End If
                    ChkVer = ""
                    LGRemark = Empty
                    ChkVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
                    If IsNumeric(frmGLMenu.txtCASLGAmt) = False Then
                        frmGLMenu.txtCASLGAmt = 0
                    End If
                    strsqlCASLG = ""
                    If ClaimLGStatus = False Then
                        strsqlCASLG = "EXEC Case_AdjInsertCASLG_2011 '" & ChkStr(CaseNumber) & "', '" & Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-"))) & "', '" & TempSysDate & "', " & frmGLMenu.txtCASLGAmt & ", '" & Logon & "', '" & ChkStr(Trim(frmGLMenu.txtremarks)) & "', '" & Letter & "', '" & tmpPLNPerVisit & "', '" & tmpRBType & "' , '" & tmpLGTime & "', '" & tmpSysTime & "', '" & Format(frmGLMenu.txtRequestDate, "yyyy/mm/dd") & "', '" & frmGLMenu.txtRequestTime & "', '" & frmGLMenu.cboRequestMode & "', '" & frmGLMenu.cboRequestBy & "', '" & IIf(IsDate(frmGLMenu.txtFaxDate), Format(frmGLMenu.txtFaxDate, "yyyy/mm/dd"), "") & "', '" & frmGLMenu.txtFaxTime & "', '" & frmGLMenu.cboDelayreason & "', '" & frmGLMenu.cbPurpose & "'"
                    Else
                        strsqlCASLG = "EXEC Case_AdjUpdateCASLG '" & ChkStr(CaseNumber) & "', '" & Trim(tmpLGVer) & "', '" & TempSysDate & "', " & frmGLMenu.txtCASLGAmt & ", '" & Logon & "', '" & ChkStr(Trim(frmGLMenu.txtremarks)) & "', '" & Letter & "', '" & tmpPLNPerVisit & "', '" & tmpRBType & "' , '" & tmpLGTime & "', '" & tmpSysTime & "', '" & Format(frmGLMenu.txtRequestDate, "yyyy/mm/dd") & "', '" & frmGLMenu.txtRequestTime & "', '" & frmGLMenu.cboRequestMode & "', '" & frmGLMenu.cboRequestBy & "', '" & IIf(IsDate(frmGLMenu.txtFaxDate), Format(frmGLMenu.txtFaxDate, "yyyy/mm/dd"), "") & "', '" & frmGLMenu.txtFaxTime & "', '" & frmGLMenu.cboDelayreason & "', '" & frmGLMenu.cbPurpose & "'"
                    End If
                    medixmasterconnMaint.Execute strsqlCASLG
                Dim UPRS As New ADODB.Recordset
                tmpStatusTime = Format$(Now, "dd/mm/yyyy hh:mm")
                strsqlCASLG = ""
                strsqlCASLG = " update tblRecord set " & _
                         " IsLGissued = '1', CaseProgress = 'Case covered. GL issued'" & _
                         " Where CaseNo = '" & Trim(txtCaseNo) & "'"
                UPRS.Open strsqlCASLG, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
                
                stringmessage = "AXA Affin Life Claims:" & " GL Request : " & frmGLMenu.txtRequestDate & "," & frmGLMenu.txtRequestTime & " Dear " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & " " & tmptxtMBMPolNo & " ,we are pleased to inform that your Guarantee Letter has been approved. We wish you a speedy recovery."
                Set UPRS = Nothing
                    Call funcword
                    Call MobileNotification(txtMBMNumber, LGNumber)
                   If Mid(txtCaseNo, 1, 2) = "AY" And txtDateAdmitted >= "03/08/2015" Then
                        If tmpCASLGSMS <> "1" Then
                            If tmpMBMTelNoM <> "" Then
                                stringmessage = "AXA Affin Life Claims:" & " GL Request : " & frmGLMenu.txtRequestDate & "," & frmGLMenu.txtRequestTime & " Dear " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & " " & tmptxtMBMPolNo & " ,we are pleased to inform that your Guarantee Letter has been approved. We wish you a speedy recovery."
                                Call SendSMS(stringmessage)
                            End If
                            strsqlCASLG = ""
                            strsqlCASLG = " update CAS set  CASLGSMS = '1', CASPendingSMS = '0', CASRejectSMS = '0' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                            UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                            Set UPRS = Nothing
                            tmpCASLGSMS = 1
                        End If
                    End If
                    
                   If Mid(txtCaseNo, 1, 2) = "MX" Then
                        If tmpCASLGSMS <> "1" Then
                            If tmpMBMTelNoM <> "" Then
                                Call SendSMSMAATakaful
                            End If
                            strsqlCASLG = ""
                            strsqlCASLG = " update CAS set  CASLGSMS = '1', CASPendingSMS = '0', CASRejectSMS = '0' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                            UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                            Set UPRS = Nothing
                            tmpCASLGSMS = 1
                        End If
                    End If
                    
                   Dim IC, ICnew As String
                   Dim X As Integer
                   IC = txtPatientIC
                   If Len(IC) = 12 Then
                        ICnew = "******" + "-" + "**" + "-" + Right(IC, 4)
                    Else
                        For X = 0 To Len(IC) - 5
                            ICnew = ICnew + "*"
                        Next
                        ICnew = ICnew + Right(IC, 4)
                   End If
                   If Mid(txtCaseNo, 1, 2) = "ET" Or Mid(txtCaseNo, 1, 2) = "EQ" Or Mid(txtCaseNo, 1, 2) = "ES" Or Mid(txtCaseNo, 1, 2) = "EP" Then
                        If tmpCASLGSMS <> "1" Then
                            tmpMBMTelNoM = frmGLMenu.txtHPNumber
                            If tmpMBMTelNoM <> "" Then
                                Dim Ch As String
                                Ch = CStr(Chr(38))
                                stringmessage = "ETIQA: " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & " " & ICnew & " " & " IGL Approved " & Date & "," & Time & " RB: " & tmpRMAmt & " " & "per day, if exceed, co-payment apply depend on policy TC. TQ"
                                Call SendSMS(stringmessage)
                            End If
                            strsqlCASLG = ""
                            strsqlCASLG = " update CAS set  CASLGSMS = '1', CASPendingSMS = '0', CASRejectSMS = '0' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                            UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                            Set UPRS = Nothing
                            tmpCASLGSMS = 1
                        End If
                    End If
                    
                    'Code New For MAA
                    If Mid(txtCaseNo, 1, 2) = "MX" Then
                        If tmpInitialLG <> "1" Then
                            'Call MaaInitialLGNotification
                            Dim ToAddress As String
                            ToAddress = "Norlaila@maa.my;Farihan@maa.my;Zamri@maa.my;DrKasthureeBai@maa.my "
                            Call MailAdmInitialLG(ToAddress, "", "")
                            strsqlCASLG = ""
                            strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                            UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                            Set UPRS = Nothing
                        End If
                    Else
                        If Mid(txtCaseNo, 1, 2) <> "EM" And Mid(txtCaseNo, 1, 2) <> "ES" And Mid(txtCaseNo, 1, 2) <> "EQ" And Mid(txtCaseNo, 1, 2) <> "ET" And Mid(txtCaseNo, 1, 2) <> "EP" Then
                            If tmpInitialLG <> "1" Then
                            
                            'Dim strsqlPLN As String
                            'Dim PLN As New ADODB.Recordset
                            Dim cnt As Integer
                            Dim Mailtype As String
                            strsqlPLN = ""
                            
                            strsqlPLN = "Select Count(*) as Cnt,Mailtype from Tbl_CISEmailMaster where Grpcompany='" + Trim(tmpMBMGroupCompany) + "' Group by Mailtype"
                            PLN.Open strsqlPLN, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                            If Not PLN.EOF Then
                                cnt = IIf(Len(PLN!cnt), CInt(PLN!cnt), 0)
                                Mailtype = PLN!Mailtype
                            End If
                            PLN.Close
                            Set PLN = Nothing
                            
                            If Mailtype <> "G" And cnt > 1 Then
                            
                                'tmpMBMNum
                                Dim Subsidary As String
                                Dim CostCenter As String
                                Dim Department As String
                                
                                Subsidary = ""
                                CostCenter = ""
                                Department = ""
                                strsqlPLN = ""
                                strsqlPLN = "Select top 1  MBMGroupCompany, Subsidiary , MBMCostCenter , MBMDepartment   from MBMOthers where MBMNumber ='" & CStr(txtMBMNumber) & "'"
                                PLN.Open strsqlPLN, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                                If Not PLN.EOF Then
                                    Subsidary = PLN!Subsidiary
                                    CostCenter = PLN!MBMCostCenter
                                    Department = PLN!MBMDepartment
                                End If
                                PLN.Close
                                Set PLN = Nothing
                        
                                Dim Email As String
                                Dim BCC As String
                                Dim CC As String
                                
                                strsqlPLN = ""
                                If Trim(Mailtype) = Trim("G") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsHIS='Y' and Grpcompany='" + Trim(tmpMBMGroupCompany) + "'"
                                ElseIf Trim(Mailtype) = Trim("S") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsHIS='Y' and Grpcompany='" + Trim(tmpMBMGroupCompany) + "' and  Subsidary='" + Subsidary + "'"
                                ElseIf Trim(Mailtype) = Trim("D") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsHIS='Y' and Grpcompany='" + Trim(tmpMBMGroupCompany) + "' and Department='" + Department + "'"
                                ElseIf Trim(Mailtype) = Trim("C") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsHIS='Y' and Grpcompany='" + Trim(tmpMBMGroupCompany) + "' and Costcenter='" + CostCenter + "'"
                                End If
                                
                                If Len(strsqlPLN) Then
                                
                                    PLN.Open strsqlPLN, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                                    If Not PLN.EOF Then
                                        Email = IIf(Len(PLN!EmailAddress), PLN!EmailAddress, "")
                                        BCC = IIf(Len(PLN!BCC), PLN!BCC, "")
                                        CC = IIf(Len(PLN!CC), PLN!CC, "")
                                    End If
                                    PLN.Close
                                    Set PLN = Nothing
                                    
                                    If Len(Email) Then
                                        Call MailAdmInitialLG(Email, BCC, CC)
                                    End If
                                    
                                    strsqlCASLG = ""
                                    strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                    UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                    Set UPRS = Nothing
                                
                                End If
                                
                            Else
                                strsqlPLN = ""
                                If Trim(Mailtype) = Trim("G") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsHIS='Y' and Grpcompany='" + Trim(tmpMBMGroupCompany) + "'"
                                ElseIf Trim(Mailtype) = Trim("S") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsHIS='Y' and Grpcompany='" + Trim(tmpMBMGroupCompany) + "' and  Subsidary='" + Subsidary + "'"
                                ElseIf Trim(Mailtype) = Trim("D") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsHIS='Y' and Grpcompany='" + Trim(tmpMBMGroupCompany) + "' and Department='" + Department + "'"
                                ElseIf Trim(Mailtype) = Trim("C") Then
                                    strsqlPLN = "Select Grpcompany,Subsidary,Department,Costcenter ,EmailAddress ,BCC,CC    from Tbl_CISEmailMaster  where LGInd= 1 and IsHIS='Y' and Grpcompany='" + Trim(tmpMBMGroupCompany) + "' and Costcenter='" + CostCenter + "'"
                                End If
                                
                                If Len(strsqlPLN) Then
                                
                                    PLN.Open strsqlPLN, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                                    If Not PLN.EOF Then
                                        Email = IIf(Len(PLN!EmailAddress), PLN!EmailAddress, "")
                                        BCC = IIf(Len(PLN!BCC), PLN!BCC, "")
                                        CC = IIf(Len(PLN!CC), PLN!CC, "")
                                    End If
                                    PLN.Close
                                    Set PLN = Nothing
                                    
                                    If Len(Email) Then
                                        Call MailAdmInitialLG(Email, BCC, CC)
                                    End If
                                    
                                        strsqlCASLG = ""
                                        strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                        UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                        Set UPRS = Nothing
                                End If
                            End If
                            
                            'Dim sqlQry As String
                            'Dim PVRec As New ADODB.Recordset
                            'sqlQry = ""
                            'sqlQry = "Select Top 1  GrpName, EmailID , IsInitial from Tbl_EmailMaintainance where GrpName='" + Trim(tmpMBMGroupCompany) + "' and IsInitial= 1"
                            'PVRec.Open sqlQry, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                            'Do While Not PVRec.EOF
                            '    If Len(PVRec!EmailID) Then
                            '        Call MailAdmInitialLG(PVRec!EmailID)
                            '    End If
                            '    PVRec.MoveNext
                            'Loop
                            'PVRec.Close
                            'Set PVRec = Nothing
                            
                           
                            
                            End If
                            
                        ElseIf Trim(tmpMBMGroupCompany) = Trim("MALAYSIAN RESOURCES CORPORATION BERHAD") Then
                            If tmpInitialLG <> "1" Then
                                Call MailAdmInitialLG("peyling@mrcb.com;rozaina@mrcb.com.my", "", "")
                                strsqlCASLG = ""
                                strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                Set UPRS = Nothing
                            End If
                        ElseIf Trim(tmpMBMGroupCompany) = Trim("GOODYEAR MALAYSIA BERHAD") Then
                            If tmpInitialLG <> "1" Then
                                Call MailAdmInitialLG("ponnaiah_palany@goodyear.com;ushaletchumy@bib.com.my", "", "")
                                strsqlCASLG = ""
                                strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                Set UPRS = Nothing
                            End If
                         ElseIf Trim(tmpMBMGroupCompany) = Trim("ECO WORLD DEVELOPMENT GROUP BERHAD AND ALL ITS SUBSIDIARY F.T.R.R & I") Then
                            If tmpInitialLG <> "1" Then
                                Call MailAdmInitialLG("yikfung.choo@ecoworld.my;ushaletchumy@bib.com.my;yhlee@bib.com.my", "", "")
                                strsqlCASLG = ""
                                strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                Set UPRS = Nothing
                            End If
                        ElseIf Trim(tmpMBMGroupCompany) = Trim("LEMBAGA PEMBANGUNAN INDUSTRI PEMBINAAN M") Then
                            If tmpInitialLG <> "1" Then
                                Call MailAdmInitialLG("masliza@cidb.gov..my;rafidah@cidb.gov..my;zaini.arifin@takaful-malaysia.com.my ;aziekasmadi.ag@isfics.com.my;aziekasmadi81@gmail.com", "", "")
                                strsqlCASLG = ""
                                strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                Set UPRS = Nothing
                            End If
                         ElseIf Trim(tmpMBMGroupCompany) = Trim("CANON MARKETING (MALAYSIA) SDN BHD") Then
                            If tmpInitialLG <> "1" Then
                                Call MailAdmInitialLG("Sally_Lim@cmm.canon.com.my;ChoiYee_Woo@cmm.canon.com.my;", "", "")
                                strsqlCASLG = ""
                                strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                Set UPRS = Nothing
                            End If
                         ElseIf Trim(tmpMBMGroupCompany) = Trim("SAINT-GOBAIN MALAYSIA SDN BHD") Then
                            If tmpInitialLG <> "1" Then
                                Call MailAdmInitialLG("Shahrani.MohamedRaffi@saint-gobain.com;Norhidayah.Munawar@saint-gobain.com;MohammadArif.ABDULLAH@saint-gobain.com;Sitiyuniza.ABKADIR@saint-gobain.com;", "", "")
                                strsqlCASLG = ""
                                strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                Set UPRS = Nothing
                            End If
                         ElseIf Trim(tmpMBMGroupCompany) = Trim("LEMBAGA PEMBANGUNAN INDUSTRI PEMBINAAN MALAYSIA") Then
                            If tmpInitialLG <> "1" Then
                                Call MailAdmInitialLG("masliza@cidb.gov.my;rafidah@cidb.gov.my ;zaini.arifin@takaful-malaysia.com.my ;aziekasmadi.ag@isfics.com.my; aziekasmadi81@gmail.com ;", "", "")
                                strsqlCASLG = ""
                                strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                Set UPRS = Nothing
                            End If
                            
                        ElseIf Mid(txtCaseNo, 1, 2) = "EM" Then
                            If frmGLMenu.txtEmailGL <> "" Then
                                If tmpMBMEmail <> "" Then
                                    tmpMBMEmail = tmpMBMEmail & ";" & frmGLMenu.txtEmailGL
                                Else
                                    tmpMBMEmail = frmGLMenu.txtEmailGL
                                End If
                            End If
                            
                            If tmpMBMEmail <> "" Then
                                    ToAddress = tmpMBMEmail
                                    Call MailEtiqaAdmLG(ToAddress)
                                    strsqlCASLG = ""
                                    strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                    UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                    Set UPRS = Nothing
                            End If
                            
                        ElseIf Mid(txtCaseNo, 1, 2) = "ES" Then
                            
                            If frmGLMenu.txtEmailGL <> "" Then
                                If tmpMBMEmail <> "" Then
                                    tmpMBMEmail = tmpMBMEmail & ";" & frmGLMenu.txtEmailGL
                                Else
                                    tmpMBMEmail = frmGLMenu.txtEmailGL
                                End If
                            End If
                            
                            If tmpMBMEmail <> "" Then
                                    ToAddress = tmpMBMEmail
                                    Call MailEtiqaAdmLG(ToAddress)
                                    strsqlCASLG = ""
                                    strsqlCASLG = " update CAS set  InitialLGMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                                    UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                                    Set UPRS = Nothing
                            End If
                            
                        End If
                    End If
                    cmdLG.Enabled = False
                End If
            End If
        
            SndLG = False
            cmdShow.Enabled = True
            cmdLG.Enabled = True
            cmdRepudiate.Enabled = True
            cmdSaveAdd.Enabled = True
            cmdPrior.Enabled = True
            cmdPrev.Enabled = True
            CmdNext.Enabled = True
            cmdUpdate1.Enabled = True
            frmGLMenu.GLForm = False
            Unload frmGLMenu
            
            If txtPayCode = "SI" Then
            Dim ToAdd As String
            Dim Body As String
            
            ToAdd = "sjtung@berjayasompo.com.my;dzulhizam@berjayasompo.com.my;farkadir@berjayasompo.com.my;bhyeap@berjayasompo.com.my"
            'ToAdd = "bryan@medix.com.my;trinath@medix.com.my;trinath18@gmail.com"
            Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
                            "NOTIFICATION OF ADMISSION" & vbCrLf & vbCrLf & _
                            "Please be advised that the above-mentioned member is admitted to hospital. Below is the case details for your attention." & vbCrLf & vbCrLf & _
                            "Company Name : " & tmpMBMGroupCompany & vbCrLf & _
                            "Patient Name : " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbCrLf & _
                            "Policy Number: " & tmptxtMBMPolNo & vbCrLf & _
                            "LG Number : " & LGNumber & vbCrLf & _
                            "Date of Admission : " & txtDateAdmitted & vbCrLf & _
                            "Hospital : " & txtHOSCode & vbCrLf & _
                            "Diagnosis : " & txtReason & vbCrLf & _
                            vbCrLf & vbCrLf & _
                            "Please contact the us at 603-7803 2003 if you require any further clarification." & vbCrLf & vbCrLf & _
                            vbCrLf & vbCrLf & _
                            "Thank you." & vbCrLf & _
                            "Regards, " & vbCrLf & _
                            "MediExpress (Malaysia) Sdn Bhd"
                            
            Call MailCDOMsg(ToAdd, Body)
            
            ElseIf tmpMBMGroupCompany = "PELABUHAN TANJUNG PELEPAS SDN BHD" Then
               ' Dim ToAdd As String
               ' Dim Body As String
                
                'ToAdd = "bryan@medix.com.my;sashi@medix.com.my"
                
                ToAdd = "rosliza@ptp.com.my;norma@ptp.com.my"
                'ToAdd = "bryan@medix.com.my;trinath@medix.com.my;trinath18@gmail.com"
                Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
                                "NOTIFICATION OF ADMISSION" & vbCrLf & vbCrLf & _
                                "Please be advised that the above-mentioned member is admitted to hospital. Below is the case details for your attention." & vbCrLf & vbCrLf & _
                                "Company Name : " & tmpMBMGroupCompany & vbCrLf & _
                                "Patient Name : " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbCrLf & _
                                "Policy Number: " & tmptxtMBMPolNo & vbCrLf & _
                                "LG Number : " & LGNumber & vbCrLf & _
                                "Date of Admission : " & txtDateAdmitted & vbCrLf & _
                                "Hospital : " & txtHOSCode & vbCrLf & _
                                "Diagnosis : " & txtReason & vbCrLf & _
                                vbCrLf & vbCrLf & _
                                "Please contact the us at 603-7803 2003 if you require any further clarification." & vbCrLf & vbCrLf & _
                                vbCrLf & vbCrLf & _
                                "Thank you." & vbCrLf & _
                                "Regards, " & vbCrLf & _
                                "MediExpress (Malaysia) Sdn Bhd"
                                
                Call MailCDOMsg(ToAdd, Body)

                'Set oOApp = CreateObject("Outlook.Application")
                'Set oOMail = oOApp.CreateItem(olMailItem)
    
                'With oOMail
                 '   '.SenderEmailAddress = "admission@medix.com.my"
                 '   '.To = "sjtung@berjayasompo.com.my;dzulhizam@berjayasompo.com.my;farkadir@berjayasompo.com.my"
                 '   .To = "bryan@medix.com.my"
                 '   .BCC = "bryan@medix.com.my"
                 '   .Subject = "Notification of Admission" & "-" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                 '   .Body = "Dear Sir/Madam," & vbNewLine & vbNewLine & _
                 '           "NOTIFICATION OF ADMISSION" & vbNewLine & vbNewLine & _
                 '           "Please be advised that the above-mentioned member is admitted to hospital. Below is the case details for your attention." & vbNewLine & vbNewLine & _
                 '           "Company Name : " & tmpMBMGroupCompany & vbNewLine & _
                 '           "Patient Name : " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbNewLine & _
                 '           "Policy Number: " & tmptxtMBMPolNo & vbNewLine & _
                 '           "LG Number : " & LGNumber & vbNewLine & _
                 '           "Date of Admission : " & txtDateAdmitted & vbNewLine & _
                 '           "Hospital : " & txtHOSCode & vbNewLine & _
                 '           "Diagnosis : " & txtReason & vbNewLine & _
                 '           vbNewLine & vbNewLine & _
                 '           "Please contact the us at 603-7803 2003 if you require any further clarification." & vbNewLine & vbNewLine & _
                 '           vbNewLine & vbNewLine & _
                 '           "Thank you." & vbNewLine & _
                 '           "Regards, " & vbNewLine & _
                 '           "MediExpress (Malaysia) Sdn Bhd"
                '.Send
                
                'End With
            
            
            ElseIf tmpAgentCode = "T04372K" Then
                'Set oOApp = CreateObject("Outlook.Application")
                'Set oOMail = oOApp.CreateItem(olMailItem)
    
                'With oOMail
                    '.SenderEmailAddress = "admission@medix.com.my"
                    ToAdd = ""
                    Body = ""
                    ToAdd = "setiarm@streamyx.com"
                    '.To = "bryan@medix.com.my"
                    '.BCC = "bryan@medix.com.my;siewping@medix.com.my"
                    '.Subject = "Notification of Admission" & "-" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                    Body = "Dear Sir/Madam," & vbNewLine & vbNewLine & _
                            "NOTIFICATION OF ADMISSION" & vbNewLine & vbNewLine & _
                            "Please be advised that the above-mentioned member is admitted to hospital. Below is the case details for your kind attention." & vbNewLine & vbNewLine & _
                            "Patient Name             :       " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbNewLine & _
                            "Policy Number            :       " & tmptxtMBMPolNo & vbNewLine & _
                            "LG Number                :       " & LGNumber & vbNewLine & _
                            "Date of Admission      :       " & txtDateAdmitted & vbNewLine & _
                            "Hospital                       :       " & txtHOSCode & vbNewLine & _
                            vbNewLine & vbNewLine & _
                            "Please contact the us at 603-7803 2003 if you require any further clarification." & vbNewLine & vbNewLine & _
                            vbNewLine & vbNewLine & _
                            "Thank you." & vbNewLine & _
                            "Regards, " & vbNewLine & _
                            "MediExpress (Malaysia) Sdn Bhd"
            
                  Call MailCDOMsg(ToAdd, Body)
            ElseIf txtPayCode = "AY" Then
            
                    'Set oOApp = CreateObject("Outlook.Application")
                    'Set oOMail = oOApp.CreateItem(olMailItem)
                    'With oOMail
                        '.SenderEmailAddress = "admission@medix.com.my"
                        '.To = "nurshahirah.ng@axa-life.com.my"
                       ToAdd = ""
                       ToAdd = "claims@axa-life.com.my"
                       '.To = "bryan@medix.com.my"
                       ' .BCC = "bryan@medix.com.my;siewping@medix.com.my"
                       ' .Subject = "Notification of Admission" & "-" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                       Body = "Dear Sir/Madam," & vbNewLine & vbNewLine & _
                                "NOTIFICATION OF ADMISSION" & vbNewLine & vbNewLine & _
                                "Please be advised that the above-mentioned member is admitted to hospital. Below is the case details for your kind attention." & vbNewLine & vbNewLine & _
                                "Patient Name             :       " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbNewLine & _
                                "Policy Number            :       " & tmptxtMBMPolNo & vbNewLine & _
                                "LG Number                :       " & LGNumber & vbNewLine & _
                                "Date of Admission      :       " & txtDateAdmitted & vbNewLine & _
                                "Hospital                       :       " & txtHOSCode & vbNewLine & _
                                vbNewLine & vbNewLine & _
                                "Please contact the us at 603-7803 2003 if you require any further clarification." & vbNewLine & vbNewLine & _
                                vbNewLine & vbNewLine & _
                                "Thank you." & vbNewLine & _
                                "Regards, " & vbNewLine & _
                                "MediExpress (Malaysia) Sdn Bhd"
                    '.Send
                    'End With
                    
                     Call MailCDOMsg(ToAdd, Body)
                    
        End If
    End If
End Sub
Public Sub MailCDOMsg(ToAdd As String, Body As String)
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object
On Error GoTo debugs

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "Notification of Admission" & "-" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
'objMessage.From = "adm@medix.com.my"
objMessage.From = "notifications@medix.com.my"
objMessage.To = ToAdd
'objMessage.BCC = "bryan@medix.com.my"
'objMessage.TextBody = "This is some sample message text.." & vbCrLf & "It was sent using SMTP authentication."
objMessage.TextBody = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
                            "NOTIFICATION OF ADMISSION" & vbCrLf & vbCrLf & _
                            "Please be advised that the above-mentioned member is admitted to hospital. Below is the case details for your attention." & vbCrLf & vbCrLf & _
                            "Company Name : " & tmpMBMGroupCompany & vbCrLf & vbCrLf & _
                            "Patient Name : " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbCrLf & vbCrLf & _
                            "Policy Number: " & tmptxtMBMPolNo & vbCrLf & vbCrLf & _
                            "LG Number : " & LGNumber & vbCrLf & vbCrLf & _
                            "Date of Admission : " & txtDateAdmitted & vbCrLf & vbCrLf & _
                            "Hospital : " & txtHOSCode & vbCrLf & vbCrLf & _
                            "Diagnosis : " & txtReason & vbCrLf & vbCrLf & _
                            "Please contact us at 603-7803 2003 if you require any further clarification." & vbCrLf & vbCrLf & _
                            "Thank you." & vbCrLf & _
                            "Regards, " & vbCrLf & _
                            "MediExpress (Malaysia) Sdn Bhd"

objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "notifications@medix.com.my"
'objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "adm@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "medix123"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update
objMessage.Send
debugs:
If Err.Description <> "" Then MsgBox Err.Description

End Sub

Private Sub SendSMSEtiqa()
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXml As String
  Dim Messagestring As String
  
  'Dear <policyholder's name> <policy number>, we are pleased to inform that your Guarantee Letter has been approved. We wish you a speedy recovery
    'tmpMBMTelNoM = "0123652738"
    
    strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    'strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    strSoapAction = "http://tempuri.org/SendSMS"
    Dim strxml2 As String
    
    Messagestring = "ETIQA:" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & txtPatientIC & " IGL Approved" & txtDateAdmitted & " R&B:" & tmpRMAmt & "per day, if exceed, co-payment apply depend on policy TC. TQ" & ""
    
    
    strXml = ""
    strxml2 = ""
               ' frmGLMenu.txtGLDate = Format(Date, "dd/mm/yyyy")
            'frmGLMenu.txtLGTime = Format(Time, "hh:mm")
                                   ' frmGLMenu.txtRequestDate = IIf(Len(CLMLG!CaseDate), Format(CLMLG!CaseDate, "dd/mm/yyyy"), "__/__/____")
                        'frmGLMenu.txtRequestTime = IIf(Len(CLMLG!CaseTime), Mid(CLMLG!CaseTime, 1, 5), "__:__")
                        
                        
   
             
             
    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<SendSMS xmlns=""http://tempuri.org/"">"
             
             

    strxml2 = "<MobileNo>" & tmpMBMTelNoM & " </MobileNo>" & _
             "<Msg>" & Messagestring & "</Msg>" & _
             "</SendSMS>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"
             
             Dim stringxml As String
             
        
    Debug.Print PostWebservice(strUrl, strSoapAction, strXml + strxml2)
   ' rst2.MoveNext
    'Loop
    
    'Set LGDR = Nothing
    
End Sub
Public Sub MobileNotification(MBMNumber As String, LGNumber As String)
      Dim sqlstr As String
      Dim LGstr As String
      Dim MBMId As String
      Dim NotMsg As String
      Dim LGDR As New ADODB.Recordset
      Dim rst2 As New ADODB.Recordset
      Dim LGInfo As New ADODB.Recordset
      Dim NotificationDR As New ADODB.Recordset
      Dim isFirst As Boolean
      isFirst = True
      sqlstr = ""
      MBMId = ""
   
             
      sqlstr = "select  MBMAndroidID.RegistrationID  , MBMAndroidID.MemberID from MBMLoginDetails inner join MBMAndroidID on MBMAndroidID.MemberID= MBMLoginDetails.MBMUserName where MBMLoginDetails.MBMNumber ='" & MBMNumber & "' and Status= 1"
      rst2.Open sqlstr, medixmasterconn, adOpenForwardOnly, adLockOptimistic
      NotMsg = ""
      
      Do Until rst2.EOF
      MBMId = rst2!MemberID
      
      If isFirst Then
       Dim cmd As New ADODB.Command
       Dim Prm1 As New ADODB.Parameter
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Sp_AndroidNotification"
            cmd.CommandTimeout = 300
            cmd.CommandType = adCmdStoredProc
            Set Prm1 = cmd.CreateParameter("LGNo", adVarChar, adParamInput, 20, LGNumber)
            cmd.Parameters.Append Prm1
            LGInfo.CacheSize = 100
            LGInfo.CursorLocation = adUseClient
            LGInfo.CursorType = adOpenForwardOnly
            Set LGInfo = cmd.Execute
      LGstr = ""
     
       NotMsg = "Notification Alert From MediExpress!" + LGInfo!LGPurpose + " GL Issued for " + LGInfo!PatientName + " on " + Format(LGInfo!LGDate, "dd/mm.yyyy")
        
       NotificationDR.Open "MBMAndroidNotification", medixmasterconn, adOpenKeyset, adLockOptimistic
       NotificationDR.AddNew
          With NotificationDR
                    !UserName = MBMId
                    !MBMNumber = MBMNumber
                    !MBMCoverId = LGInfo!MBMPatientCovID
                    !Notifications = "Notification Alert From MediExpress!" + LGInfo!LGPurpose + " GL Issued on " + Format(LGInfo!LGDate, "dd/mm.yyyy")
                    !NotificationDt = LGInfo!LGDate
                    !LGPurpose = LGInfo!LGPurpose
                    !LGNumber = LGNumber
                    !NotificationSeen = 0
          End With
          NotificationDR.Update
        Set NotificationDR = Nothing
        rst2.MoveNext
        Set cmd = Nothing
        Set LGInfo = Nothing
        
        isFirst = False
     End If
     
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXml As String
  
    'strUrl = "http://etiqahc.com.my/androidservice.asmx"
    strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    strSoapAction = "http://tempuri.org/InvokeNotification"
    Dim strxml2 As String

    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<InvokeNotification xmlns=""http://tempuri.org/"">"

    strxml2 = "<UserName>" & MBMId & " </UserName>" & _
             "<Msg>" & NotMsg & " </Msg>" & _
             "</InvokeNotification>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"
        
    Debug.Print PostWebservice(strUrl, strSoapAction, strXml + strxml2)
   ' rst2.MoveNext
    Loop
     Set LGDR = Nothing
     
End Sub
Private Function PostWebservice(ByVal AsmxUrl As String, ByVal SoapActionUrl As String, ByVal XmlBody As String) As String
    Dim objDom As Object
    Dim objXmlHttp As Object
    Dim strRet As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    On Error GoTo Err_PW
    
    ' Create objects to DOMDocument and XMLHTTP
    Set objDom = CreateObject("MSXML2.DOMDocument")
    Set objXmlHttp = CreateObject("MSXML2.XMLHTTP")
    objDom.async = True
    objDom.loadXML XmlBody
    objXmlHttp.Open "POST", AsmxUrl, False
    objXmlHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objXmlHttp.setRequestHeader "SOAPAction", SoapActionUrl
    objXmlHttp.Send objDom.XML
    strRet = objXmlHttp.responseText
    Set objXmlHttp = Nothing
    intPos1 = InStr(strRet, "Result>") + 7
    intPos2 = InStr(strRet, "</")
    If intPos1 > 7 And intPos2 > 0 Then
        strRet = Mid(strRet, intPos1, intPos2 - intPos1)
    End If
    PostWebservice = strRet
    Exit Function
Err_PW:
    PostWebservice = "Error: " & Err.Number & " - " & Err.Description
End Function

Private Sub SendSMS(message As String)
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXml As String
  
  
  'Dear <policyholder's name> <policy number>, we are pleased to inform that your Guarantee Letter has been approved. We wish you a speedy recovery
    'tmpMBMTelNoM = "0123652738"
    
    strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    'strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    strSoapAction = "http://tempuri.org/SendSMS"
    Dim strxml2 As String
    
    strXml = ""
    strxml2 = ""
               ' frmGLMenu.txtGLDate = Format(Date, "dd/mm/yyyy")
            'frmGLMenu.txtLGTime = Format(Time, "hh:mm")
                                   ' frmGLMenu.txtRequestDate = IIf(Len(CLMLG!CaseDate), Format(CLMLG!CaseDate, "dd/mm/yyyy"), "__/__/____")
                        'frmGLMenu.txtRequestTime = IIf(Len(CLMLG!CaseTime), Mid(CLMLG!CaseTime, 1, 5), "__:__")
    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<SendSMS xmlns=""http://tempuri.org/"">"

    'strxml2 = "<MobileNo>" & tmpMBMTelNoM & " </MobileNo>" & _
     '        "<Msg>" & "AXA Affin Life Claims:" & " GL Request : " & frmGLMenu.txtRequestDate & "," & frmGLMenu.txtRequestTime & " Dear " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & " " & tmptxtMBMPolNo & " ,we are pleased to inform that your Guarantee Letter has been approved. We wish you a speedy recovery." & " </Msg>" & _
     '        "</SendSMS>" & _
     '        "</soap:Body>" & _
     '        "</soap:Envelope>"
        
    strxml2 = "<MobileNo>" & tmpMBMTelNoM & " </MobileNo>" & _
             "<Msg>" & message & " </Msg>" & _
             "</SendSMS>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"
            
        
    Debug.Print PostWebservice(strUrl, strSoapAction, strXml + strxml2)
   ' rst2.MoveNext
    'Loop
    
    'Set LGDR = Nothing
    
End Sub

Private Sub SendSMSMAATakaful()
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXml As String
  
  
  'Dear <policyholder's name> <policy number>, we are pleased to inform that your Guarantee Letter has been approved. We wish you a speedy recovery
    'tmpMBMTelNoM = "0123652738"
    
    strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    'strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    strSoapAction = "http://tempuri.org/SendSMS"
    Dim strxml2 As String
    
    strXml = ""
    strxml2 = ""
    
    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<SendSMS xmlns=""http://tempuri.org/"">"

    strxml2 = "<MobileNo>" & tmpMBMTelNoM & " </MobileNo>" & _
             "<Msg>" & "MAA Takaful Berhad:" & "Dear " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & " " & tmptxtMBMPolNo & " ,we are pleased to inform that your Guarantee Letter has been approved. We wish you a speedy recovery." & " </Msg>" & _
             "</SendSMS>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"
        
    Debug.Print PostWebservice(strUrl, strSoapAction, strXml + strxml2)
   ' rst2.MoveNext
    'Loop
    
    'Set LGDR = Nothing
    
End Sub

Private Sub SendSMSReject()
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXml As String
  
  
  'tmpMBMTelNoM = "0123652738"
  'Dear <policyholder's name> <policy number>, we regret to inform that we are unable to approve your request for a Guarantee Letter due to <decline reason>.  You may proceed to make payment directly and file an appeal. Please do not hesitate to contact us at 03 - 7884 1814 for further clarification. We wish you a speedy recovery
  'tmpMBMTelNoM = "0192245720,0173701934,0166000941,0123652738"
    strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    'strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    strSoapAction = "http://tempuri.org/SendSMS"
    Dim strxml2 As String
    
    
    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<SendSMS xmlns=""http://tempuri.org/"">"
Dim msg As String
msg = "AXA Affin Life Claims:" & "Dear " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & " " & tmptxtMBMPolNo & ", we regret to inform that we are unable to approve your request for a Guarantee Letter due to " & tmpRejectCat & " . Please do not hesitate to contact us at 03 - 7884 1818 for further clarification. We wish you a speedy recovery."
    strxml2 = "<MobileNo>" & tmpMBMTelNoM & " </MobileNo>" & _
            "<Msg>" & msg & "</Msg>" & _
             "</SendSMS>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"
        
    Debug.Print PostWebservice(strUrl, strSoapAction, strXml + strxml2)
   ' rst2.MoveNext
    'Loop
    
    'Set LGDR = Nothing
    
End Sub

Private Sub SendSMSCancel()
    Dim strSoapAction As String
    Dim strUrl As String
    Dim strXml As String
  
  
  
  'Dear <policyholder's name> <policy number>, we regret to inform that we are unable to approve your request for a Guarantee Letter due to <decline reason>.  You may proceed to make payment directly and file an appeal. Please do not hesitate to contact us at 03 - 7884 1814 for further clarification. We wish you a speedy recovery
  'tmpMBMTelNoM = "0123652738"
    strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    'strUrl = "http://97.1.1.9:8884/pushnotification.asmx"
    strSoapAction = "http://tempuri.org/SendSMS"
    Dim strxml2 As String
    
    strXml = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
             "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
             "<soap:Body>" & _
             "<SendSMS xmlns=""http://tempuri.org/"">"
Dim msg As String
msg = "AXA Affin Life Claims:" & "Dear " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & " " & tmptxtMBMPolNo & ", this is to inform you that the earlier request for a Guarantee Letter has now been cancelled. Please do not hesitate to contact us at 03-7884 1818 for any further clarification. Thank you. "
    strxml2 = "<MobileNo>" & tmpMBMTelNoM & " </MobileNo>" & _
            "<Msg>" & msg & "</Msg>" & _
             "</SendSMS>" & _
             "</soap:Body>" & _
             "</soap:Envelope>"
        
    Debug.Print PostWebservice(strUrl, strSoapAction, strXml + strxml2)
   ' rst2.MoveNext
    'Loop
    
    'Set LGDR = Nothing
    
End Sub

Private Sub cmdRepudiate_Click()
Dim REP As New ADODB.Recordset
Dim strsql As String
Dim Isrepudate As Boolean
Dim RecordSaved
Isrepudate = False

tmpAgent = ""
repremarks = ""


    tmpRejectCat = ""
    Document = ""
    Personnel = ""
    Attention = ""
    tmpCashType = ""
    RepType = ""
    
    Dim strrep() As String
    If InStr(cboREPCode1, "||") Then
        If Len(cboREPCode1) Then
            strrep = Split(cboREPCode1, "||")
            Isrepudate = True
            RepType = Mid(strrep(1), 1, 1)
        ElseIf Len(cboREPCode2) Then
            strrep = Split(cboREPCode1, "||")
            Isrepudate = True
            RepType = Mid(strrep(1), 1, 1)
        ElseIf Len(cboREPCode3) Then
            strrep = Split(cboREPCode1, "||")
            Isrepudate = True
            RepType = Mid(strrep(1), 1, 1)
        End If
    Else
            RepType = ""
    End If
    
    If Trim(txtMBMNumber) = "" Then
        MsgBox "Please Enter Membership Information", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
    ElseIf Trim(cboPatientList) = "" Then
        MsgBox "Please Enter Patient Information", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
    ElseIf Trim(txtHOSCode) = "" Then
        MsgBox "Please Enter Hospital Information", vbOKOnly + vbCritical, DialogBoxTitle
        txtHOSCode.SetFocus
    ElseIf txtDateAdmitted = "__/__/____" Then
        MsgBox "Please Enter Admission Date", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
    ElseIf cboPatientList = "" Then
        MsgBox "Please Enter Patient's ID", vbOKOnly + vbCritical, DialogBoxTitle
        cboPatientList.SetFocus
    ElseIf Trim(txtCASReserveAmount) = "" Then
        MsgBox "Please Enter Reserve Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASReserveAmount.SetFocus
    ElseIf Isrepudate = False Then
        MsgBox "Please Enter Repudiation Reason", vbOKOnly + vbCritical, DialogBoxTitle
        cboREPCode1.SetFocus
    Else
        tmpCashType = Mid(CboPayType, 1, 1)
        If Trim(tmpCashType) <> "R" Then
            strsql = "Exec Case_Adjustment '" & Trim(txtCaseNo) & "', 22"
            REP.Open strsql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
            If Not REP.EOF Then
                If REP!CASDocument <> "" And REP!CASPersonnel <> "" Then
                    With REP
                        frmRepudiationSelection.Document = !CASDocument
                        frmRepudiationSelection.Personnel = !CASPersonnel
                        frmRepudiationSelection.Attention = !CASAttention
                        If Len(!Agent) Then
                            frmRepudiationSelection.txtAgent = !Agent
                        End If
                        If Len(!Remarks) Then
                            frmRepudiationSelection.txtremarks = !Remarks
                        End If
                        tmpAgent = frmRepudiationSelection.txtAgent
                        repremarks = frmRepudiationSelection.txtremarks
                    End With
                Else
                    frmRepudiationSelection.Document = ""
                    frmRepudiationSelection.Personnel = ""
                    frmRepudiationSelection.Attention = ""
                    frmRepudiationSelection.txtAgent = ""
                    frmRepudiationSelection.txtremarks = ""
                    tmpAgent = frmRepudiationSelection.txtAgent
                    repremarks = frmRepudiationSelection.txtremarks
                    AdmStatus = ""
                End If
                
                If REP!CASUnderlying = "Y" Then
                    frmRepudiationSelection.CASUnderlying = True
                Else
                    frmRepudiationSelection.CASUnderlying = False
                End If
            Else
                frmRepudiationSelection.Document = ""
                frmRepudiationSelection.Personnel = ""
                frmRepudiationSelection.Attention = ""
                frmRepudiationSelection.CASUnderlying = False
            End If
            REP.Close
            Set REP = Nothing
            frmRepudiationSelection.Source = 1
            frmRepudiationSelection.Show vbModal
            
           
            
            If Document <> "" And Personnel <> "" Then
                RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
                    If RecordSaved = vbYes Then
                        Screen.MousePointer = vbHourglass
                        Letter = ""
                        Letter = "R"
                        
                        'tmpAgent = frmRepudiationSelection.txtAgent
                        'repremarks = frmRepudiationSelection.txtRemarks
                        Call SaveRepudiation
                        cmdRepudiate.Enabled = False
                        cmdLG.Enabled = False
                        Screen.MousePointer = vbDefault
                    End If
            End If
            
                Dim UPRS As New ADODB.Recordset
                tmpStatusTime = Format$(Now, "dd/mm/yyyy hh:mm")
                strsqlCASLG = ""
                strsqlCASLG = " update tblRecord set " & _
                         " IsRejected = '1', IsLGissued = '0', CaseProgress = 'Case declined'" & _
                         " Where CaseNo = '" & Trim(txtCaseNo) & "'"
                UPRS.Open strsqlCASLG, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
                
            If Mid(txtCaseNo, 1, 2) = "AY" And txtDateAdmitted > "03/08/2015" Then
            strsqlCASLG = ""
            strsqlCASLG = "SELECT RepCode.RepCode, RepCodeCategory.RepCatCode, RepCodeCategory.RepCatName FROM RepCode INNER JOIN RepCodeCategory ON RepCode.RepCatCode = RepCodeCategory.RepCatCode where RepCode.RepCode = '" & Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - IIf(InStr(1, cboREPCode1, "-") = 0, 0, 1))) & "'"
            UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                
                If UPRS.EOF = False Then
                    
                    tmpRejectCat = IIf(Len(UPRS!RepCatName), UPRS!RepCatName, "")
                Else
                    tmpRejectCat = ""
                End If
            UPRS.Close
            Set UPRS = Nothing
                     If tmpCASRejectSMS <> "1" Then
                         If tmpMBMTelNoM <> "" Then
                             Call SendSMSReject
                         End If
                         strsqlCASLG = ""
                         strsqlCASLG = " update CAS set  CASLGSMS = '0', CASPendingSMS = '0', CASRejectSMS = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                         UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                        tmpCASRejectSMS = 1
                     End If
            End If
            
            Call funcwordRepudation(RepType)
    Else
        Screen.MousePointer = vbHourglass
        txtCASReserveAmount.SetFocus
        Letter = ""
        Letter = "R"
        Call SaveRepudiation
        'Call funcword
        Call funcwordRepudation(RepType)
        cmdRepudiate.Enabled = False
        cmdLG.Enabled = False
        Screen.MousePointer = vbDefault
    End If
    End If
    cmdShow.Enabled = True
    cmdRepudiate.Enabled = True
    cmdSaveAdd.Enabled = True
    cmdPrior.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdUpdate1.Enabled = True
End Sub

Public Sub funcwordRepudation(RepType As String)
Dim X As Boolean
Dim DocSavePath As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim TemplateName As String
TemplateName = ""
Set objword = Nothing

'select Count(*) as Cont from Tbl_RepTemplate  where PayCode='ST'
''txtPayCode

Dim sqlqryrep As String
Dim RepRS As New ADODB.Recordset
Dim RepRS1 As New ADODB.Recordset
sqlqryrep = ""
sqlqryrep = "select Count(*) as Cont from Tbl_RepTemplate  where PayCode='" + Trim(txtPayCode) + "'"
RepRS.Open sqlqryrep, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
If CInt(RepRS!Cont) > 0 Then
    sqlqryrep = ""
    sqlqryrep = "select TemplateName from Tbl_RepTemplate  where PayCode='" + Trim(txtPayCode) + "'"
    RepRS1.Open sqlqryrep, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
    TemplateName = RepRS1!TemplateName
    RepRS1.Close
    Set RepRS1 = Nothing
    IsEtiqa = True
Else
        If RepType = "D" Then
            TemplateName = "DisclouserTemplate.dot"
        ElseIf RepType = "S" Then
            TemplateName = "DisclouserTemplate.dot"
        ElseIf RepType = "N" Then
            TemplateName = "NonDisclouserTemplate.dot"
        End If
        IsEtiqa = False
End If
RepRS.Close
Set RepRS = Nothing

        Set objword = New Word.Application
        objword.Visible = True
        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & TemplateName)
        bookmarkDisclouser (TemplateName)
        objdoc.SaveAs (DocSavePath & "R" & Trim(txtCaseNo) & ".doc")
        X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
        objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
        objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & "R" & Trim(txtCaseNo) & ".pdf", FileFormat:=Word.wdFormatPDF
        Me.Show

End Sub

Private Function bookmarkDisclouser(Optional strAppend As String) As Integer
Dim RepDate As String
Dim Insured As String
Dim InsuredAddress As String
Dim MBMNumber As String
Dim InsuranceCompany As String
Dim PolicyNumber As String
Dim EmployeeName As String
Dim PatinetName As String
Dim Hospital As String
Dim DOA As String
Dim Diagnosis As String
Dim illness As String
Dim Exclusion As String
Dim SubExclusion As String
Dim PIC As String
Dim claimcontact As String
Dim Insurer As String
Dim Address1 As String
Dim Address2 As String
Dim PostCode As String
Dim state As String
Dim PIC1 As String
Dim PIC2 As String
Dim InsAdd1 As String
Dim InsAdd2 As String
Dim InsAdd3 As String
Dim strsqlPay As String
Dim PAY As New ADODB.Recordset
strsqlPay = ""

strsqlPay = "select ClmContact,claimPIC,ClmContact, PAYCode,PAYGRPCode,PAYCompanyName,PAYAddress1,PAYAddress2,PAYAddress3,PAYMainContact,PAYTELno from PAY where PAYCode ='" & Trim(txtPayCode) & "'"
PAY.Open strsqlPay, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    Do Until PAY.EOF
        'Insured = PAY!PAYCompanyName
        'InsuredAddress = IIf(Len(PAY!PAYAddress1), PAY!PAYAddress1, "")
        InsuranceCompany = UCase(IIf(Len(PAY!PAYCompanyName), PAY!PAYCompanyName, ""))
        Insurer = UCase(IIf(Len(PAY!PAYCompanyName), PAY!PAYCompanyName, ""))
        Address1 = UCase(IIf(Len(PAY!PAYAddress1), PAY!PAYAddress1, ""))
        Address2 = UCase(IIf(Len(PAY!PAYAddress2), PAY!PAYAddress2, ""))
        PostCode = UCase(IIf(Len(PAY!PAYAddress3), PAY!PAYAddress3, ""))
        If Len(PAY!claimPIC) Then
            PIC1 = UCase(IIf(Len(PAY!claimPIC), PAY!claimPIC, ""))
        Else
            PIC1 = UCase(IIf(Len(PAY!PAYMAINContact), PAY!PAYMAINContact, ""))
        End If
        claimcontact = UCase(IIf(Len(PAY!ClmContact), PAY!ClmContact, ""))
        PIC2 = Insurer
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    'Get Member address
    
    strsqlPay = "select  INSCode, MBMName, MBMAddress1 ,MBMAddress2,MBMAddress3 , MBMNumber    from MBM  where MBMNumber ='" + txtMBMNumber + "'"
    PAY.Open strsqlPay, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    Do Until PAY.EOF
    If PAY!INSCode = "I" Then
        Insured = UCase(IIf(Len(PAY!MBMName), PAY!MBMName, ""))
        If PAY!MBMAddress1 = "NA" Then
            InsAdd1 = "-"
        Else
            InsAdd1 = UCase(IIf(Len(PAY!MBMAddress1), PAY!MBMAddress1, ""))
        End If
        If PAY!MBMAddress2 = "NA" Then
            InsAdd2 = "-"
        Else
            InsAdd2 = UCase(IIf(Len(PAY!MBMAddress2), PAY!MBMAddress2, ""))
        End If
        If PAY!MBMAddress3 = "NA" Then
            InsAdd3 = "-"
        Else
            InsAdd3 = UCase(IIf(Len(PAY!MBMAddress3), PAY!MBMAddress3, ""))
        End If
    Else
        Dim grpQry As String
        Dim GrpRS As New ADODB.Recordset
        
        grpQry = "Select CompanyGRP.CoGRPName , CompanyGRP.CoGRPAdd1 , CompanyGRP.CoGRPAdd2 , CompanyGRP.CoGRPAdd3  from MBMOthers " & _
                 " inner join MBM on MBM.MBMNumber = MBMOthers.MBMNumber " & _
                 " inner join CompanyGRP on CompanyGRP.CoGRPName = MBMOthers.MBMGroupCompany   and CompanyGRP.CoGRPPaycode = MBM.PAYCode " & _
                 " Where MBMOthers.MBMNumber ='" + txtMBMNumber + "'"
        GrpRS.Open grpQry, medixmasterconn, adOpenForwardOnly, adLockOptimistic
        Do Until GrpRS.EOF
            Insured = UCase(IIf(Len(GrpRS!CoGRPName), GrpRS!CoGRPName, ""))
                If GrpRS!CoGRPAdd1 = "NA" Then
                    InsAdd1 = "-"
                Else
                    InsAdd1 = UCase(IIf(Len(GrpRS!CoGRPAdd1), GrpRS!CoGRPAdd1, ""))
                End If
                If GrpRS!CoGRPAdd2 = "NA" Then
                    InsAdd2 = "-"
                Else
                    InsAdd2 = UCase(IIf(Len(GrpRS!CoGRPAdd2), GrpRS!CoGRPAdd2, ""))
                End If
                If GrpRS!CoGRPAdd3 = "NA" Then
                    InsAdd3 = "-"
                Else
                    InsAdd3 = UCase(IIf(Len(GrpRS!CoGRPAdd3), GrpRS!CoGRPAdd3, ""))
                End If
           GrpRS.MoveNext
        Loop
        GrpRS.Close
        Set GrpRS = Nothing
    End If
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    InsuredAddress = ""
    If Len(InsAdd1) Then
        InsuredAddress = InsuredAddress + InsAdd1 + " "
    End If
    
    If Len(InsAdd2) Then
        InsuredAddress = InsuredAddress + InsAdd2 + " "
    End If
    
    If Len(InsAdd3) Then
        InsuredAddress = InsuredAddress + InsAdd3 + " "
    End If
    
       If objdoc.Bookmarks.Exists("RepDate") = True Then
           objdoc.Bookmarks("RepDate").Range.Text = CStr(Format(DateTime.Now, "DD/MM/YYYY"))
       End If
       If objdoc.Bookmarks.Exists("CasNumber") = True Then
           objdoc.Bookmarks("CasNumber").Range.Text = txtCaseNo
       End If
       
       If objdoc.Bookmarks.Exists("Insured") = True Then
           objdoc.Bookmarks("Insured").Range.Text = Insured
       End If
       If objdoc.Bookmarks.Exists("InsuredAddress") = True Then
           objdoc.Bookmarks("InsuredAddress").Range.Text = InsuredAddress
       End If
       If objdoc.Bookmarks.Exists("MBMNumber") = True Then
           objdoc.Bookmarks("MBMNumber").Range.Text = txtMBMNumber
       End If
       If objdoc.Bookmarks.Exists("InsuranceCompany") = True Then
           objdoc.Bookmarks("InsuranceCompany").Range.Text = InsuranceCompany
       End If
       If objdoc.Bookmarks.Exists("PolicyNumber") = True Then
           objdoc.Bookmarks("PolicyNumber").Range.Text = tmptxtMBMPolNo
       End If
       If objdoc.Bookmarks.Exists("EmployeeName") = True Then
            objdoc.Bookmarks("EmployeeName").Range.Text = txtMBMName
       End If
       
       If objdoc.Bookmarks.Exists("PatinetName") = True Then
           Dim patient() As String
           If InStr(cboPatientList, "-") Then
                patient = Split(cboPatientList, "-")
                objdoc.Bookmarks("PatinetName").Range.Text = patient(1)
           Else
                objdoc.Bookmarks("PatinetName").Range.Text = cboPatientList
           End If
           'objdoc.Bookmarks("PatinetName").Range.Text = cboPatientList
       End If
       
       'txtMBMIcBcPP
       If objdoc.Bookmarks.Exists("ICNumber") = True Then
           objdoc.Bookmarks("ICNumber").Range.Text = txtMBMIcBcPP
       End If
       
       If objdoc.Bookmarks.Exists("PFNumber") = True Then
           objdoc.Bookmarks("PFNumber").Range.Text = ""
       End If
       
       If objdoc.Bookmarks.Exists("Hospital") = True Then
           objdoc.Bookmarks("Hospital").Range.Text = txtHOSCode
       End If
       
       If objdoc.Bookmarks.Exists("DOA") = True Then
           objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
       End If
       
       If objdoc.Bookmarks.Exists("Remarks") = True Then
           objdoc.Bookmarks("Remarks").Range.Text = repremarks
       End If
       
       If objdoc.Bookmarks.Exists("HospAdd1") = True Then
           objdoc.Bookmarks("HospAdd1").Range.Text = txtHOSCode
       End If
       
       'Remarks
       
        Dim strsql5 As String
        If txtFinalDiag1(0) <> "" Then
            strsql5 = strsql5 + txtFinalDiag1(0)
            If txtFinalDiag1(1) <> "" Then
                strsql5 = strsql5 + ", " + txtFinalDiag1(1)
                If txtFinalDiag1(2) <> "" Then
                    strsql5 = strsql5 + ", " + txtFinalDiag1(2)
                End If
            End If
        End If
       If objdoc.Bookmarks.Exists("Diagnosis") = True Then
           objdoc.Bookmarks("Diagnosis").Range.Text = strsql5
       End If
       
       
       illness = ""
       
       If Len(txtUD1(0)) Then
        illness = illness + " " + txtUD1(0)
       End If
       
       If Len(txtUD2) Then
        illness = illness + " " + txtUD2
       End If
       
       If Len(txtUD3) Then
        illness = illness + " " + txtUD3
       End If
       If objdoc.Bookmarks.Exists("illness") = True Then
           objdoc.Bookmarks("illness").Range.Text = illness
       End If
       
        Dim strrep() As String
        Dim RepType As String
        If Len(cboREPCode1) Then
            strrep = Split(cboREPCode1, "||")
            RepType = strrep(1)
        ElseIf Len(cboREPCode2) Then
            strrep = Split(cboREPCode1, "||")
            RepType = strrep(1)
        ElseIf Len(cboREPCode3) Then
            strrep = Split(cboREPCode1, "||")
            RepType = strrep(1)
        End If
        
        
       If Mid(RepType, 1, 1) <> "S" Then
            If objdoc.Bookmarks.Exists("Exclusion") = True Then
                objdoc.Bookmarks("Exclusion").Range.Text = strrep(0)
            End If
            If objdoc.Bookmarks.Exists("SubExclusion") = True Then
                objdoc.Bookmarks("SubExclusion").Range.Text = ""
            End If
       Else
            'Get Exclusion
            
            Dim sqlQry As String
            Dim DisClouser As String
            sqlQry = ""
            DisClouser = ""
            sqlQry = " Select Tbl_Disclouser.DisClouser, Tbl_Subdisclouser.DisclouserID,Tbl_Subdisclouser.SubDisclouser from Tbl_PaycodeDisclouser " & _
                    " inner join Tbl_Disclouser on Tbl_Disclouser.DisclouserID = Tbl_PaycodeDisclouser.DisclouserID  inner join Tbl_Subdisclouser on  Tbl_Subdisclouser.DID = Tbl_Disclouser.DID and Tbl_Subdisclouser.DisclouserID = Tbl_Disclouser.DisclouserID" & _
                    " Where Tbl_Subdisclouser.SubDisclouser='" + strrep(0) + "'  and PayCode ='" & Trim(txtPayCode) & "'"
            PAY.Open sqlQry, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
            Do Until PAY.EOF
                DisClouser = PAY!DisClouser
                PAY.MoveNext
            Loop
            PAY.Close
            Set PAY = Nothing

            If objdoc.Bookmarks.Exists("Exclusion") = True Then
                objdoc.Bookmarks("Exclusion").Range.Text = DisClouser
            End If
            If objdoc.Bookmarks.Exists("SubExclusion") = True Then
                objdoc.Bookmarks("SubExclusion").Range.Text = strrep(0)
            End If
       End If
       
       If objdoc.Bookmarks.Exists("PIC") = True Then
           objdoc.Bookmarks("PIC").Range.Text = UCase(claimcontact)
       End If
       If objdoc.Bookmarks.Exists("Insurer") = True Then
           objdoc.Bookmarks("Insurer").Range.Text = UCase(Insurer)
       End If
       If objdoc.Bookmarks.Exists("Address1") = True Then
           objdoc.Bookmarks("Address1").Range.Text = UCase(Address1)
       End If
       If objdoc.Bookmarks.Exists("Address2") = True Then
           objdoc.Bookmarks("Address2").Range.Text = Address2
       End If
       If objdoc.Bookmarks.Exists("PostCode") = True Then
           objdoc.Bookmarks("PostCode").Range.Text = PostCode
       End If
       If objdoc.Bookmarks.Exists("PIC1") = True Then
           objdoc.Bookmarks("PIC1").Range.Text = UCase(Insurer) + "-" + UCase(PIC1)
       End If
       If objdoc.Bookmarks.Exists("PIC2") = True Then
           objdoc.Bookmarks("PIC2").Range.Text = UCase(tmpAgent)
       End If
       'If objdoc.Bookmarks.Exists("AgentName") = True Then
       '    objdoc.Bookmarks("AgentName").Range.Text = UCase(tmpAgent)
       'End If
      bookmarkDisclouser = True
    End Function
Private Sub CmdNext_Click()
    Dim CAS As New ADODB.Recordset
    Dim sql As String
    Dim CASStatus As Boolean
    Dim CasNumber As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strOrderby As String
    CAS.MaxRecords = 1
    strAppend = ""

    strAppendCase = "2"
    strAppend = " AND CASNumber > '" & RemStr(txtCaseNo) & "'"
    strOrderby = " Order by CASnumber "
    strAppend = strAppend + strOrderby
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    CAS.CursorLocation = adUseClient
    CAS.CursorType = adOpenForwardOnly
    CAS.CacheSize = 100
    Set CAS = cmd.Execute
    
    If CAS.EOF = False Then
        CASStatus = True
        CasNumber = CAS!CasNumber
    Else
        CASStatus = False
    End If
    
    CAS.Close
    Set CAS = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    cmdUpdate1.Visible = True

    If CASStatus = True Then
        txtCaseNo.Text = CasNumber
        cmdShow_Click
    Else
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
    End If
    cmdShow.Enabled = True
    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    Else
        cmdRepudiate.Enabled = False
        'cmdLG.Enabled = True
    End If
    cmdSaveAdd.Enabled = True
    cmdPrior.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdUpdate1.Enabled = True

End Sub

Private Sub cmdPrev_Click()
    Dim CAS As New ADODB.Recordset
    Dim sql As String
    Dim CASStatus As Boolean
    Dim CasNumber  As String
    Dim StrsqlAppend As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strOrderby As String
    
    
  '  lblTopupStatus.Visible = False
    'Call SendSMS
    If Trim(txtCaseNo) <> "" Then
       ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        'sql = " select CASNumber from CASCrossReference where CASnumber < '" & RemStr(txtCaseNo) & "' " & _
            "Order by CASNumber DESC "
        strAppend = ""

        CAS.MaxRecords = 1
        'CAS.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        strAppendCase = "2"
        strAppend = " AND CASnumber < '" & RemStr(txtCaseNo) & "'"
        strOrderby = "Order by CASNumber DESC "
        strAppend = strAppend + strOrderby
        Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        CAS.CursorLocation = adUseClient
        CAS.CursorType = adOpenForwardOnly
        CAS.CacheSize = 100
        Set CAS = cmd.Execute
        
        If CAS.EOF = False Then
            CASStatus = True
            CasNumber = CAS!CasNumber
        Else
            CASStatus = False
        End If
        
        CAS.Close
        Set CAS = Nothing
       ' medixmasterconnMaint.Close
       ' Set medixmasterconnMaint = Nothing
        cmdUpdate1.Visible = True

        If CASStatus = True Then
            txtCaseNo.Text = CasNumber
            cmdShow_Click
        Else
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Case Number Format! ", vbCritical + vbOKOnly
        End If
    End If
    cmdShow.Enabled = True
    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
        cmdRepudiate.Enabled = True
        cmdLG.Enabled = False
    Else
        cmdRepudiate.Enabled = False
        'cmdLG.Enabled = True
    End If
    cmdSaveAdd.Enabled = True
    cmdPrior.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdUpdate1.Enabled = True
    
End Sub

Private Sub CmdSearch_Click()
  '  lblTopupStatus.Visible = False
    frmCASList.Source = 3
    frmCASList.Show vbModal
End Sub

Public Sub cmdShow_Click()
On Error GoTo ErrMsg

Dim getRecord As Integer

txtATTDoc(0).Clear
txtATTDoc(1).Clear
txtATTDoc(2).Clear
txtATTDoc(3).Clear
txtATTDoc(4).Clear


    SSTab1.Tab = 1
    RedAmt = 0
    REDAmtInd = False
    MBMCPLNCode = ""
    LGNo = ""
    LGVer = 0
    tmpPAYLGDesc = ""
    frmGLMenu.GLForm = False
    frmGLMenu.GLSource = 1
    
    'TxtLGRemark = ""
    tmpRemarks = ""
    lblFaxDocStatus.Visible = False
    Screen.MousePointer = vbHourglass

    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
    Else
        ScanCode = Trim(txtCaseNo)
        If Mid(ScanCode, 1, 2) = "LO" Then
            txtCaseNo = ""
            Call SearchForm(ScanCode)
        Else
                TempCASNo = txtCaseNo
                patient = cboPatientList
                Call ClearForm(Me)
                Call ClearEntries
                lstLGListing.ListItems.Clear
                lstWorksheetListing.ListItems.Clear
                cboPatientList.Clear
                txtCaseNo = TempCASNo
                MMARateCal = 0
                txtHOSName.Enabled = False
                getRecord = showCAS
                MMARateCal = 1
                If tmpMBMTermination = True Then
                    cmdLG.Enabled = False
                End If
                
                If tmpMBMGLIssuance = "N" Or tmpMBMStatus <> "A" Then
                    cmdLG.Enabled = False
                    MsgBox "GL is NOT allowed for this member", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                Call CheckFiles
            cmdUpdate1.Enabled = True
            If Trim(tmpUSRBlockGL) = "Y" Then
                cmdLG.Enabled = False
            End If
            ChkDOA.Value = 1
            ChkDOD.Value = 1
        End If
    End If
    Screen.MousePointer = Default
    Exit Sub
ErrMsg:
    MsgBox Err.Description
    Screen.MousePointer = Default
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
End Sub

Private Sub CheckFiles()
Dim B As String
List1.Clear

    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & ""
    '\Medical Rpt"
    B = Dir(PictureFileName & "\*.*")
    'b = Dir(Dir1.Path & "\*.*")
    If B = "" Then
        lblFaxDocStatus.Visible = True
    Else
        lblFaxDocStatus.Visible = False
        Exit Sub
    End If
End Sub
Private Function showCAS()
Dim CoverID As String
Dim HOS As New ADODB.Recordset
Dim strsqlHOS As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim RecordSaved
Dim StrSqlMsgProm As String
Dim MsgProm As New ADODB.Recordset
Dim tmpPLNPrinNoGL As Boolean
Dim tmpPLNSuppNoGL As Boolean
        
        
        tmpCasPendingMail = ""
        tmpCasRejectMail = ""
        tmpInitialLG = ""
        tmpFinalLG = ""
        tmpPLNPrinNoGL = False
        tmpPLNSuppNoGL = False
        CaseNumber = ""
        strAppend = ""

        sqlstrCAS = "Exec SearchCasAdjustment '" & Trim(txtCaseNo) & "'"
        Set CASSearch = New ADODB.Recordset
        CASSearch.Open sqlstrCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
    If Not CASSearch.EOF Then
        showCAS = 1
        With CASSearch
            If Mid(CASSearch!VIPInd, 1, 1) = "Y" Then
                MsgBox CASSearch!PLNSpecialNote, vbOKOnly + vbInformation
            End If
            cboOPChemo = IIf(Len(!CASOPChemo), !CASOPChemo, "")
            tmpCASLGSMS = IIf(Len(!CASLGSMS), !CASLGSMS, "")
            tmpCASPendingSMS = IIf(Len(!CASPendingSMS), !CASPendingSMS, "")
            tmpCASRejectSMS = IIf(Len(!CASRejectSMS), !CASRejectSMS, "")
            tmpCasPendingMail = IIf(Len(!CASPendingMail), !CASPendingMail, "")
            tmpCasRejectMail = IIf(Len(!CASRejectMail), !CASRejectMail, "")
            tmpInitialLG = IIf(Len(!InitialLGMail), !InitialLGMail, "")
            tmpFinalLG = IIf(Len(!FinalLGMail), !FinalLGMail, "")
            tmpCASCancelSMS = IIf(Len(!CASCancelSMS), !CASCancelSMS, "")
            tmpMBMEmail = IIf(Len(!MBMEmail), !MBMEmail, "")
            If Trim(!PLNPrincipalNoGL) = "Y" Then
                tmpPLNPrinNoGL = True
            Else
                tmpPLNPrinNoGL = False
            End If
            
            If Trim(!PLNSuppNoGL) = "Y" Then
                tmpPLNSuppNoGL = True
            Else
                tmpPLNSuppNoGL = False
            End If
            
            txtCaseNo = !CasNumber
            CaseNumber = Trim(!CasNumber)
            txtMBMNumber = Trim(!MBMNumber)
            If Len(!MBMCPLNCode) Then
                MBMCPLNCode = !MBMCPLNCode
            End If
            If !MBMPatientCovID = "00" Then
                cboPatientList = Trim(!MBMPatientCovID) & " - " & Trim(!MBMName)
                PatientName = Trim(!MBMName)
            Else
                cboPatientList = Trim(!MBMPatientCovID) & " - " & Trim(!MBMCName)
                cboSuppExc = Trim(!MBMPatientCovID) & " - " & Trim(!MBMCName)
                PatientName = Trim(!MBMCName)
            End If
            If IsDate(!CaseRegDate) = False Then
                CaseRegDate = ""
            Else
                CaseRegDate = !CaseRegDate
            End If
            
            'CaseRegDate = IIf(IsNull(!CaseRegDate), CDate(!CaseRegDate), "")
                txtDateAdmitted.mask = ""
                txtDateAdmitted.mask = "##/##/####"
            If Len(!CASDateAdmitted) Then
                txtDateAdmitted = Format(!CASDateAdmitted, "dd/mm/yyyy")
            End If
                txtDateDischarge.mask = ""
                txtDateDischarge.mask = "##/##/####"

            If Len(!CASDischargeDate) Then
                txtDateDischarge = Format(!CASDischargeDate, "dd/mm/yyyy")
            End If
            'If !CASQuery <> "" Then
            '    cboQuery = !CASQuery
            'End If
            
            'If !RoomBoardDay <> "" Then
            '    txtRoomNBoardDays = !RoomBoardDay
            'End If
            If Trim(!MBMTerminateStatus) = "Y" Then
                tmpMBMTermination = True
            Else
                tmpMBMTermination = False
            End If
            
            
            If !MBMPatientCovID = "00" Then
                tmpPatientDOB = !MBMDOB
            Else
                tmpPatientDOB = !MBMCDOB
            End If
            
            tmpMBMTelNoM = IIf(Len(!MBMTelNoM), !MBMTelNoM, "")
            
            If Len(!CASAuditReview) Then
                txtPrincipalExc = Replace(!CASAuditReview, "~", vbNewLine)
            Else
                txtPrincipalExc = ""
            End If
            
            If Len(!CASSpecialNote) Then
                txtSpecialRemark = Replace(!CASSpecialNote, "~", vbNewLine)
            Else
                txtSpecialRemark = ""
            End If
            
            
            If Len(!CASInvNote) Then
                txtInvRemark = Replace(!CASInvNote, "~", vbNewLine)
            Else
                txtInvRemark = ""
            End If
            

        strAppendCase = "6"
        strAppend = ""

        strAppend = " AND HOSCode = '" & Trim(!CASHOSCode) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        HOS.CursorLocation = adUseClient
        HOS.CursorType = adOpenForwardOnly
        HOS.CacheSize = 100
        Set HOS = cmd.Execute
        
            If HOS.EOF = False Then
                 If Len(!CASHOSCode) Then
                     txtHOSCode = HOS!HOSName
                 End If
            End If
        HOS.Close
        Set HOS = Nothing
            
            If Len(!CASStayType) Then
                cboStayType = !CASStayType
            End If

            If Len(!CASPAYType) Then
                CboPayType = !CASPAYType
            End If
            
            If Len(!CASAdmissionReason) Then
                txtReason = !CASAdmissionReason
            End If
            If Len(!ClaimNatureCode) Then
                cboNature = !ClaimNatureCode
            End If
            If Len(!ModalityCode) Then
                cboModality = !ModalityCode
            End If
            If Len(!ReferedBy) Then
                txtREFBy = !ReferedBy
            End If
            If Len(!CASReserveAmount) Then
                txtCASReserveAmount = !CASReserveAmount
            End If
            If Len(!CASLGAmt) Then
                txtCASLGAmt = !CASLGAmt
            End If
            If Trim(!CASStatus) = "O" Then
                cboStatus1.ListIndex = 0
            ElseIf Trim(!CASStatus) = "C" Then
                cboStatus1.ListIndex = 1
            ElseIf Trim(!CASStatus) = "D" Then
                cboStatus1.ListIndex = 2
            End If
            
            If Trim(!CASClaimability) = "AA" Then
                cboClaimability.Text = "AA - Accepted"
             ElseIf Trim(!CASClaimability) = "RR" Then
                cboClaimability.Text = "RR - Rejected"
             ElseIf Trim(!CASClaimability) = "CC" Then
                cboClaimability.Text = "CC - Cancelled"
             ElseIf Trim(!CASClaimability) = "SA" Then
                cboClaimability.Text = "SA - Special Approval"
             ElseIf Trim(!CASClaimability) = "PP" Then
                cboClaimability.Text = "PP - Pay & Claim"
             ElseIf Trim(!CASClaimability) = "NN" Then
                cboClaimability.Text = "NN - Not Decided"
             ElseIf Trim(!CASClaimability) = "NA" Then
                cboClaimability.Text = "NA - Not Decided/Accepted"
             ElseIf Trim(!CASClaimability) = "NR" Then
                cboClaimability.Text = "NR - Not Decided/Rejected"
            ElseIf Trim(!CASClaimability) = "PA" Then
                cboClaimability.Text = "PA - Pay&Claim/Accepted"
            ElseIf Trim(!CASClaimability) = "PR" Then
                cboClaimability.Text = "PR - Pay&Claim/Rejected"
            ElseIf Trim(!CASClaimability) = "AR" Then
                cboClaimability.Text = "AR - Accepted/Rejected"
            ElseIf Trim(!CASClaimability) = "RA" Then
                cboClaimability.Text = "RA - Rejected/Accepted"
            ElseIf Trim(!CASClaimability) = "BR" Then
                cboClaimability.Text = "BR - BTBG/Rejected"
            ElseIf Trim(!CASClaimability) = "ER" Then
                cboClaimability.Text = "ER - Expired/Rejected"
            ElseIf Trim(!CASClaimability) = "AI" Then
                cboClaimability.Text = "AI - Accepted/Insurance"
            'ElseIf Trim(!CASClaimability) = "DB" Then
            '    cboClaimability.Text = "DB - Death Benefit"
            End If
             
             If Trim(!CASPending) = "Y" Then
                cboPending.ListIndex = 0
             ElseIf Trim(!CASPending) = "N" Then
                cboPending.ListIndex = 1
             End If
            'If Len(!CASInd) Then
            '    cboCASInd = !CASInd
            'End If
            If Len(!LINKCASNo) Then
                TxtLinkCASNo = !LINKCASNo
            End If
            
            'If Len(!CASTopupLmt) Then
            '    cboTopUpLmt = !CASTopupLmt
            'End If
            
            txtCASTakeOverAmt = IIf(Len(!CASTakeOverAmt), !CASTakeOverAmt, "")
            
            cboCASTakeOveInd = IIf(Len(!CASTakeOverInd), !CASTakeOverInd, "")
            
            txtCASTOPlan = IIf(Len(!CASTakeOverPlan), !CASTakeOverPlan, "")
            
            If Len(!CASLGTime) Then
                txtLGTime = !CASLGTime
            End If
            
            If Trim(!PAYRepLetter) <> "" Then
                tmpRepStatus = !PAYRepLetter
            Else
                tmpRepStatus = ""
            End If
            txtMBMPolYr = IIf(Len(!CASPolYear), !CASPolYear, "")
            tmpProductName = IIf(Len(!ProductName), !ProductName, "")
            tmpPAYLGDesc = IIf(Len(!PAYLGDesc), !PAYLGDesc, "")
            
            If !PLNFolupDays <> "" Then
                tmpFolDays = !PLNFolupDays
            Else
                tmpFolDays = ""
            End If
            
            txtPolEffDate = IIf(Len(!MBMPolEffDate), Format(!MBMPolEffDate, "dd/mm/yyyy"), "__/__/____")
            
            tmpMBMGLIssuance = IIf(Len(!MBMGLIssuance), !MBMGLIssuance, "")
            
            tmpCASDocDate = IIf(Len(!CASDocDate), !CASDocDate, "")
            tmpCASDocType = IIf(Len(!CASDocType), !CASDocType, "")
            tmpCaseOpenBy = IIf(Len(!CaseOpenBy), !CaseOpenBy, "")
            'txtSpecialRemark
        End With
        
            Call showMBM
            MBMNumber = Trim(txtMBMNumber)
            Call CallLimit
            Call showCASRemarks
            Call showOther
            Call showOther2
            Call ShowFirstDiag
            Call ShowFinalDiag
            Call ShowUD
            Call ShowRepudiation
            Call showMBMListing
            Call showMBMCovListing
            Call BNFListing
            
            StrSqlMsgProm = ""
            If Trim(tmpHltCode) = "S" Then
                StrSqlMsgProm = "Select MessagePromt, PLNPolicyWording from PLN where PLNCode = '" & Trim(txtPlanCode) & "'"
            Else
                StrSqlMsgProm = "Select MessagePromt, PLNPolicyWording from PLN where PLNCode = '" & Trim(txtPlanCode) & "' And PAYCode = '" & Trim(txtPayCode) & "' "
            End If
            
            MsgProm.Open StrSqlMsgProm, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                Do While Not MsgProm.EOF
                    tmpPolicyWording = IIf(Len(MsgProm!PLNPolicyWording), MsgProm!PLNPolicyWording, "")
                    tmpMsgprom = IIf(Len(MsgProm!MessagePromt), MsgProm!MessagePromt, "")
                MsgProm.MoveNext
                Loop
            MsgProm.Close
            Set MsgProm = Nothing
            
            
            StrSqlMsgProm = ""
            StrSqlMsgProm = "Select * from tblRecord where CASENo   ='" & Trim(txtCaseNo) & "'"
            MsgProm.Open StrSqlMsgProm, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
                Do While Not MsgProm.EOF
                    cboPendingReason = IIf(Len(MsgProm!DocRemark), MsgProm!DocRemark, "")
                MsgProm.MoveNext
                Loop
            MsgProm.Close
            Set MsgProm = Nothing
            
            

            If tmpMsgprom <> "" Then
                MsgBox tmpMsgprom, vbOKOnly + vbInformation
            End If
            
        If Mid(cboStatus1, 1, 1) = "O" Then
            'cboStatus1.Enabled = False
            If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
                cmdRepudiate.Enabled = True
                cmdLG.Enabled = False
            Else
                cmdRepudiate.Enabled = False
                cmdLG.Enabled = True
            End If
            cmdSaveAdd.Enabled = True
            cmdUpdate1.Enabled = True
            txtDateDischarge.Enabled = True
            txtCASRemark.Enabled = True
            lblPage.Caption = "Page 1"
            lblPage1.FontBold = True
            lblPage2.FontBold = False
            lblPage3.FontBold = False
            lblPage4.FontBold = False
            lblPage5.FontBold = False
        ElseIf Mid(cboStatus1, 1, 1) = "C" Or Mid(cboStatus1, 1, 1) = "D" Then
            'cboStatus1.Enabled = False
            cmdLG.Enabled = False
            cmdRepudiate.Enabled = False
            cmdSaveAdd.Enabled = False
            cmdUpdate1.Enabled = False
            Frame6.Enabled = False
            txtDateDischarge.Enabled = False
            txtCASRemark.Enabled = False
        End If
        If Mid(cboClaimability, 1, 2) = "AA" And Mid(cboPending, 1, 1) = "Y" Then
            RecordSaved = MsgBox("Do you want to change Claimability Status?", vbYesNo + vbDefaultButton2 + vbExclamation)
            If RecordSaved = vbYes Then
                Call UpdateClaimability
                MsgBox "Status has been updated sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            End If
        End If
        
            
            If tmpPLNPrinNoGL = True And Mid(cboPatientList, 1, 2) = "00" Then
                cmdLG.Enabled = False
                MsgBox "No GL for Principal", vbOKOnly + vbInformation, DialogBoxTitle
            Else
                cmdLG.Enabled = True
            End If
            
            If tmpPLNSuppNoGL = True And Mid(cboPatientList, 1, 2) <> "00" Then
                cmdLG.Enabled = False
                MsgBox "No GL for Dependent", vbOKOnly + vbInformation, DialogBoxTitle
            Else
                cmdLG.Enabled = True
            End If
            
           ' StrSqlMsgProm = "Select MessagePromt,PLNPolicyWording,productname,PLNRBType,PLNDescription  from PLN where PLNCode = '" & Trim(txtPlanCode) & "'"
           ' MsgProm.Open StrSqlMsgProm, medixmasterconn, adOpenForwardOnly, adLockReadOnly
           ' If MsgProm.EOF = False Then
           '     If MsgProm!MessagePromt <> "" Then
           '         tmpMsgprom = MsgProm!MessagePromt
           '     End If
           ' End If
           ' MsgProm.Close
           ' Set MsgProm = Nothing
            
           ' If tmpMsgprom <> "" Then
           '     MsgBox tmpMsgprom, vbOKOnly + vbInformation
           ' End If
    Else
        cmdSaveAdd.Enabled = False
        cmdUpdate1.Enabled = False
        cmdLG.Enabled = False
        cmdRepudiate.Enabled = False
        MsgBox "No record found!", vbCritical + vbOKOnly
    End If
   
    CASSearch.Close
    Set CASSearch = Nothing
End Function

Private Function showCASRemarks()
    txtCASRemark = ""
    lblPage1.ForeColor = vbBlack
    lblPage2.ForeColor = vbBlack
    lblPage3.ForeColor = vbBlack
    lblPage4.ForeColor = vbBlack
    lblPage5.ForeColor = vbBlack
    If Not CASSearch.EOF Then
        If Len(CASSearch!CASORemarks) Then
            txtCASRemark = Replace(CASSearch!CASORemarks, "~", Chr(13))
            lblPage1.ForeColor = vbRed
        End If
        If Len(CASSearch!CASORemarksTwo) Then
            txtCASRemark = txtCASRemark + Replace(CASSearch!CASORemarksTwo, "~", Chr(13))
            lblPage1.ForeColor = vbRed
        End If
        
        If Len(CASSearch!CASOTwoRemarks) Then
            txtCASRemarkTwo = Replace(CASSearch!CASOTwoRemarks, "~", Chr(13))
            lblPage2.ForeColor = vbRed
        End If
        If Len(CASSearch!CASOTwoRemarksTwo) Then
            txtCASRemarkTwo = txtCASRemarkTwo + Replace(CASSearch!CASOTwoRemarksTwo, "~", Chr(13))
            lblPage2.ForeColor = vbRed
        End If
        
        If Len(CASSearch!Casremarks) Then
            txtCASRemark3 = Replace(CASSearch!Casremarks, "~", Chr(13))
            lblPage3.ForeColor = vbRed
        End If
        
        If Len(CASSearch!CASRemarksFour) Then
            txtCASRemark4 = Replace(CASSearch!CASRemarksFour, "~", Chr(13))
            lblPage4.ForeColor = vbRed
        End If
        
        If Len(CASSearch!CASRemarksFive) Then
            txtCASRemark5 = Replace(CASSearch!CASRemarksFive, "~", Chr(13))
            lblPage5.ForeColor = vbRed
        End If
        
        'If Len(CASSearch!CASORemarksInt) Then
        '    txtIntCASRemark = Replace(CASSearch!CASORemarksInt, "~", Chr(13))
        'End If
    End If
    
End Function

Private Function showMBM()
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim sqlstr1 As String
Dim strsqlMBMExc As String
Dim MBMCov As New ADODB.Recordset
Dim MBMExc As New ADODB.Recordset

    'sqlstr = " Select MBMNumber, HLTCode, MBMPolicyNo,  MBMName ," & _
    '        " MBMIcBcPp,  MBMMemberType , " & _
    '        " MBMStatus, PAYCode, AGECode , " & _
    '        " INSCode, PLNCode, " & _
    '        " MBMAddress1, MBMAddress2, " & _
    '        " MBMAddress3, MBMPostCode, CTYCode, STACode, " & _
    '        " MBMDOB , BRNCode, MBMTakeOver , MBMRenewFlag,  MBMAvailableLimit, " & _
    '        " MBMPayorEffDate, MBMPayorExpDate, MBMENDtEffDate , MBMNewPlanCode, MBMNewPolEffDate "
    'sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic

    'strsqlMBMExc = " Select MBMNumber, MBMExclusion From MBMOthersTwo Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    'MBMExc.Open strsqlMBMExc, medixmasterconn, adOpenKeyset, adLockOptimistic

    sqlstr = "Exec SearchMBMCas '" & Trim(txtMBMNumber) & "'"
    
    Set MBM = New ADODB.Recordset
    MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    If Not MBM.EOF Then
        With MBM
        
            If Trim(!PAYChkHosp) = "Y" Then
                tmpHospChk = "Y"
            Else
                tmpHospChk = ""
            End If
        
            If Len(!MBMNumber) Then
                txtMBMNumber = Trim(!MBMNumber)
            End If
            If Len(!MBMName) Then
                txtMBMName = Trim(!MBMName)
            End If
            If Len(!MBMPolicyNo) Then
                tmptxtMBMPolNo = Trim(!MBMPolicyNo)
            End If
            If Len(!MBMIcBcPp) Then
                txtMBMIcBcPP = Trim(!MBMIcBcPp)
            End If
            If Len(!MBMAddress1) Then
                txtAdd1 = Trim(!MBMAddress1)
            Else
                txtAdd1 = ""
            End If
            If Len(!MBMAddress2) Then
                txtAdd2 = Trim(!MBMAddress2)
            Else
                txtAdd2 = ""
            End If
            If Len(!MBMAddress3) Then
                txtAdd3 = Trim(!MBMAddress3)
            Else
                txtAdd3 = ""
            End If
            If Len(!CTYCode) Then
                txtCity = Trim(!CTYCode)
            Else
                txtCity = ""
            End If
            If Len(!MBMPostCode) Then
                txtPostCode = Trim(!MBMPostCode)
            Else
                txtPostCode = ""
            End If
            If Len(Trim(!STACode)) Then
                txtState = Trim(!STACode)
            Else
                txtState = ""
            End If
            
            If Len(Trim(!MBMAppName)) Then
                tmpMBMAppName = Trim(!MBMAppName)
            Else
                tmpMBMAppName = ""
            End If
            If Len(!PAYCode) Then
                txtPayCode = Trim(!PAYCode)
            End If
            
            If Len(Format(!MBMPayorEffDate, "dd/mm/yyyy")) Then
                txtMBMEffDate.mask = ""
                txtMBMEffDate.Text = Trim(Format(MBM!MBMPayorEffDate, "dd/mm/yyyy"))
                txtMBMEffDate.mask = "##/##/####"
            End If
            If Len(!HLTCode) Then
                tmpHltCode = Trim(!HLTCode)
            End If
            If Len(Format(!MBMPayorEXPDate, "dd/mm/yyyy")) Then
                txtMBMExpiryDate.mask = ""
                txtMBMExpiryDate = Trim(Format(!MBMPayorEXPDate, "dd/mm/yyyy"))
                txtMBMExpiryDate.mask = "##/##/####"
            End If
            
            If Len(!PLNCode) Then
                tmpPlnCode = Trim(!PLNCode)
            End If
            
            If Len(!INSCode) Then
                tmpINSCode = Trim(MBM!INSCode)
            End If
            If Len(Format(!MBMDOB, "dd/mm/yyyy")) Then
                tmpPrincipalDOB = Trim(Format(!MBMDOB, "dd/mm/yyyy"))
            Else
                tmpPrincipalDOB = ""
            End If
                
                txtPlanCode = Trim(!PLNCode)
                txtPolStatus = Trim(!MBMStatus)
                TmpExclusion = IIf(Len(!MBMExclusion), !MBMExclusion, "")
                cboPatientList.AddItem "00 -" & txtMBMName
                
                tmpMBMGroupCompany = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
                tmpMBMPymtMode = IIf(Len(!MBMInstallment), !MBMInstallment, "")
                tmpMBMEmpNo = IIf(Len(!MBMEmployeeNo), !MBMEmployeeNo, "")
                
                tmpAgentCode = IIf(Len(!MBMAgentCode), !MBMAgentCode, "")
                
                If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
                    Do While Not MBM.EOF
                        cboPatientList.AddItem MBM!MBMCCoverID & "-" & MBM!MBMCName
                        cboSuppExc.AddItem MBM!MBMCCoverID & "-" & MBM!MBMCName
                        MBM.MoveNext
                    Loop
                End If
            End With
          End If
    MBM.Close
    Set MBM = Nothing
End Function

Public Sub CallLimit()
Dim sqlstr As String
Dim AnnualLimit As String
Dim PLN As String
Dim INS As String
Dim HLT As String
Dim Effdate As Date
Dim MBM As New ADODB.Recordset
    'sqlstr = "Select PLNCode, INSCode, HLTCode, MBMPayorEffDate from MBM Where MBMNumber = '" & RemStr(txtMBMNumber) & "'"
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'PLN = Trim(MBM!PLNCode)
    'INS = Trim(MBM!INSCode)
    'HLT = Trim(MBM!HLTCode)
    
    Effdate = CDate(txtMBMEffDate)
    
    'txtAnnualLimit = ""
    'If Len(Trim(cboPLNCode)) Then
    '    txtAnnualLimit = Format(GetAnnualLimit(tmpINSCode, tmpPLNCode, tmpHltCode, Effdate), "#,##0.00")
    'End If
            
        'txtAnnualLimit = ""
        'txtAnnualLimit = Format(GetAnnualLimit(INS, PLN, HLT, EffDate), "#,##0.00")

End Sub

Private Function showOther()
Dim sqlstr As String
Dim Other As New ADODB.Recordset
Dim i As Integer
Dim CPTList As ListItem
Dim CPTStatus As Boolean
Dim ProcReq As String
Dim CPTCode As String
Dim ActSurgFees As Double
Dim ActAnaeFees As Double
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
Dim TotalActSurgFees As Double
Dim TotalActAnaeFees As Double
Dim TotalMMASurgFees As Double
Dim TotalMMAAnaeFees As Double
    
    'txtCASRemark = ""
    'txtCASAdmissionReason = ""
    'txtCASRepudiation = ""
    
    'sqlstr = "Select * From CASOthers where CASNumber = '" & RemStr(txtCaseNo) & "'"
    'Other.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TotalActSurgFees = 0
    TotalActAnaeFees = 0
    TotalMMASurgFees = 0
    TotalMMAAnaeFees = 0
    
    If Not CASSearch.EOF Then
        With CASSearch
            If Len(Trim(!CASOCondWillRecur)) Then
                 cboCondwillOcc = !CASOCondWillRecur
            End If
            If Len(Trim(!CASOFollowUpRequired)) Then
                 cboFLWTreatment = !CASOFollowUpRequired
            End If
            If Len(Trim(!CASFolUp)) Then
                 cboFolUp = !CASFolUp
            End If
            
            If Len(!CASFirstCall) Then
                txtFirstCall.mask = ""
                txtFirstCall = Format(!CASFirstCall, "dd/mm/yyyy")
                txtFirstCall.mask = "##/##/####"
            End If
            If Len(!CASFirstTreated) Then
                txtFirstTreated.mask = ""
                txtFirstTreated = Format(!CASFirstTreated, "dd/mm/yyyy")
                txtFirstTreated.mask = "##/##/####"
            End If
            If Len(!CASOPROCRequired1) Then
                txtProcedure1 = !CASOPROCRequired1
            End If
            
            If Len(!CASOPROCRequired2) Then
                txtProcedure2 = !CASOPROCRequired2
            End If

            If Len(!CASOPROCRequired3) Then
                 txtProcedure3 = !CASOPROCRequired3
            End If

            If Len(!CASOPROCRequired4) Then
               txtProcedure4 = !CASOPROCRequired4
            End If

            If Len(!CASOPROCRequired5) Then
                 txtProcedure5 = !CASOPROCRequired5
            End If

            If Len(!CASOPROCRequired6) Then
                 txtProcedure6 = !CASOPROCRequired6
            End If

            If Len(!CASOProcInvestigationDone1) Then
                 txtPROCINVDone1 = !CASOProcInvestigationDone1
            End If
            
            If Len(!CASOProcInvestigationDone2) Then
                 txtPROCINVDone2 = !CASOProcInvestigationDone2
            End If

            If Len(!CASOProcInvestigationDone3) Then
                 txtPROCINVDone3 = !CASOProcInvestigationDone3
            End If

            If Len(!CASOMedSurMgmt1) Then
                 txtMEDSURDone1 = !CASOMedSurMgmt1
            End If
            
            If Len(!CASOMedSurMgmt2) Then
                 txtMEDSURDone2 = !CASOMedSurMgmt2
            End If

            If Len(!CASOMedSurMgmt3) Then
                 txtMEDSURDone3 = !CASOMedSurMgmt3
            End If
            
            If Len(!CASDRD) Then
                txtDateRECDoc = CDate(!CASDRD)
            End If
            
            'If (!CASRedAmt) Then
            '    txtREDAmt = !CASRedAmt
            '    REDAmtInd = True
            '    RedAmt = !CASRedAmt
            'End If
            
            If Len(!CASCptCode1) Then
                txtProcCode1 = !CASCptCode1
                'Call SearchMMARate(!CASCptCode1)
                'txtSMMARate1 = MMASurgFees
                'txtAnaesMMARate1 = MMAAnaeFees
            End If
            
            If Len(!CASCptCode2) Then
                txtProcCode2 = !CASCptCode2
                'Call SearchMMARate(!CASCptCode2)
                'txtSMMARate2 = MMASurgFees
                'txtAnaesMMARate2 = MMAAnaeFees
            End If
            
            If Len(!CASCptCode3) Then
                txtProcCode3 = !CASCptCode3
                'Call SearchMMARate(!CASCptCode3)
                'txtSMMARate3 = MMASurgFees
                'txtAnaesMMARate3 = MMAAnaeFees
            End If
            
            If Len(!CASCptCode4) Then
                txtProcCode4 = !CASCptCode4
                'Call SearchMMARate(!CASCptCode4)
                'txtSMMARate4 = MMASurgFees
                'txtAnaesMMARate4 = MMAAnaeFees
            End If
            
            If Len(!CASCptCode5) Then
                txtProcCode5 = !CASCptCode5
                'Call SearchMMARate(!CASCptCode5)
                'txtSMMARate5 = MMASurgFees
                'txtAnaesMMARate5 = MMAAnaeFees
            End If
            
            If Len(!CASCptCode6) Then
                txtProcCode6 = !CASCptCode6
                'Call SearchMMARate(!CASCptCode5)
                'txtSMMARate5 = MMASurgFees
                'txtAnaesMMARate5 = MMAAnaeFees
            End If
            
            If Len(!CASInvCode1) Then
                txtINVCode1 = !CASInvCode1
            End If
            
            If Len(!CASInvCode2) Then
                txtINVCode2 = !CASInvCode2
            End If
            
            If Len(!CASInvCode3) Then
                txtINVCode3 = !CASInvCode3
            End If
            
            If Len(!CASPayPending) Then
                cboPayPending = !CASPayPending
            End If
            
            If Len(!CASONonSurgPROCRequired1) Then
                txtNSProcedure1(0) = !CASONonSurgPROCRequired1
            End If
            
            If Len(!CASONonSurgPROCRequired2) Then
               txtNSProcedure1(1) = !CASONonSurgPROCRequired2
            End If

            If Len(!CASONonSurgPROCRequired3) Then
                 txtNSProcedure1(2) = !CASONonSurgPROCRequired3
            End If

            If Len(!CASONonSurgPROCRequired4) Then
               txtNSProcedure1(3) = !CASONonSurgPROCRequired4
            End If

            If Len(!CASONonSurgPROCRequired5) Then
                 txtNSProcedure1(4) = !CASONonSurgPROCRequired5
            End If
            
            If Len(!CASONonSurgPROCRequired6) Then
                 txtNSProcedure1(5) = !CASONonSurgPROCRequired6
            End If
            
            If Len(!CASNonSurgCptCode1) Then
                txtNSProcCode1 = !CASNonSurgCptCode1
            End If
            
            If Len(!CASNonSurgCptCode2) Then
                txtNSProcCode2 = !CASNonSurgCptCode2
            End If
            
            If Len(!CASNonSurgCptCode3) Then
                txtNSProcCode3 = !CASNonSurgCptCode3
            End If
            
            If Len(!CASNonSurgCptCode4) Then
                txtNSProcCode4 = !CASNonSurgCptCode4
            End If
            
            If Len(!CASNonSurgCptCode5) Then
                txtNSProcCode5 = !CASNonSurgCptCode5
            End If
            
            If Len(!CASNonSurgCptCode6) Then
                txtNSProcCode6 = !CASNonSurgCptCode6
            End If
            
            txtSActFee1 = IIf(Len(!CASProcActSurgFees1), !CASProcActSurgFees1, 0)
            txtSActFee2 = IIf(Len(!CASProcActSurgFees2), !CASProcActSurgFees2, 0)
            txtSActFee3 = IIf(Len(!CASProcActSurgFees3), !CASProcActSurgFees3, 0)
            txtSActFee4 = IIf(Len(!CASProcActSurgFees4), !CASProcActSurgFees4, 0)
            txtSActFee5 = IIf(Len(!CASProcActSurgFees5), !CASProcActSurgFees5, 0)
            txtSActFee6 = IIf(Len(!CASProcActSurgFees6), !CASProcActSurgFees6, 0)
            
            txtAnaesActFees1 = IIf(Len(!CASProcActAnaeFees1), !CASProcActAnaeFees1, 0)
            txtAnaesActFees2 = IIf(Len(!CASProcActAnaeFees2), !CASProcActAnaeFees2, 0)
            txtAnaesActFees3 = IIf(Len(!CASProcActAnaeFees3), !CASProcActAnaeFees3, 0)
            txtAnaesActFees4 = IIf(Len(!CASProcActAnaeFees4), !CASProcActAnaeFees4, 0)
            txtAnaesActFees5 = IIf(Len(!CASProcActAnaeFees5), !CASProcActAnaeFees5, 0)
            txtAnaesActFees6 = IIf(Len(!CASProcActAnaeFees6), !CASProcActAnaeFees6, 0)
            
            txtNSActFee1 = IIf(Len(!CASProcActNonSurgFees1), !CASProcActNonSurgFees1, 0)
            txtNSActFee2 = IIf(Len(!CASProcActNonSurgFees2), !CASProcActNonSurgFees2, 0)
            txtNSActFee3 = IIf(Len(!CASProcActNonSurgFees3), !CASProcActNonSurgFees3, 0)
            txtNSActFee4 = IIf(Len(!CASProcActNonSurgFees4), !CASProcActNonSurgFees4, 0)
            txtNSActFee5 = IIf(Len(!CASProcActNonSurgFees5), !CASProcActNonSurgFees5, 0)
            txtNSActFee6 = IIf(Len(!CASProcActNonSurgFees6), !CASProcActNonSurgFees6, 0)
            
            txtPatientAge = IIf(Len(!CASPatientAge), !CASPatientAge, "")
            cboYearDay = IIf(Len(!CASPatientDayYr), !CASPatientDayYr, "")
        End With
    
    End If
End Function

Private Function showOther2()
    
    If Not CASSearch.EOF Then
        
        With CASSearch
            If Len(!CASAttendDoctor1) Then
                txtATTDoc(0) = Trim(!CASAttendDoctor1)
            End If
            If Len(!CASAttendDoctor2) Then
                txtATTDoc(1) = Trim(!CASAttendDoctor2)
            End If
            If Len(!CASAttendDoctor3) Then
                txtATTDoc(2) = Trim(!CASAttendDoctor3)
            End If
            If Len(!CASAttendDoctor4) Then
                txtATTDoc(3) = Trim(!CASAttendDoctor4)
            End If
            If Len(!CASAttendDoctor5) Then
                txtATTDoc(4) = Trim(!CASAttendDoctor5)
            End If
            If Len(!CASAnaesName1) Then
                txtAnaesDoc(0) = Trim(!CASAnaesName1)
            End If
            If Len(!CASAnaesName2) Then
                txtAnaesDoc(1) = Trim(!CASAnaesName2)
            End If
            If Len(!CASAnaesName3) Then
                txtAnaesDoc(2) = Trim(!CASAnaesName3)
            End If
                                
            If Len(!CASFirstDocDate) Then
                txtFirstDocDate = !CASFirstDocDate
            End If
            If Len(!CASSubSeqDocDate) Then
                txtSubDocDate = !CASSubSeqDocDate
            End If
            cboIntExtTO = IIf(Len(!CASIntExtTO), !CASIntExtTO, "")
            txtHOSName = IIf(Len(!CASHOSName), !CASHOSName, "")
                                
        End With
        
    End If
    
End Function

Private Sub cmdUpdate1_Click()
Dim strsqlCAS As String
Dim CAS As New ADODB.Recordset
Dim CASHIS As New ADODB.Recordset
Dim CASHOS As New ADODB.Recordset
Dim DIAHIS As New ADODB.Recordset
Dim DIA As New ADODB.Recordset
Dim AUT As New ADODB.Recordset
Dim sqlstr As String
Dim sqlstr1 As String
Dim diafirstsql As String
Dim diafinalsql As String
Dim diafirst As New ADODB.Recordset
Dim diafinal As New ADODB.Recordset
Dim listrecord As Integer
Dim listloop As Integer
Dim diagnosistr As String
Dim hoshistory As String
Dim RecordSaved
Dim AdjustmentSeqNo As Integer
Dim ListCount As Integer
Dim CaseNumber As String
Dim CASDRD As String
Dim strsqlRep As String
Dim REP As New ADODB.Recordset
Dim StrSqlCRef As String
Dim CASCRef As New ADODB.Recordset
Dim CASMBMNumber As String
Dim CASMBMID As String
Dim HOSCode As String
Dim CASRMK As New ADODB.Recordset
Dim strsqlRMK As String
Dim cmd As New ADODB.Command
Dim cmd1 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim Prm3 As ADODB.Parameter
Dim strAppendCase As String
Dim LINKCASNo, CASInd, DateAdmitted, DateDischarge, LogonDate
Dim CASRemark, CASRemarkTwo, Reason, RefBy, CASRemarkThree, CASRemarkFour, CASRemarkFive
Dim TempTakeOverLmt As Double
Dim strSearchRMK As String
Dim CASRemarksInt As String
Dim RecordSave
Dim tempdays, tempday1, tempday2 As Integer
Dim isSent As Boolean
Dim stringmessage As String
Dim CASSpecialNote As String
Dim CASInvNote  As String

    isSent = False
    UD1 = 0
    FirstDiag = 0
    FinalDiag = 0
    CASMBMNumber = ""
    CASMBMID = ""
    HOSCode = ""
    TempTakeOverLmt = 0
    CASSpecialNote = ""
    If IsDate(txtDateAdmitted) And IsDate(txtDateDischarge) Then
        tempdays = DateDiff("d", CDate(txtDateAdmitted), CDate(txtDateDischarge))
        tempday1 = tempdays + 1
        tempday2 = tempdays - 1
    End If
    If tempdays > 5 Then
        RecordSave = MsgBox("Admission Period Is More Than 5 Days.", vbYesNo + vbExclamation)
            If RecordSave = vbNo Then
                Exit Sub
            End If
    End If
    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No!", vbCritical + vbOKOnly, DialogBoxTitle
        txtCaseNo.SetFocus
    ElseIf cboClaimability = "" Then
        MsgBox "Please Select Claimability!", vbCritical + vbOKOnly, DialogBoxTitle
        cboClaimability.SetFocus
    ElseIf txtMBMNumber = "" Then
        MsgBox "Please Enter Member No!", vbCritical + vbOKOnly, DialogBoxTitle
        txtMBMNumber.SetFocus
    ElseIf txtMBMName = "" Then
        MsgBox "Please Enter Member Name!", vbCritical + vbOKOnly, DialogBoxTitle
        txtMBMName.SetFocus
    ElseIf txtDateAdmitted = "__/__/____" Then
        MsgBox "Please Enter Admission Date", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
        txtDateAdmitted.SetFocus
    ElseIf Mid(cboIntExtTO, 1, 1) = "I" And Trim(txtCASTOPlan) = "" Then
        MsgBox "Please Enter Takeover Plan", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASTOPlan.SetFocus
    ElseIf Trim(txtMBMPolYr) = "" Then
        MsgBox "Please Enter Policy Year", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMPolYr.SetFocus
    ElseIf Trim(txtPatientAge) = "" Then
        MsgBox "Please Enter Patient's Age", vbOKOnly + vbCritical, DialogBoxTitle
        txtPatientAge.SetFocus
    ElseIf Trim(cboYearDay) = "" Then
        MsgBox "Please Select Day/Year", vbOKOnly + vbCritical, DialogBoxTitle
        cboYearDay.SetFocus
    ElseIf (txtHOSCode = "DENTAL CLINIC" Or txtHOSCode = "DIALYSIS CENTRE" Or txtHOSCode = "FOREIGN HOSPITAL" Or txtHOSCode = "GENERAL PRACTITIONER" Or txtHOSCode = "GOVERNMENT HOSPITAL" Or txtHOSCode = "NON-PANEL HOSPITAL" Or txtHOSCode = "SPECIALIST CLINIC") And Trim(txtHOSName) = "" Then
        MsgBox "Please Enter Hospital Name or Clinic Name!", vbOKOnly + vbInformation
        txtHOSName.Enabled = True
    ElseIf cboREPCode1 = "" And Mid(cboClaimability, 1, 2) = "RR" Then
        MsgBox "Please Enter Repudiation Reason", vbOKOnly + vbCritical, DialogBoxTitle
        cboREPCode1.SetFocus
    'ElseIf cboREPCode1 = "" And Mid(cboClaimability, 1, 2) = "PP" Then
        'MsgBox "Please Enter Repudiation Reason", vbOKOnly + vbCritical, DialogBoxTitle
        'cboREPCode1.SetFocus
    ElseIf (txtPayCode = "XL" Or txtPayCode = "XB" Or txtPayCode = "XF") And cboOPChemo = "" And Mid(CboPayType, 1, 1) = "R" Then
        MsgBox "Please Select Outpatient Chemotheraphy Status", vbOKOnly + vbCritical, DialogBoxTitle
        cboOPChemo.SetFocus
    ElseIf (Mid(txtCaseNo, 1, 2) = "ET" Or Mid(txtCaseNo, 1, 2) = "EQ" Or Mid(txtCaseNo, 1, 2) = "ES" Or Mid(txtCaseNo, 1, 2) = "EP") And txtHPNumber = "" And cboPendingReason <> "" Then
        MsgBox "Please enter Patient's Mobile Number", vbOKOnly + vbCritical, DialogBoxTitle
        txtHPNumber.SetFocus
    Else
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                TempTakeOverLmt = IIf(IsNumeric(txtCASTakeOverAmt) = True, txtCASTakeOverAmt, 0)
                
                'AdjustmentSeqNo = GetAdjustmentNo(txtMBMNumber, txtMBMPolNo, txtCaseNo, "C") ' Get Adjustment Sequence No
                '---------------------------------------------------------------------------------------------------------------------------------------
                'Duplicate Copy for Case and put into Case History CAS -> CASHIS
    
                'updatecase = "Select * from CAS Where CASNumber ='" & Trim(tempcasNo) & "' And MBMNumber ='" & Trim(MBMNumber) & " '"
                'CAS.Open updatecase, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "7"
                strAppend = ""
                strAppend = " AND CASNumber ='" & Trim(TempCASNo) & "'"
                Set cmd.ActiveConnection = medixmasterconn
                cmd.CommandText = "Case_Adjustment"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd.Parameters.Append Prm1
                Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd.Parameters.Append Prm2
                CAS.CursorLocation = adUseClient
                CAS.CursorType = adOpenForwardOnly
                CAS.CacheSize = 100
                Set CAS = cmd.Execute
                With CAS
                    CASMBMNumber = !MBMNumber
                    CASMBMID = !MBMPatientCovID
                    'HOSCode = !CASHOSCode
                    HOSCode = HOSArr.Item(txtHOSCode)
                End With
            '----------7-----------------------------------------------------------------------------------------------------------------------------
            'Update the Cas Fields
            
                    If txtDateAdmitted <> "__/__/____" Then
                        DateAdmitted = Format(txtDateAdmitted, "mm/dd/yyyy")
                    Else
                        DateAdmitted = Empty
                    End If
                    If txtDateDischarge <> "__/__/____" Then
                        DateDischarge = Format(txtDateDischarge, "mm/dd/yyyy")
                    Else
                        DateDischarge = Empty
                    End If
                    If txtDateRECDoc <> "__/__/____" Then
                        CASDRD = Format(txtDateRECDoc, "mm/dd/yyyy")
                    Else
                        CASDRD = Empty
                    End If
                    If TxtLinkCASNo <> "" Then
                        LINKCASNo = ChkStr(TxtLinkCASNo)
                    Else
                        LINKCASNo = Null
                    End If
                    'If Len(cboCASInd) Then
                    '    CASInd = Mid(cboCASInd, 1, 1)
                    'Else
                        CASInd = Null
                    'End If
                    LogonDate = Format(GetServerTime, "mm/dd/yyyy")
                    CASRemark = ChkStr(Replace(txtCASRemark, "'", "''"))
                    CASRemark = ChkStr(Replace(CASRemark, Chr(13), "~"))
                    CASRemarkTwo = ChkStr(Replace(txtCASRemarkTwo, "'", "''"))
                    CASRemarkTwo = ChkStr(Replace(CASRemarkTwo, Chr(13), "~"))
                    CASRemarkThree = ChkStr(Replace(txtCASRemark3, "'", "''"))
                    CASRemarkThree = ChkStr(Replace(CASRemarkThree, Chr(13), "~"))
                    CASRemarkFour = ChkStr(Replace(txtCASRemark4, "'", "''"))
                    CASRemarkFour = ChkStr(Replace(CASRemarkFour, Chr(13), "~"))
                    CASRemarkFive = ChkStr(Replace(txtCASRemark5, "'", "''"))
                    CASRemarkFive = ChkStr(Replace(CASRemarkFive, Chr(13), "~"))
                    CASRemarksInt = ChkStr(Replace(CASRemarksInt, Chr(13), "~"))
                    If txtLGTime <> "__:__" Then
                        tmpLGTime = Format(txtLGTime, "hh:mm")
                    Else
                        tmpLGTime = Empty
                    End If
                    
                    CASSpecialNote = ChkStr(Replace(txtSpecialRemark, "'", "''"))
                    CASSpecialNote = ChkStr(Replace(txtSpecialRemark, Chr(13), "~"))
                    
                    CASInvNote = ChkStr(Replace(txtInvRemark, "'", "''"))
                    CASInvNote = ChkStr(Replace(txtInvRemark, Chr(13), "~"))
                    
                strsqlCAS = ""
                strsqlCAS = "EXEC Case_AdjUpdateCAS_2016_V2 '" & Trim(TempCASNo) & "', " & _
                            "'" & Mid(cboStatus1, 1, 1) & "', '" & Mid(cboClaimability, 1, 2) & "', '" & DateAdmitted & "', '" & DateDischarge & "', " & _
                            "'" & Mid(cboModality, 1, 1) & "', " & _
                            "'" & Mid(cboStayType, 1, 1) & "', '" & Mid(CboPayType, 1, 1) & "', " & _
                            "'" & Mid(cboNature, 1, 3) & "', '" & Mid(cboPending, 1, 1) & "', " & _
                            "" & txtCASReserveAmount & ", '" & LINKCASNo & "', " & _
                            "'" & CASInd & "', '" & Logon & "', '" & LogonDate & "', " & txtCASLGAmt & ", '', " & _
                            "'" & CASDRD & "', 'N', " & TempTakeOverLmt & ", '" & Mid(cboCASTakeOveInd, 1, 1) & "', '" & txtCASTOPlan & "'," & _
                            "" & 0 & ", '" & tmpLGTime & "', '" & txtMBMPolYr & "', '" & cboOPChemo & "', '" & CASSpecialNote & "', '" & CASInvNote & "'"
                medixmasterconn.Execute strsqlCAS
                
                strAppendCase = "8"
                strAppend = ""
                strAppend = " AND CASNumber ='" & Trim(TempCASNo) & "'"
                Set cmd1.ActiveConnection = medixmasterconn
                cmd1.CommandText = "Case_Adjustment"
                cmd1.CommandType = adCmdStoredProc
                cmd1.CommandTimeout = 300
                Set Prm1 = cmd1.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd1.Parameters.Append Prm1
                Set Prm2 = cmd1.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd1.Parameters.Append Prm2
                CASRMK.CursorLocation = adUseClient
                CASRMK.CursorType = adOpenForwardOnly
                CASRMK.CacheSize = 100
                Set CASRMK = cmd1.Execute
                
                If CASRMK.EOF = True Then
                    If Len(CASRemark) <= 1900 Then
                        strsqlRMK = "EXEC Case_AdjInsertCASRemark '" & Trim(TempCASNo) & "', '" & CASRemark & "', ''"
                    ElseIf Len(CASRemark) > 1900 Then
                        strsqlRMK = "EXEC Case_AdjInsertCASRemark '" & Trim(TempCASNo) & "', '" & Mid(CASRemark, 1, 1900) & "', '" & Mid(CASRemark, 1901, Len(CASRemark) - 1900) & "'"
                    End If
                Else
                    If Len(CASRemark) <= 1900 Then
                        strsqlRMK = "EXEC Case_AdjUpdateCASRemark '" & Trim(TempCASNo) & "', '" & CASRemark & "', ''"
                    ElseIf Len(CASRemark) > 1900 Then
                        strsqlRMK = "EXEC Case_AdjUpdateCASRemark '" & Trim(TempCASNo) & "', '" & Mid(CASRemark, 1, 1900) & "', '" & Mid(CASRemark, 1901, Len(CASRemark) - 1900) & "'"
                    End If
                End If
                medixmasterconn.Execute strsqlRMK
                CASRMK.Close
                Set CASRMK = Nothing
                strSearchRMK = "Select CASNumber from CASRemarksTwo Where CASNumber ='" & Trim(TempCASNo) & "'"
                CASRMK.Open strSearchRMK, medixmasterconn, adOpenKeyset, adLockOptimistic
                If CASRMK.recordCount Then
                    If Len(CASRemarkTwo) <= 1900 Then
                        strsqlRMK = "EXEC Case_AdjUpdateCASRemark2 '" & Trim(TempCASNo) & "', '" & CASRemarkTwo & "', ''"
                    ElseIf Len(CASRemarkTwo) > 1900 Then
                        strsqlRMK = "EXEC Case_AdjUpdateCASRemark2 '" & Trim(TempCASNo) & "', '" & Mid(CASRemarkTwo, 1, 1900) & "', '" & Mid(CASRemarkTwo, 1901, Len(CASRemarkTwo) - 1900) & "'"
                    End If
                Else
                    If Len(CASRemarkTwo) <= 1900 Then
                        strsqlRMK = "EXEC Case_AdjInsertCASRemark2 '" & Trim(TempCASNo) & "', '" & CASRemarkTwo & "', ''"
                    ElseIf Len(CASRemarkTwo) > 1900 Then
                        strsqlRMK = "EXEC Case_AdjInsertCASRemark2 '" & Trim(TempCASNo) & "', '" & Mid(CASRemarkTwo, 1, 1900) & "', '" & Mid(CASRemarkTwo, 1901, Len(CASRemarkTwo) - 1900) & "'"
                    End If
                    
                End If
                    medixmasterconn.Execute strsqlRMK
                CASRMK.Close
                Set CASRMK = Nothing
                
                strsqlRMK = ""
                strSearchRMK = "Select CASNumber from CASRemarkThree Where CASNumber ='" & Trim(TempCASNo) & "'"
                CASRMK.Open strSearchRMK, medixmasterconn, adOpenKeyset, adLockOptimistic
                If CASRMK.recordCount Then
                    strsqlRMK = "EXEC Case_AdjUpdateCASRemark3 '" & Trim(TempCASNo) & "', '" & CASRemarkThree & "', '" & CASRemarkFour & "', '" & CASRemarkFive & "'"
                Else
                    strsqlRMK = "EXEC Case_AdjInsertCASRemark3 '" & Trim(TempCASNo) & "', '" & CASRemarkThree & "', '" & CASRemarkFour & "', '" & CASRemarkFive & "'"
                End If
                    medixmasterconn.Execute strsqlRMK
                CASRMK.Close
                Set CASRMK = Nothing
                
                strAppendCase = "3"
                strAppend = ""
                strAppend = " AND CASNumber ='" & Trim(TempCASNo) & "'"
                Set cmd2.ActiveConnection = medixmasterconnMaint
                cmd2.CommandText = "Case_Adjustment"
                cmd2.CommandType = adCmdStoredProc
                cmd2.CommandTimeout = 300
                Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd2.Parameters.Append Prm1
                Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd2.Parameters.Append Prm2
                CASCRef.CursorLocation = adUseClient
                CASCRef.CursorType = adOpenForwardOnly
                CASCRef.CacheSize = 100
                Set CASCRef = cmd2.Execute
                
                        If txtDateAdmitted <> "__/__/____" Then
                            DateAdmitted = Format(txtDateAdmitted, "mm/dd/yyyy")
                        Else
                            DateAdmitted = Empty
                        End If
                        If txtDateDischarge <> "__/__/____" Then
                            DateDischarge = Format(txtDateDischarge, "mm/dd/yyyy")
                        Else
                            DateDischarge = Empty
                        End If
                        If TxtLinkCASNo <> "" Then
                            LINKCASNo = ChkStr(TxtLinkCASNo)
                        Else
                            LINKCASNo = Null
                        End If
                    
                    StrSqlCRef = "'" & Trim(TempCASNo) & "', '" & Mid(cboStatus1, 1, 1) & "', " & _
                                "'" & Mid(cboClaimability, 1, 2) & "', '" & DateAdmitted & "', '" & DateDischarge & "', " & _
                                "'" & Mid(cboPending, 1, 1) & "', '" & LINKCASNo & "', '', " & _
                                "'" & CASMBMNumber & "', '" & CASMBMID & "', '" & Mid(cboPayPending, 1, 1) & "'"
                    
                    If CASCRef.EOF = True Then
                        StrSqlCRef = "EXEC Case_AdjInsertCASRef " + StrSqlCRef
                    Else
                        StrSqlCRef = "EXEC Case_AdjUpdateCASRef " + StrSqlCRef
                    End If
                medixmasterconnMaint.Execute StrSqlCRef
                Call SaveFirstDiag
                Call SaveFinalDiag
                Call SaveUD
                Call SaveCASOthers
                Call SaveCASOthers2
                Call saveHos
               ' Call UpdateBalLmt
                Call UpdateLG
                Dim UPRS As New ADODB.Recordset
                
                
                If Mid(txtCaseNo, 1, 2) = "AY" And txtDateAdmitted > "03/08/2015" And (Mid(cboClaimability, 1, 2) = "PP" Or Mid(cboClaimability, 1, 2) = "RR") Then
                    strsqlCASLG = ""
                    strsqlCASLG = "SELECT        dbo.RepCodeCategory.RepCatCode, dbo.RepCodeCategory.RepCatName, dbo.RepCode.RepCatCode AS Expr1, dbo.RepCode.RepCode FROM dbo.RepCode INNER JOIN  dbo.RepCodeCategory ON dbo.RepCode.RepCatCode = dbo.RepCodeCategory.RepCatCode  WHERE dbo.RepCode.RepCode  = '" & Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - IIf(InStr(1, cboREPCode1, "-") = 0, 0, 1))) & "'"
                    UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                        
                        If UPRS.EOF = False Then
                            tmpRejectCat = IIf(Len(UPRS!RepCatName), UPRS!RepCatName, "")
                        Else
                            tmpRejectCat = ""
                        End If
                    UPRS.Close
                    Set UPRS = Nothing
                    
                    If tmpCASRejectSMS <> "1" Then
                       If tmpMBMTelNoM <> "" Then
                            Call SendSMSReject
                            isSent = True
                        End If
                        strsqlCASLG = ""
                        strsqlCASLG = " update CAS set  CASLGSMS = '0', CASPendingSMS = '0', CASRejectSMS = '1',CASCancelSMS = '0' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                        UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                        tmpCASLGSMS = 1
                    End If
                End If
                
                If Mid(txtCaseNo, 1, 2) = "AY" And (Mid(cboClaimability, 1, 2) = "CC") And txtDateAdmitted > "03/08/2015" Then
                    If tmpCASCancelSMS <> "1" Then
                       If tmpMBMTelNoM <> "" Then
                            Call SendSMSCancel
                            isSent = True
                        End If
                        strsqlCASLG = ""
                        strsqlCASLG = " update CAS set  CASLGSMS = '0', CASPendingSMS = '0', CASRejectSMS = '0', CASCancelSMS = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                        UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                        tmpCASCancelSMS = 1
                    End If
                End If
                
                
                If Mid(txtCaseNo, 1, 2) = "MX" And (Mid(cboClaimability, 1, 2) = "PP") Then
                    'Code For MAA Takafule
                    If tmpCasPendingMail <> "1" Then
                        Call MailAdmPendingNotification("Norlaila@maa.my;Farihan@maa.my;Zamri@maa.my;DrKasthureeBai@maa.my")
                        strsqlCASLG = ""
                        strsqlCASLG = " update CAS set  CASPendingMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                        UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                        isSent = True
                    End If
                Else
                    'code For MRCB
                    If Trim(tmpMBMGroupCompany) = Trim("MALAYSIAN RESOURCES CORPORATION BERHAD") Then
                        If tmpCasPendingMail <> "1" Then
                            Call MailAdmPendingNotification("rozaina@mrcb.com.my;")
                            strsqlCASLG = ""
                            strsqlCASLG = " update CAS set  CASPendingMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                            UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                            isSent = True
                        End If
                    ElseIf Trim(tmpMBMGroupCompany) = Trim("PADU") Then
                        If tmpCasPendingMail <> "1" Then
                            Call MailAdmPendingNotification("rozaina@mrcb.com.my;")
                            strsqlCASLG = ""
                            strsqlCASLG = " update CAS set  CASPendingMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                            UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                            isSent = True
                        End If
                    End If
                    
                End If
                
                If Mid(txtCaseNo, 1, 2) = "MX" And (Mid(cboClaimability, 1, 2) = "RR") Then
                    'code For Maa
                    If tmpCasRejectMail <> "1" Then
                        Call MailAdmNotifiReject("Norlaila@maa.my;Farihan@maa.my;Zamri@maa.my;DrKasthureeBai@maa.my")
                        'Call MaaRejectNotification
                        strsqlCASLG = ""
                        strsqlCASLG = " update CAS set  CASRejectMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                        UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                        isSent = True
                    End If
                Else
                    'code For MRCB
                    If Trim(tmpMBMGroupCompany) = Trim("MALAYSIAN RESOURCES CORPORATION BERHAD") Then
                        If tmpCasRejectMail <> "1" Then
                            Call MailAdmPendingNotification("rozaina@mrcb.com.my;")
                            strsqlCASLG = ""
                            strsqlCASLG = " update CAS set  CASPendingMail = '1' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                            UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                            isSent = True
                        End If
                    End If
                End If
                
                
                If txtPayCode = "AY" And Trim(Mid(cboClaimability, 2, 2)) = Trim("P") Then
                    Call AYPendingNotify
                    isSent = True
                End If
                 
                If Mid(cboClaimability, 1, 2) <> "AA" Then
                    Call SaveRepudiation
                    StrSqlCRef = ""
                    StrSqlCRef = " update tblRecord set " & _
                             " IsRejected = '1', IsLGIssued = '0' , CaseProgress = 'Case declined'" & _
                             " Where CaseNo = '" & Trim(txtCaseNo) & "'"
                    UPRS.Open StrSqlCRef, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
                End If
                StrSqlCRef = ""
                StrSqlCRef = " update tblRecord set " & _
                         " DocRemark = '" & cboPendingReason & "'" & _
                         " Where CaseNo = '" & Trim(txtCaseNo) & "'"
                UPRS.Open StrSqlCRef, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
                CAS.Close
                Set CAS = Nothing
                CASCRef.Close
                Set CASCRef = Nothing
                cmdUpdate1.Enabled = False
                Dim X As Integer
                X = 0
                
                   Dim IC, ICnew As String
                   IC = txtPatientIC
                   If Len(IC) = 12 Then
                        ICnew = "******" + "-" + "**" + "-" + Right(IC, 4)
                    Else
                        For X = 0 To Len(IC) - 5
                            ICnew = ICnew + "*"
                        Next
                        
                        ICnew = ICnew + Right(IC, 4)
                        
                   End If
                
                If Len(cboPendingReason) Then
                   If Mid(txtCaseNo, 1, 2) = "ET" Or Mid(txtCaseNo, 1, 2) = "EQ" Or Mid(txtCaseNo, 1, 2) = "ES" Or Mid(txtCaseNo, 1, 2) = "EP" Then
                        If tmpCASPendingSMS <> "1" Then
                            tmpMBMTelNoM = txtHPNumber
                            
                            If tmpMBMTelNoM <> "" Then
                                stringmessage = "ETIQA: " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & " " & ICnew & " " & " Incomplete Document. We are liasing with hospital. GL Pending. " & Date & "," & Time & " TQ"
                                'tmpMBMTelNoM = "0123652738"
                                'tmpMBMTelNoM = "0123652738,0128829232,0122982285,0195494513"
                                Call SendSMS(stringmessage)
                            End If
                            
                            strsqlCASLG = ""
                            strsqlCASLG = " update CAS set  CASLGSMS = '0', CASPendingSMS = '1', CASRejectSMS = '0' Where CASNumber = '" & Trim(txtCaseNo) & "'"
                            UPRS.Open strsqlCASLG, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                            Set UPRS = Nothing
                            tmpCASPendingSMS = 1
                        End If
                    End If
                
                End If
                
                MsgBox "Record Updated Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            End If

            cmdUpdate1.Enabled = True
            
        End If
        
        
    If isSent Then
        Call cmdShow_Click
    End If
    
            
        cmdShow.Enabled = True
        'cmdLG.Enabled = True
        cmdRepudiate.Enabled = True
        cmdSaveAdd.Enabled = True
        cmdPrior.Enabled = True
        cmdPrev.Enabled = True
        CmdNext.Enabled = True
        cmdUpdate1.Enabled = True
        
    '    medixmasterconn.Close
    '    Set medixmasterconn = Nothing
    '    medixmasterconnMaint.Close
    '    Set medixmasterconnMaint = Nothing
End Sub

Public Sub AYPendingNotify()
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object
On Error GoTo debugs
Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "Notification of Pending Admission " & "-" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
objMessage.From = "notifications@medix.com.my"
objMessage.To = "claims@axa-life.com.my"
objMessage.TextBody = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
                      "NOTIFICATION OF PENDING ADMISSION" & vbCrLf & vbCrLf & _
                      "Please be advised that the above-mentioned member is admitted to hospital. Below is the case details for your kind attention." & vbCrLf & vbCrLf & _
                      "Patient Name             :       " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbCrLf & _
                      "Policy Number            :       " & tmptxtMBMPolNo & vbCrLf & _
                      "LG Number                :       " & LGNumber & vbCrLf & _
                      "Date of Admission      :       " & txtDateAdmitted & vbCrLf & _
                      "Hospital                       :       " & txtHOSCode & vbCrLf & _
                      "Please contact the us at 603-7803 2003 if you require any further clarification." & vbCrLf & vbCrLf & _
                      "Thank you." & vbCrLf & _
                      "Regards, " & vbCrLf & _
                      "MediExpress (Malaysia) Sdn Bhd" & vbCrLf & vbCrLf
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "notifications@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "medix123"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update
objMessage.Send
debugs:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

'Public Sub AYPendingNotify()
'Dim oOApp As Outlook.Application
'Dim oOMail As Outlook.MailItem'

'        Set oOApp = CreateObject("Outlook.Application")
'        Set oOMail = oOApp.CreateItem(olMailItem)
'        With oOMail
'            '.To = "nurshahirah.ng@axa-life.com.my"
'            .To = "claims@axa-life.com.my"
'            '.BCC = "bryan@medix.com.my;siewping@medix.com.my"
'            .Subject = "Notification of Pending Admission " & "-" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
'            .Body = "Dear Sir/Madam," & vbNewLine & vbNewLine & _
'                    "NOTIFICATION OF PENDING ADMISSION" & vbNewLine & vbNewLine & _
'                    "Please be advised that the above-mentioned member is admitted to hospital. Below is the case details for your kind attention." & vbNewLine & vbNewLine & _
'                    "Patient Name             :       " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbNewLine & _
'                    "Policy Number            :       " & tmptxtMBMPolNo & vbNewLine & _
'                    "LG Number                :       " & LGNumber & vbNewLine & _
'                    "Date of Admission      :       " & txtDateAdmitted & vbNewLine & _
'                    "Hospital                       :       " & txtHOSCode & vbNewLine & _
'                    vbNewLine & vbNewLine & _
'                    "Please contact the us at 603-7803 2003 if you require any further clarification." & vbNewLine & vbNewLine & _
'                    vbNewLine & vbNewLine & _
'                    "Thank you." & vbNewLine & _
'                    "Regards, " & vbNewLine & _
'                    "MediExpress (Malaysia) Sdn Bhd"
'        .Send
'        End With
'End Sub
Public Sub MailAdmInitialLG(ToAddress As String, BCC As String, CC As String)

'get PLNPolicny  Name
Dim PayorPlan As String
Dim sqlstrP As String
Dim PLNRS As New ADODB.Recordset
sqlstrP = "Select top 1   PLN.PLNPayorPLNDesc,PLN.PLNPayorPlan, PLN.PLNPayorCode ,PLN.PLNCode   from PLN inner join MBM on MBM.PLNCode = PLN.PLNCode and MBM.PAYCode = PLN.PAYCode   Where MBM.MBMNumber ='" + txtMBMNumber + "'"
PLNRS.Open sqlstrP, medixmasterconn, adOpenForwardOnly, adLockOptimistic
Do While Not PLNRS.EOF
    If Len(PLNRS!PLNPayorPLNDesc) Then
        PayorPlan = PLNRS!PLNPayorPLNDesc
    Else
        PayorPlan = PLNRS!PLNCode
    End If
PLNRS.MoveNext
Loop
PLNRS.Close
Set PLNRS = Nothing

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object
On Error GoTo debugs
Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "Notification of Initial LG Issued: " & "-" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
'objMessage.From = "adm@medix.com.my"

If Len(sFilePath) Then
objMessage.AddAttachment sFilePath
End If

objMessage.From = "notifications@medix.com.my"
'objMessage.To = "trinath@medix.com.my"
objMessage.To = ToAddress
objMessage.BCC = BCC
objMessage.CC = CC
objMessage.TextBody = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
                            "NOTIFICATION FOR INITIAL LG ISSUED: " & vbCrLf & vbCrLf & _
                            "Please be advised that the above-mentioned member is admitted to hospital." & vbCrLf & vbCrLf & _
                            "Below is the case details for your kind attention." & vbCrLf & vbCrLf & _
                            "Patient Name             :       " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbCrLf & vbCrLf & _
                            "Employee Name            :       " & txtMBMName & vbCrLf & vbCrLf & _
                            "Policy Number            :       " & tmptxtMBMPolNo & vbCrLf & vbCrLf & _
                            "Benefit Plan             :       " & PayorPlan & vbCrLf & vbCrLf & _
                            "Initial GL Number        :       " & LGNumber & vbCrLf & vbCrLf & _
                            "Date of Admission        :       " & txtDateAdmitted & vbCrLf & vbCrLf & _
                            "Hospital                 :       " & txtHOSCode & vbCrLf & vbCrLf & _
                            "Admitting Diagnosis      :       " & txtReason & vbCrLf & vbCrLf & _
                            "Action by                :       " & Logon & vbCrLf & vbCrLf & _
                            " " & vbCrLf & vbCrLf & _
                            "Please contact the us at 603-7884 1818 if you require any further clarification." & vbCrLf & vbCrLf & _
                            " " & vbCrLf & vbCrLf & _
                            "Thank you." & vbCrLf & vbCrLf & _
                            "Regards, " & vbCrLf & vbCrLf & _
                            "MediExpress (Malaysia) Sdn Bhd" & vbCrLf & vbCrLf
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "notifications@medix.com.my"
'objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "adm@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "medix123"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update
objMessage.Send
debugs:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub MailEtiqaAdmLG(ToAddress As String)
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object
On Error GoTo debugs
Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "Notification of LG Issued: " & "-" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & "( PF No : " & tmpMBMEmpNo & " )"
objMessage.From = "etiqa@medix.com.my"
objMessage.To = ToAddress
'objMessage.BCC = "bryan@medix.com.my"
'objMessage.BCC = "network.mgt@etiqa.com.my;etiqaglemail@medix.com.my"
objMessage.BCC = "etiqaglemail@medix.com.my"
objMessage.AddAttachment DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"


'1800229988

objMessage.TextBody = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
                            "We are pleased to inform you that the letter of guarantee has been issued. Below is the case details for your kind attention." & vbCrLf & vbCrLf & _
                            "Patient Name             :       " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbCrLf & vbCrLf & _
                            "Policy Number            :       " & tmptxtMBMPolNo & vbCrLf & vbCrLf & _
                            "LG Number                :       " & LGNumber & vbCrLf & vbCrLf & _
                            "Date of Visit            :       " & txtDateAdmitted & vbCrLf & vbCrLf & _
                            "Hospital                 :       " & txtHOSCode & vbCrLf & vbCrLf & _
                            " " & vbCrLf & vbCrLf & _
                            "Please contact us at 1800889998 if you require any further clarification." & vbCrLf & vbCrLf & _
                            " " & vbCrLf & vbCrLf & _
                            "Thank you." & vbCrLf & vbCrLf & _
                            "Regards, " & vbCrLf & vbCrLf & _
                            "Etiqa Healthcare" & vbCrLf & vbCrLf
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "etiqa@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "etiqa0903"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update
objMessage.Send
debugs:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub MailAdmPendingNotification(ToAddresss As String)
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object
On Error GoTo debugs
Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "Notification of Pending Admission " & "-" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
objMessage.From = "notifications@medix.com.my"
'objMessage.From = "adm@medix.com.my"
objMessage.To = ToAddresss
'objMessage.BCC = "bryan@medix.com.my;siewping@medix.com.my;"
objMessage.TextBody = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
                            "NOTIFICATION OF PENDING ADMISSION " & vbCrLf & vbCrLf & _
                            "Please be advised that the above-mentioned member is admitted to hospital. Below is the case details for your kind attention." & vbCrLf & vbCrLf & _
                            "Patient Name             :       " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbCrLf & vbCrLf & _
                            "Policy Number            :       " & tmptxtMBMPolNo & vbCrLf & vbCrLf & _
                            "LG Number                :       " & LGNumber & vbCrLf & vbCrLf & _
                            "Date of Admission      :       " & txtDateAdmitted & vbCrLf & vbCrLf & _
                            "Hospital                       :       " & txtHOSCode & vbCrLf & vbCrLf & _
                            " " & vbCrLf & vbCrLf & _
                            "Please contact the us at 603-7803 2003 if you require any further clarification." & vbCrLf & vbCrLf & _
                            " " & vbCrLf & vbCrLf & _
                            "Thank you." & vbCrLf & vbCrLf & _
                            "Regards, " & vbCrLf & vbCrLf & _
                            "MediExpress (Malaysia) Sdn Bhd" & vbCrLf & vbCrLf
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "notifications@medix.com.my"
'objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "adm@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "medix123"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update
objMessage.Send
debugs:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub
Public Sub MailAdmNotifiReject(ToAddress As String)
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object
On Error GoTo debugs

        Set objMessage = CreateObject("CDO.Message")
        objMessage.Subject = "Notification of Declaine Admission: " & "-" & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
        'objMessage.From = "adm@medix.com.my"
        objMessage.From = "notifications@medix.com.my"
        objMessage.To = ToAddress
        'objMessage.BCC = "bryan@medix.com.my;siewping@medix.com.my;"
        objMessage.TextBody = "Dear Sir/Madam," & vbCrLf & vbCrLf & _
        "NOTIFICATION  OF DECLAINE ADMISSION: " & vbCrLf & vbCrLf & _
        "Please be advised that the above-mentioned member is admitted to hospital. Below is the case details for your kind attention." & vbCrLf & vbCrLf & _
        "Patient Name             :       " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))) & vbCrLf & vbCrLf & _
        "Policy Number            :       " & tmptxtMBMPolNo & vbCrLf & vbCrLf & _
        "LG Number                :       " & LGNumber & vbCrLf & vbCrLf & _
        "Date of Admission      :       " & txtDateAdmitted & vbCrLf & vbCrLf & _
        "Hospital                       :       " & txtHOSCode & vbCrLf & vbCrLf & _
        " " & vbCrLf & vbCrLf & _
        "Please contact the us at 603-7803 2003 if you require any further clarification." & vbCrLf & vbCrLf & _
        " " & vbCrLf & vbCrLf & _
        "Thank you." & vbCrLf & vbCrLf & _
        "Regards, " & vbCrLf & vbCrLf & _
        "MediExpress (Malaysia) Sdn Bhd" & vbCrLf & vbCrLf

objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "notifications@medix.com.my"
'objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "adm@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "medix123"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update
objMessage.Send
debugs:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub SaveFirstDiag()
On Error Resume Next
Dim diafirst As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strsqlDIAFirst As String
Dim DIADOS1, DIADOS2, DIADOS3
Dim PrelimDiag1, PrelimDiag2, PrelimDiag3
    strAppendCase = "9"
    strAppend = ""
    strAppend = " AND CasNumber = '" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    diafirst.CursorLocation = adUseClient
    diafirst.CursorType = adOpenForwardOnly
    diafirst.CacheSize = 100
    Set diafirst = cmd.Execute
        
                If txtPDDOS(0) <> "__/__/____" Then
                    DIADOS1 = Format(Trim(txtPDDOS(0)), "mm/dd/yyyy")
                    txtPDDOS(0).mask = "##/##/####"
                Else
                    DIADOS1 = Empty
                End If
    
        '======================== Bug ========================
                If txtPDDOS(1) <> "__/__/____" Then
                    DIADOS2 = Format(Trim(txtPDDOS(1)), "mm/dd/yyyy")
                Else
                    DIADOS2 = Empty
                End If
        '======================== Bug =========================
        '        !DIAPrelimDiag2 = Trim(ChkStr(txtPrelimDiag2))
        '    If Trim(Mid(cboICDCode3, 1, 3)) <> "" Then
        '        !ICDCode3 = Mid(cboICDCode3, 1, 3)
        '    End If
                If txtPDDOS(2) <> "__/__/____" Then
                    DIADOS3 = Format(Trim(txtPDDOS(2)), "mm/dd/yyyy")
                Else
                    DIADOS3 = Empty
                End If
        '        !DIAPrelimDiag3 = Trim(ChkStr(txtPrelimDiag3))
        'diafirst.Update
        'End With
        PrelimDiag1 = ChkStr(Replace(txtPrelimDiag(0), "'", "''"))
        PrelimDiag2 = ChkStr(Replace(txtPrelimDiag(1), "'", "''"))
        PrelimDiag3 = ChkStr(Replace(txtPrelimDiag(2), "'", "''"))
        
        strsqlDIAFirst = "'" & Trim(txtCaseNo) & "', '" & Mid(cboICDCode(0), 1, 3) & "', '" & Mid(cboICDCode(1), 1, 3) & "', " & _
                        "'" & Mid(cboICDCode(2), 1, 3) & "', '" & Trim(PrelimDiag1) & "', '" & Trim(PrelimDiag2) & "', " & _
                        "'" & Trim(PrelimDiag3) & "', '" & DIADOS1 & "', '" & DIADOS2 & "', '" & DIADOS3 & "'"
        
    If diafirst.EOF = True Then
        strsqlDIAFirst = "EXEC Case_AdjInsertDIAFirst " + strsqlDIAFirst
    Else
        strsqlDIAFirst = "EXEC Case_AdjUpdateDIAFirst " + strsqlDIAFirst
    End If
    medixmasterconn.Execute strsqlDIAFirst
    diafirst.Close
    Set diafirst = Nothing
    
End Sub

Private Sub SaveRepudiation()
Dim REP As New ADODB.Recordset
Dim CAS As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cnn As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strsqlRep As String
Dim strsqlCAS As String
Dim LogonDate

Dim StrArr() As String

Dim RIPcode1 As String
Dim RIPcode2 As String
Dim RIPcode3 As String

RIPcode1 = ""
RIPcode2 = ""
RIPcode3 = ""

If Len(cboREPCode1) Then
    If InStr(cboREPCode1, "||") Then
        
        StrArr = Split(cboREPCode1, "||")
        RIPcode1 = StrArr(1)
    Else
        RIPcode1 = IIf(Len(cboREPCode1), Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - IIf(InStr(1, cboREPCode1, "-") = 0, 0, 1))), "")
    End If
End If

If Len(cboREPCode2) Then
    If InStr(cboREPCode2, "||") Then
        StrArr = Split(cboREPCode2, "||")
        RIPcode2 = StrArr(1)
    Else
        RIPcode2 = IIf(Len(cboREPCode2), Trim(Mid(cboREPCode2, 1, InStr(1, cboREPCode2, "-") - IIf(InStr(1, cboREPCode2, "-") = 0, 0, 1))), "")
    End If
End If

If Len(cboREPCode3) Then
    If InStr(cboREPCode3, "||") Then
        StrArr = Split(cboREPCode3, "||")
        RIPcode3 = StrArr(1)
    Else
        RIPcode3 = IIf(Len(cboREPCode3), Trim(Mid(cboREPCode3, 1, InStr(1, cboREPCode3, "-") - IIf(InStr(1, cboREPCode3, "-") = 0, 0, 1))), "")
    End If
End If

    ' INSERT INTO CASRepudiation TABLE
    strAppendCase = "10"
    strAppend = " AND CASNumber = '" & Trim(TempCASNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    REP.CursorLocation = adUseClient
    REP.CursorType = adOpenForwardOnly
    REP.CacheSize = 100
    Set REP = cmd.Execute
    strAppendCase = "7"
    strAppend = ""
    strAppend = " AND CASNumber = '" & Trim(txtCaseNo) & "'"
    Set cnn.ActiveConnection = medixmasterconn
    cnn.CommandText = "Case_Adjustment"
    cnn.CommandType = adCmdStoredProc
    cnn.CommandTimeout = 300
    Set Prm1 = cnn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cnn.Parameters.Append Prm1
    Set Prm2 = cnn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cnn.Parameters.Append Prm2
    CAS.CursorLocation = adUseClient
    CAS.CursorType = adOpenForwardOnly
    CAS.CacheSize = 100
    Set CAS = cnn.Execute
    LogonDate = Format(GetServerTime, "mm/dd/yyyy")
            strsqlRep = "'" & Trim(TempCASNo) & "', '" & RIPcode1 & "', " & _
                        "'" & RIPcode2 & "', '" & RIPcode3 & "', " & _
                        "'" & Logon & "', '" & LogonDate & "', '" & IIf(Len(Document), IIf(InStr(1, Document, "-") = 0, Trim(Document), Trim(Mid(Document, 1, InStr(1, Document, "-") - IIf(InStr(1, Document, "-") = 0, 0, 1)))), "") & "', " & _
                        "'" & IIf(Len(Personnel), IIf(InStr(1, Personnel, "-") = 0, Trim(Personnel), Trim(Mid(Personnel, 1, InStr(1, Personnel, "-") - IIf(InStr(1, Personnel, "-") = 0, 0, 1)))), "") & "', " & _
                        "'" & IIf(Len(Attention), IIf(InStr(1, Attention, "-") = 0, Trim(Attention), Trim(Mid(Attention, 1, InStr(1, Attention, "-") - IIf(InStr(1, Attention, "-") = 0, 0, 1)))), "") & "', " & _
                        "'" & CASUnderlying & "','" & tmpAgent & "','" & repremarks & "'"
            
            strsqlCAS = "EXEC Case_AdjUpdateCAS1 '" & Trim(txtCaseNo) & "', '" & Mid(cboClaimability, 1, 2) & "'"
            medixmasterconn.Execute strsqlCAS
            
        If REP.EOF = True Then
            strsqlRep = "EXEC Case_AdjInsertCASRep1 " + strsqlRep
        Else
            strsqlRep = "EXEC Case_AdjUpdateCASRep2016 " + strsqlRep
        End If
        medixmasterconn.Execute strsqlRep
    REP.Close
    CAS.Close
    Set REP = Nothing
    Set CAS = Nothing
End Sub

Private Sub SaveFinalDiag()
On Error Resume Next
Dim diafinal As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strsqlDIAFinal As String
Dim LogonDate
Dim DIADOS1, DIADOS2, DIADOS3 As String
Dim FinalDiag1, FinalDiag2, FinalDiag3
    
TotalFinalDiag = 0

' INSERT INTO DIAFinal TABLE
    'strsql = "select *  from DIAFinal where casNumber ='" & txtCaseNo & "'"
    'diafinal.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "11"
    strAppend = ""
    strAppend = " AND casNumber ='" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    diafinal.CursorLocation = adUseClient
    diafinal.CursorType = adOpenForwardOnly
    diafinal.CacheSize = 100
    Set diafinal = cmd.Execute
    
    LogonDate = Format(GetServerTime, "mm/dd/yyyy")
    FinalDiag1 = ChkStr(Replace(txtFinalDiag1(0), "'", "''"))
    FinalDiag2 = ChkStr(Replace(txtFinalDiag1(1), "'", "''"))
    FinalDiag3 = ChkStr(Replace(txtFinalDiag1(2), "'", "''"))
    
    'If diafinal.EOF = True Then
        'diafinal.AddNew
        'diafinal!CasNumber = ChkStr(txtCaseNo)
    'End if
    
     'With diafinal
     '    !DIALastAdjustmentUser = Logon
     '    !DIALastAdjustmentDate = Date
         
         If cboFDICDCode1 <> "" Then
     '       !ICDCode1 = Mid(ChkStr(cboFDICDCode1), 1, 3)
            TotalFinalDiag = CDbl(TotalFinalDiag) + 1
         End If
         If txtFDDOS1 <> "__/__/____" Then
             DIADOS1 = Format(txtFDDOS1, "mm/dd/yyyy")
         Else
             DIADOS1 = Empty
         End If
     '    !DIAFinalDiag1 = ChkStr(txtFinalDiag1)
         If cboFDICDCode2 <> "" Then
     '       !ICDCode2 = Mid(ChkStr(cboFDICDCode2), 1, 3)
            TotalFinalDiag = CDbl(TotalFinalDiag) + 1
         End If
         If txtFDDOS2 <> "__/__/____" Then
             DIADOS2 = Format(txtFDDOS2, "mm/dd/yyyy")
         Else
             DIADOS2 = Empty
         End If
     '    !DIAFinalDiag2 = ChkStr(txtFinalDiag2)
         If cboFDICDCode3 <> "" Then
     '       !ICDCode3 = Mid(ChkStr(cboFDICDCode3), 1, 3)
            TotalFinalDiag = CDbl(TotalFinalDiag) + 1
         End If
         If txtFDDOS3 <> "__/__/____" Then
             DIADOS3 = Format(txtFDDOS3, "mm/dd/yyyy")
         Else
             DIADOS3 = Empty
         End If
      '   !DIAFinalDiag3 = ChkStr(txtFinalDiag3)
     'End With
     'diafinal.Update
    strsqlDIAFinal = "'" & ChkStr(txtCaseNo) & "', '" & Mid(ChkStr(cboFDICDCode1), 1, 3) & "', " & _
                        "'" & Mid(ChkStr(cboFDICDCode2), 1, 3) & "', '" & Mid(ChkStr(cboFDICDCode3), 1, 3) & "', '" & FinalDiag1 & "', " & _
                        "'" & FinalDiag2 & "', '" & FinalDiag3 & "', '" & DIADOS1 & "', '" & DIADOS2 & "', '" & DIADOS3 & "', " & _
                        "'" & Logon & "', '" & LogonDate & "'"
    
   
    If diafinal.EOF = True Then
        strsqlDIAFinal = "EXEC Case_AdjInsertDIAFinal " + strsqlDIAFinal
    Else
        strsqlDIAFinal = "EXEC Case_AdjUpdateDIAFinal " + strsqlDIAFinal
    End If
    medixmasterconn.Execute strsqlDIAFinal
    diafinal.Close
    Set diafinal = Nothing

End Sub

Private Sub SaveUD()
On Error Resume Next
Dim CASUD As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strsqlCASUD As String
Dim CASUDDOS1, CASUDDOS2, CASUDDOS3
Dim UD1, UD2, UD3
    
    'strsql = "select * from CAS_UD where CASNumber = '" & Trim(tempcasNo) & "'"
    'CASUD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "12"
    strAppend = ""

    strAppend = " AND CASNumber = '" & Trim(TempCASNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    CASUD.CursorLocation = adUseClient
    CASUD.CursorType = adOpenForwardOnly
    CASUD.CacheSize = 100
    Set CASUD = cmd.Execute
        
    'If CASUD.EOF = True Then
        'CASUD.AddNew
        'CASUD!CasNumber = ChkStr(txtCaseNo)
    'End If
        'strsqlCASUD = "EXEC Case_AdjUpdateUD '" & Trim(tempcasNo) & "', '" & Mid(cboUDICDCode1, 1, 3) & "', '" & Mid(cboUDICDCode2, 1, 3) & "', " & _
                    "'" & Mid(cboUDICDCode3, 1, 3) & "', '" & ChkStr(txtUD1) & "', '" & ChkStr(txtUD2) & "', '" & ChkStr(txtUD3) & "', " & _
                    "'" & txtUDDOS1 & "', '" & txtUDDOS2 & "', '" & txtUDDOS3 & "'"
    'With CASUD
    '    If Trim(cboUDICDCode1) <> "" Then
    '        !CASUDCode1 = Mid(cboUDICDCode1, 1, 3)
    '    End If
        If txtUDDOS1 <> "__/__/____" Then
            CASUDDOS1 = Format(txtUDDOS1, "mm/dd/yyyy")
        Else
            CASUDDOS1 = Empty
        End If
    '        !CASUDReasons1 = ChkStr(txtUD1)
    '    If Trim(cboUDICDCode2) <> "" Then
    '        !CASUDCode2 = Mid(cboUDICDCode2, 1, 3)
    '    End If
        If txtUDDOS2 <> "__/__/____" Then
            CASUDDOS2 = Format(txtUDDOS2, "mm/dd/yyyy")
        Else
            CASUDDOS2 = Empty
        End If
    '    !CASUDReasons2 = txtUD2
    '    If Trim(cboUDICDCode3) <> "" Then
    '        !CASUDCode3 = Mid(cboUDICDCode3, 1, 3)
    '    End If
        If txtUDDOS3 <> "__/__/____" Then
            CASUDDOS3 = Format(txtUDDOS3, "mm/dd/yyyy")
        Else
            CASUDDOS3 = Empty
        End If
    '        !CASUDReasons3 = ChkStr(txtUD3)
    'CASUD.Update
    'End With
    UD1 = ChkStr(Replace(txtUD1(0), "'", "''"))
    UD2 = ChkStr(Replace(txtUD2, "'", "''"))
    UD3 = ChkStr(Replace(txtUD3, "'", "''"))
    
    strsqlCASUD = "'" & Trim(TempCASNo) & "', '" & Mid(cboUDICDCode1, 1, 3) & "', '" & Mid(cboUDICDCode2, 1, 3) & "', " & _
                    "'" & Mid(cboUDICDCode3, 1, 3) & "', '" & UD1 & "', '" & UD2 & "', '" & UD3 & "', " & _
                    "'" & CASUDDOS1 & "', '" & CASUDDOS2 & "', '" & CASUDDOS3 & "'"
    
    If CASUD.EOF = True Then
        strsqlCASUD = "EXEC Case_AdjInsertUD " + strsqlCASUD
    Else
        strsqlCASUD = "EXEC Case_AdjUpdateUD " + strsqlCASUD
    End If
    medixmasterconn.Execute strsqlCASUD
    CASUD.Close
    Set CASUD = Nothing
End Sub

Private Sub SaveCASOthers()
Dim sqlstrCASOthers As String
Dim Other As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim CASRedAmt, FirstCall, FirstTreated As String
Dim strsqlCASOthers As String
Dim MEDSURDone1, MEDSURDone2, MEDSURDone3
Dim PROCINVDone1, PROCINVDone2, PROCINVDone3
Dim PROCRequired1, PROCRequired2, PROCRequired3, PROCRequired4, PROCRequired5, PROCRequired6
Dim CASCptCode1, CASCptCode2, CASCptCode3, CASCptCode4, CASCptCode5, CASCptCode6 As String
Dim CASONonSurgPROCRequired1, CASONonSurgPROCRequired2, CASONonSurgPROCRequired3, CASONonSurgPROCRequired4, CASONonSurgPROCRequired5, CASONonSurgPROCRequired6 As String
Dim CASNonSurgCptCode1, CASNonSurgCptCode2, CASNonSurgCptCode3, CASNonSurgCptCode4, CASNonSurgCptCode5, CASNonSurgCptCode6 As String
    
    TotalProc = 0
    
    'sqlstr = "Select * From CASOthers where CASNumber = '" & RemStr(txtCaseNo) & "'"
    'Other.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "13"
    strAppend = ""

    strAppend = " AND CASNumber = '" & RemStr(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    Other.CursorLocation = adUseClient
    Other.CursorType = adOpenForwardOnly
    Other.CacheSize = 100
    Set Other = cmd.Execute


        If txtFirstCall <> "__/__/____" Then
            FirstCall = Format(txtFirstCall, "mm/dd/yyyy")
        Else
            FirstCall = Empty
        End If
        
        If txtFirstTreated <> "__/__/____" Then
             FirstTreated = Format(txtFirstTreated, "mm/dd/yyyy")
        Else
             FirstTreated = Empty
        End If
        
        If txtProcedure1 <> "" Then
            TotalProc = CDbl(TotalProc) + 1
        End If

        If txtProcedure2 <> "" Then
            TotalProc = CDbl(TotalProc) + 1
        End If

        If txtProcedure3 <> "" Then
            TotalProc = CDbl(TotalProc) + 1
        End If
        
        If txtProcedure4 <> "" Then
            TotalProc = CDbl(TotalProc) + 1
        End If
            
        If txtProcedure5 <> "" Then
            TotalProc = CDbl(TotalProc) + 1
        End If
            
        'If IsNumeric(txtREDAmt) = True Then
        '    CASRedAmt = txtREDAmt
        'Else
            CASRedAmt = 0
        'End If
        
        MEDSURDone1 = ChkStr(Replace(txtMEDSURDone1, "'", "''"))
        MEDSURDone2 = ChkStr(Replace(txtMEDSURDone2, "'", "''"))
        MEDSURDone3 = ChkStr(Replace(txtMEDSURDone3, "'", "''"))
        PROCINVDone1 = ChkStr(Replace(txtPROCINVDone1, "'", "''"))
        PROCINVDone2 = ChkStr(Replace(txtPROCINVDone2, "'", "''"))
        PROCINVDone3 = ChkStr(Replace(txtPROCINVDone3, "'", "''"))
        
        PROCRequired1 = ChkStr(Replace(txtProcedure1, "'", "''"))
        PROCRequired2 = ChkStr(Replace(txtProcedure2, "'", "''"))
        PROCRequired3 = ChkStr(Replace(txtProcedure3, "'", "''"))
        PROCRequired4 = ChkStr(Replace(txtProcedure4, "'", "''"))
        PROCRequired5 = ChkStr(Replace(txtProcedure5, "'", "''"))
        PROCRequired6 = ChkStr(Replace(txtProcedure6, "'", "''"))
        
        CASCptCode1 = IIf(Len(txtProcCode1), txtProcCode1, "")
        CASCptCode2 = IIf(Len(txtProcCode2), txtProcCode2, "")
        CASCptCode3 = IIf(Len(txtProcCode3), txtProcCode3, "")
        CASCptCode4 = IIf(Len(txtProcCode4), txtProcCode4, "")
        CASCptCode5 = IIf(Len(txtProcCode5), txtProcCode5, "")
        CASCptCode6 = IIf(Len(txtProcCode6), txtProcCode6, "")
        
        CASONonSurgPROCRequired1 = ChkStr(Replace(txtNSProcedure1(0), "'", "''"))
        CASONonSurgPROCRequired2 = ChkStr(Replace(txtNSProcedure1(1), "'", "''"))
        CASONonSurgPROCRequired3 = ChkStr(Replace(txtNSProcedure1(2), "'", "''"))
        CASONonSurgPROCRequired4 = ChkStr(Replace(txtNSProcedure1(3), "'", "''"))
        CASONonSurgPROCRequired5 = ChkStr(Replace(txtNSProcedure1(4), "'", "''"))
        CASONonSurgPROCRequired6 = ChkStr(Replace(txtNSProcedure1(5), "'", "''"))
        
        CASNonSurgCptCode1 = IIf(Len(txtNSProcCode1), txtNSProcCode1, "")
        CASNonSurgCptCode2 = IIf(Len(txtNSProcCode2), txtNSProcCode2, "")
        CASNonSurgCptCode3 = IIf(Len(txtNSProcCode3), txtNSProcCode3, "")
        CASNonSurgCptCode4 = IIf(Len(txtNSProcCode4), txtNSProcCode4, "")
        CASNonSurgCptCode5 = IIf(Len(txtNSProcCode5), txtNSProcCode5, "")
        CASNonSurgCptCode6 = IIf(Len(txtNSProcCode6), txtNSProcCode6, "")
        
        strsqlCASOthers = "'" & ChkStr(txtCaseNo) & "', " & _
                        "'" & PROCINVDone1 & "', '" & PROCINVDone2 & "', '" & PROCINVDone3 & "', " & _
                        "'" & MEDSURDone1 & "', '" & MEDSURDone2 & "', '" & MEDSURDone3 & "', " & _
                        "'" & Mid(cboCondwillOcc, 1, 1) & "', '" & Mid(cboFLWTreatment, 1, 1) & "', '" & FirstCall & "', " & _
                        "'" & FirstTreated & "', " & CASRedAmt & ", " & _
                        "'" & ChkStr(txtINVCode1) & "','" & ChkStr(txtINVCode2) & "', '" & ChkStr(txtINVCode3) & "', '" & Mid(cboFolUp, 1, 1) & "', '" & Mid(cboPayPending, 1, 1) & "', " & _
                        "'" & PROCRequired1 & "', '" & PROCRequired2 & "', '" & PROCRequired3 & "', '" & PROCRequired4 & "', '" & PROCRequired5 & "', '" & PROCRequired6 & "'," & _
                        "'" & CASCptCode1 & "', '" & CASCptCode2 & "', '" & CASCptCode3 & "', '" & CASCptCode4 & "', '" & CASCptCode5 & "','" & CASCptCode6 & "', " & _
                        "'" & CASONonSurgPROCRequired1 & "', '" & CASONonSurgPROCRequired2 & "', '" & CASONonSurgPROCRequired3 & "', '" & CASONonSurgPROCRequired4 & "', '" & CASONonSurgPROCRequired5 & "', '" & CASONonSurgPROCRequired6 & "'," & _
                        "'" & CASNonSurgCptCode1 & "', '" & CASNonSurgCptCode2 & "', '" & CASNonSurgCptCode3 & "', '" & CASNonSurgCptCode4 & "', '" & CASNonSurgCptCode5 & "', '" & CASNonSurgCptCode6 & "','" & txtPatientAge & "', '" & Mid(cboYearDay, 1, 1) & "'"
        

    
    If Other.EOF = True Then
        strsqlCASOthers = "EXEC Case_AdjInsertCASOthers " + strsqlCASOthers
    Else
        strsqlCASOthers = "EXEC Case_AdjUpdateCASOthers " + strsqlCASOthers
    End If
    
    medixmasterconn.Execute strsqlCASOthers
    
    Other.Close
    Set Other = Nothing
End Sub

Private Sub SaveCASOthers2()
Dim CASOthersTwo As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim TotalDoctor As String
Dim TotalAnaes As String
Dim TempDRName As String
Dim TempDRName2 As String
Dim TempDRName3 As String
Dim TempDRName4 As String
Dim TempDRName5 As String
Dim TempAnaesName As String
Dim TempAnaesName2 As String
Dim TempAnaesName3 As String
Dim strsqlCASOthers2 As String
Dim Reason, RefBy As String
Dim TmpFirstDocDate As String
Dim TmpSubSeqDocDate As String
Dim TmpIntExtTO As String
Dim TmpCASDOATime As String
Dim TmpCASDODTime As String
Dim TmpCASDOALGTime As String
Dim TmpCASDODLGTime As String
Dim TmpCASDOAComment As String
Dim TmpCASDODComment As String
Dim TmpCASComplaintDate As String
Dim TmpCASComplaintIssue As String
Dim TmpCASComplainant As String
Dim TmpCASCompAction As String
Dim tmpHospName As String

TotalDoctor = 0
TotalAnaes = 0

TempDRName = Replace(txtATTDoc(0), "'", "''")
TempDRName2 = Replace(txtATTDoc(1), "'", "''")
TempDRName3 = Replace(txtATTDoc(2), "'", "''")
TempDRName4 = Replace(txtATTDoc(3), "'", "''")
TempDRName5 = Replace(txtATTDoc(4), "'", "''")
TempAnaesName = Replace(txtAnaesDoc(0), "'", "''")
TempAnaesName2 = Replace(txtAnaesDoc(1), "'", "''")
TempAnaesName3 = Replace(txtAnaesDoc(2), "'", "''")
Reason = ChkStr(Replace(txtReason, "'", "''"))
RefBy = ChkStr(Replace(txtREFBy, "'", "''"))
    If Trim(txtHOSName) <> "" Then
        tmpHospName = Replace(txtHOSName, "'", "''")
    Else
        tmpHospName = ""
    End If
    ' INSERT INTO CASOthersTwo TABLE
    'CASOthersTwo.Open "Select CASNumber, * from CASOthersTwo where CASNumber = '" & Trim(txtCaseNo) & "'", medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "14"
    strAppend = ""

    strAppend = " AND CASNumber = '" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    CASOthersTwo.CursorLocation = adUseClient
    CASOthersTwo.CursorType = adOpenForwardOnly
    CASOthersTwo.CacheSize = 100
    Set CASOthersTwo = cmd.Execute
        
    'If CASOthersTwo.EOF = True Then
        'CASOthersTwo.AddNew
        'CASOthersTwo!CasNumber = ChkStr(txtCaseNo)
    'End If
        
        'With CASOthersTwo
            If TempDRName <> "" Then
        '        !CASAttendDoctor1 = ChkStr(TempDRName)
                TotalDoctor = CDbl(TotalDoctor) + 1
        '    Else
        '        !CASAttendDoctor1 = Null
            End If
            If TempDRName2 <> "" Then
        '        !CASAttendDoctor2 = ChkStr(TempDRName2)
                TotalDoctor = CDbl(TotalDoctor) + 1
        '    Else
        '        !CASAttendDoctor2 = Null
            End If
            If TempDRName3 <> "" Then
        '        !CASAttendDoctor3 = ChkStr(TempDRName3)
                TotalDoctor = CDbl(TotalDoctor) + 1
        '    Else
        '        !CASAttendDoctor3 = Null
            End If
            If TempDRName4 <> "" Then
        '        !CASAttendDoctor4 = ChkStr(TempDRName4)
                TotalDoctor = CDbl(TotalDoctor) + 1
        '    Else
        '        !CASAttendDoctor4 = Null
            End If
            If TempDRName5 <> "" Then
        '        !CASAttendDoctor5 = ChkStr(TempDRName5)
                TotalDoctor = CDbl(TotalDoctor) + 1
        '    Else
        '        !CASAttendDoctor5 = Null
            End If
            If TempAnaesName <> "" Then
        '        !CASAnaesName1 = ChkStr(TempAnaesName)
                TotalAnaes = CDbl(TotalAnaes) + 1
        '    Else
        '        !CASAnaesName1 = Null
            End If
            If TempAnaesName2 <> "" Then
        '        !CASAnaesName2 = ChkStr(TempAnaesName2)
                TotalAnaes = CDbl(TotalAnaes) + 1
        '    Else
        '        !CASAnaesName2 = Null
            End If
            If TempAnaesName3 <> "" Then
        '        !CASAnaesName3 = ChkStr(TempAnaesName3)
                TotalAnaes = CDbl(TotalAnaes) + 1
        '    Else
        '        !CASAnaesName3 = Null
            End If
                
            If txtFirstDocDate <> "__/__/____" Then
                TmpFirstDocDate = Format(txtFirstDocDate, "mm/dd/yyyy")
            Else
                TmpFirstDocDate = Empty
            End If
        
            If txtSubDocDate <> "__/__/____" Then
                TmpSubSeqDocDate = Format(txtSubDocDate, "mm/dd/yyyy")
            Else
                TmpSubSeqDocDate = Empty
            End If
            TmpIntExtTO = IIf(Len(cboIntExtTO), Mid(cboIntExtTO, 1, 1), "")
                
            'If txtDOAComment <> "__/__/____" Then
            '        TmpCASDOAComment = Format(txtDOAComment, "mm/dd/yyyy")
            'Else
                TmpCASDOAComment = Empty
            'End If
            
           ' If txtDODComment <> "__/__/____" Then
           '     TmpCASDODComment = Format(txtDODComment, "mm/dd/yyyy")
           ' Else
                TmpCASDODComment = Empty
           ' End If
            
            'If txtComplaintDate <> "__/__/____" Then
           '     TmpCASComplaintDate = Format(txtComplaintDate, "mm/dd/yyyy")
            'Else
                TmpCASComplaintDate = Empty
            'End If
            
            'If txtDOATime <> "__:__" Then
            '        TmpCASDOATime = Format(txtDOATime, "HH:MM")
            'Else
                TmpCASDOATime = Empty
            'End If
            
            'If txtDODTime <> "__:__" Then
            '    TmpCASDODTime = Format(txtDODTime, "HH:MM")
            'Else
                TmpCASDODTime = Empty
            'End If

            'If txtDOALGTime <> "__:__" Then
            '    TmpCASDOALGTime = Format(txtDOALGTime, "HH:MM")
            'Else
                TmpCASDOALGTime = Empty
            'End If
            
            'If txtDODLGTime <> "__:__" Then
            '    TmpCASDODLGTime = Format(txtDODLGTime, "HH:MM")
            'Else
                TmpCASDODLGTime = Empty
            'End If

            
        '    !CASTotalDoctors = TotalDoctor
        '    !CASTotalAnaes = TotalAnaes
        '    !CASTotalFinalDiag = TotalFinalDiag
        '    !CASTotalProc = TotalProc
            
        'End With
        'CASOthersTwo.Update
        
        strsqlCASOthers2 = "'" & ChkStr(txtCaseNo) & "', '" & ChkStr(TempDRName) & "', " & _
                        "'" & ChkStr(TempDRName2) & "', '" & ChkStr(TempDRName3) & "', '" & ChkStr(TempDRName4) & "', " & _
                        "'" & ChkStr(TempDRName5) & "', '" & ChkStr(TempAnaesName) & "', '" & ChkStr(TempAnaesName2) & "', " & _
                        "'" & ChkStr(TempAnaesName3) & "', '" & TotalDoctor & "', '" & TotalAnaes & "', '" & TotalFinalDiag & "', " & _
                        "'" & TotalProc & "', '" & Replace(Reason, "'", "''") & "', '" & RefBy & "', " & _
                        "'" & TmpFirstDocDate & "', '" & TmpSubSeqDocDate & "', '" & TmpIntExtTO & "', '" & _
                        TmpCASDOATime & "','" & TmpCASDODTime & "','" & TmpCASDODLGTime & "','" & TmpCASDODLGTime & "','" & _
                        TmpCASDOAComment & "','" & TmpCASDODComment & "','" & TmpCASComplaintDate & "','" & _
                        TmpCASComplaintIssue & "','" & TmpCASComplainant & "','" & TmpCASCompAction & "', '" & tmpHospName & "','" & Replace(Replace(txtPrincipalExc, "'", "''"), vbNewLine, "~") & "'"
                        
    If CASOthersTwo.EOF = True Then
        'strsqlCASOthers2 = "EXEC Case_AdjInsertCASOthers2 " + strsqlCASOthers2
        strsqlCASOthers2 = "Case_AdjInsertCASOthers_2015 " + strsqlCASOthers2
        
    Else
        'strsqlCASOthers2 = "EXEC Case_AdjUpdateCASOthers2 " + strsqlCASOthers2
        strsqlCASOthers2 = "Case_AdjUpdateCASOthers_2015 " + strsqlCASOthers2
    End If
    medixmasterconn.Execute strsqlCASOthers2
    CASOthersTwo.Close
    Set CASOthersTwo = Nothing
End Sub

Public Sub bindDocs()
Dim sqlQry As String
Dim Doc As New ADODB.Recordset
sqlQry = ""

sqlQry = "Select  DocName   from Tbl_Doctors Order by DocName "
Doc.Open sqlQry, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
Do Until Doc.EOF
    If Len((Doc!DocName)) Then
        txtATTDoc(0).AddItem (Doc!DocName)
        txtATTDoc(1).AddItem (Doc!DocName)
        txtATTDoc(2).AddItem (Doc!DocName)
        txtATTDoc(3).AddItem (Doc!DocName)
        txtATTDoc(4).AddItem (Doc!DocName)
    End If
    Doc.MoveNext
Loop
Doc.Close
Set Doc = Nothing
End Sub

Private Sub ComboListing()
Dim HOS As New ADODB.Recordset
Dim SRV As New ADODB.Recordset
Dim Nature As New ADODB.Recordset
Dim Modality As New ADODB.Recordset
Dim PAYMode As New ADODB.Recordset
Dim PAYTYPE As New ADODB.Recordset
Dim STAYTYPE As New ADODB.Recordset
Dim CASRep As New ADODB.Recordset
Dim SPEAPP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim tmpHOSCode As String
Dim tmpHosName As String
Dim Pending As New ADODB.Recordset


    
'    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
   ' Set cmd.ActiveConnection = medixmasterconnMaint
   ' cmd.CommandText = "SearchSpecApp"
   ' cmd.CommandType = adCmdStoredProc
  '  cmd.CommandTimeout = 15
  '  SPEAPP.CursorLocation = adUseClient
  '  SPEAPP.CursorType = adOpenForwardOnly
  '  SPEAPP.CacheSize = 100
  '  Set SPEAPP = cmd.Execute
 
'    Do Until SPEAPP.EOF
'        With SPEAPP
'            cboCASInd.AddItem !SpecAppCode & " - " & !SpecAppDesc
'            SPEAPP.MoveNext
'        End With
'    Loop
'    SPEAPP.Close
'    Set SPEAPP = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchNatureCode"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Nature.CursorLocation = adUseClient
    Nature.CursorType = adOpenForwardOnly
    Nature.CacheSize = 100
    Set Nature = cmd.Execute
     
    Do Until Nature.EOF
        With Nature
            cboNature.AddItem !ClaimNatureCode & " - " & !ClaimNatureName
            Nature.MoveNext
        End With
    Loop
    Nature.Close
    Set Nature = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPAYType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAYTYPE.CursorLocation = adUseClient
    PAYTYPE.CursorType = adOpenForwardOnly
    PAYTYPE.CacheSize = 100
    Set PAYTYPE = cmd.Execute
    
    Do Until PAYTYPE.EOF
        With PAYTYPE
            CboPayType.AddItem !PayTypeCode & " - " & Trim(!PayTypeName)
            PAYTYPE.MoveNext
        End With
    Loop
    PAYTYPE.Close
    Set PAYTYPE = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchStayType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    STAYTYPE.CursorLocation = adUseClient
    STAYTYPE.CursorType = adOpenForwardOnly
    STAYTYPE.CacheSize = 100
    Set STAYTYPE = cmd.Execute
    
    Do Until STAYTYPE.EOF
        With STAYTYPE
            cboStayType.AddItem !PayTypeCode & " - " & Trim(!PayTypeName)
            STAYTYPE.MoveNext
        End With
    Loop

    STAYTYPE.Close
    Set STAYTYPE = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchTreatMode"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Modality.CursorLocation = adUseClient
    Modality.CursorType = adOpenForwardOnly
    Modality.CacheSize = 100
    Set Modality = cmd.Execute

    Do Until Modality.EOF
        With Modality
            cboModality.AddItem !ModalityCode & " - " & Trim(!ModalityName)
            Modality.MoveNext
        End With
    Loop
    Modality.Close
    Set Modality = Nothing

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHOS"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cmd.Execute
    Set HOSArr = New Collection
    
    Do Until HOS.EOF
        With HOS
            tmpHOSCode = HOS!HOSCode
            tmpHosName = HOS!HOSName
            txtHOSCode.AddItem Trim(!HOSName)
            HOSArr.Add tmpHOSCode, tmpHosName
            HOS.MoveNext
        End With
    Loop
    HOS.Close
    Set HOS = Nothing
    
    cboStatus1.AddItem "O" & " - " & "Open"
    cboStatus1.AddItem "C" & " - " & "Close"
    cboStatus1.AddItem "D" & " - " & "Delete"

    cboClaimability.AddItem "AA" & " - " & "Accepted"
    cboClaimability.AddItem "RR" & " - " & "Repudiated"
    cboClaimability.AddItem "CC" & " - " & "Cancelled"
    cboClaimability.AddItem "SA" & " - " & "Special Approval"
    cboClaimability.AddItem "PP" & " - " & "Pay & Claim"
    cboClaimability.AddItem "NN" & " - " & "Not Decided"
    cboClaimability.AddItem "NA" & " - " & "Not Decided/Accepted"
    cboClaimability.AddItem "NR" & " - " & "Not Decided/Repudiated"
    cboClaimability.AddItem "PA" & " - " & "Pay&Claim/Accepted"
    cboClaimability.AddItem "PR" & " - " & "Pay&Claim/Rejected"
    cboClaimability.AddItem "AR" & " - " & "Accepted/Rejected"
    cboClaimability.AddItem "RA" & " - " & "Rejected/Accepted"
    cboClaimability.AddItem "BR" & " - " & "BTBG/Rejected"
    cboClaimability.AddItem "ER" & " - " & "Expired/Rejected"
    cboClaimability.AddItem "AI" & " - " & "Accepted/Insurance"
    'cboClaimability.AddItem "DB" & " - " & "Death Benefit"
    
    cboPending.AddItem "Y - Yes"
    cboPending.AddItem "N - No"
    
    cboCondwillOcc.AddItem "Y - YES"
    cboCondwillOcc.AddItem "N - NO"
    
    cboFLWTreatment.AddItem "Y - YES"
    cboFLWTreatment.AddItem "N - NO"

    'cboTopUpLmt.AddItem "N - No"
    'cboTopUpLmt.AddItem "Y - Yes"

    'cboQuery.AddItem "Y - YES"
    'cboQuery.AddItem "N - No"
    
    cboFolUp.AddItem "Y - YES"
    cboFolUp.AddItem "N - NO"
    
    'With cboDiagRepCode
    '    .AddItem "Prelim. Diag"
    '    .AddItem "Final Diag"
    '    .AddItem "Underlying Disease"
    '    If Mid(cboDiagRepCode, 1, 1) = "R" Then
    '        .AddItem "Repudiation Reason"
    '    End If
    'End With
    
    'With cboInvestProcCode
    '    .AddItem "Investigation"
    '    .AddItem "Procedures"
    'End With
    
    With cboPayPending
        .AddItem "Y - Yes"
        .AddItem "N - No"
    End With

    cboCASTakeOveInd.AddItem "Y - Yes"
    cboCASTakeOveInd.AddItem "N - No"
   
    cboIntExtTO.AddItem "I - Internal"
    cboIntExtTO.AddItem "E - External"
    
    cboAppStatus.AddItem "N - No"
    cboAppStatus.AddItem "Y - Yes"
    
    cboREVBordxInd.AddItem "N - No"
    cboREVBordxInd.AddItem "Y - Yes"

    cboBordxRevStatus.AddItem "N - No"
    cboBordxRevStatus.AddItem "Y - Yes"
    
    cboYearDay.AddItem "W - Weeks"
    cboYearDay.AddItem "M - Months"
    cboYearDay.AddItem "Y - Year"
    
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchPendingReason"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Pending.CursorLocation = adUseClient
    Pending.CursorType = adOpenForwardOnly
    Pending.CacheSize = 100
    Set Pending = cmd.Execute
 
    Do Until Pending.EOF
        With Pending
            cboPendingReason.AddItem !CASPendingReason
            Pending.MoveNext
        End With
    Loop
    Pending.Close
    Set Pending = Nothing
    
    cboOPChemo.AddItem "Y - Yes"
    cboOPChemo.AddItem "N - No"
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End Sub

Private Sub CheckClaim(ErrorStatus As String)
On Error Resume Next
    Dim CLM As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String

    'strclm = "SELECT CASNumber, MBMNumber, CLMStatus From CLM WHERE (CASNumber = '" & txtCaseNo & "') AND (MBMNumber = '" & txtMBMNumber & "') AND (CLMStatus = 'O')"
    'CLM.Open strclm, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "16"
    strAppend = ""

    strAppend = " AND (CASNumber = '" & txtCaseNo & "') AND (MBMNumber = '" & txtMBMNumber & "') AND (CLMStatus = 'O')"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        CLM.CursorLocation = adUseClient
        CLM.CursorType = adOpenForwardOnly
        CLM.CacheSize = 100
        Set CLM = cmd.Execute
        
    If CLM.EOF = False Then
        If ErrorStatus = "D" Then
            MsgBox "Claim Worksheet Already Registered Please Use Close Case", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf ErrorStatus = "C" Then
            MsgBox "Close All Claim Before Close The Case", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    End If
    CLM.Close
End Sub

Private Sub Form_Load()
Dim intFullHeight As Integer
Dim intDisplayHeight As Integer

   
    SndLG = False
    Call ComboListing
    FirstDiag = 0
    FinalDiag = 0

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
    
End Sub

Private Sub ShowFinalDiag()
    Dim diafinal As New ADODB.Recordset
    Dim strsql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String

    'strsql = "select * from DIAFinal where CASNumber = '" & txtCaseNo & "'"
    'diafinal.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "11"
    strAppend = ""

    strAppend = " AND CASNumber = '" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        diafinal.CursorLocation = adUseClient
        diafinal.CursorType = adOpenForwardOnly
        diafinal.CacheSize = 100
        Set diafinal = cmd.Execute
        
        If diafinal.EOF = False Then
                With diafinal
                    If Len(!ICDCode1) Then
                        cboFDICDCode1 = !ICDCode1
                    End If
                    If Len(!DIADOS1) Then
                        txtFDDOS1.mask = ""
                        txtFDDOS1 = Format(!DIADOS1, "dd/mm/yyyy")
                        txtFDDOS1.mask = "##/##/####"
                    End If
                    If Len(!DIAFinalDiag1) Then
                        txtFinalDiag1(0) = !DIAFinalDiag1
                    End If
                    If Len(!ICDCode2) Then
                        cboFDICDCode2 = !ICDCode2
                    End If
                    If Len(!DIADOS2) Then
                        txtFDDOS2.mask = ""
                        txtFDDOS2 = Format(!DIADOS2, "dd/mm/yyyy")
                        txtFDDOS2.mask = "##/##/####"
                    End If
                    If Len(!DIAFinalDiag2) Then
                        txtFinalDiag1(1) = !DIAFinalDiag2
                    End If
                    If Len(!ICDCode3) Then
                        cboFDICDCode3 = !ICDCode3
                    End If
                    If Len(!DIADOS3) Then
                        txtFDDOS3.mask = ""
                        txtFDDOS3 = Format(!DIADOS3, "dd/mm/yyyy")
                        txtFDDOS3.mask = "##/##/####"
                    End If
                    If Len(!DIAFinalDiag3) Then
                        txtFinalDiag1(2) = !DIAFinalDiag3
                    End If
        
                End With
            End If
    diafinal.Close
    Set diafinal = Nothing
End Sub

Private Sub ShowFirstDiag()
    Dim diafirst As New ADODB.Recordset
    Dim strsql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    
    'strsql = "Select * from DIAFirst where CasNumber = '" & Trim(txtCaseNo) & "'"
    'diafirst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "9"
    strAppend = ""
    
    strAppend = " AND CasNumber = '" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        diafirst.CursorLocation = adUseClient
        diafirst.CursorType = adOpenForwardOnly
        diafirst.CacheSize = 100
        Set diafirst = cmd.Execute
        
        If diafirst.EOF = False Then
            Do Until diafirst.EOF
            With diafirst
                If Len(!ICDCode1) Then
                    cboICDCode(0) = !ICDCode1
                End If
                If Len(!DIADOS1) Then
                    txtPDDOS(0).mask = ""
                    txtPDDOS(0) = Format(!DIADOS1, "dd/mm/yyyy")
                    txtPDDOS(0).mask = "##/##/####"
                End If
                If Len(!DIAPrelimDiag1) Then
                    txtPrelimDiag(0) = !DIAPrelimDiag1
                End If
                If Len(!ICDCode2) Then
                    cboICDCode(1) = !ICDCode2
                End If
                If Len(!DIADOS2) Then
                    txtPDDOS(1).mask = ""
                    txtPDDOS(1) = Format(!DIADOS2, "dd/mm/yyyy")
                    txtPDDOS(1).mask = "##/##/####"
                End If
                If Len(!DIAPrelimDiag2) Then
                    txtPrelimDiag(1) = !DIAPrelimDiag2
                End If
                If Len(!ICDCode3) Then
                    cboICDCode(2) = !ICDCode3
                End If
                If Len(!DIADOS3) Then
                    txtPDDOS(2).mask = ""
                    txtPDDOS(2) = Format(!DIADOS3, "dd/mm/yyyy")
                    txtPDDOS(2).mask = "##/##/####"
                End If
                If Len(!DIAPrelimDiag3) Then
                    txtPrelimDiag(2) = !DIAPrelimDiag3
                End If
            End With
            diafirst.MoveNext
            Loop
            End If
    diafirst.Close
    Set diafirst = Nothing
End Sub

Private Sub ShowUD()
    Dim CASUD As New ADODB.Recordset
    Dim strsql As String
    
        If Not CASSearch.EOF Then
            With CASSearch
                If Len(!CASUDCode1) Then
                    cboUDICDCode1 = !CASUDCode1
                End If
                If Len(!CASUDDOS1) Then
                    txtUDDOS1.mask = ""
                    txtUDDOS1 = Format(!CASUDDOS1, "dd/mm/yyyy")
                    txtUDDOS1.mask = "##/##/####"
                End If
                If Len(!CASUDReasons1) Then
                    txtUD1(0) = !CASUDReasons1
                End If
                If Len(!CASUDCode2) Then
                    cboUDICDCode2 = !CASUDCode2
                End If
                If Len(!CASUDDOS2) Then
                    txtUDDOS2.mask = ""
                    txtUDDOS2 = Format(!CASUDDOS2, "dd/mm/yyyy")
                    txtUDDOS2.mask = "##/##/####"
                End If
                If Len(!CASUDReasons2) Then
                    txtUD2 = !CASUDReasons2
                End If
                If Len(!CASUDCode3) Then
                    cboUDICDCode3 = !CASUDCode3
                End If
                If Len(!CASUDDOS3) Then
                    txtUDDOS3.mask = ""
                    txtUDDOS3 = Format(!CASUDDOS3, "dd/mm/yyyy")
                    txtUDDOS3.mask = "##/##/####"
                End If
                If Len(!CASUDReasons3) Then
                    txtUD3 = !CASUDReasons3
                End If
               
            End With
        End If
End Sub



Private Sub lblPage1_Click()
        txtCASRemark.Visible = True
        txtCASRemarkTwo.Visible = False
        txtCASRemark3.Visible = False
        txtCASRemark4.Visible = False
        txtCASRemark5.Visible = False
        lblPage.Caption = "Page 1"
        lblPage1.FontBold = True
        lblPage2.FontBold = False
        lblPage3.FontBold = False
        lblPage4.FontBold = False
        lblPage5.FontBold = False
End Sub

Private Sub lblPage2_Click()
        txtCASRemark.Visible = False
        txtCASRemarkTwo.Visible = True
        txtCASRemark3.Visible = False
        txtCASRemark4.Visible = False
        txtCASRemark5.Visible = False
        lblPage.Caption = "Page 2"
        lblPage1.FontBold = False
        lblPage2.FontBold = True
        lblPage3.FontBold = False
        lblPage4.FontBold = False
        lblPage5.FontBold = False
End Sub

Private Sub lblPage3_Click()
    txtCASRemark.Visible = False
    txtCASRemarkTwo.Visible = False
    txtCASRemark3.Visible = True
    txtCASRemark4.Visible = False
    txtCASRemark5.Visible = False
    lblPage.Caption = "Page 3"
    lblPage1.FontBold = False
    lblPage2.FontBold = False
    lblPage3.FontBold = True
    lblPage4.FontBold = False
    lblPage5.FontBold = False
End Sub

Private Sub lblPage4_Click()
    txtCASRemark.Visible = False
    txtCASRemarkTwo.Visible = False
    txtCASRemark3.Visible = False
    txtCASRemark4.Visible = True
    txtCASRemark5.Visible = False
    lblPage.Caption = "Page 4"
    lblPage1.FontBold = False
    lblPage2.FontBold = False
    lblPage3.FontBold = False
    lblPage4.FontBold = True
    lblPage5.FontBold = False
    
End Sub

Private Sub lblPage5_Click()
    txtCASRemark.Visible = False
    txtCASRemarkTwo.Visible = False
    txtCASRemark3.Visible = False
    txtCASRemark4.Visible = False
    txtCASRemark5.Visible = True
    lblPage.Caption = "Page 5"
    lblPage1.FontBold = False
    lblPage2.FontBold = False
    lblPage3.FontBold = False
    lblPage4.FontBold = False
    lblPage5.FontBold = True
    
End Sub
Sub crt(pth As String)
Dim f
f = FreeFile
Open App.path + "\temp.html" For Output As #f
    Print #f, "<html><body><center><img src=""" & pth & """></center></body></html>"
Close #f
DoEvents
'WebBrowser1.Navigate App.Path & "\temp.html"
End Sub


Private Sub List1_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    
    PictureFileName = DocPath & "ScannedDoc\" & txtCaseNo & ""
    '\Medical Rpt"
    crt PictureFileName & "\" & List1.List(List1.ListIndex)

    'If Index = 0 Then
        sFile = PictureFileName & "\" & List1.List(List1.ListIndex)
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
       ' Picture2.Picture = LoadPicture()
       ' Picture2.Picture = LoadPicture(sFile)
End Sub

Private Sub List2_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    
    PictureFileName = DocPath & "InternalDocuments\" & txtCaseNo & ""
    '\Medical Rpt"
    crt PictureFileName & "\" & List2.List(List2.ListIndex)

    'If Index = 0 Then
        sFile = PictureFileName & "\" & List2.List(List2.ListIndex)
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
       ' Picture2.Picture = LoadPicture()
       ' Picture2.Picture = LoadPicture(sFile)

End Sub

Private Sub List2_ItemCheck(Item As Integer)
    tmpDeleteGL = DocPath & "InternalDocuments\" & txtCaseNo & "\" & List2.List(List2.ListIndex)
End Sub

Private Sub List3_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    
    PictureFileName = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & ""
    '\Medical Rpt"
    crt PictureFileName & "\" & List3.List(List3.ListIndex)

    'If Index = 0 Then
        sFile = PictureFileName & "\" & List3.List(List3.ListIndex)
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
       ' Picture2.Picture = LoadPicture()
       ' Picture2.Picture = LoadPicture(sFile)

End Sub

Private Sub List3_ItemCheck(Item As Integer)
    tmpDelGLWeb = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & List3.List(List3.ListIndex)
End Sub

Private Sub lstLGListing_Click()

    If lstLGListing.ListItems.Count Then
        
        LGNo = ""
        LGVer = 0
        
        LGNo = Trim(Mid(lstLGListing.SelectedItem.SubItems(9), 1, IIf(InStr(1, lstLGListing.SelectedItem.SubItems(9), "-") = 0, 10, InStr(1, lstLGListing.SelectedItem.SubItems(9), "-") - 1)))
        LGVer = Trim(Mid(lstLGListing.SelectedItem.SubItems(9), InStr(1, lstLGListing.SelectedItem.SubItems(9), "-") + 1, Len(lstLGListing.SelectedItem.SubItems(9)) - InStr(1, lstLGListing.SelectedItem.SubItems(9), "-")))
             
        'TxtLGRemark = Trim(lstLGListing.SelectedItem.SubItems(8))
        
    End If
End Sub



Private Sub lstPolicyWording_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    
    PictureFileName = DocPath & "PolicyWording\" & tmpPolicyWording & ""
    '\Medical Rpt"
    crt PictureFileName & "\" & lstPolicyWording.List(lstPolicyWording.ListIndex)

    'If Index = 0 Then
        sFile = PictureFileName & "\" & lstPolicyWording.List(lstPolicyWording.ListIndex)
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
       ' Picture2.Picture = LoadPicture()
       ' Picture2.Picture = LoadPicture(sFile)
End Sub


Private Sub ListCases1()
Dim CASList As ListItem
Dim strsql As String
Dim CAS As New ADODB.Recordset
Dim tmpPrin, tmpPatient, tmpAge, tmpGender, tmpDiag As String
Dim tmpUndIll, tmpInv, TmpEffDate, tmpDOb As String
Dim tmpCaseHis As String
Dim cnt As Integer
Dim tmpHosName As String
Dim tmpCaseHis1 As String
Dim tmpDOA, tmpDOD As String
Dim tmpMBMNumber, tmpCASNumber As String
Dim tmpAppAmt As Double
Dim AppStatus As String
Dim Supp As New ADODB.Recordset
Dim strsqlSupp As String
Dim cntSupp As Integer
Dim TotalCnt As Integer
Dim tmpCaseSuppHistory As String
Dim tmpCaseSuppHistory1 As String
Dim tmpClaimVer As String

    tmpAppAmt = 0
    AppStatus = ""
    cntSupp = 0
    TotalCnt = 0
    tmpCaseHistory = ""
    tmpCaseSuppHistory = ""
    tmpCaseSuppHistory1 = ""
    tmpCaseHis = ""
    tmpCaseHis1 = ""
    tmpClaimVer = ""
    strsql = ""
    strsql = "Exec SearchCasHistorymgmt '" & Trim(txtPatientIC) & "' , '" & Format(txtDateAdmitted, "yyyy/mm/DD") & "'"
    
    Set CAS = New ADODB.Recordset
    CAS.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly

    Screen.MousePointer = vbHourglass
    Do While Not CAS.EOF
        
        'If CAS!ReceiptNo <> "" Then
        '    AppStatus = "APPROVED"
        'Else
        '    AppStatus = "NOT APPROVED"
        'End If
        tmpMBMNumber = IIf(Len(Trim(CAS!MBMNumber)), Trim(CAS!MBMNumber), "")
        tmpCASNumber = IIf(Len(Trim(CAS!CasNumber)), Trim(CAS!CasNumber), "")
        'tmpClaimVer = IIf(Len(Trim(CAS!claimversion)), Trim(CAS!claimversion), "")
        'If tmpClaimVer <> "" Then
        '    tmpCASNumber = tmpCASNumber & "-" & tmpClaimVer
        'End If
        If CAS!MBMPatientCovID = "00" Then
            tmpPrin = IIf(Len(Trim(CAS!MBMName)), Trim(CAS!MBMName), "")
            tmpPatient = IIf(Len(Trim(CAS!MBMName)), Trim(CAS!MBMName), "")
        Else
            tmpPatient = IIf(Len(Trim(CAS!MBMCName)), Trim(CAS!MBMCName), "")
        End If
        If CAS!MBMPatientCovID = "00" Then
            tmpDOb = IIf(Len(Trim(CAS!MBMDOB)), Trim(CAS!MBMDOB), "")
        Else
            tmpDOb = IIf(Len(Trim(CAS!MBMCDOB)), Trim(CAS!MBMCDOB), "")
        End If
        
        tmpAppAmt = IIf(IsNumeric(CAS!CLMBordxAmt), CAS!CLMBordxAmt, 0)
        
        TmpEffDate = IIf(Len(Trim(CAS!MBMPayorEffDate)), Trim(CAS!MBMPayorEffDate), "")
        If Len(tmpDOb) Then
            tmpAge = GetAgeCase(CDate(TmpEffDate), CDate(tmpDOb))
        End If
        If Trim(CAS!DIAFinalDiag1) <> "" Then
            tmpDiag = CAS!DIAFinalDiag1
        End If
        If Trim(CAS!DIAFinalDiag1) <> "" And Trim(CAS!DIAFinalDiag2) <> "" Then
            tmpDiag = CAS!DIAFinalDiag1 & ";" & CAS!DIAFinalDiag2
        End If
        
        If Trim(CAS!DIAFinalDiag1) <> "" And Trim(CAS!DIAFinalDiag2) <> "" And Trim(CAS!DIAFinalDiag3) <> "" Then
            tmpDiag = CAS!DIAFinalDiag1 & ";" & CAS!DIAFinalDiag2 & ";" & CAS!DIAFinalDiag3
        End If
                      
        If Trim(txtUD1(0)) <> "" Then
            tmpUndIll = txtUD1(0)
        End If
        
        'If Trim(txtUD1) <> "" And Trim(txtUD2) <> "" Then
        '    tmpUndIll = txtUD1 & ";" & txtUD2
        'End If
       
        'If Trim(txtUD1) <> "" And Trim(txtUD2) <> "" And Trim(txtUD3) <> "" Then
        '    tmpUndIll = txtUD1 & ";" & txtUD2 & ";" & txtUD3
        'End If
        
        'If Trim(txtPROCINVDone1) <> "" Then
        '    tmpInv = txtPROCINVDone1
        'End If
        
        'If Trim(txtPROCINVDone1) <> "" And Trim(txtPROCINVDone2) <> "" Then
        '    tmpInv = txtPROCINVDone1 & ";" & txtPROCINVDone2
        'End If
       
        'If Trim(txtPROCINVDone1) <> "" And Trim(txtPROCINVDone2) <> "" And Trim(txtPROCINVDone3) <> "" Then
        '    tmpInv = txtPROCINVDone1 & ";" & txtPROCINVDone2 & ";" & txtPROCINVDone3
        'End If
        
        'tmpHosName = IIf(Len(CAS!HOSName), CAS!HOSName, "")
        
        tmpDOA = IIf(Len(CAS!CASDateAdmitted), CAS!CASDateAdmitted, "")
        
        tmpDOD = IIf(Len(CAS!CASDischargeDate), CAS!CASDischargeDate, "")
        
            cnt = cnt + 1
            If cnt = 1 Then
                tmpCaseHistory = "" & _
                    vbNewLine & "                        CLAIM HISTORY" & _
                    vbNewLine & "                        [" & cnt & "]  " & tmpCASNumber & "       " & tmpPatient & "         " & tmpDOA & " - " & tmpDOD & " " & _
                    vbNewLine & "                        " & tmpDiag & "                                    " & tmpAppAmt & "     "
            Else
                tmpCaseHis1 = ""
                tmpCaseHis1 = "" & _
                    vbNewLine & "                        [" & cnt & "] " & tmpCASNumber & "        " & tmpPatient & "          " & tmpDOA & " - " & tmpDOD & " " & _
                    vbNewLine & "                        " & tmpDiag & "                                     " & tmpAppAmt & "     "
                 tmpCaseHistory = tmpCaseHistory + tmpCaseHis1
            End If
        CAS.MoveNext
    Loop
    CAS.Close
    Set CAS = Nothing
    
            
        'strsqlSupp = "Select MBMCICBCPPNo from MBMCoveredPersons where MBMCNumber '" & Trim(txtMBMIcBcPP) & "'"
        'Supp.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        

            'Do While Not Supp.EOF
                strsql = "Exec SearchCasHistorySuppmgmt '" & Trim(txtMBMIcBcPP) & "', '" & Format(txtDateAdmitted, "yyyy/mm/DD") & "'"
            
                Set CAS = New ADODB.Recordset
                CAS.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            
            
                Screen.MousePointer = vbHourglass
                Do While Not CAS.EOF
                
                     'If CAS!ReceiptNo <> "" Then
                     '    AppStatus = "APPROVED"
                     'Else
                     '    AppStatus = "NOT APPROVED"
                     'End If
                     tmpMBMNumber = IIf(Len(Trim(CAS!MBMCNumber)), Trim(CAS!MBMCNumber), "")
                     tmpCASNumber = IIf(Len(Trim(CAS!CasNumber)), Trim(CAS!CasNumber), "")
                     If CAS!MBMPatientCovID = "00" Then
                         tmpPrin = IIf(Len(Trim(CAS!MBMName)), Trim(CAS!MBMName), "")
                         tmpPatient = IIf(Len(Trim(CAS!MBMName)), Trim(CAS!MBMName), "")
                     Else
                         tmpPatient = IIf(Len(Trim(CAS!MBMCName)), Trim(CAS!MBMCName), "")
                     End If
                     If CAS!MBMPatientCovID = "00" Then
                         tmpDOb = IIf(Len(Trim(CAS!MBMDOB)), Trim(CAS!MBMDOB), "")
                     Else
                         tmpDOb = IIf(Len(Trim(CAS!MBMCDOB)), Trim(CAS!MBMCDOB), "")
                     End If
                     
                     tmpAppAmt = IIf(IsNumeric(CAS!CLMBordxAmt), CAS!CLMBordxAmt, 0)
                     
                     TmpEffDate = IIf(Len(Trim(CAS!MBMPayorEffDate)), Trim(CAS!MBMPayorEffDate), "")
                     
                     tmpAge = GetAgeCase(CDate(TmpEffDate), CDate(tmpDOb))
                    
                     If Trim(CAS!DIAFinalDiag1) <> "" Then
                         tmpDiag = CAS!DIAFinalDiag1
                     End If
                     
                     If Trim(CAS!DIAFinalDiag1) <> "" And Trim(CAS!DIAFinalDiag2) <> "" Then
                         tmpDiag = CAS!DIAFinalDiag1 & ";" & CAS!DIAFinalDiag2
                     End If
                     
                     If Trim(CAS!DIAFinalDiag1) <> "" And Trim(CAS!DIAFinalDiag2) <> "" And Trim(CAS!DIAFinalDiag3) <> "" Then
                         tmpDiag = CAS!DIAFinalDiag1 & ";" & CAS!DIAFinalDiag2 & ";" & CAS!DIAFinalDiag3
                     End If
                                   
                     If Trim(txtUD1(0)) <> "" Then
                         tmpUndIll = txtUD1(0)
                     End If
                     
                     'If Trim(txtUD1) <> "" And Trim(txtUD2) <> "" Then
                     '    tmpUndIll = txtUD1 & ";" & txtUD2
                     'End If
                    
                     'If Trim(txtUD1) <> "" And Trim(txtUD2) <> "" And Trim(txtUD3) <> "" Then
                     '    tmpUndIll = txtUD1 & ";" & txtUD2 & ";" & txtUD3
                     'End If
                     
                     'If Trim(txtPROCINVDone1) <> "" Then
                     '    tmpInv = txtPROCINVDone1
                     'End If
                     
                     'If Trim(txtPROCINVDone1) <> "" And Trim(txtPROCINVDone2) <> "" Then
                     '    tmpInv = txtPROCINVDone1 & ";" & txtPROCINVDone2
                     'End If
                    
                     'If Trim(txtPROCINVDone1) <> "" And Trim(txtPROCINVDone2) <> "" And Trim(txtPROCINVDone3) <> "" Then
                     '    tmpInv = txtPROCINVDone1 & ";" & txtPROCINVDone2 & ";" & txtPROCINVDone3
                     'End If
                     
                     'tmpHosName = IIf(Len(CAS!HOSName), CAS!HOSName, "")
                     
                     tmpDOA = IIf(Len(CAS!CASDateAdmitted), CAS!CASDateAdmitted, "")
                     
                     tmpDOD = IIf(Len(CAS!CASDischargeDate), CAS!CASDischargeDate, "")
                
                    'cntSupp = 1
                    If cntSupp = 0 Then
                        TotalCnt = 1 + cnt
                        cntSupp = 1
                        tmpCaseSuppHistory = "" & _
                            vbNewLine & "                        CLAIM HISTORY" & _
                            vbNewLine & "                        [" & TotalCnt & "]  " & tmpCASNumber & "       " & tmpPatient & "         " & tmpDOA & " - " & tmpDOD & " " & _
                            vbNewLine & "                        " & tmpDiag & "                                    " & tmpAppAmt & ""
                    Else
                        TotalCnt = cntSupp + TotalCnt
                        tmpCaseSuppHistory1 = ""
                        tmpCaseSuppHistory1 = "" & _
                            vbNewLine & "                        [" & TotalCnt & "] " & tmpCASNumber & "        " & tmpPatient & "          " & tmpDOA & " - " & tmpDOD & " " & _
                            vbNewLine & "                        " & tmpDiag & "                                     " & tmpAppAmt & ""
                         tmpCaseSuppHistory = tmpCaseSuppHistory + tmpCaseSuppHistory1
                    End If
                    CAS.MoveNext
                Loop
                CAS.Close
                Set CAS = Nothing
                tmpCaseHistory = tmpCaseHistory + tmpCaseSuppHistory
         '   Supp.MoveNext
         '   Loop
        'Supp.Close
        'Set Supp = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub ListALLCasesMgmt()
Dim ForLoop As Integer
Dim TempIcNo As String
Dim StrSqlMBMsupp As String
Dim MBMSupp As New ADODB.Recordset
Dim MBMSuppRecStatus As Boolean
    'cntCasesHisIC = 0
    ForLoop = 0
    TempIcNo = ""
    lstNote1(1).ListItems.Clear
    'MBMSuppRecStatus = False
    If Len(Trim(txtMBMIcBcPP)) > 5 Then
        Call ListCases1Mgmt(True, "", txtMBMIcBcPP, False)
        'StrSqlMBMsupp = "Select Top 1 MBMCICBCPPNo from MBMCoveredPersons where MBMCICBCPPNo = '" & Trim(txtMBMIcBcPP) & "'"
        'MBMSupp.Open StrSqlMBMsupp, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        'MBMSuppRecStatus = False
        'Do While Not MBMSupp.EOF
           MBMSuppRecStatus = True
        '   MBMSupp.MoveNext
        '   Exit Do
        'Loop
        'MBMSupp.Close
        'Set MBMSupp = Nothing
       
       ' If MBMSuppRecStatus = True Then
       '     Call ListCases1(True, "", txtMBMIcBcPP, True)
       ' End If
    End If
    'If Mid(cboINSCode, 1, 1) = "F" Or Mid(cboINSCode, 1, 1) = "H" Then
    '    For ForLoop = 1 To lstCovAdd.ListItems.Count
    '        MBMSuppRecStatus = False
    '        TempIcNo = lstCovAdd.ListItems(ForLoop).SubItems(2)
        '    If Len(Trim(TempIcNo)) > 5 Then
        '        StrSqlMBMsupp = "Select Top 1 MBMIcBcPp from MBMCrossReference where MBMIcBcPp = '" & Trim(TempIcNo) & "'"
        '        MBMSupp.Open StrSqlMBMsupp, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
        '        Do While Not MBMSupp.EOF
          '          MBMSuppRecStatus = True
        '            MBMSupp.MoveNext
        '            Exit Do
        '        Loop
        '        MBMSupp.Close
        '        Set MBMSupp = Nothing
        '    End If
    '        If MBMSuppRecStatus = True Then
    '            Call ListCases1(True, "", TempIcNo, False)
    '        End If
    '    Next ForLoop
    'End If
    'lblTotalCasesIC = cntCasesHisIC
End Sub
Private Sub ListCases1Mgmt(tmpHistory As Boolean, tmpMBMNo As String, tmpICNo As String, tmpSuppCas As Boolean)
Dim CASList As ListItem
Dim strsql As String
Dim CAS As New ADODB.Recordset
   ' If tmpHistory = False Then
   '     strsql = "Exec SearchCas '" & Trim(tmpMBMNo) & "'"
   '     lstCaseList.ListItems.Clear
   ' Else
        'If tmpSuppCas = False Then
            strsql = "Exec SearchCASHistory '" & Trim(tmpICNo) & "'"
        'Else
        '    strsql = "Exec SearchCasHistorySupp '" & Trim(tmpICNo) & "'"
        'End If
    'End If
    Set CAS = New ADODB.Recordset
    CAS.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly

    Screen.MousePointer = vbHourglass
        
    Do While Not CAS.EOF
        'If tmpHistory = False Then
            Set CASList = lstNote1(1).ListItems.Add(, , CAS!CasNumber)
        'Else
        '    Set CASList = lstCASHistory.ListItems.Add(, , CAS!CasNumber)
        'End If
        CASList.SubItems(1) = CAS!MBMNumber
        If CAS!MBMPatientCovID = "00" Then
            CASList.SubItems(2) = IIf(Len(Trim(CAS!MBMName)), Trim(CAS!MBMName), "")
            CASList.SubItems(3) = "P"
        Else
            CASList.SubItems(2) = IIf(Len(Trim(CAS!MBMCName)), Trim(CAS!MBMCName), "")
            CASList.SubItems(3) = IIf(Len(CAS!MBMCRELCode), CAS!MBMCRELCode, "")
        End If
        
        If Len(Trim(CAS!CASDateAdmitted)) Then
            CASList.SubItems(4) = CAS!CASDateAdmitted
        End If
        If Len(Trim(CAS!CASDischargeDate)) Then
            CASList.SubItems(5) = CAS!CASDischargeDate
        End If
        CASList.SubItems(6) = IIf(Len(CAS!CASClaimability), CAS!CASClaimability, "")
        CASList.SubItems(7) = IIf(Len(CAS!DIAFinalDiag1), CAS!DIAFinalDiag1, "")
        CASList.SubItems(8) = IIf(Len(CAS!CASStatus), CAS!CASStatus, "")
        CASList.SubItems(9) = IIf(Len(CAS!HOSName), CAS!HOSName, "")
        CAS.MoveNext
        
    Loop
    
    CAS.Close
    Set CAS = Nothing
    Screen.MousePointer = vbDefault
End Sub

Public Function GetAgeCase(EffectiveDate As Date, DateOfBirth As Date)
    Dim CheckDate As Long
    Dim CheckYear As Integer
    Dim CheckDOB As Integer
    
    CheckDate = Year(EffectiveDate) - Year(DateOfBirth)
    CheckYear = CheckDate
    GetAgeCase = CheckYear
End Function
Private Sub ListALLCases()
Dim ForLoop As Integer
Dim TempIcNo As String
Dim StrSqlMBMsupp As String
Dim MBMSupp As New ADODB.Recordset
Dim MBMSuppRecStatus As Boolean
       
    If Len(Trim(txtPatientIC)) > 5 Then
        Call ListCases1
       ' StrSqlMBMsupp = "Select MBMCICBCPPNo from MBMCoveredPersons where MBMCICBCPPNo = '" & Trim(txtMBMIcBcPP) & "'"
       ' MBMSupp.Open StrSqlMBMsupp, medixmasterconn, adOpenForwardOnly, adLockReadOnly
       ' MBMSuppRecStatus = False
       ' Do While Not MBMSupp.EOF
       '    MBMSuppRecStatus = True
       '    MBMSupp.MoveNext
       '    Exit Do
       ' Loop
       ' MBMSupp.Close
       ' Set MBMSupp = Nothing
        'If MBMSuppRecStatus = True Then
        '    Call ListCases1(True, "", txtMBMIcBcPP, True)
        'End If
    End If
    'If Mid(cboINSCode, 1, 1) = "F" Or Mid(cboINSCode, 1, 1) = "H" Then
    '    For ForLoop = 1 To lstCovAdd.listitems.Count
    '        MBMSuppRecStatus = False
    '        TempIcNo = lstCovAdd.listitems(ForLoop).SubItems(2)
    '        If Len(Trim(TempIcNo)) > 5 Then
    '            StrSqlMBMsupp = "Select MBMIcBcPp from MBMCrossReference where MBMIcBcPp = '" & Trim(TempIcNo) & "'"
    '            MBMSupp.Open StrSqlMBMsupp, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    '            Do While Not MBMSupp.EOF
    '                MBMSuppRecStatus = True
    '                MBMSupp.MoveNext
    '                Exit Do
    '            Loop
    '            MBMSupp.Close
    '            Set MBMSupp = Nothing
    '        End If
    '        If MBMSuppRecStatus = True Then
    '            Call ListCases1(True, "", TempIcNo, False)
    '        End If
    '    Next ForLoop
    'End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Dim frmTermsStatus As Integer
Dim tmpAge As String
Dim INVList As ListItem
    If SSTab1.Tab = 8 Then
        lstNote1(1).ListItems.Clear
        lstNote1(5).ListItems.Clear
        lstNote1(0).ListItems.Clear
        lstNote1(2).ListItems.Clear
        lstNote1(7).ListItems.Clear
        lstNote1(8).ListItems.Clear
        lstINVListing.ListItems.Clear
        'VScroll1.SmallChange = 25 'how much to change value on arrow click
        'VScroll1.LargeChange = 200 'how much to change value on scroll
        'VScroll1.Min = 0 'minimum value (set to text box top)
        'VScroll1.Max = 6000 'maximum value (set to text box height
        
        'Dim i As Integer
        'Dim lTop As Long
                
        'Picture1.Top = 500
        'Picture1.Left = 80
        
        'Picture2.Top = 30
        'Picture2.Height = 12800
        'Picture1.BackColor = Picture2.BackColor
        'Picture1.Height = 13000
        '
        'VScroll1.Top = 20
        'i = 1
        'VScroll1.Height = SSTab1.Height - 1000
            
            Call MBMHistoryMgmt("Princ")
            Call MBMHistoryMgmt("Supp")
            Call ListALLCasesMgmt
            Call INVListing
            Call WSListingMgmt
            Call ExcessListingMgmt
            
            Call ReceiptFRMBMListingMgmt
            Call ReceiptFRMPYRListingMgmt
            Call PaymentHOSListingMgmt
            Call PaymentMBMListingMgmt
            'Call PaymentReissuedListing
            'Call ReceiptFRMNRBListing
            
            If txtATTDoc(0) <> "" Then
                Set INVList = lstNote1(2).ListItems.Add(, , txtATTDoc(0))
                    If txtATTDoc(1) <> "" Then
                        Set INVList = lstNote1(2).ListItems.Add(, , txtATTDoc(1))
                    End If
                    If txtATTDoc(2) <> "" Then
                        Set INVList = lstNote1(2).ListItems.Add(, , txtATTDoc(2))
                    End If
                    
                    If txtATTDoc(3) <> "" Then
                        Set INVList = lstNote1(2).ListItems.Add(, , txtATTDoc(3))
                    End If
                    
                    If txtATTDoc(4) <> "" Then
                        Set INVList = lstNote1(2).ListItems.Add(, , txtATTDoc(4))
                    End If
            End If
            
            If txtAnaesDoc(0) <> "" Then
                Set INVList = lstNote1(8).ListItems.Add(, , txtAnaesDoc(0))
                If txtAnaesDoc(1) <> "" Then
                    Set INVList = lstNote1(8).ListItems.Add(, , txtAnaesDoc(1))
                End If
                If txtAnaesDoc(2) <> "" Then
                    Set INVList = lstNote1(8).ListItems.Add(, , txtAnaesDoc(2))
                End If
            End If
            
            Call LGListing
    End If
    

    If SSTab1.Tab = 8 Then
    '    lstNote1(2).ListItems.Clear
    '    lstNote1(7).ListItems.Clear
    '    lstNote1(8).ListItems.Clear
    '    lstINVListing.ListItems.Clear
        
    '        Call INVListing
    '        Call ReceiptFRMBMListingMgmt
    '        Call ReceiptFRMPYRListingMgmt
    '        Call PaymentHOSListingMgmt
    '        Call PaymentMBMListingMgmt
            

            'Call PaymentReissuedListing
            'Call ReceiptFRMNRBListing
    'End If

       ' If SSTab1.Tab = 6 Then
        
       '     If Trim(txtCASRemark) = "" And txtCaseNo <> "" Then
       '                 Call MBMHistory("Princ", True)
       '         Call ListALLCases
       '         tmpDiagnosis = ""
       '         tmpFinalDiagnosis = ""
       '         If Trim(txtPrelimDiag(0)) <> "" And txtPrelimDiag(1) <> "" And txtPrelimDiag(2) <> "" Then
       '             tmpDiagnosis = txtPrelimDiag(0) & ";" & txtPrelimDiag(1) & ";" & txtPrelimDiag(1)
       '         ElseIf Trim(txtPrelimDiag(0)) <> "" And txtPrelimDiag(1) <> "" Then
       '             tmpDiagnosis = txtPrelimDiag(0) & ";" & txtPrelimDiag(1)
       '         ElseIf Trim(txtPrelimDiag(0)) <> "" Then
       '             tmpDiagnosis = txtPrelimDiag(0)
       '         End If
       '
       '         If Trim(txtFinalDiag1) <> "" And txtFinalDiag2 <> "" And txtFinalDiag3 <> "" Then
       '             tmpFinalDiagnosis = txtFinalDiag1 & ";" & txtFinalDiag2 & ";" & txtFinalDiag3
       '         ElseIf Trim(txtFinalDiag1) <> "" And txtFinalDiag2 <> "" Then
       '             tmpFinalDiagnosis = txtFinalDiag1 & ";" & txtFinalDiag2
       '         ElseIf Trim(txtFinalDiag1) <> "" Then
       '             tmpFinalDiagnosis = txtFinalDiag1
       '         End If
       '
       '
       '         tmpAge = GetAgeCase(CDate(txtMBMEffDate), CDate(tmpPatientDOB))
                'TmpExclusion = IIf(Len(txtPrincipalExc), txtPrincipalExc, "NIL")
                
       '         If Mid(CboPayType, 1, 1) = "C" Then
       '             tmpRemarks = "                        PT Name: " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, " - ") + 1, Len(cboPatientList) - InStr(1, cboPatientList, " - "))) & "    Age : " & tmpAge & "" & _
       '                  vbNewLine & "                        DOA: " & txtDateAdmitted & " " & _
       '                  vbNewLine & "                        HOSP: " & txtHOSCode & " " & _
       '                  vbNewLine & "" & GetServerTime & "   RCVD M/FORM @                      REGISTERED @       " & _
       '                  vbNewLine & "" & _
       '                  vbNewLine & "                        PYMT MODE: " & tmpMBMPymtMode & " NEXT DUE :                                                                                  " & Logon & " " & _
       '                  vbNewLine & "" & _
       '                  vbNewLine & "                        GROUP CO: " & tmpMBMGroupCompany & " " & _
       '                  vbNewLine & "                        RECOVERY: " & _
       '                  vbNewLine & "" & _
       '                  vbNewLine & "                         EXCLUSION : " & TmpExclusion & "" & _
       '                  vbNewLine & "                          " & tmpPolHistory & "" & _
       '                  vbNewLine & "" & _
       '                  vbNewLine & "                          " & tmpCaseHistory & "" & _
       '                  vbNewLine & "" & _
       '                  vbNewLine & "                        PROV DIAG: " & tmpDiagnosis & " " & _
       '                  vbNewLine & "                        SYMPTOM: " & _
       '                  vbNewLine & "                        DURATION: " & _
       '                  vbNewLine & "                        U/ILLNESS 1: " & txtUD1(0) & " " & _
       '                  vbNewLine & "                        U/ILLNESS 2: " & txtUD2 & " " & _
       '                  vbNewLine & "                        U/ILLNESS 3: " & txtUD3 & " " & _
       '                  vbNewLine & "" & _
       '                  vbNewLine & "                        CASE COVERED, LG FAXED @" & _
       '                  vbNewLine & " " & _
       '                  vbNewLine & "                        RCVD FINAL BILL RM @ "''

                'tmpRemarks2 = vbNewLine & "" & _
                '         vbNewLine & "                        TOTAL EXCESS : RM " & _
                '         vbNewLine & "                        ADM KIT:             TTA DRUGS:" & _
                '         vbNewLine & "                        EXCESS LETTER & FINAL LG FAXED @ " & _
                '         vbNewLine & "" & _
                '         vbNewLine & "                        FINAL DIAG: " & tmpFinalDiagnosis & "" & _
                '         vbNewLine & "                        DR 1) " & txtATTDoc(0) & "" & _
                '         vbNewLine & "                        DR 2) " & txtATTDoc(1) & "" & _
                '         vbNewLine & "                        DR 3) " & txtATTDoc(2) & "" & _
                '         vbNewLine & "                        DR 4) " & txtATTDoc(3) & "" & _
                '         vbNewLine & "                        DR 5) " & txtATTDoc(4) & "" & _
                '         vbNewLine & "" & _
                '         vbNewLine & "                        PROCEDURE: " & _
                '         vbNewLine & "                                             MMA RATE         HOSP   " & _
                '         vbNewLine & "                        S/FEE:    " & _
                '         vbNewLine & "                        A/FEE:   " & _
                '         vbNewLine & "                        DOA : " & txtDateAdmitted & " DOD : " & txtDateDischarge & "" & _
                '         vbNewLine & "                        Annual Limit:   " & _
                '         vbNewLine & "                        Balance Limit:   "
           '
           '
            
            'ElseIf Mid(CboPayType, 1, 1) = "R" Then
            '    tmpRemarks = "                        Document : " & tmpCASDocType & "  Received On " & tmpCASDocDate & " assigned to " & tmpCaseOpenBy & "                                            " & Logon & "" & _
            '         vbNewLine & "" & _
            '         vbNewLine & "                        PT Name: " & Trim(Mid(cboPatientList, InStr(1, cboPatientList, " - ") + 1, Len(cboPatientList) - InStr(1, cboPatientList, " - "))) & "     Age : " & tmpAge & "  " & _
            '         vbNewLine & "                        DOA: " & txtDateAdmitted & " " & _
            '         vbNewLine & "                        HOSP: " & txtHOSCode & " " & _
            '         vbNewLine & "" & GetServerTime & "   RCVD M/FORM @                      REGISTERED @       " & _
            '         vbNewLine & "" & _
            '         vbNewLine & "                        PYMT MODE:" & tmpMBMPymtMode & " NEXT DUE :                                                                                        " & Logon & "" & _
            '         vbNewLine & "                        RECOVERY: " & _
            '         vbNewLine & "" & _
            '         vbNewLine & "                         EXCLUSION : " & TmpExclusion & "" & _
            '         vbNewLine & "" & _
            '         vbNewLine & "                          " & tmpPolHistory & "" & _
            '         vbNewLine & "" & _
            '         vbNewLine & "                          " & tmpCaseHistory & "" & _
            '         vbNewLine & "" & _
            '         vbNewLine & "                        GROUP CO: " & tmpMBMGroupCompany & "" & _
            '         vbNewLine & "" & _
            '         vbNewLine & "                        PROV DIAG: " & tmpDiagnosis & "" & _
             '        vbNewLine & "                        SYMPTOM: " & _
            '         vbNewLine & "                        DURATION: " & _
            '         vbNewLine & "                        U/ILLNESS: " & _
            '         vbNewLine & "" & _
             '        vbNewLine & "                        ADM KIT:             TTA DRUGS:"
            '
            '
            'tmpRemarks2 = vbNewLine & "" & _
            '         vbNewLine & "                        FINAL DIAG: " & tmpFinalDiagnosis & "" & _
            '         vbNewLine & "                        DR 1) " & txtATTDoc(0) & "" & _
            '         vbNewLine & "                        DR 2) " & txtATTDoc(1) & "" & _
            '         vbNewLine & "                        DR 3) " & txtATTDoc(2) & "" & _
            '         vbNewLine & "                        DR 4) " & txtATTDoc(3) & "" & _
            '         vbNewLine & "                        DR 5) " & txtATTDoc(4) & "" & _
             '        vbNewLine & "" & _
            '         vbNewLine & "                        PROCEDURE:" & _
            '         vbNewLine & "                                             MMA RATE         HOSP   " & _
            '         vbNewLine & "                        S/FEE:    " & _
            '         vbNewLine & "                        A/FEE:   " & _
             '        vbNewLine & "                        DOA : " & txtDateAdmitted & " DOD : " & txtDateDischarge & " "
            'End If
             '   tmpRemarks = tmpRemarks + tmpRemarks2
             '   If Len(tmpRemarks) > 3750 Then
             '       txtCASRemark = Mid(tmpRemarks, 1, 3750)
             '       txtCASRemarkTwo = Mid(tmpRemarks, 3751, 8000)
             '   Else
             '       txtCASRemark = tmpRemarks
             '   End If
            'End If

    'ElseIf SSTab1.Tab = 9 Then
        
    ElseIf SSTab1.Tab = 10 Then
        Call WSListing
    'ElseIf SSTab1.Tab = 11 Then
    '    If tmpPlnCode <> "" Then
    '        frmTermsStatus = 2
    '        Call PolicyTermsListing(tmpPlnCode, frmTermsStatus, txtPayCode, tmpHltCode)
    '    End If
    ' 'ElseIf SSTab1.Tab = 12 Then
     '   If tmpPlnCode <> "" Then
     '       If lstSurgProc.ListItems.Count = 0 Then
     '           frmTermsStatus = 2
     '           Call ScheduleSurg(txtPayCode, tmpPlnCode, frmTermsStatus)
     '       End If
     '   End If
    ElseIf SSTab1.Tab = 11 Then
        Call LoadDoc
        Call LoadIntDoc
        Call LoadWebDoc
    ElseIf SSTab1.Tab = 12 Then
        If Trim(tmpPolicyWording) <> "" Then
            Call LoadPolicy
        End If
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub LoadIntDoc()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
Call addTheIntFiles



'f = FreeFile
'Open App.Path & "\temp.html" For Output As #f
'    Print #f, "<html><body><div align=center><table border=0><tr><td valign=center><center><a href=" & Chr(34) & "Please Click File to view" & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">visit xeek.net</a></td></tr></table></body>"
'Close #f
'WebBrowser1.Navigate App.Path & "\temp.html"
Exit Sub
End Sub

Private Sub INVListing()
Dim INV As New ADODB.Recordset
Dim INVList As ListItem
Dim SqlstrINV As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

If Len(Trim(txtMBMNumber)) <> 0 Then
    If lstINVListing.ListItems.Count = 0 Then
            'SqlstrINV = " Select CASNumber, INVWSversion, InvNo, InvDate, " & _
            "INVPayTo,INVPayee, BordxRefNo, INVAmount,BordxDate,CLMInvPending " & _
            " From VIEW_INV_Listing " & _
            " where CASNumber ='" & Trim(txtCaseNo) & "'"
                
        'INV.Open SqlstrINV, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        strAppendCase = "3"
        strAppend = " AND CASNumber ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        INV.CursorLocation = adUseClient
        INV.CursorType = adOpenForwardOnly
        INV.CacheSize = 100
        Set INV = cmd.Execute
        
        If Not INV.EOF Then
        Do Until INV.EOF
            With INV
                Set INVList = lstINVListing.ListItems.Add(, , !CasNumber & "-" & !INVWSversion)
                    If Len(!INVNo) Then
                        INVList.SubItems(1) = !INVNo
                    End If
                    If Len(!InvDate) Then
                        INVList.SubItems(2) = CDate(!InvDate)
                    End If
                    If Len(!InvPayTo) Then
                        INVList.SubItems(3) = !InvPayTo
                    End If
                    
                    If Len(!INVPayee) Then
                        If !InvPayTo = "H" Then
                            INVList.SubItems(4) = FindHosName(!INVPayee)
                        Else
                            INVList.SubItems(4) = !INVPayee
                        End If
                    End If
                    If Len(!INVAmount) Then
                        INVList.SubItems(5) = !INVAmount
                    End If
                    If Len(!BordxRefNo) Then
                        INVList.SubItems(6) = !BordxRefNo
                    End If
                    If Len(!BordxDate) Then
                        INVList.SubItems(7) = !BordxDate
                    End If
                    If Len(!CLMInvPending) Then
                        INVList.SubItems(8) = !CLMInvPending
                    End If
                .MoveNext
                End With
        Loop
        INV.Close
        Set INV = Nothing
    End If
End If
End If
End Sub
Private Sub WSListingMgmt()
Dim ws As New ADODB.Recordset
Dim WSList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim amntb4GST As Double
Dim GST As Double
Dim amntWithGST As Double


If Len(Trim(txtMBMNumber)) <> 0 Then
    If lstNote1(5).ListItems.Count = 0 Then
            'sql = " Select CASNumber, ClaimVersion, CLMTotalApproved, CLMTotalActual,CLMPending, " & _
            "CLMPayableByMedix2H,CLMPayableByMedix2M, BordxRefNo, ClaimCloseDate, BordxDate " & _
            " From VIEW_WS_Listing " & _
            " where CASNumber ='" & Trim(txtCaseNo) & "'"
                
        'ws.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        amntb4GST = 0
        GST = 0
        
        strAppendCase = "2"
        strAppend = " AND CASNumber ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        ws.CursorLocation = adUseClient
        ws.CursorType = adOpenForwardOnly
        ws.CacheSize = 100
        Set ws = cmd.Execute
        
        If Not ws.EOF Then
        Do Until ws.EOF
            With ws
                Set WSList = lstNote1(5).ListItems.Add(, , !CasNumber & "-" & !claimversion)
                
                    If Len(!BordxRefNo) Then
                        WSList.SubItems(1) = !BordxRefNo
                    End If
                    If Len(!BordxDate) Then
                        WSList.SubItems(2) = !BordxDate
                    End If
                
                    If Len(!ClaimCloseDate) Then
                        WSList.SubItems(3) = !ClaimCloseDate
                    End If
                    If Len(!CLMTotalActual) Then
                        WSList.SubItems(4) = Format(!CLMTotalActual, "#,##0.00")
                        amntb4GST = !CLMTotalActual
                    End If
                    
                    'GST AMount
                    
                    If Len(!TotGST) Then
                        WSList.SubItems(5) = Format(!TotGST, "#,##0.00")
                        GST = !TotGST
                    End If
                    
                   'Total Actual With GST
                    amntWithGST = amntb4GST + GST
                   If Len(amntWithGST) Then
                      WSList.SubItems(6) = Format(amntWithGST, "#,##0.00")
                    End If
                    
                    If Len(!CLMTotalApproved) Then
                        WSList.SubItems(7) = !CLMTotalApproved
                    End If
                    If Len(!CLMPayableByMedix2H) Then
                        WSList.SubItems(8) = !CLMPayableByMedix2H
                    End If
                    If Len(!CLMPayableByMedix2M) Then
                        WSList.SubItems(9) = !CLMPayableByMedix2M
                    End If
                    If Len(!CLMPending) Then
                        WSList.SubItems(10) = !CLMPending
                    End If
                    If Len(!CLMPayorAppStatus) Then
                        WSList.SubItems(11) = !CLMPayorAppStatus
                    End If
                    If Len(!CLMPayorAppDate) Then
                        WSList.SubItems(12) = !CLMPayorAppDate
                    End If
                    If Len(!CLMPayorAppCode) Then
                        WSList.SubItems(13) = !CLMPayorAppCode
                    End If
                    If Len(!CLMBordxRevInd) Then
                        WSList.SubItems(14) = !CLMBordxRevInd
                    End If
                    
                    'shan 18/02/2009
                     'If Trim(!CRTCartonNo) Then
                        'WSList.SubItems(13) = !CRTCartonNo
                    'End If
                    
                    'shan 18/02/2009
                    If Len(!CRTCartonNo) Then
                        WSList.SubItems(15) = !CRTCartonNo
                    End If
                    
                    If Len(!DoSend) Then
                        WSList.SubItems(16) = Format(!DoSend, "dd/mm/yyyy")
                    End If
                    
                    If Len(!DebitCreditRefNo) Then
                        WSList.SubItems(17) = Trim(!DebitCreditRefNo)
                    End If
                    
                    If Len(!DebitCreditDate) Then
                        WSList.SubItems(18) = Format(!DebitCreditDate, "dd/mm/yyyy")
                    End If
                .MoveNext
                End With
        Loop
       
        ws.Close
        Set ws = Nothing
    End If
End If
End If
End Sub

Private Sub ExcessListingMgmt()
Dim ws1 As New ADODB.Recordset
Dim WSList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim amntb4GST As Double
Dim GST As Double
Dim amntWithGST As Double

lstNote1(3).ListItems.Clear

If Len(Trim(txtMBMNumber)) <> 0 Then
    If lstNote1(3).ListItems.Count = 0 Then
            'sql = " Select CASNumber, ClaimVersion, CLMTotalApproved, CLMTotalActual,CLMPending, " & _
            "CLMPayableByMedix2H,CLMPayableByMedix2M, BordxRefNo, ClaimCloseDate, BordxDate " & _
            " From VIEW_WS_Listing " & _
            " where CASNumber ='" & Trim(txtCaseNo) & "'"
                
        'ws.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        amntb4GST = 0
        GST = 0
        
        strAppendCase = "4"
        strAppend = " AND dbo.ClaimMember.CASNumber ='" & Trim(txtCaseNo) & "'"
        'strAppend = " AND CASNumber ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        ws1.CursorLocation = adUseClient
        ws1.CursorType = adOpenForwardOnly
        ws1.CacheSize = 100
        Set ws1 = cmd.Execute
        
        If Not ws1.EOF Then
        Do Until ws1.EOF
            With ws1
            
                            
               
                
                    If !CLMBoardAmt <> "-" Then
                        If !CLMBoardAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                                WSList.SubItems(1) = "R&B"
                                
                                If Len(!CLMBoardAmt) Then
                                    WSList.SubItems(2) = Format(!CLMBoardAmt * -1, "#,##0.00")
                                End If
                            
                        End If
                    End If
                
                
                    If !CLMICUAmt <> "-" Then
                        If !CLMICUAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "ICU"
                            
                            If Len(!CLMICUAmt) Then
                                WSList.SubItems(2) = Format(!CLMICUAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMServices <> "-" Then
                        If !CLMServices < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "HSS"
                            If Len(!CLMServices) Then
                                WSList.SubItems(2) = Format(!CLMServices * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMTheatre <> "-" Then
                        If !CLMTheatre < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "OT"
                            If Len(!CLMTheatre) Then
                                WSList.SubItems(2) = Format(!CLMTheatre * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMPreHSAmt <> "-" Then
                        If !CLMPreHSAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Pre-Hosp"
                            
                            If Len(!CLMPreHSAmt) Then
                                WSList.SubItems(2) = Format(!CLMPreHSAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMSurgicalAmt <> "-" Then
                        If !CLMSurgicalAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Surgical"
                            
                            If Len(!CLMSurgicalAmt) Then
                                WSList.SubItems(2) = Format(!CLMSurgicalAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMAna <> "-" Then
                        If !CLMAna < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Anae"
                            If Len(!CLMAna) Then
                                WSList.SubItems(2) = Format(!CLMAna * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMNSPreHos <> "-" Then
                        If !CLMNSPreHos < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Pre-Hosp"
                            If Len(!CLMNSPreHos) Then
                                WSList.SubItems(2) = Format(!CLMNSPreHos * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMPhysician <> "-" Then
                        If !CLMPhysician < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Physician"
                            If Len(!CLMPhysician) Then
                                WSList.SubItems(2) = Format(!CLMPhysician * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMPost <> "-" Then
                        If !CLMPost < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Post-Hosp"
                            If Len(!CLMPost) Then
                                WSList.SubItems(2) = Format(!CLMPost * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMEmergency <> "-" Then
                        If !CLMEmergency < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Emergengy Accident"
                            If Len(!CLMEmergency) Then
                                WSList.SubItems(2) = Format(!CLMEmergency * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMAmbulance <> "-" Then
                        If !CLMAmbulance < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Ambulance"
                            If Len(!CLMAmbulance) Then
                                WSList.SubItems(2) = Format(!CLMAmbulance * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMOutpatient <> "-" Then
                        If !CLMOutpatient < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Outpatient"
                            If Len(!CLMOutpatient) Then
                                WSList.SubItems(2) = Format(!CLMOutpatient * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    If !CLMMonthly <> "-" Then
                        If !CLMMonthly < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Cancer Treatment"
                            If Len(!CLMMonthly) Then
                                WSList.SubItems(2) = Format(!CLMMonthly * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    If !CLMOrgan <> "-" Then
                        If !CLMOrgan < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Organ Transplant"
                            If Len(!CLMOrgan) Then
                                WSList.SubItems(2) = Format(!CLMOrgan * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    If !CLMMedical <> "-" Then
                        If !CLMMedical < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Medical Report"
                            If Len(!CLMMedical) Then
                                WSList.SubItems(2) = Format(!CLMMedical * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                         
                    If !CLMPhone <> "-" Then
                        If !CLMPhone < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Telephone"
                            If Len(!CLMPhone) Then
                                WSList.SubItems(2) = Format(!CLMPhone * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                 
                    
                    If !CLMGuardianAmt <> "-" Then
                        If !CLMGuardianAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Guardian Allowance"
                            If Len(!CLMGuardianAmt) Then
                                WSList.SubItems(2) = Format(!CLMGuardianAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMAdmin <> "-" Then
                        If !CLMAdmin < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Admin"
                            If Len(!CLMAdmin) Then
                                WSList.SubItems(2) = Format(!CLMAdmin * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMTax <> "-" Then
                        If !CLMTax < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Tax"
                            If Len(!CLMTax) Then
                                WSList.SubItems(2) = Format(!CLMTax * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                 
                    If !CLMOthers1 <> "-" Then
                        If !CLMOthers1 < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Others" & "-" & !CLMOther1Desc
                            If Len(!CLMOthers1) Then
                                WSList.SubItems(2) = Format(!CLMOthers1 * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMOthers2 <> "-" Then
                        If !CLMOthers2 < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Others" & "-" & !CLMOthers2Desc
                            If Len(!CLMOthers2) Then
                                WSList.SubItems(2) = Format(!CLMOthers2 * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    If !CLMCashAmt <> "-" Then
                        If !CLMCashAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Cash Allowance"
                            If Len(!CLMCashAmt) Then
                                WSList.SubItems(2) = Format(!CLMCashAmt * -1, "#,##0.00")
                            End If
                        End If
                    End If
                    
                    If !CLMBoardTaxAmt <> "-" Then
                        If !CLMBoardTaxAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "R&B Tax"
                            If Len(!CLMBoardTaxAmt) Then
                                WSList.SubItems(2) = Format(!CLMBoardTaxAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMICUTaxAmt <> "-" Then
                        If !CLMICUTaxAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "ICU Tax"
                            If Len(!CLMICUTaxAmt) Then
                                WSList.SubItems(2) = Format(!CLMICUTaxAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMLodgerTaxAmt <> "-" Then
                        If !CLMLodgerTaxAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Lodger Tax"
                            If Len(!CLMLodgerTaxAmt) Then
                                WSList.SubItems(2) = Format(!CLMLodgerTaxAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMMRI <> "-" Then
                        If !CLMMRI < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "MRI"
                            If Len(!CLMMRI) Then
                                WSList.SubItems(2) = Format(!CLMMRI * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                
                    
                    If !CLMMPreHSAmt <> "-" Then
                        If !CLMMPreHSAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Pre-Hosp"
                            If Len(!CLMMPreHSAmt) Then
                                WSList.SubItems(2) = Format(!CLMMPreHSAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If IsNumeric(!CLMCOPreHSSum) = True Then
                        If !CLMCOPreHSSum < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Pre-Hosp"
                            If Len(!CLMCOPreHSSum) Then
                                WSList.SubItems(2) = Format(!CLMCOPreHSSum * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If IsNumeric(!CLMCOMPreHSSum) = True Then
                        If !CLMCOMPreHSSum < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Pre-Hosp"
                            If Len(!CLMCOMPreHSSum) Then
                                WSList.SubItems(2) = Format(!CLMCOMPreHSSum * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMAccPostAmt <> "-" Then
                        If !CLMAccPostAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Accident"
                            If Len(!CLMAccPostAmt) Then
                                WSList.SubItems(2) = Format(!CLMAccPostAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMPreHosp <> "-" Then
                        If !CLMPreHosp < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Pre-Hosp"
                            If Len(!CLMPreHosp) Then
                                WSList.SubItems(2) = Format(!CLMPreHosp * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMOthersNotCoPay <> "-" Then
                        If !CLMOthersNotCoPay < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Others (Not Copay)"
                            If Len(!CLMOthersNotCoPay) Then
                                WSList.SubItems(2) = Format(!CLMOthersNotCoPay * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMHomeNCAmt <> "-" Then
                        If !CLMHomeNCAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Home Nursing"
                            If Len(!CLMHomeNCAmt) Then
                                WSList.SubItems(2) = Format(!CLMHomeNCAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    If !CLMHomeNCAmt <> "-" Then
                        If !CLMHomeNCAmt < "0" Then
                            Set WSList = lstNote1(3).ListItems.Add(, , !CasNumber & "-" & !clmVersion)
                            WSList.SubItems(1) = "Home Nursing"
                            If Len(!CLMHomeNCAmt) Then
                                WSList.SubItems(2) = Format(!CLMHomeNCAmt * -1, "#,##0.00")
                            End If
                            
                        End If
                    End If
                    
                    'dbo.ClaimMember. , dbo.ClaimMember.CLMIndKidneyDial, dbo.ClaimMember.CLMIndCancerTreatment,
                .MoveNext
                End With
        Loop
       
        ws1.Close
        Set ws1 = Nothing
    End If
End If
End If
End Sub

Private Sub ReceiptFRMBMListingMgmt()
Dim rst As New ADODB.Recordset
Dim masterlist As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

If Len(Trim(txtCaseNo)) <> 0 Then
   ' If lstNote1(7).ListItems.Count = 0 Then
        'sql = " Select CASNumber, ClaimVersion, ReceiptNo, ReceiptDate, " & _
              " ReceiptChequeNo,ReceiptAmt,MBMRemarks " & _
              " From ReceiptFrMBM" & _
              " where CASNumber ='" & Trim(txtCaseNo) & "'"
                
        'rst.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
        strAppendCase = "4"
        strAppend = " AND CASNumber ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst.CursorLocation = adUseClient
        rst.CursorType = adOpenForwardOnly
        rst.CacheSize = 100
        Set rst = cmd.Execute
        
        If Not rst.EOF Then
         
            Do Until rst.EOF
                With rst
            
                    Set masterlist = lstNote1(7).ListItems.Add(, , "Receipt")
                    
                    masterlist.SubItems(1) = !CasNumber & "-" & !claimversion
                    If Len(!ReceiptNo) Then
                        masterlist.SubItems(2) = !ReceiptNo
                    End If
                    If Len(!ReceiptDate) Then
                        masterlist.SubItems(3) = !ReceiptDate
                    End If
                    If Len(!ReceiptChequeNo) Then
                        masterlist.SubItems(4) = !ReceiptChequeNo
                    End If
                    If Len(!ReceiptAmt) Then
                        masterlist.SubItems(5) = !ReceiptAmt
                    End If
                    masterlist.SubItems(6) = ""
                    If Len(!MBMRemarks) Then
                        masterlist.SubItems(7) = !MBMRemarks
                    End If
                    'If Len(!MBMRemarks) Then
                    '    masterlist.SubItems(7) = !MBMRemarks
                    'End If
                         
                .MoveNext
                End With
            Loop
       
        rst.Close
        Set rst = Nothing
        End If
    'End If
End If

End Sub
Private Sub ReceiptFRMPYRListingMgmt()
Dim rst As New ADODB.Recordset
Dim masterlist As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

If Len(Trim(txtCaseNo)) <> 0 Then
    'If lstNote1(7).ListItems.Count = 0 Then
        'sql = " Select CASNumber, ClaimVersion, ReceiptNo, ReceiptDate, " & _
              " ReceiptChequeNo,ReceiptAmt,MBMRemarks " & _
              " From ReceiptFrMBM" & _
              " where CASNumber ='" & Trim(txtCaseNo) & "'"
                
        'rst.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
        strAppendCase = "5"
        strAppend = " AND CASNumber ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst.CursorLocation = adUseClient
        rst.CursorType = adOpenForwardOnly
        rst.CacheSize = 100
        Set rst = cmd.Execute
        
        If Not rst.EOF Then
         
            Do Until rst.EOF
                With rst
            
                    Set masterlist = lstNote1(7).ListItems.Add(, , "Receipt")
                    
                    masterlist.SubItems(1) = !CasNumber & "-" & !claimversion
                    If Len(!ReceiptNo) Then
                        masterlist.SubItems(2) = !ReceiptNo
                    End If
                    If Len(!ReceiptDate) Then
                        masterlist.SubItems(3) = !ReceiptDate
                    End If
                    If Len(!ChequeNo) Then
                        masterlist.SubItems(4) = !ChequeNo
                    End If
                    If Len(!CheckAmt) Then
                        masterlist.SubItems(5) = !CheckAmt
                    End If
                    masterlist.SubItems(6) = ""
                    If Len(!Remark) Then
                        masterlist.SubItems(7) = !Remark
                    End If
                    'If Len(!Remark) Then
                    '    masterlist.SubItems(7) = !Remark
                    'End If
                    '
                .MoveNext
                End With
            Loop
       
        rst.Close
        Set rst = Nothing
        End If
    'End If
End If

End Sub
Private Sub PaymentHOSListingMgmt()
Dim PH As New ADODB.Recordset
Dim PHList As ListItem
Dim sql As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

'If Len(Trim(txtCaseNo)) <> 0 Then
    'If lstNote1(7).ListItems.Count = 0 Then
            'sql = " Select CASNumber, ClaimVersion, PVHOSNo, PVHOSDate, PVStatus," & _
            " PVHOSChequeNo,PVHOSChequeDate,HOSPaidAmt, PVPayHOSStatus,HOSRemark " & _
            " From Payment2HOS " & _
            " where CASNumber ='" & Trim(txtCaseNo) & "'"
                
        'PH.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
        strAppendCase = "1"
        strAppend = " AND CASNumber ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        PH.CursorLocation = adUseClient
        PH.CursorType = adOpenForwardOnly
        PH.CacheSize = 100
        Set PH = cmd.Execute
        
        If Not PH.EOF Then
        Do Until PH.EOF
        With PH
            Set PHList = lstNote1(7).ListItems.Add(, , "Payment")
                     
                     'If Len(!PVHOSNo) Then
                         PHList.SubItems(1) = !CasNumber & "-" & !claimversion
                     'End If
                     
                     If Len(!PVHOSNo) Then
                         PHList.SubItems(2) = !PVHOSNo 'Format(!CLMTotalActual, "#,##0.00")
                     End If
                     If Len(!PVHOSDate) Then
                         PHList.SubItems(3) = !PVHOSDate
                     End If
                     If Len(!PVHOSChequeNo) Then
                         PHList.SubItems(4) = !PVHOSChequeNo
                     End If
                     If Len(!HOSPaidAmt) Then
                         PHList.SubItems(5) = Format(!HOSPaidAmt, "#,##0.00")
                     End If
                     If Len(!PVPayHOSStatus) Then
                         PHList.SubItems(6) = !PVPayHOSStatus
                     End If
                     
                    If Len(!HOSRemark) Then
                        PHList.SubItems(7) = !HOSRemark
                     End If

                     'If Len(!HOSRemark) Then
                     '   PHList.SubItems(7) = Format(!HOSRemark, "#,##0.00")
                     'End If
                     'If Len(!HOSRemark) Then
                     '   PHList.SubItems(8) = Format(!HOSRemark, "#,##0.00")
                     'End If
                .MoveNext
                End With
        Loop
        PH.Close
        Set PH = Nothing
    End If
'End If
'End If
End Sub
Private Sub PaymentMBMListingMgmt()
Dim PM As New ADODB.Recordset
Dim PMList As ListItem
Dim sql As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

If Len(Trim(txtCaseNo)) <> 0 Then
    'If lstNote1(7).ListItems.Count = 0 Then
            'sql = " Select CASNumber, ClaimVersion, PVNo, PVDate, " & _
            " PVChequeNo,PVChequeDate,MBMPaidAmt, PVPayMBMStatus,Remark,PVStatus " & _
            " From Payment2MBM " & _
            " where CASNumber ='" & Trim(txtCaseNo) & "'"
                
        'PM.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
        strAppendCase = "2"
        strAppend = " AND CASNumber ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        PM.CursorLocation = adUseClient
        PM.CursorType = adOpenForwardOnly
        PM.CacheSize = 100
        Set PM = cmd.Execute
        
        If Not PM.EOF Then
        Do Until PM.EOF
        With PM
            Set PMList = lstNote1(7).ListItems.Add(, , "Payment")
                     
                     PMList.SubItems(1) = !CasNumber & "-" & !claimversion
                     If Len(!PVNo) Then
                         PMList.SubItems(2) = !PVNo 'Format(!CLMTotalActual, "#,##0.00")
                     End If
                     If Len(!PVDate) Then
                         PMList.SubItems(3) = !PVDate
                     End If
                     If Len(!PVChequeNo) Then
                         PMList.SubItems(4) = !PVChequeNo
                     End If
                     If Len(!MBMPaidAmt) Then
                         PMList.SubItems(5) = Format(!MBMPaidAmt, "#,##0.00")
                     End If
                     If Len(!PVPayMBMStatus) Then
                         PMList.SubItems(6) = !PVPayMBMStatus
                     End If
                     If Len(!Remark) Then
                        PMList.SubItems(7) = !Remark
                    End If
                    'If Len(!PVStatus) Then
                    '    PMList.SubItems(7) = !Remark
                    'End If
                    
                    'If Len(!Remark) Then
                    '    PMList.SubItems(8) = !Remark
                    'End If

                .MoveNext
                End With
        Loop
        PM.Close
        Set PM = Nothing
    End If
'End If
End If
End Sub

Sub addTheIntFiles()
Dim B As String
List2.Clear

PictureFileName = DocPath & "InternalDocuments\" & txtCaseNo & ""
'\Medical Rpt"
B = Dir(PictureFileName & "\*.*")
'b = Dir(Dir1.Path & "\*.*")
If B <> "" Then
    If Right(B, 4) = ".jpg" Then
        List2.AddItem B
    End If
    If Right(B, 4) = ".tif" Then
        List2.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".gif" Then
        List2.AddItem B
    End If
    
    If Right(B, 4) = ".JPG" Then
        List2.AddItem B
    End If
    If Right(B, 4) = ".TIF" Then
        List2.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".GIF" Then
        List2.AddItem B
    End If
    If Right(B, 4) = ".xls" Then
        List2.AddItem B
    End If
    If Right(B, 4) = ".doc" Then
        List2.AddItem B
    End If

    If Right(B, 4) = ".pdf" Then
        List2.AddItem B
    End If
Else
    Exit Sub
End If
Do: DoEvents
    B = Dir
    If B <> "" Then
        If Right(B, 4) = ".jpg" Then
            List2.AddItem B
        End If
        If Right(B, 4) = ".tif" Then
            List2.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".gif" Then
            List2.AddItem B
        End If
        
        If Right(B, 4) = ".JPG" Then
            List2.AddItem B
        End If
        If Right(B, 4) = ".TIF" Then
            List2.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".GIF" Then
            List2.AddItem B
        End If
        
        If Right(B, 4) = ".xls" Then
            List2.AddItem B
        End If
        
        If Right(B, 4) = ".doc" Then
            List2.AddItem B
        End If
        
        If Right(B, 4) = ".pdf" Then
            List2.AddItem B
        End If
    Else
        Exit Do
    End If
Loop
'List1.Height = Me.Height - 2520 - 350
End Sub
'Private Sub LoadEmails()
'On Error Resume Next
'Dim B As String, f As Integer, D As Integer, e As Integer
'Call addemails


Private Sub LoadPolicy()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
Call addPolicy



'f = FreeFile
'Open App.Path & "\temp.html" For Output As #f
'    Print #f, "<html><body><div align=center><table border=0><tr><td valign=center><center><a href=" & Chr(34) & "Please Click File to view" & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">visit xeek.net</a></td></tr></table></body>"
'Close #f
'WebBrowser1.Navigate App.Path & "\temp.html"
Exit Sub
End Sub

Sub addPolicy()
Dim B As String
lstPolicyWording.Clear

PictureFileName = DocPath & "PolicyWording\" & tmpPolicyWording & ""
'\Medical Rpt"
B = Dir(PictureFileName & "\*.*")
'b = Dir(Dir1.Path & "\*.*")
If B <> "" Then
    If Right(B, 4) = ".jpg" Then
        lstPolicyWording.AddItem B
    End If
    If Right(B, 4) = ".tif" Then
        lstPolicyWording.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".gif" Then
        lstPolicyWording.AddItem B
    End If
    
    If Right(B, 4) = ".JPG" Then
        lstPolicyWording.AddItem B
    End If
    If Right(B, 4) = ".TIF" Then
        lstPolicyWording.AddItem B
    End If
    'If Right(b, 4) = ".art" Then
    '    List1.AddItem b
    'End If
    If Right(B, 4) = ".GIF" Then
        lstPolicyWording.AddItem B
    End If
    
Else
    Exit Sub
End If
Do: DoEvents
    B = Dir
    If B <> "" Then
        If Right(B, 4) = ".jpg" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".tif" Then
            lstPolicyWording.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".gif" Then
            lstPolicyWording.AddItem B
        End If
        
        If Right(B, 4) = ".JPG" Then
            lstPolicyWording.AddItem B
        End If
        If Right(B, 4) = ".TIF" Then
            lstPolicyWording.AddItem B
        End If
        'If Right(b, 4) = ".art" Then
        '    List1.AddItem b
        'End If
        If Right(B, 4) = ".GIF" Then
            lstPolicyWording.AddItem B
        End If
        
    Else
        Exit Do
    End If
Loop
'List1.Height = Me.Height - 2520 - 350
End Sub
Private Sub MBMHistory(TmpMBMType, SearchByICNo As Boolean)
Dim TmpREc1 As New ADODB.Recordset
Dim HISList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpName As String
Dim tmpDOb As String
Dim cnt As Integer
Dim tmpPolHistory1 As String
    tmpPolHistory = ""
    tmpPolHistory1 = ""
    tmpName = Trim(Replace(txtMBMName, "'", "''"))
    'tmpDOb = Format(txtMBMBirthDate, "yyyy/mm/dd")
    cnt = 0
    strAppend = ""

        If SearchByICNo = True Then
            strAppend = " And HISMaintenance.dbo.MBMCrossReference.MBMIcBcPp = '" & Trim(txtMBMIcBcPP) & "' And HISMaintenance.dbo.MBMCrossReference.MBMStatus <> 'D'"
        End If
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchMBMHistory06"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("MBMType", adVarChar, adParamInput, 10, TmpMBMType)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm2

    TmpREc1.CursorLocation = adUseClient
    TmpREc1.CursorType = adOpenForwardOnly
    TmpREc1.CacheSize = 100
    Set TmpREc1 = cmd.Execute
    Do While Not TmpREc1.EOF
        cnt = cnt + 1
        If cnt = 1 Then
            tmpPolHistory = "" & _
                vbNewLine & "                         Policy History" & _
                vbNewLine & "                         " & TmpREc1!MBMNumber & " - " & TmpREc1!MBMPolicyNo & " - " & TmpREc1!MBMPayorEffDate & " - " & TmpREc1!MBMPayorEXPDate & ""
        Else
            tmpPolHistory1 = ""
            
            tmpPolHistory1 = "" & _
                vbNewLine & "                         " & TmpREc1!MBMNumber & " - " & TmpREc1!MBMPolicyNo & " - " & TmpREc1!MBMPayorEffDate & " - " & TmpREc1!MBMPayorEXPDate & ""
             tmpPolHistory = tmpPolHistory + tmpPolHistory1
        End If
        TmpREc1.MoveNext
    Loop
    TmpREc1.Close
    Set TmpREc1 = Nothing

End Sub
Private Sub MBMHistoryMgmt(TmpMBMType)
Dim TmpREc1 As New ADODB.Recordset
Dim HISList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpName As String
Dim strAppend As String
Dim tmpDOb As String

    'tmpName = Trim(Replace(txtMBMName, "'", "''"))
    'tmpDOb = Format(txtMBMBirthDate, "yyyy/mm/dd")
    If TmpMBMType = "Princ" Then
        'If SearchByICNo = True Then
            'strAppend = " And Replace(HISMaintenance.dbo.MBMCrossReference.MBMIcBcPp,'-','') = Replace('" & Trim(txtMBMIcBcPP) & "','-','')"
            strAppend = " And HISMaintenance.dbo.MBMCrossReference.MBMIcBcPp = '" & Trim(txtMBMIcBcPP) & "'"
        'Else
        '    strAppend = " And HISMaintenance.dbo.MBMCrossReference.MBMName = '" & tmpName
            'strAppend = strAppend & "' And HISMaintenance.dbo.MBMCrossReference.MBMDOB = '" & tmpDOb & "'"
        'End If
    Else
        'If SearchByICNo = True Then
            'strAppend = " And Replace(dbo.MBMCoveredPersons.MBMCICBCPPNo,'-','') = Replace('" & Trim(txtMBMIcBcPP) & "','-','')"
            strAppend = " And dbo.MBMCoveredPersons.MBMCICBCPPNo = '" & Trim(txtMBMIcBcPP) & "'"
        'Else
        '    strAppend = " And dbo.MBMCoveredPersons.MBMCName = '" & tmpName
        '    strAppend = strAppend & "' And dbo.MBMCoveredPersons.MBMCDOB='" & tmpDOb & "'"
        'End If
    End If
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchMBMHistory06"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("MBMType", adVarChar, adParamInput, 10, TmpMBMType)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm2

    TmpREc1.CursorLocation = adUseClient
    TmpREc1.CursorType = adOpenForwardOnly
    TmpREc1.CacheSize = 100
    Set TmpREc1 = cmd.Execute
    Do While Not TmpREc1.EOF
        Set HISList = lstNote1(0).ListItems.Add(, , TmpREc1!MBMNumber)
        HISList.SubItems(1) = IIf(Len(TmpREc1!MBMCOVID), TmpREc1!MBMCOVID, "")
        HISList.SubItems(2) = IIf(Len(TmpREc1!MBMPolicyNo), TmpREc1!MBMPolicyNo, "")
        
        If TmpMBMType = "Princ" Then
            HISList.SubItems(3) = IIf(Len(TmpREc1!MBMName), TmpREc1!MBMName, "")
        Else
            HISList.SubItems(3) = IIf(Len(TmpREc1!MBMCName), TmpREc1!MBMCName, "")
        End If
        'If TmpMBMType = "Princ" Then
            HISList.SubItems(4) = IIf(TmpREc1!MBMPayorEffDate <> "", TmpREc1!MBMPayorEffDate, "")
            HISList.SubItems(5) = IIf(TmpREc1!MBMPayorEXPDate <> "", TmpREc1!MBMPayorEXPDate, "")
       ' Else
        '    If TmpREc1!MBMPayorEffDate = "" Then
        '        HISList.SubItems(4) = IIf(TmpREc1!MBMPayorEffDate <> "", TmpREc1!MBMPayorEffDate, "")
        '    Else
        '        HISList.SubItems(4) = IIf(TmpREc1!MBMPayorEffDate <> "", TmpREc1!MBMPayorEffDate, "")
        '    End If
        '    HISList.SubItems(5) = IIf(TmpREc1!MBMPayorEXPDate <> "", TmpREc1!MBMPayorEXPDate, "")
        'End If
        HISList.SubItems(6) = IIf(Len(TmpREc1!MBMDataStatus), TmpREc1!MBMDataStatus, "")
        HISList.SubItems(7) = IIf(Len(TmpREc1!MBMStatus), TmpREc1!MBMStatus, "")
        'HISList.SubItems(8) = IIf(Len(TmpREc1!MBMBordxDate), TmpREc1!MBMBordxDate, "")
        TmpREc1.MoveNext
    Loop
    TmpREc1.Close
    Set TmpREc1 = Nothing

End Sub

Private Sub SSTab1_DblClick()
    If SSTab1.Tab = 7 Then
       ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Call LGListing
        'Call FollowUpListing
      '  medixmasterconnMaint.Close
      '  Set medixmasterconnMaint = Nothing
    End If
End Sub



Private Sub Text6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)

    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesActFees1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesActFees2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesActFees3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesActFees4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesActFees5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesActFees6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesActFees1), txtAnaesActFees1, 0)
    B = IIf(IsNumeric(txtAnaesActFees2), txtAnaesActFees2, 0)
    C = IIf(IsNumeric(txtAnaesActFees3), txtAnaesActFees3, 0)
    d = IIf(IsNumeric(txtAnaesActFees4), txtAnaesActFees4, 0)
    e = IIf(IsNumeric(txtAnaesActFees5), txtAnaesActFees5, 0)
    g = IIf(IsNumeric(txtAnaesActFees6), txtAnaesActFees6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesMMARate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesMMARate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesMMARate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesMMARate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesMMARate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAnaesMMARate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtAnaesMMARate1), txtAnaesMMARate1, 0)
    B = IIf(IsNumeric(txtAnaesMMARate2), txtAnaesMMARate2, 0)
    C = IIf(IsNumeric(txtAnaesMMARate3), txtAnaesMMARate3, 0)
    d = IIf(IsNumeric(txtAnaesMMARate4), txtAnaesMMARate4, 0)
    e = IIf(IsNumeric(txtAnaesMMARate5), txtAnaesMMARate5, 0)
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtAppDate_LostFocus()
    If txtAppDate <> "__/__/____" Then
        If IsDate(txtAppDate) = False Then
            MsgBox "Invalid date format!", vbCritical
        End If
    End If
End Sub

Private Sub txtCaseNo_GotFocus()
    cmdShow.Default = True
    cmdUpdate1.Enabled = False
End Sub

Private Sub TxtCaseNo_LostFocus()
txtCaseNo = ChkStr(txtCaseNo)
cmdShow.Default = False
End Sub

Private Sub txtCASRemark_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub

Private Sub txtCASRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 126 Then
        MsgBox "You are not allow to use this symbol as a Keyword", vbOKOnly + vbCritical, DialogBoxTitle
        KeyAscii = 0
    End If
End Sub

Private Sub txtCasREmark_LostFocus()
    KeyPreview = True
    cmdShow.Default = True
End Sub


Private Sub txtCASRemark3_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub



Private Sub txtCASRemark4_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub

Private Sub txtCASRemark5_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub

Private Sub txtCASRemarkTwo_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub

Private Sub txtCASReserveAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtCASLGAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub



Private Sub txtCLMAppCode_LostFocus()
    If Trim(Len(txtCLMAppCode)) Then
        cboAppStatus = "Y"
    Else
        cboAppStatus = "N"
    End If
End Sub

Private Sub txtDateAdmitted_GotFocus()
    PrevDOA = txtDateAdmitted
End Sub

Private Sub txtDateAdmitted_LostFocus()
    If IsDate(txtDateAdmitted) Then
        If (txtDateDischarge <> "__/__/____") Then
            If (CDate(txtDateDischarge) < CDate(txtDateAdmitted)) Then
                MsgBox "Admission Date cannot greater than Discharge Date! ", vbCritical
                txtDateAdmitted = PrevDOA
                txtDateAdmitted.SetFocus
            End If
        End If
        If txtDateAdmitted <> "__/__/____" Then
            If CDate(txtDateAdmitted) > Date And ChkDOA.Value = 1 Then
                MsgBox "Admission Date cannot be greater than system date! ", vbCritical
                txtDateAdmitted.SetFocus
            End If
        End If
        If (txtDateAdmitted <> "__/__/____") Then
            If (CDate(txtDateAdmitted) < CDate(txtMBMEffDate)) Then
                MsgBox "Admission Date cannot be less than Effective Date! ", vbCritical
                'txtDateAdmitted = PrevDOA
                txtDateAdmitted.SetFocus
            End If
        End If
        
        If (txtDateAdmitted <> "__/__/____") Then
            If Mid(CboPayType, 1, 1) = "C" Then
                If (CDate(txtMBMExpiryDate) < CDate(txtDateAdmitted)) Then
                    MsgBox "Admission Date cannot be greater than Expiry Date! ", vbCritical
                    'txtDateAdmitted = PrevDOA
                    txtDateAdmitted.SetFocus
                End If
            End If
        End If
        
    Else
        If txtDateAdmitted <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateAdmitted.SetFocus
        End If
    End If
End Sub

Private Sub txtDateDischarge_LostFocus()
    If IsDate(txtDateDischarge) Then
        If txtDateAdmitted <> "__/__/____" Then
            If (CDate(txtDateDischarge) < CDate(txtDateAdmitted)) Then
                MsgBox "Discharged Date cannot less then Admission Date ! ", vbCritical
                txtDateDischarge.SetFocus
            End If
        End If
        If txtDateDischarge <> "__/__/____" Then
            If CDate(txtDateDischarge) > CDate(GetServerTime) And ChkDOD.Value = 1 Then
                MsgBox "Discharge Date cannot greater than system date! ", vbCritical
                txtDateDischarge.SetFocus
            End If
        End If
        
        If (txtDateDischarge <> "__/__/____") Then
            If (CDate(txtDateDischarge) < CDate(txtMBMEffDate)) Then
                MsgBox "Discharge Date cannot be less than Effective Date! ", vbCritical
                'txtDateAdmitted = PrevDOA
                txtDateDischarge.SetFocus
            End If
        End If
        
        
    Else
        If txtDateDischarge <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateDischarge.SetFocus
        End If
    End If
End Sub

Private Sub txtDateRECDoc_LostFocus()
    If IsDate(txtDateRECDoc) Then
    Else
    If txtDateRECDoc <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtDateRECDoc.SetFocus
    End If
    End If
End Sub

Private Sub txtFDDOS1_LostFocus()
    If IsDate(txtFDDOS1) Then
    Else
    If txtFDDOS1 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtFDDOS1.SetFocus
    End If
    End If
End Sub

Private Sub txtFDDOS2_LostFocus()
    If IsDate(txtFDDOS2) Then
    Else
    If txtFDDOS2 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtFDDOS2.SetFocus
    End If
    End If
End Sub

Private Sub txtFDDOS3_LostFocus()
    If IsDate(txtFDDOS3) Then
    Else
    If txtFDDOS3 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtFDDOS3.SetFocus
    End If
    End If
End Sub

Private Sub txtFirstCall_LostFocus()
    If IsDate(txtFirstCall) Then
    Else
    If txtFirstCall <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtFirstCall.SetFocus
    End If
    End If
End Sub

Private Sub txtFirstDocDate_LostFocus()
    If IsDate(txtFirstDocDate) Then
        If txtSubDocDate <> "__/__/____" Then
            If (CDate(txtSubDocDate) < CDate(txtFirstDocDate)) Then
                MsgBox "First Doc. Date cannot greater than subsequent Doc. Date! ", vbCritical
                txtFirstDocDate.SetFocus
            End If
        End If
    Else
        If txtFirstDocDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtFirstDocDate.SetFocus
        End If
    End If
End Sub

Private Sub txtFirstTreated_LostFocus()
    If IsDate(txtFirstTreated) Then
    Else
    If txtFirstTreated <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtFirstTreated.SetFocus
    End If
    End If
End Sub

Private Sub txtHOSCode_LostFocus()
Dim tmpchkHopsCode As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpAXAClerical As String
Dim tmpPrivate As Boolean
    
    tmpPrivate = True
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If (txtHOSCode = "DENTAL CLINIC" Or txtHOSCode = "DIALYSIS CENTRE" Or txtHOSCode = "FOREIGN HOSPITAL" Or txtHOSCode = "GENERAL PRACTITIONER" Or txtHOSCode = "GOVERNMENT HOSPITAL" Or txtHOSCode = "NON-PANEL HOSPITAL" Or txtHOSCode = "SPECIALIST CLINIC") And Trim(txtHOSName) = "" Then
        MsgBox "Please Enter Hospital Name or Clinic Name!", vbOKOnly + vbInformation
        txtHOSName.Enabled = True
    ElseIf (txtHOSCode <> "DENTAL CLINIC" And txtHOSCode <> "DIALYSIS CENTRE" And txtHOSCode <> "FOREIGN HOSPITAL" And txtHOSCode <> "GENERAL PRACTITIONER" And txtHOSCode <> "GOVERNMENT HOSPITAL" And txtHOSCode <> "NON-PANEL HOSPITAL" And txtHOSCode <> "SPECIALIST CLINIC") And Trim(txtHOSName) <> "" Then
        txtHOSName.Enabled = False
        txtHOSName = ""
    'Else
    
    End If
    
    If (txtHOSCode = "DENTAL CLINIC" Or txtHOSCode = "DIALYSIS CENTRE" Or txtHOSCode = "FOREIGN HOSPITAL" Or txtHOSCode = "GENERAL PRACTITIONER" Or txtHOSCode = "GOVERNMENT HOSPITAL" Or txtHOSCode = "NON-PANEL HOSPITAL" Or txtHOSCode = "SPECIALIST CLINIC") Then
    
    Else
        If Mid(CboPayType, 1, 1) = "C" Then
            tmpchkHopsCode = HOSArr.Item(txtHOSCode)
            strsqlHOS = ""
    
            strsqlHOS = "select * from HOS where HOSCode = '" & Trim(tmpchkHopsCode) & "'"
            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                If HOS!HOSStatus = "P" Then
                    tmpPrivate = True
                Else
                    tmpPrivate = False
                End If
            HOS.Close
            Set HOS = Nothing
            
            If tmpPrivate = True Then
                If tmpHospChk = "Y" Then
                    If txtHOSCode <> "" Then
                        'tmpchkHopsCode = HOSArr.Item(txtHOSCode)
                        
                        If txtPayCode = "AX" And tmptxtMBMPolNo = "03215666" Then
                            tmpAXAClerical = "AX - AXA AFFIN STAFF (NON-EXECUTIVE)"
                            strsqlHOS = ""
                            strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(txtPayCode) & "' and HosPanelCode = '" & tmpAXAClerical & "'"
                            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                        ElseIf txtPayCode = "AX" And (tmptxtMBMPolNo = "03215814" Or tmptxtMBMPolNo = "03215886") Then
                            tmpAXAClerical = "AX - SPNB"
                            strsqlHOS = ""
                            strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(txtPayCode) & "' and HosPanelCode = '" & tmpAXAClerical & "'"
                            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                        ElseIf txtPayCode = "AX" And tmptxtMBMPolNo = "03312049" Then
                            tmpAXAClerical = "AX - EMBASSY OF LIBYA (STAFF&MT)"
                            strsqlHOS = ""
                            strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(txtPayCode) & "' and HosPanelCode = '" & tmpAXAClerical & "'"
                            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
                        
                        ElseIf tmpAXAClerical = "AX - AXA AFFIN" Then
                        
                            
                            strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(txtPayCode) & "' and HosPanelCode = '" & tmpAXAClerical & "'"
                            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                        Else
                        
                            strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(txtPayCode) & "'"
                            HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                            
                        End If
                        
                        If HOS.EOF = True Then
                            MsgBox "The selected hospital is not on the panel!", vbOKOnly + vbInformation
                            cmdLG.Enabled = False
                        Else
                            cmdLG.Enabled = True
                        End If
                        HOS.Close
                        Set HOS = Nothing
                    End If
                End If
            End If
        End If
    End If

End Sub


Private Sub txtHOSName_LostFocus()
    txtHOSName = ChkStr(txtHOSName)
End Sub

Private Sub txtInvRemark_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub

Private Sub txtInvRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 126 Then
        MsgBox "You are not allow to use this symbol as a Keyword", vbOKOnly + vbCritical, DialogBoxTitle
        KeyAscii = 0
    End If
End Sub

Private Sub txtLGTime_LostFocus()
    If isTime(txtLGTime) Then
    
    Else
        If txtLGTime <> "__:__" Then
            MsgBox "Invalid Time Format! ", vbCritical
            txtLGTime.SetFocus
        End If
    End If
End Sub



Private Sub txtMBMPolYr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNSActFee1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtNSActFee2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtNSActFee3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtNSActFee4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtNSActFee5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtNSActFee6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSActFee1), txtNSActFee1, 0)
    B = IIf(IsNumeric(txtNSActFee2), txtNSActFee2, 0)
    C = IIf(IsNumeric(txtNSActFee3), txtNSActFee3, 0)
    d = IIf(IsNumeric(txtNSActFee4), txtNSActFee4, 0)
    e = IIf(IsNumeric(txtNSActFee5), txtNSActFee5, 0)
    g = IIf(IsNumeric(txtNSActFee6), txtNSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtNSMMARate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)

    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtNSMMARate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)

    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)

End Sub

Private Sub txtNSMMARate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)

    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)

End Sub

Private Sub txtNSMMARate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)

    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)

End Sub

Private Sub txtNSMMARate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)

    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)

End Sub

Private Sub txtNSMMARate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtNSMMARate1), txtNSMMARate1, 0)
    B = IIf(IsNumeric(txtNSMMARate2), txtNSMMARate2, 0)
    C = IIf(IsNumeric(txtNSMMARate3), txtNSMMARate3, 0)
    d = IIf(IsNumeric(txtNSMMARate4), txtNSMMARate4, 0)
    e = IIf(IsNumeric(txtNSMMARate5), txtNSMMARate5, 0)
    g = IIf(IsNumeric(txtNSMMARate6), txtNSMMARate6, 0)

    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)

End Sub

Private Sub txtNSProcCode1_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtNSProcCode1)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtNSProcCode1)
            Else
                Call SearchMMARateCal(txtNSProcCode1)
            End If
            txtNSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtNS13Rate1 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtNS13MinRate1 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
                Call Search14thRate(txtNSProcCode1)
                Text1(11).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(0).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
        Else
            txtNSMMARate1 = 0
            txtNS13Rate1 = 0
            txtNS13MinRate1 = 0
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtNSProcCode1_DblClick()
    If Len(txtNSProcCode1) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtNSProcCode1
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtNSProcCode2_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtNSProcCode2)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtNSProcCode2)
            Else
                Call SearchMMARateCal(txtNSProcCode2)
            End If
            txtNSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtNS13Rate2 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtNS13MinRate2 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
            
                Call Search14thRate(txtNSProcCode2)
                Text1(10).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(1).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
        Else
            txtNSMMARate2 = 0
            txtNS13Rate2 = 0
            txtNS13MinRate2 = 0
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtNSProcCode2_DblClick()
    If Len(txtNSProcCode2) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtNSProcCode2
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtNSProcCode3_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtNSProcCode3)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtNSProcCode3)
            Else
                Call SearchMMARateCal(txtNSProcCode3)
            End If
            txtNSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtNS13Rate3 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtNS13MinRate3 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
            
                Call Search14thRate(txtNSProcCode3)
                Text1(8).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(3).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
            
        Else
            txtNSMMARate3 = 0
            txtNS13Rate3 = 0
            txtNS13MinRate3 = 0
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtNSProcCode3_DblClick()
    If Len(txtNSProcCode3) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtNSProcCode3
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtNSProcCode4_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtNSProcCode4)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtNSProcCode4)
            Else
                Call SearchMMARateCal(txtNSProcCode4)
            End If
            txtNSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtNS13Rate4 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtNS13MinRate4 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
            
                Call Search14thRate(txtNSProcCode4)
                Text1(9).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(2).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
        Else
            txtNSMMARate4 = 0
            txtNS13Rate4 = 0
            txtNS13MinRate4 = 0
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtNSProcCode4_DblClick()
    If Len(txtNSProcCode4) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtNSProcCode4
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtNSProcCode5_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtNSProcCode5)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtNSProcCode5)
            Else
                Call SearchMMARateCal(txtNSProcCode5)
            End If
            txtNSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtNS13Rate5 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtNS13MinRate5 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
            
                Call Search14thRate(txtNSProcCode5)
                Text1(7).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(4).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
        Else
            txtNSMMARate5 = 0
            txtNS13Rate5 = 0
            txtNS13MinRate5 = 0
            
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtNSProcCode5_DblClick()
    If Len(txtNSProcCode5) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtNSProcCode5
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtNSProcCode6_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtNSProcCode6)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtNSProcCode6)
            Else
                Call SearchMMARateCal(txtNSProcCode6)
            End If
            txtNSMMARate6 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtNS13Rate6 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtNS13MinRate6 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
            
                Call Search14thRate(txtNSProcCode6)
                Text1(6).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(5).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
        Else
            txtNSMMARate6 = 0
            txtNS13Rate6 = 0
            txtNS13MinRate6 = 0
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtNSProcCode6_DblClick()
    If Len(txtNSProcCode6) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtNSProcCode6
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtPatientAge_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPDDOS_LostFocus(Index As Integer)
    
    If IsDate(txtPDDOS(0)) Then
    Else
        If txtPDDOS(0) <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPDDOS(0).SetFocus
        End If
    End If
    
    If IsDate(txtPDDOS(1)) Then
    Else
        If txtPDDOS(1) <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPDDOS(1).SetFocus
        End If
    End If

    If IsDate(txtPDDOS(2)) Then
    Else
        If txtPDDOS(2) <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPDDOS(2).SetFocus
        End If
    End If

End Sub

Private Sub txtPrincipalExc_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub

Private Sub txtProcCode1_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtProcCode1)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtProcCode1)
            Else
                Call SearchMMARateCal(txtProcCode1)
            End If
                txtSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
                txtAnaesMMARate1 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
                txtS13Rate1 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                txtAnae13Rate1 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                txtS13MinRate1 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
                
                Call Search14thRate(txtProcCode1)
                
                Text1(18).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(24) = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                Text1(17).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0) '= IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        Else
            txtSMMARate1 = 0
            txtAnaesMMARate1 = 0
            txtS13Rate1 = 0
            txtAnae13Rate1 = 0
            txtS13MinRate1 = 0
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtProcCode1_DblClick()
    If Len(txtProcCode1) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtProcCode1
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtProcCode2_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtProcCode2)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtProcCode2)
            Else
                Call SearchMMARateCal(txtProcCode2)
            End If
            txtSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtAnaesMMARate2 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
            txtS13Rate2 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtAnae13Rate2 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
            txtS13MinRate2 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
            
                Call Search14thRate(txtProcCode2)
                
                Text1(19).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(25) = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                Text1(16).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
        Else
            txtSMMARate2 = 0
            txtAnaesMMARate2 = 0
            txtS13Rate2 = 0
            txtAnae13Rate2 = 0
            txtS13MinRate2 = 0
             
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtProcCode2_DblClick()
    If Len(txtProcCode2) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtProcCode2
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtProcCode3_Change()

    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtProcCode3)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtProcCode3)
            Else
                Call SearchMMARateCal(txtProcCode3)
            End If
            txtSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtAnaesMMARate3 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
            txtS13Rate3 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtAnae13Rate3 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
            txtS13MinRate3 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
            
                Call Search14thRate(txtProcCode3)
                
                Text1(21).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(27) = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                Text1(14).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
        Else
            txtSMMARate3 = 0
            txtAnaesMMARate3 = 0
            txtS13Rate3 = 0
            txtAnae13Rate3 = 0
            txtS13MinRate3 = 0
             
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    
End Sub

Private Sub txtProcCode3_DblClick()
    If Len(txtProcCode3) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtProcCode3
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtProcCode4_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtProcCode4)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtProcCode4)
            Else
                Call SearchMMARateCal(txtProcCode4)
            End If
            txtSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtAnaesMMARate4 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
            txtS13Rate4 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtAnae13Rate4 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
            txtS13MinRate4 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
            
            
                Call Search14thRate(txtProcCode4)
                
                Text1(20).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(26) = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                Text1(15).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
        Else
            txtSMMARate4 = 0
            txtAnaesMMARate4 = 0
            txtS13Rate4 = 0
            txtAnae13Rate4 = 0
            txtS13MinRate4 = 0
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtProcCode4_DblClick()
    If Len(txtProcCode4) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtProcCode4
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtProcCode5_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtProcCode5)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtProcCode5)
            Else
                Call SearchMMARateCal(txtProcCode5)
            End If
            txtSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtAnaesMMARate5 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
            txtS13Rate5 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtAnae13Rate5 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
            txtS13MinRate5 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
            
                Call Search14thRate(txtProcCode5)
                
                Text1(22).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(28) = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                Text1(13).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
        Else
            txtSMMARate5 = 0
            txtAnaesMMARate5 = 0
            txtS13Rate5 = 0
            txtAnae13Rate5 = 0
            txtS13MinRate5 = 0
             
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtProcCode5_DblClick()
    If Len(txtProcCode1) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtProcCode5
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtProcCode6_Change()
    If Trim(TempCASNo) <> "" Then
        If Len(Trim(txtProcCode6)) Then
            If MMARateCal = 0 Then
                Call SearchMMARate(txtProcCode6)
            Else
                Call SearchMMARateCal(txtProcCode6)
            End If
            txtSMMARate6 = IIf(Len(MMASurgFees), MMASurgFees, 0)
            txtAnaesMMARate6 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
            txtS13Rate6 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
            txtAnae13Rate6 = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
            txtS13MinRate6 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
            
                Call Search14thRate(txtProcCode6)
                
                Text1(23).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
                Text1(29) = IIf(Len(MMAAnaeFees13th), MMAAnaeFees13th, 0)
                Text1(12).Text = IIf(Len(MMASurgFees13th), MMASurgFees13th * 1.06, 0)
        Else
            txtSMMARate6 = 0
            txtAnaesMMARate6 = 0
            txtS13Rate6 = 0
            txtAnae13Rate6 = 0
            txtS13MinRate6 = 0
        End If
    Else
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub txtProcCode6_DblClick()
    If Len(txtProcCode6) Then
        frmShowInvestigation.Source = "2"
        frmShowInvestigation.tmpOPCS = txtProcCode6
        frmShowInvestigation.Show vbModal
    End If
End Sub

Private Sub txtReason_Click()
    KeyPreview = False
End Sub

Private Sub txtReason_LostFocus()
    KeyPreview = True
End Sub

Private Sub txtSActFee1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)

End Sub

Private Sub txtSActFee2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)

End Sub

Private Sub txtSActFee3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)

End Sub

Private Sub txtSActFee4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)

End Sub

Private Sub txtSActFee5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)

End Sub

Private Sub txtSActFee6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSActFee1), txtSActFee1, 0)
    B = IIf(IsNumeric(txtSActFee2), txtSActFee2, 0)
    C = IIf(IsNumeric(txtSActFee3), txtSActFee3, 0)
    d = IIf(IsNumeric(txtSActFee4), txtSActFee4, 0)
    e = IIf(IsNumeric(txtSActFee5), txtSActFee5, 0)
    g = IIf(IsNumeric(txtSActFee6), txtSActFee6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtSMMARate1_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtSMMARate2_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtSMMARate3_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtSMMARate4_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtSMMARate5_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtSMMARate6_Change()
Dim a, B, C, d, e, f, g As Double

    a = IIf(IsNumeric(txtSMMARate1), txtSMMARate1, 0)
    B = IIf(IsNumeric(txtSMMARate2), txtSMMARate2, 0)
    C = IIf(IsNumeric(txtSMMARate3), txtSMMARate3, 0)
    d = IIf(IsNumeric(txtSMMARate4), txtSMMARate4, 0)
    e = IIf(IsNumeric(txtSMMARate5), txtSMMARate5, 0)
    g = IIf(IsNumeric(txtSMMARate6), txtSMMARate6, 0)
    
    f = CDbl(a) + CDbl(B) + CDbl(C) + CDbl(d) + CDbl(e) + CDbl(g)
End Sub

Private Sub txtState_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SSTab1.Tab = 1
        cboPatientList.SetFocus
    End If
End Sub

Private Sub txtState_LostFocus()
    SSTab1.Tab = 1
End Sub

Private Sub txtSpecialRemark_Click()
    KeyPreview = False
    cmdShow.Default = False
End Sub

Private Sub txtSpecialRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 126 Then
        MsgBox "You are not allow to use this symbol as a Keyword", vbOKOnly + vbCritical, DialogBoxTitle
        KeyAscii = 0
    End If
End Sub

Private Sub txtSubDocDate_LostFocus()
    If IsDate(txtSubDocDate) Then
        If txtFirstDocDate = "__/__/____" Then
            MsgBox "Please Enter First Doc. Date! ", vbCritical
            txtSubDocDate.Text = "__/__/____"
            txtFirstDocDate.SetFocus
        Else
            If (CDate(txtSubDocDate) < CDate(txtFirstDocDate)) Then
                MsgBox "First Doc. Date cannot greater than subsequent Doc. Date! ", vbCritical
                txtFirstDocDate.SetFocus
            End If
        End If
    Else
        If txtSubDocDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtSubDocDate.SetFocus
        End If
    End If

End Sub

Private Sub txtUDDOS1_LostFocus()
    If IsDate(txtUDDOS1) Then
    Else
    If txtUDDOS1 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtUDDOS1.SetFocus
    End If
    End If
End Sub

Private Sub txtUDDOS2_LostFocus()
    If IsDate(txtUDDOS2) Then
    Else
    If txtUDDOS2 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtUDDOS2.SetFocus
    End If
    End If
End Sub

Private Sub txtUDDOS3_LostFocus()
    If IsDate(txtUDDOS3) Then
    Else
    If txtUDDOS3 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtUDDOS3.SetFocus
    End If
    End If
End Sub

Private Sub saveHos()
On Error Resume Next
    Dim strsql As String
    Dim strsqlCAS As String
    Dim CAS As New ADODB.Recordset
    Dim strsqlHOS As String
    Dim HOS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim cnn As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    
    strAppendCase = "7"
    strAppend = ""
    
    strAppend = " AND CASNumber = '" & Trim(TempCASNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Adjustment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    CAS.CursorLocation = adUseClient
    CAS.CursorType = adOpenForwardOnly
    CAS.CacheSize = 100
    Set CAS = cmd.Execute
        
    HOspitalName = Replace(txtHOSCode, "'", "''")
    
    'strsqlHOS = "select HosCode, HosName from HOS where HosName = '" & Trim(HOspitalName) & "'"
    'HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "23"
    strAppend = ""

    strAppend = " AND HosName = '" & Trim(HOspitalName) & "'"
    Set cnn.ActiveConnection = medixmasterconn
    cnn.CommandText = "Case_Adjustment"
    cnn.CommandType = adCmdStoredProc
    cnn.CommandTimeout = 300
    Set Prm1 = cnn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cnn.Parameters.Append Prm1
    Set Prm2 = cnn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cnn.Parameters.Append Prm2
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cnn.Execute
    
    If CAS.EOF = False Then
        If HOS.EOF = False Then
            strsqlHOS = "EXEC Case_AdjUpdateHOS '" & Trim(TempCASNo) & "', '" & ChkStr(HOS!HOSCode) & "'"
            medixmasterconn.Execute strsqlHOS
            'CAS!CASHOSCode = ChkStr(HOS!HOSCode)
            'CAS.Update
        End If
    End If
    
    CAS.Close
    Set CAS = Nothing
    HOS.Close
    Set HOS = Nothing
End Sub

Private Sub BNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim bnf As ListItem
    Dim sql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    
    strAppend = ""
    If Len(Trim(tmpHltCode)) <> 0 _
     And Len(Trim(tmpINSCode)) <> 0 And Len(Trim(tmpPlnCode)) <> 0 Then
        If Len(MBMCPLNCode) And Mid(cboPatientList, 1, 2) <> "00" Then
            strAppendCase = "17"
            strAppend = " AND PLNCOde ='" & Trim(MBMCPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(tmpHltCode) & "'" & _
              " AND INSCODe ='" & Trim(tmpINSCode) & "'"
            'sql = " select BNFCode, BNFNoOFDays , " & _
              " BNFclaimableAmount, BNFdescription , BNFRMPerDis " & _
              " from View_BNF where PLNCOde ='" & Trim(MBMCPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(tmpHLTCode) & "'" & _
              " AND INSCODe ='" & Trim(tmpINSCode) & "'"
        Else
            strAppendCase = "17"
            strAppend = " AND PLNCOde ='" & Trim(tmpPlnCode) & "'" & _
              " AND HLTCode ='" & Trim(tmpHltCode) & "'" & _
              " AND INSCODe ='" & Trim(tmpINSCode) & "'"
            'sql = " select BNFCode, BNFNoOFDays , " & _
              " BNFclaimableAmount, BNFdescription , BNFRMPerDis " & _
              " from View_BNF where PLNCOde ='" & Trim(tmpPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(tmpHLTCode) & "'" & _
              " AND INSCODe ='" & Trim(tmpINSCode) & "'"
        End If
            'rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            rst2.CursorLocation = adUseClient
            rst2.CursorType = adOpenForwardOnly
            rst2.CacheSize = 100
            Set rst2 = cmd.Execute
            
            Do Until rst2.EOF
            Set bnf = lstBenefits.ListItems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFdescription) <> "" Then
                    bnf.SubItems(1) = rst2!BNFdescription
                End If
                bnf.SubItems(2) = IIf(IsNull(rst2!BNFNoOfDays), 0, rst2!BNFNoOfDays)
                 If Trim(rst2!BNFClaimAbleAmount) <> "" Then
                    bnf.SubItems(3) = Format(rst2!BNFClaimAbleAmount, "###,###,###.00")
                End If
                If Trim(rst2!BNFRMperDis) <> "" Then
                    bnf.SubItems(4) = Format(rst2!BNFRMperDis, "###,###,###.00")
                End If
                If rst2!BNFCode = "101" Then
                    RB = bnf.SubItems(3)
                    tmpRMAmt = bnf.SubItems(3)
                End If
                
                If rst2!BNFCode = "102" Then
                    ICU = bnf.SubItems(3)
                    tmpICUAmt = bnf.SubItems(3)
                End If
                
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub


Private Sub LGListing()
    Dim LG As New ADODB.Recordset
    Dim LGList As ListItem
    Dim sql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    
    strAppend = ""
    If Len(Trim(txtMBMNumber)) <> 0 Then
        lstLGListing.ListItems.Clear
        
        'sql = " Select CASeNo, LGNo, LGDate, " & _
            " LGAmt, LGIssuedBy, LGRemark  From CASLG " & _
            " where CASENo ='" & Trim(txtCaseNo) & "'"
    
        'LG.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        strAppendCase = "7"
        strAppend = " AND CASENo ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        LG.CursorLocation = adUseClient
        LG.CursorType = adOpenForwardOnly
        LG.CacheSize = 100
        Set LG = cmd.Execute
        
        If Not LG.EOF Then
            Do Until LG.EOF
                    Set LGList = lstLGListing.ListItems.Add(, , "L" & LG!CaseNo & "-" & LG!LGNo & Mid(LG!LGPurpose, 1, 1))
                    With LG
                        If Len(!LGDate) Then
                            LGList.SubItems(1) = Format(!LGDate, "dd/mm/yyyy")
                        End If
                        
                        If Len(!LGFaxTime) Then
                            LGList.SubItems(2) = !LGFaxTime
                        End If

                        If Len(!LGRequestDate) Then
                            LGList.SubItems(3) = Format(!LGRequestDate, "dd/mm/yyyy")
                        End If
                        If Len(!LGRequestTime) Then
                            LGList.SubItems(4) = !LGRequestTime
                        End If
                        If Len(!LGPurpose) Then
                            LGList.SubItems(5) = Trim(!LGPurpose)
                        End If
                        If Len(!LGAmt) Then
                            LGList.SubItems(6) = Format(!LGAmt, "#,##0.00")
                        End If
                        If Len(!LGIssuedBy) Then
                            LGList.SubItems(7) = Trim(!LGIssuedBy)
                        End If
                        If Len(!LGRemark) Then
                            LGList.SubItems(8) = Trim(!LGRemark)
                        End If

                    
                    
                    .MoveNext
                End With
                Loop
                LG.Close
                Set LG = Nothing
            End If
        End If
End Sub

Private Sub ShowRepudiation()
Dim REP As New ADODB.Recordset
Dim REPCode As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String

Dim strsqlRep As String
    ' INSERT INTO CASRepudiation TABLE
     strAppend = ""
    'strsqlREP = "select * from CASRepudiation where CASNumber = '" & txtCaseNo & "'"
    'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "10"
    strAppend = " AND CASNumber = '" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        REP.CursorLocation = adUseClient
        REP.CursorType = adOpenForwardOnly
        REP.CacheSize = 100
        Set REP = cmd.Execute
        
    If Not REP.EOF Then
        With REP
            If Len(Trim(!CASREPCode1)) Then
             If Mid(!CASREPCode1, 1, 2) = "RC" Then
                strsqlRep = "EXEC Case_Inquiry '" & Trim(!CASREPCode1) & "',7"
                REPCode.Open strsqlRep, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                If REPCode.EOF = False Then
                With REPCode
                    cboREPCode1 = Trim(!REPCode) & " - " & Trim(!RepDescription)
                End With
                End If
                REPCode.Close
                Set REPCode = Nothing
             ElseIf Mid(!CASREPCode1, 1, 2) = "DC" Then
                strsqlRep = "Select Disclouser, RepCatCode   from Tbl_Disclouser where RepCatCode ='" + !CASREPCode1 + "'"
                REPCode.Open strsqlRep, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                If REPCode.EOF = False Then
                With REPCode
                    cboREPCode1 = Trim(!DisClouser) & "||" & Trim(!REPCatCode)
                End With
                End If
                REPCode.Close
                Set REPCode = Nothing
             ElseIf Mid(!CASREPCode1, 1, 2) = "SD" Then
                strsqlRep = "Select SubDisclouser, RepCatCode   from Tbl_SubDisclouser where RepCatCode ='" + !CASREPCode1 + "'"
                REPCode.Open strsqlRep, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                If REPCode.EOF = False Then
                With REPCode
                    cboREPCode1 = Trim(!SubDisclouser) & "||" & Trim(!REPCatCode)
                End With
                End If
                REPCode.Close
                Set REPCode = Nothing
             ElseIf Mid(!CASREPCode1, 1, 2) = "ND" Then
                strsqlRep = "Select NonDisclouser, RepCatCode   from Tbl_NonDisclouser where RepCatCode ='" + !CASREPCode1 + "'"
                REPCode.Open strsqlRep, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
                If REPCode.EOF = False Then
                With REPCode
                    cboREPCode1 = Trim(!NonDisclouser) & "||" & Trim(!REPCatCode)
                End With
                End If
                REPCode.Close
                Set REPCode = Nothing
             End If
                
            End If
            If Len(Trim(!CASREPCode2)) Then
                strsqlRep = "EXEC Case_Inquiry '" & Trim(!CASREPCode2) & "',7"
                REPCode.Open strsqlRep, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                If REPCode.EOF = False Then
                With REPCode
                    cboREPCode2 = Trim(!REPCode) & " - " & Trim(!RepDescription)
                End With
                End If
                REPCode.Close
                Set REPCode = Nothing
            End If
            If Len(Trim(!CASREPCode3)) Then
                strsqlRep = "EXEC Case_Inquiry '" & Trim(!CASREPCode3) & "',7"
                REPCode.Open strsqlRep, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                If REPCode.EOF = False Then
                With REPCode
                    cboREPCode3 = Trim(!REPCode) & " - " & Trim(!RepDescription)
                End With
                End If
                REPCode.Close
                Set REPCode = Nothing
            End If
            If Len(Trim(!CASDocument)) Then
                Document = !CASDocument
            End If
            If Len(Trim(!CASPersonnel)) Then
                Personnel = !CASPersonnel
            End If
            If Len(Trim(!CASAttention)) Then
                Attention = !CASAttention
            End If
        End With
    End If
    REP.Close
    Set REP = Nothing
End Sub

Public Sub ClearEntries(Optional KIV As Boolean)
   On Error Resume Next
   Dim ctl As Control
   Dim s As Integer

     For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
        ElseIf TypeOf ctl Is ComboBox Then
            If ctl.Enabled = True Then
                ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = True
        'ElseIf TypeOf ctl Is MaskEdBox Then
        '        ctl.mask = ""
        '        ctl.Text = ""
        '        ctl.mask = "##/##/####"
        End If
    Next ctl
        cboStatus1 = "O - Open"
        lstMember.ListItems.Clear
        lstBenefits.ListItems.Clear
        'lstNewBenefitDetails.listitems.Clear
        cboSuppExc.Clear
        txtCASLGAmt = 2500
        txtCASReserveAmount = 1600
        cboNature = "ILL - ILLNess"
        cboModality = "S - Surgical"
        txtCASRemark.Visible = True
        'Fgrid.Clear
        'lstSurgProc.ListItems.Clear
        'Fgrid.col = 1
        'For s = 1 To 10
        '    Fgrid.row = (s)
        'Fgrid = ""
        'Next
        txtLGTime.mask = ""
        txtLGTime.Text = ""
        txtLGTime.mask = "##:##"
        lstPolicyWording.Clear
End Sub

Public Function ClearForm(frmCall)
On Error Resume Next
Dim ctl As Control

For Each ctl In frmCall.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    ElseIf TypeOf ctl Is ListBox Then
            ctl.Clear
    ElseIf TypeOf ctl Is ComboBox Then
            ctl.Text = ""
    ElseIf TypeOf ctl Is CheckBox Then
            ctl.Value = False
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    
    End If
Next ctl
End Function

Private Function bookmark(Optional strAppend As String) As Integer
Dim strsqlPay As String
Dim PAY As New ADODB.Recordset
Dim strsqlRep As String
Dim REP As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim strsqlHOS As String
Dim strsqlLG As String
Dim LG As New ADODB.Recordset
Dim REPCat As New ADODB.Recordset
Dim strsqlREPCat As String
Dim LGNo As String
Dim LGVersion As String
Dim ContactPerson As String
Dim MBMExpiryDate As Long
Dim MBM As New ADODB.Recordset
Dim strsqlMBM As String
Dim strsqlDiag As String
Dim Including As String
Dim Terms As String
Dim COPay As Boolean
Dim MRI As Boolean
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
Dim strsqlRepRea As String
Dim LGCoPayPerc As String
Dim tmpMMBPerc As String
    LGCoPayPerc = ""
    HOspitalName = Replace(txtHOSCode, "'", "''")
    strAppendCase = "6"
    StrAppend1 = " AND HOSName = '" & Trim(HOspitalName) & "'"
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "Case_Adjustment"
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd1.Parameters.Append Prm2
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cmd1.Execute
    
    
    If Letter = "S" Or Letter = "A" Or Letter = "D" Then
        
        If objdoc.Bookmarks.Exists("patientname2") = True Then
            objdoc.Bookmarks("patientname2").Range.Text = IIf(Len(tmpMBMAppName), tmpMBMAppName, txtMBMName)
        End If
        If objdoc.Bookmarks.Exists("CaseNo") = True Then
            objdoc.Bookmarks("CaseNo").Range.Text = IIf(Len(txtCaseNo), txtCaseNo, "")
        End If
        If objdoc.Bookmarks.Exists("Add1") = True Then
            objdoc.Bookmarks("Add1").Range.Text = IIf(Len(txtAdd1), txtAdd1, "")
        End If
        If objdoc.Bookmarks.Exists("Add2") = True Then
            objdoc.Bookmarks("Add2").Range.Text = IIf(Len(txtAdd2), txtAdd2, "")
        End If
        If objdoc.Bookmarks.Exists("Add3") = True Then
            objdoc.Bookmarks("Add3").Range.Text = IIf(Len(txtAdd3), txtAdd3, "")
        End If
        If objdoc.Bookmarks.Exists("PatientIC") = True Then
            objdoc.Bookmarks("PatientIC").Range.Text = IIf(Len(txtPatientIC), txtPatientIC, "")
        End If
        If objdoc.Bookmarks.Exists("Postcode") = True Then
            objdoc.Bookmarks("Postcode").Range.Text = IIf(Len(txtPostCode), txtPostCode, "")
        End If
        If objdoc.Bookmarks.Exists("State") = True Then
            objdoc.Bookmarks("State").Range.Text = IIf(Len(txtCity), txtCity & "" & txtState, "")
        End If
        If objdoc.Bookmarks.Exists("patientname") = True Then
            objdoc.Bookmarks("patientname").Range.Text = IIf(Trim(cboPatientList) <> "", Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))), " Please key in patient name here!")
        End If
        If objdoc.Bookmarks.Exists("hospital") = True Then
            objdoc.Bookmarks("hospital").Range.Text = IIf(Len(Trim(txtHOSCode)), Trim(txtHOSCode), "")
        End If
        If objdoc.Bookmarks.Exists("DOA") = True Then
            objdoc.Bookmarks("DOA").Range.Text = IIf(Len(txtDateAdmitted), txtDateAdmitted, "")
        End If
        If objdoc.Bookmarks.Exists("dateofadmission") = True Then
            objdoc.Bookmarks("dateofadmission").Range.Text = IIf(Len(txtDateAdmitted), txtDateAdmitted, "")
        End If
        If objdoc.Bookmarks.Exists("MedixStaffName") = True Then
            objdoc.Bookmarks("MedixStaffName").Range.Text = IIf(Len(LogonName), LogonName, "")
        End If
        If objdoc.Bookmarks.Exists("LGID") = True Then
            objdoc.Bookmarks("LGID").Range.Text = IIf(Len(LogonName), LogonName, "")
        End If
        If objdoc.Bookmarks.Exists("SystemDate") = True Then
            objdoc.Bookmarks("SystemDate").Range.Text = Format(GetServerTime, "DD/MM/YYYY")
        End If
        
        If objdoc.Bookmarks.Exists("LGDate") = True Then
            objdoc.Bookmarks("LGDate").Range.Text = Format(GetServerTime, "DD/MM/YYYY")
        End If
        If objdoc.Bookmarks.Exists("HOSName") = True Then
            objdoc.Bookmarks("HOSName").Range.Text = IIf(Len(txtHOSCode), txtHOSCode, "")
        End If
        If objdoc.Bookmarks.Exists("Hospital") = True Then
            objdoc.Bookmarks("Hospital").Range.Text = IIf(Len(txtHOSCode), txtHOSCode, "")
        End If
        If objdoc.Bookmarks.Exists("HOSStaff") = True Then
            objdoc.Bookmarks("HOSStaff").Range.Text = IIf(Len(HOS!HOSContact), HOS!HOSContact, "")
        End If
        If objdoc.Bookmarks.Exists("HospFax") = True Then
            objdoc.Bookmarks("HospFax").Range.Text = IIf(Len(HOS!HOSFaxNo), HOS!HOSFaxNo, "")
        End If
        If objdoc.Bookmarks.Exists("patient") = True Then
            objdoc.Bookmarks("patient").Range.Text = IIf(Trim(cboPatientList) <> "", Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))), "")
        End If
        If objdoc.Bookmarks.Exists("mbmno") = True Then
            objdoc.Bookmarks("mbmno").Range.Text = IIf(Len(txtMBMNumber), txtMBMNumber, "")
        End If
        If objdoc.Bookmarks.Exists("mbmnumber") = True Then
            objdoc.Bookmarks("mbmnumber").Range.Text = IIf(Len(txtMBMNumber), txtMBMNumber, "")
        End If
        
        If objdoc.Bookmarks.Exists("DOD") = True Then
            objdoc.Bookmarks("DOD").Range.Text = IIf(Len(txtDateDischarge), txtDateDischarge, "")
        End If
        bookmark = True
    Else
    
            ContactPerson = ""
            LGNo = ""
            LGVersion = ""
            Including = ""
            Terms = ""
           
            strAppendCase = "18"
            StrAppend1 = " AND PAYCode = '" & Trim(txtPayCode) & "'"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Adjustment"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            PAY.CursorLocation = adUseClient
            PAY.CursorType = adOpenForwardOnly
            PAY.CacheSize = 100
            Set PAY = cmd.Execute
       '
        If Letter = "L" Or Letter = "C" Then
            bookmark = False
            
            If tmpHltCode = "S" Then
                strAppendCase = "19"
                StrAppend1 = " AND PLNCode = '" & Trim(txtPlanCode) & "'"
            Else
                strAppendCase = "19"
                StrAppend1 = " AND PLNCode = '" & Trim(txtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
            End If
        
            Set cmd2.ActiveConnection = medixmasterconn
            cmd2.CommandText = "Case_Adjustment"
            cmd2.CommandType = adCmdStoredProc
            cmd2.CommandTimeout = 300
            Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
            cmd2.Parameters.Append Prm1
            Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd2.Parameters.Append Prm2
            MBM.CursorLocation = adUseClient
            MBM.CursorType = adOpenForwardOnly
            MBM.CacheSize = 100
            Set MBM = cmd2.Execute
        
            If MBM.EOF = False Then
        
                If MBM!PLNMeal = "Y" Then
                    Including = Including + "Meal"
                End If
                If MBM!PLNNursing = "Y" Then
                    Including = Including + " " & ",Nursing Charges"
                End If
                If MBM!PLNTax = "Y" Then
                    Including = Including + " " & ",Tax"
                End If
                If MBM!CoPaymentInd = "Y" Then
                    COPay = True
                Else
                    COPay = False
                End If
                If MBM!PLNMRI = "Y" Then
                    MRI = True
                Else
                    MRI = False
                End If
                tmpRBType = IIf(Len(MBM!PLNRBType), MBM!PLNRBType, "")
                If IsNumeric(MBM!PLNCoPayPerc) Then
                    LGCoPayPerc = 100 - CDbl(MBM!PLNCoPayPerc)
                Else
                    LGCoPayPerc = ""
                End If
                If MBM!PLNMMIndperc <> "" Then
                    tmpMMBPerc = MBM!PLNMMIndperc * 100
                    tmpMMBPerc = 100 - tmpMMBPerc
                Else
                    tmpMMBPerc = IIf(Len(MBM!PLNMMIndperc), MBM!PLNMMIndperc, "")
                End If
            Else
                COPay = False
                MRI = False
            End If
            
        
                If Trim(HOSType) = "P" Then
                    If Mid(frmGLMenu.cbPurpose, 1, 1) = "O" And Trim(tmpPLNLGPrivate) <> "4" And Trim(tmpPLNLGPrivate) <> "5" And Trim(tmpPLNLGPrivate) <> "6" And Trim(tmpPLNLGPrivate) <> "7" And Trim(tmpPLNLGPrivate) <> "8" And Trim(tmpPLNLGPrivate) <> "26" And Trim(tmpPLNLGPrivate) <> "34" And (Trim(tmpPLNLGPrivate) = "5" Or Trim(tmpPLNLGPrivate) = "6" Or Trim(tmpPLNLGPrivate) = "7" Or Trim(tmpPLNLGPrivate) = "8") Then
                            
                            If objdoc.Bookmarks.Exists("patient") = True Then
                                objdoc.Bookmarks("patient").Range.Text = IIf(Trim(cboPatientList) <> "", Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))), " Please key in patient name here!")
                            End If
                            If objdoc.Bookmarks.Exists("GLAmt") = True Then
                                objdoc.Bookmarks("GLAmt").Range.Text = IIf(txtCASLGAmt <> "", Format(frmGLMenu.txtCASLGAmt, "#,##0.00"), 0)
                            End If
                            If objdoc.Bookmarks.Exists("LGno") = True Then
                                objdoc.Bookmarks("LGno").Range.Text = IIf(Len(LGNumber), LGNumber, "")
                            End If
                            LGVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
                            If objdoc.Bookmarks.Exists("GrpCompany") = True Then
                                objdoc.Bookmarks("GrpCompany").Range.Text = IIf(Len(tmpMBMGroupCompany), tmpMBMGroupCompany, "")
                            End If
                            If objdoc.Bookmarks.Exists("Principal") = True Then
                                objdoc.Bookmarks("Principal").Range.Text = IIf(Len(txtMBMName), txtMBMName, "")
                            End If
                            If objdoc.Bookmarks.Exists("LGID") = True Then
                                objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
                            End If
                            If objdoc.Bookmarks.Exists("Diagnosis") = True Then
                                objdoc.Bookmarks("Diagnosis").Range.Text = IIf(Len(txtReason), txtReason, "")
                            End If
                            If objdoc.Bookmarks.Exists("mbmnumber") = True Then
                                objdoc.Bookmarks("mbmnumber").Range.Text = IIf(Len(txtMBMNumber), txtMBMNumber, "")
                            End If
                            If objdoc.Bookmarks.Exists("hospital") = True Then
                                objdoc.Bookmarks("hospital").Range.Text = IIf(Len(txtHOSCode), Trim(txtHOSCode), "")
                            End If
                            If objdoc.Bookmarks.Exists("dateofadmission") = True Then
                                objdoc.Bookmarks("dateofadmission").Range.Text = IIf(Len(txtDateAdmitted), txtDateAdmitted, "")
                            End If
                            If objdoc.Bookmarks.Exists("Fax") = True Then
                                objdoc.Bookmarks("Fax").Range.Text = IIf(Len(HOS!HOSFaxNo), HOS!HOSFaxNo, "")
                            End If
                            If objdoc.Bookmarks.Exists("Remark") = True Then
                                objdoc.Bookmarks("Remark").Range.Text = IIf(Len(frmGLMenu.txtremarks), frmGLMenu.txtremarks, "")
                            End If
                            If objdoc.Bookmarks.Exists("LGDate") = True Then
                                objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
                            End If
                            If objdoc.Bookmarks.Exists("AttDr") = True Then
                                objdoc.Bookmarks("AttDr").Range.Text = IIf(Len(txtATTDoc(0)), txtATTDoc(0), "")
                            End If
                            
                            'GL 5,6,7,8
                            
                            If objdoc.Bookmarks.Exists("amount") = True Then
                                objdoc.Bookmarks("amount").Range.Text = IIf(Len(frmGLMenu.txtCASLGAmt), Format(frmGLMenu.txtCASLGAmt, "#,##0.00"), 0)
                            End If
                            If txtDateDischarge <> "__/__/____" Then
                                If objdoc.Bookmarks.Exists("DOD") = True Then
                                    objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
                                End If
                            End If
                            
                            If tmptxtMBMPolNo <> "" Then
                                If objdoc.Bookmarks.Exists("policyno") = True Then
                                    objdoc.Bookmarks("policyno").Range.Text = tmptxtMBMPolNo
                                End If
                            End If
                            
                            If tmpRBType <> "" Then
                                If objdoc.Bookmarks.Exists("RB") = True Then
                                    objdoc.Bookmarks("RB").Range.Text = tmpRBType
                                End If
                            ElseIf RB <> "" Then
                                If IsNumeric(ICU) Then
                                    If objdoc.Bookmarks.Exists("RB") = True Then
                                        objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & " /" & " ICU - RM" & ICU
                                    End If
                                Else
                                    If objdoc.Bookmarks.Exists("RB") = True Then
                                        objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - FULL COVER"
                                    End If
                                End If
                            End If
                            
                             If objdoc.Bookmarks.Exists("Principal") = True Then
                                objdoc.Bookmarks("Principal").Range.Text = txtMBMName
                            End If
                            If objdoc.Bookmarks.Exists("InsuredBy") = True Then
                                objdoc.Bookmarks("InsuredBy").Range.Text = tmpMBMGroupCompany
                            End If
                            If objdoc.Bookmarks.Exists("PatientIC") = True Then
                                objdoc.Bookmarks("PatientIC").Range.Text = IIf(Len(txtPatientIC), txtPatientIC, "")
                            End If
                            If objdoc.Bookmarks.Exists("Attention") = True Then
                                objdoc.Bookmarks("Attention").Range.Text = IIf(Len(HOS!HOSContact), HOS!HOSContact, "")
                            End If
                            If objdoc.Bookmarks.Exists("Fax") = True Then
                                objdoc.Bookmarks("Fax").Range.Text = IIf(Len(HOS!HOSFaxNo), HOS!HOSFaxNo, "")
                            End If
                            If objdoc.Bookmarks.Exists("Remark") = True Then
                                objdoc.Bookmarks("Remark").Range.Text = IIf(Len(txtReason), txtReason, "")
                            End If
                            If objdoc.Bookmarks.Exists("LGRemark") = True Then
                                objdoc.Bookmarks("LGRemark").Range.Text = IIf(Len(frmGLMenu.txtremarks), frmGLMenu.txtremarks, "")
                            End If
                            bookmark = True
                            
                    ElseIf Trim(tmpPLNLGPrivate) = "5" Or Trim(tmpPLNLGPrivate) = "6" Or Trim(tmpPLNLGPrivate) = "7" Or Trim(tmpPLNLGPrivate) = "8" Then
                        If objdoc.Bookmarks.Exists("hospital") = True Then
                                objdoc.Bookmarks("hospital").Range.Text = IIf(txtHOSCode <> "", Trim(txtHOSCode), "")
                            End If
                            If objdoc.Bookmarks.Exists("patient") = True Then
                              objdoc.Bookmarks("patient").Range.Text = IIf(Trim(cboPatientList) <> "", Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))), " Please key in patient name here!")
                            End If
                            If objdoc.Bookmarks.Exists("Principal") = True Then
                                objdoc.Bookmarks("Principal").Range.Text = txtMBMName
                            End If
                            If objdoc.Bookmarks.Exists("PatientIC") = True Then
                                objdoc.Bookmarks("PatientIC").Range.Text = IIf(Len(txtPatientIC), txtPatientIC, "")
                            End If
                            If objdoc.Bookmarks.Exists("dateofadmission") = True Then
                                objdoc.Bookmarks("dateofadmission").Range.Text = IIf(Len(txtDateAdmitted), txtDateAdmitted, "")
                            End If
                            If objdoc.Bookmarks.Exists("Remark") = True Then
                                objdoc.Bookmarks("Remark").Range.Text = IIf(Len(txtReason), txtReason, "")
                            End If
                            If objdoc.Bookmarks.Exists("LGRemark") = True Then
                                objdoc.Bookmarks("LGRemark").Range.Text = IIf(Len(frmGLMenu.txtremarks), frmGLMenu.txtremarks, "")
                            End If
                            If objdoc.Bookmarks.Exists("Attention") = True Then
                                objdoc.Bookmarks("Attention").Range.Text = IIf(Len(HOS!HOSContact), HOS!HOSContact, "")
                            End If
                            If objdoc.Bookmarks.Exists("Fax") = True Then
                                objdoc.Bookmarks("Fax").Range.Text = IIf(Len(HOS!HOSFaxNo), HOS!HOSFaxNo, "")
                            End If
                            If objdoc.Bookmarks.Exists("LGID") = True Then
                                objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
                            End If
                            If objdoc.Bookmarks.Exists("LGno") = True Then
                                objdoc.Bookmarks("LGno").Range.Text = IIf(Len(LGNumber), LGNumber, "")
                            End If
                            If objdoc.Bookmarks.Exists("Doctor") = True Then
                                objdoc.Bookmarks("Doctor").Range.Text = txtATTDoc(0)
                            End If
                            If objdoc.Bookmarks.Exists("LGDate") = True Then
                                objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
                            End If
                            If objdoc.Bookmarks.Exists("policyno") = True Then
                                    objdoc.Bookmarks("policyno").Range.Text = IIf(Len(tmptxtMBMPolNo), tmptxtMBMPolNo, "")
                                End If
                            If txtDateDischarge <> "__/__/____" Then
                            If objdoc.Bookmarks.Exists("lblDOD") = True Then
                                objdoc.Bookmarks("lblDOD").Range.Text = "Discharge Date"
                            End If
                            If objdoc.Bookmarks.Exists("lblDODDot") = True Then
                                objdoc.Bookmarks("lblDODDot").Range.Text = ":"
                            End If
                            If objdoc.Bookmarks.Exists("DOD") = True Then
                                objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
                            End If
                        End If
                         If frmGLMenu.txtCASLGAmt <> "" Then
                    If objdoc.Bookmarks.Exists("amount") = True Then
                        objdoc.Bookmarks("amount").Range.Text = Format(frmGLMenu.txtCASLGAmt, "#,##0.00")
                    End If
                End If
                 If tmpRBType <> "" Then
                                If objdoc.Bookmarks.Exists("RB") = True Then
                                    objdoc.Bookmarks("RB").Range.Text = IIf(Len(tmpRBType), tmpRBType, "")
                                End If
                            ElseIf RB <> "" Then
                                If IsNumeric(ICU) Then
                                    If objdoc.Bookmarks.Exists("RB") = True Then
                                        objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - RM" & ICU
                                    End If
                                Else
                                    If objdoc.Bookmarks.Exists("RB") = True Then
                                        objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - FULL COVER"
                                    End If
                                End If
                            End If
                            bookmark = True
                        'maybank/etiqa staff
                    ElseIf Trim(tmpPLNLGPrivate) = "4" Or Trim(tmpPLNLGPrivate) = "34" Then
                        If Mid(frmGLMenu.cbPurpose, 1, 1) = "O" Then
                            
                            If objdoc.Bookmarks.Exists("hospital") = True Then
                                objdoc.Bookmarks("hospital").Range.Text = IIf(txtHOSCode <> "", Trim(txtHOSCode), "")
                            End If
                            If objdoc.Bookmarks.Exists("patient") = True Then
                              objdoc.Bookmarks("patient").Range.Text = IIf(Trim(cboPatientList) <> "", Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))), " Please key in patient name here!")
                            End If
                            If objdoc.Bookmarks.Exists("Principal") = True Then
                                objdoc.Bookmarks("Principal").Range.Text = txtMBMName
                            End If
                            If objdoc.Bookmarks.Exists("PatientIC") = True Then
                                objdoc.Bookmarks("PatientIC").Range.Text = IIf(Len(txtPatientIC), txtPatientIC, "")
                            End If
                            If objdoc.Bookmarks.Exists("dateofadmission") = True Then
                                objdoc.Bookmarks("dateofadmission").Range.Text = IIf(Len(txtDateAdmitted), txtDateAdmitted, "")
                            End If
                            If objdoc.Bookmarks.Exists("Remark") = True Then
                                objdoc.Bookmarks("Remark").Range.Text = IIf(Len(txtReason), txtReason, "")
                            End If
                            If objdoc.Bookmarks.Exists("LGRemark") = True Then
                                objdoc.Bookmarks("LGRemark").Range.Text = IIf(Len(frmGLMenu.txtremarks), frmGLMenu.txtremarks, "")
                            End If
                            If objdoc.Bookmarks.Exists("Attention") = True Then
                                objdoc.Bookmarks("Attention").Range.Text = IIf(Len(HOS!HOSContact), HOS!HOSContact, "")
                            End If
                            If objdoc.Bookmarks.Exists("Fax") = True Then
                                objdoc.Bookmarks("Fax").Range.Text = IIf(Len(HOS!HOSFaxNo), HOS!HOSFaxNo, "")
                            End If
                            If objdoc.Bookmarks.Exists("LGID") = True Then
                                objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
                            End If
                            If objdoc.Bookmarks.Exists("LGno") = True Then
                                objdoc.Bookmarks("LGno").Range.Text = IIf(Len(LGNumber), LGNumber, "")
                            End If
                            'If objdoc.Bookmarks.Exists("Status") = True Then
                            '    objdoc.Bookmarks("Status").Range.Text = "(Outpatient SP)"
                            'End If
                            If objdoc.Bookmarks.Exists("Doctor") = True Then
                                objdoc.Bookmarks("Doctor").Range.Text = txtATTDoc(0)
                            End If
                            If objdoc.Bookmarks.Exists("LGDate") = True Then
                                objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
                            End If
                        ElseIf Mid(frmGLMenu.cbPurpose, 1, 1) = "E" Then
                            If objdoc.Bookmarks.Exists("hospital") = True Then
                                objdoc.Bookmarks("hospital").Range.Text = IIf(Len(Trim(txtHOSCode)), Trim(txtHOSCode), "")
                            End If
                            If objdoc.Bookmarks.Exists("patient") = True Then
                                objdoc.Bookmarks("patient").Range.Text = IIf(Trim(cboPatientList) <> "", Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))), " Please key in patient name here!")
                            End If
                            If objdoc.Bookmarks.Exists("PatientIC") = True Then
                                objdoc.Bookmarks("PatientIC").Range.Text = IIf(Len(txtPatientIC), txtPatientIC, "")
                            End If
                            If txtDateAdmitted <> "" Then
                                If objdoc.Bookmarks.Exists("dateofadmission") = True Then
                                    objdoc.Bookmarks("dateofadmission").Range.Text = txtDateAdmitted
                                End If
                            End If
                            If objdoc.Bookmarks.Exists("Attention") = True Then
                                objdoc.Bookmarks("Attention").Range.Text = IIf(Len(HOS!HOSContact), HOS!HOSContact, "")
                            End If
                            If objdoc.Bookmarks.Exists("Fax") = True Then
                                objdoc.Bookmarks("Fax").Range.Text = IIf(Len(HOS!HOSFaxNo), HOS!HOSFaxNo, "")
                            End If
                            If objdoc.Bookmarks.Exists("LGID") = True Then
                                objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
                            End If
                            If objdoc.Bookmarks.Exists("LGno") = True Then
                                objdoc.Bookmarks("LGno").Range.Text = IIf(Len(LGNumber), LGNumber, "")
                            End If
                            'If objdoc.Bookmarks.Exists("Status") = True Then
                            '    objdoc.Bookmarks("Status").Range.Text = "(Exec Med Check)"
                            'End If
                            If objdoc.Bookmarks.Exists("EmpNum") = True Then
                                objdoc.Bookmarks("EmpNum").Range.Text = tmpMBMEmpNo
                            End If
                            If objdoc.Bookmarks.Exists("LGDate") = True Then
                                objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
                            End If
                            
                        Else
                            
                            If objdoc.Bookmarks.Exists("patient") = True Then
                              objdoc.Bookmarks("patient").Range.Text = IIf(Trim(cboPatientList) <> "", Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))), " Please key in patient name here!")
                            End If
                            If objdoc.Bookmarks.Exists("LGno") = True Then
                                objdoc.Bookmarks("LGno").Range.Text = IIf(Len(LGNumber), LGNumber, "")
                            End If
                            LGVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
                            
                            If Letter = "C" Then
                                If objdoc.Bookmarks.Exists("GrpCompany") = True Then
                                    objdoc.Bookmarks("GrpCompany").Range.Text = " - " & tmpMBMGroupCompany
                                End If
                            End If
                            If objdoc.Bookmarks.Exists("LGID") = True Then
                                objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
                            End If
                            If objdoc.Bookmarks.Exists("hospital") = True Then
                                objdoc.Bookmarks("hospital").Range.Text = IIf(txtHOSCode <> "", Trim(txtHOSCode), "")
                            End If
                            If objdoc.Bookmarks.Exists("dateofadmission") = True Then
                                objdoc.Bookmarks("dateofadmission").Range.Text = IIf(Len(txtDateAdmitted), txtDateAdmitted, "")
                            End If
                            
                                If objdoc.Bookmarks.Exists("policyno") = True Then
                                    objdoc.Bookmarks("policyno").Range.Text = IIf(Len(tmptxtMBMPolNo), tmptxtMBMPolNo, "")
                                End If
                            
                            
                            If tmpRBType <> "" Then
                                If objdoc.Bookmarks.Exists("RB") = True Then
                                    objdoc.Bookmarks("RB").Range.Text = IIf(Len(tmpRBType), tmpRBType, "")
                                End If
                            ElseIf RB <> "" Then
                                If IsNumeric(ICU) Then
                                    If objdoc.Bookmarks.Exists("RB") = True Then
                                        objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - RM" & ICU
                                    End If
                                Else
                                    If objdoc.Bookmarks.Exists("RB") = True Then
                                        objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - FULL COVER"
                                    End If
                                End If
                            End If
                            If objdoc.Bookmarks.Exists("Attention") = True Then
                                objdoc.Bookmarks("Attention").Range.Text = IIf(Len(HOS!HOSContact), HOS!HOSContact, "")
                            End If
                            If objdoc.Bookmarks.Exists("Fax") = True Then
                                objdoc.Bookmarks("Fax").Range.Text = IIf(Len(HOS!HOSFaxNo), HOS!HOSFaxNo, "")
                            End If
                            If objdoc.Bookmarks.Exists("Remark") = True Then
                                objdoc.Bookmarks("Remark").Range.Text = IIf(Len(txtReason), txtReason, "")
                            End If
                            If objdoc.Bookmarks.Exists("GLRemark") = True Then
                                objdoc.Bookmarks("GLRemark").Range.Text = IIf(Len(frmGLMenu.txtremarks), frmGLMenu.txtremarks, "")
                            End If
                            If objdoc.Bookmarks.Exists("LGDate") = True Then
                                objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
                            End If
                            If objdoc.Bookmarks.Exists("Principal") = True Then
                                objdoc.Bookmarks("Principal").Range.Text = txtMBMName
                            End If
                            If objdoc.Bookmarks.Exists("PatientIC") = True Then
                                objdoc.Bookmarks("PatientIC").Range.Text = IIf(Len(txtPatientIC), txtPatientIC, "")
                            End If
                            'If objdoc.Bookmarks.Exists("Status") = True Then
                            '    objdoc.Bookmarks("Status").Range.Text = "(Admission)"
                            'End If
                        End If
                        
                    ElseIf Trim(tmpPLNLGPrivate) = "3" Then
                    Call BindBookmark(objdoc, HOS)
                    bookmark = True
            'TK GL
                    ElseIf Trim(tmpPLNLGPrivate) = "9" Or Trim(tmpPLNLGPrivate) = "12" Or Trim(tmpPLNLGPrivate) = "13" Or Trim(tmpPLNLGPrivate) = "14" Or Trim(tmpPLNLGPrivate) = "17" Or Trim(tmpPLNLGPrivate) = "18" Or Trim(tmpPLNLGPrivate) = "20" Or Trim(tmpPLNLGPrivate) = "21" Or Trim(tmpPLNLGPrivate) = "22" Or Trim(tmpPLNLGPrivate) = "24" Or Trim(tmpPLNLGPrivate) = "25" Or Trim(tmpPLNLGPrivate) = "26" Or Trim(tmpPLNLGPrivate) = "28" Or Trim(tmpPLNLGPrivate) = "31" Or Trim(tmpPLNLGPrivate) = "32" Or Trim(tmpPLNLGPrivate) = "33" Or Trim(tmpPLNLGPrivate) = "34" Or Trim(tmpPLNLGPrivate) = "35" Or Trim(tmpPLNLGPrivate) = "36" Or Trim(tmpPLNLGPrivate) = "38" Or Trim(tmpPLNLGPrivate) = "39" Or Trim(tmpPLNLGPrivate) = "40" Or Trim(tmpPLNLGPrivate) = "41" Or Trim(tmpPLNLGPrivate) = "44" Or Trim(tmpPLNLGPrivate) = "46" Then
                    Call BindBookmark(objdoc, HOS)
                    bookmark = True
                'XL
                     ElseIf Trim(tmpPLNLGPrivate) = "10" Or Trim(tmpPLNLGPrivate) = "11" Or Trim(tmpPLNLGPrivate) = "15" Or Trim(tmpPLNLGPrivate) = "16" Or Trim(tmpPLNLGPrivate) = "23" Or Trim(tmpPLNLGPrivate) = "27" Or Trim(tmpPLNLGPrivate) = "42" Then
                       Call BindBookmark(objdoc, HOS)
                        bookmark = True
                      ElseIf Trim(tmpPLNLGPrivate) = "29" Then
                        Call BindBookmark(objdoc, HOS)
                        bookmark = True
                     ElseIf Trim(tmpPLNLGPrivate) = "30" Or Trim(tmpPLNLGPrivate) = "37" Then
                       Call BindBookmark(objdoc, HOS)
                       bookmark = True
                    Else
                       Call BindBookmark(objdoc, HOS)
                       bookmark = True
                    End If
        'Goverment Hosp GL
            ElseIf Trim(HOSType) = "G" Then
                If txtHOSCode <> "" Then
                    If objdoc.Bookmarks.Exists("hospital") = True Then
                        objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
                    End If
                End If
                
                If HOS!HOSAdd1 <> "" Then
                    If objdoc.Bookmarks.Exists("Add1") = True Then
                        objdoc.Bookmarks("Add1").Range.Text = Trim(HOS!HOSAdd1)
                    End If
                End If
                
                If HOS!HOSAdd2 <> "" Then
                    If objdoc.Bookmarks.Exists("Add2") = True Then
                        objdoc.Bookmarks("Add2").Range.Text = Trim(HOS!HOSAdd2)
                    End If
                End If
                
                If HOS!HOSAdd3 <> "" Then
                    If objdoc.Bookmarks.Exists("Add3") = True Then
                        objdoc.Bookmarks("Add3").Range.Text = Trim(HOS!HOSAdd3)
                    End If
                End If
                
                If Len(HOS!HOSTelNo) Then
                    If objdoc.Bookmarks.Exists("Tel") = True Then
                        objdoc.Bookmarks("Tel").Range.Text = HOS!HOSTelNo
                    End If
                End If
                
                If Len(HOS!HOSFaxNo) Then
                    If objdoc.Bookmarks.Exists("Fax") = True Then
                        objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
                    End If
                End If
                
                If objdoc.Bookmarks.Exists("LGDate") = True Then
                    objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
                End If
                
                If txtMBMName <> "" Then
                    If objdoc.Bookmarks.Exists("Principal") = True Then
                        objdoc.Bookmarks("Principal").Range.Text = txtMBMName
                    End If
                End If
                
                If objdoc.Bookmarks.Exists("StaffNo") = True Then
                    objdoc.Bookmarks("StaffNo").Range.Text = tmpMBMEmpNo
                End If
                
                If objdoc.Bookmarks.Exists("StaffICNo") = True Then
                    objdoc.Bookmarks("StaffICNo").Range.Text = txtMBMIcBcPP
                End If
                
                If objdoc.Bookmarks.Exists("Nationality") = True Then
                    objdoc.Bookmarks("Nationality").Range.Text = "Malaysian"
                End If
                
                If Trim(cboPatientList) <> "" Then
                    If objdoc.Bookmarks.Exists("patient") = True Then
                        objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                    End If
                Else
                    If objdoc.Bookmarks.Exists("patient") = True Then
                        objdoc.Bookmarks("patient").Range.Text = " Please key in patient name here!"
                    End If
                End If
                
                If tmpRel = "Employee" Then
                    If objdoc.Bookmarks.Exists("Rel") = True Then
                        objdoc.Bookmarks("Rel").Range.Text = "Pekerja"
                    End If
                Else
                    If objdoc.Bookmarks.Exists("Rel") = True Then
                        objdoc.Bookmarks("Rel").Range.Text = "Tanggungan"
                    End If
                End If
                
                If objdoc.Bookmarks.Exists("PatientICNo") = True Then
                    objdoc.Bookmarks("PatientICNo").Range.Text = txtPatientIC
                End If
                
                If LGNumber <> "" Then
                    If objdoc.Bookmarks.Exists("LGno") = True Then
                        objdoc.Bookmarks("LGno").Range.Text = LGNumber
                    End If
                End If
                
               LGVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
        'Normal Medix Private Hosp GL
        Else
                If Trim(cboPatientList) <> "" Then
                    If objdoc.Bookmarks.Exists("patient") = True Then
                        objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                    End If
                Else
                    If objdoc.Bookmarks.Exists("patient") = True Then
                        objdoc.Bookmarks("patient").Range.Text = " Please key in patient name here!"
                    End If
                End If
                
                If LGNumber <> "" Then
                    If objdoc.Bookmarks.Exists("LGno") = True Then
                        objdoc.Bookmarks("LGno").Range.Text = LGNumber
                    End If
                End If
                    
                LGVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
                    
                If frmGLMenu.txtCASLGAmt <> "" Then
                    If objdoc.Bookmarks.Exists("amount") = True Then
                        objdoc.Bookmarks("amount").Range.Text = Format(frmGLMenu.txtCASLGAmt, "#,##0.00")
                    End If
                End If
                    
                If Letter = "C" Then
                    If objdoc.Bookmarks.Exists("GrpCompany") = True Then
                        objdoc.Bookmarks("GrpCompany").Range.Text = " - " & tmpMBMGroupCompany
                    End If
                End If
                
                If objdoc.Bookmarks.Exists("LGID") = True Then
                    objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
                End If
                
                If txtMBMNumber <> "" Then
                    If objdoc.Bookmarks.Exists("mbmnumber") = True Then
                        objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
                    End If
                End If
                
                If txtHOSCode <> "" Then
                    If objdoc.Bookmarks.Exists("hospital") = True Then
                        objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
                    End If
                End If
                
                If txtDateAdmitted <> "" Then
                    If objdoc.Bookmarks.Exists("dateofadmission") = True Then
                        objdoc.Bookmarks("dateofadmission").Range.Text = txtDateAdmitted
                    End If
                End If
                
                If tmptxtMBMPolNo <> "" Then
                    If objdoc.Bookmarks.Exists("policyno") = True Then
                        objdoc.Bookmarks("policyno").Range.Text = tmptxtMBMPolNo
                    End If
                End If
                
                If tmpRBType <> "" Then
                    If objdoc.Bookmarks.Exists("RB") = True Then
                        objdoc.Bookmarks("RB").Range.Text = tmpRBType
                    End If
                ElseIf RB <> "" Then
                    If IsNumeric(ICU) Then
                        If objdoc.Bookmarks.Exists("RB") = True Then
                            objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - RM" & ICU
                        End If
                    Else
                        If objdoc.Bookmarks.Exists("RB") = True Then
                            objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - FULL COVER"
                        End If
                    End If
                End If
                
                If Len(HOS!HOSContact) Then
                    If objdoc.Bookmarks.Exists("Attention") = True Then
                        objdoc.Bookmarks("Attention").Range.Text = HOS!HOSContact
                    End If
                End If
                
                If Len(HOS!HOSFaxNo) Then
                    If objdoc.Bookmarks.Exists("Fax") = True Then
                        objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
                    End If
                End If
                
                If lstLGListing.ListItems.Count <> 0 Then
                    If objdoc.Bookmarks.Exists("Remark") = True Then
                        objdoc.Bookmarks("Remark").Range.Text = Trim(lstLGListing.SelectedItem.SubItems(8))
                    End If
                End If
                
                If objdoc.Bookmarks.Exists("LGDate") = True Then
                    objdoc.Bookmarks("LGDate").Range.Text = Date
                End If
                
                  If objdoc.Bookmarks.Exists("Diagnosis") = True Then
                                objdoc.Bookmarks("Diagnosis").Range.Text = txtReason
                            End If
                
        End If
    ElseIf Letter = "B" Then
        bookmark = False
                If Trim(HOSType) = "P" Then
                    If Mid(frmGLMenu.cbPurpose, 1, 1) = "O" And (Trim(tmpPLNLGPrivate) <> "5" Or Trim(tmpPLNLGPrivate) <> "6" Or Trim(tmpPLNLGPrivate) <> "7" Or Trim(tmpPLNLGPrivate) Or "8") Then
                            If Trim(cboPatientList) <> "" Then
                                If objdoc.Bookmarks.Exists("patient") = True Then
                                    objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                                End If
                            Else
                                If objdoc.Bookmarks.Exists("patient") = True Then
                                    objdoc.Bookmarks("patient").Range.Text = " Please key in patient name here!"
                                End If
                            End If
                            If LGNumber <> "" Then
                                If objdoc.Bookmarks.Exists("LGno") = True Then
                                    objdoc.Bookmarks("LGno").Range.Text = LGNumber
                                End If
                            End If
                                
                            LGVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
                            If objdoc.Bookmarks.Exists("GrpCompany") = True Then
                                objdoc.Bookmarks("GrpCompany").Range.Text = tmpMBMGroupCompany
                            End If
                                
                            If txtMBMName <> "" Then
                                If objdoc.Bookmarks.Exists("Add1") = True Then
                                    objdoc.Bookmarks("Principal").Range.Text = txtMBMName
                                End If
                            End If
                            
                            If objdoc.Bookmarks.Exists("amount") = True Then
                                objdoc.Bookmarks("amount").Range.Text = Format(frmGLMenu.txtCASLGAmt, "#,##0.00")
                            End If
                            
                            If objdoc.Bookmarks.Exists("LGID") = True Then
                                objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
                            End If
                            
                            If objdoc.Bookmarks.Exists("Diagnosis") = True Then
                                objdoc.Bookmarks("Diagnosis").Range.Text = txtReason
                            End If
                            
                            If txtMBMNumber <> "" Then
                                If objdoc.Bookmarks.Exists("mbmnumber") = True Then
                                    objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
                                End If
                            End If
                            
                            If txtHOSCode <> "" Then
                                If objdoc.Bookmarks.Exists("hospital") = True Then
                                    objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
                                End If
                            End If
                            
                            If txtDateAdmitted <> "" Then
                                If objdoc.Bookmarks.Exists("dateofadmission") = True Then
                                    objdoc.Bookmarks("dateofadmission").Range.Text = txtDateAdmitted
                                End If
                            End If
                            
                            If Len(HOS!HOSFaxNo) Then
                                If objdoc.Bookmarks.Exists("Fax") = True Then
                                    objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
                                End If
                            End If
                            
                            If frmGLMenu.txtremarks <> "" Then
                                If objdoc.Bookmarks.Exists("Remark") = True Then
                                    objdoc.Bookmarks("Remark").Range.Text = frmGLMenu.txtremarks
                                End If
                            End If
                            
                            If objdoc.Bookmarks.Exists("Add1") = True Then
                                objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
                            End If
                            
                            If objdoc.Bookmarks.Exists("AttDr") = True Then
                                objdoc.Bookmarks("AttDr").Range.Text = txtATTDoc(0)
                            End If
                            
                            
                            bookmark = True
                    ElseIf Trim(tmpPLNLGPrivate) = "3" Then
                        If Trim(cboPatientList) <> "" Then
                            If objdoc.Bookmarks.Exists("patient") = True Then
                                objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                            End If
                        Else
                            If objdoc.Bookmarks.Exists("patient") = True Then
                                objdoc.Bookmarks("patient").Range.Text = " Please key in patient name here!"
                            End If
                        End If
                        
                        If txtMBMName <> "" Then
                            If objdoc.Bookmarks.Exists("Principal") = True Then
                                objdoc.Bookmarks("Principal").Range.Text = txtMBMName
                            End If
                        End If
                        
                        If objdoc.Bookmarks.Exists("ServiceData") = True Then
                            objdoc.Bookmarks("ServiceData").Range.Text = GetServerTime
                        End If
                        
                        If objdoc.Bookmarks.Exists("mbmnumber") = True Then
                            objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
                        End If
                        
                        If txtHOSCode <> "" Then
                            If objdoc.Bookmarks.Exists("hospital") = True Then
                                objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
                            End If
                        End If
                        
                        If HOS!HOSAdd1 <> "" Then
                            If objdoc.Bookmarks.Exists("Add1") = True Then
                                objdoc.Bookmarks("Add1").Range.Text = Trim(HOS!HOSAdd1)
                            End If
                        End If
                        
                        If HOS!HOSAdd2 <> "" Then
                            If objdoc.Bookmarks.Exists("Add2") = True Then
                                objdoc.Bookmarks("Add2").Range.Text = Trim(HOS!HOSAdd2)
                            End If
                        End If
                        
                        If HOS!HOSAdd3 <> "" Then
                            If objdoc.Bookmarks.Exists("Add3") = True Then
                                objdoc.Bookmarks("Add3").Range.Text = Trim(HOS!HOSAdd3)
                            End If
                        End If
                        
                        If Len(HOS!HOSContact) Then
                            If objdoc.Bookmarks.Exists("Attention") = True Then
                                objdoc.Bookmarks("Attention").Range.Text = HOS!HOSContact
                            End If
                        End If
                        
                        If Len(HOS!HOSTelNo) Then
                            If objdoc.Bookmarks.Exists("Tel") = True Then
                                objdoc.Bookmarks("Tel").Range.Text = HOS!HOSTelNo
                            End If
                        End If
                        
                        If Len(HOS!HOSFaxNo) Then
                            If objdoc.Bookmarks.Exists("Fax") = True Then
                                objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
                            End If
                        End If
                        
                        If txtDateAdmitted <> "" Then
                            If objdoc.Bookmarks.Exists("DOA") = True Then
                                objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
                            End If
                        End If
                        
                        If txtDateDischarge <> "__/__/____" Then
                            If objdoc.Bookmarks.Exists("lblDOD") = True Then
                                objdoc.Bookmarks("lblDOD").Range.Text = "Discharge Date"
                            End If
                            If objdoc.Bookmarks.Exists("lblDODDot") = True Then
                                objdoc.Bookmarks("lblDODDot").Range.Text = ":"
                            End If
                            If objdoc.Bookmarks.Exists("DOD") = True Then
                                objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
                            End If
                        End If
                        
                        If objdoc.Bookmarks.Exists("ICU") = True Then
                            objdoc.Bookmarks("ICU").Range.Text = ICU
                        End If
                        
                            If IsNumeric(ICU) Then
                                If objdoc.Bookmarks.Exists("RB") = True Then
                                    objdoc.Bookmarks("RB").Range.Text = tmpRMAmt & " (Inclusive of Meal)" & "/" & " ICU - RM" & ICU
                                End If
                            Else
                                If objdoc.Bookmarks.Exists("RB") = True Then
                                    objdoc.Bookmarks("RB").Range.Text = tmpRMAmt & " (Inclusive of Meal)" & "/" & " ICU - FULL COVER"
                                End If
                            End If
                            
                            If Len(lstLGListing.SelectedItem.Text) Then
                                If objdoc.Bookmarks.Exists("LGno") = True Then
                                    objdoc.Bookmarks("LGno").Range.Text = Trim(lstLGListing.SelectedItem.Text)
                                End If
                            End If
                            
                            LGVer = Trim(Mid(lstLGListing.SelectedItem.SubItems(9), InStr(1, lstLGListing.SelectedItem.SubItems(9), "-") + 1, Len(lstLGListing.SelectedItem.SubItems(9)) - InStr(1, lstLGListing.SelectedItem.SubItems(9), "-")))
                            If LGVer = "1" Then
                                If objdoc.Bookmarks.Exists("Initial") = True Then
                                    objdoc.Bookmarks("Initial").Range.Text = "Initial "
                                End If
                            End If
                            
                            If objdoc.Bookmarks.Exists("GLAmt") = True Then
                                objdoc.Bookmarks("GLAmt").Range.Text = Format(lstLGListing.SelectedItem.SubItems(2), "#,##0.00")
                            End If
                            
                            If objdoc.Bookmarks.Exists("GrpCompany") = True Then
                                objdoc.Bookmarks("GrpCompany").Range.Text = tmpMBMGroupCompany
                            End If
                        If frmGLMenu.txtremarks <> "" Then
                            If objdoc.Bookmarks.Exists("Remark") = True Then
                                objdoc.Bookmarks("Remark").Range.Text = frmGLMenu.txtremarks
                            End If
                        End If
                        
                    ElseIf Trim(tmpPLNLGPrivate) = "9" Or Trim(tmpPLNLGPrivate) = "17" Or Trim(tmpPLNLGPrivate) = "19" Or Trim(tmpPLNLGPrivate) = "26" Or Trim(tmpPLNLGPrivate) = "35" Or Trim(tmpPLNLGPrivate) = "36" Or Trim(tmpPLNLGPrivate) = "44" Or Trim(tmpPLNLGPrivate) = "46" Then
                        If Trim(cboPatientList) <> "" Then
                          If objdoc.Bookmarks.Exists("patient") = True Then
                            objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                          End If
                        Else
                            If objdoc.Bookmarks.Exists("patient") = True Then
                                objdoc.Bookmarks("patient").Range.Text = " Please key in patient name here!"
                            End If
                        End If
                                                                       
                        If objdoc.Bookmarks.Exists("LGno") = True Then
                            objdoc.Bookmarks("LGno").Range.Text = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
                        End If
                                                
                        If LGNumber <> "" Then
                            If objdoc.Bookmarks.Exists("LGno") = True Then
                                objdoc.Bookmarks("LGno").Range.Text = LGNumber
                            End If
                        End If
                        
                        LGVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
                            
                        If Trim(tmpPLNLGPrivate) <> "35" Then
                         If frmGLMenu.txtCASLGAmt <> "" Then
                             If objdoc.Bookmarks.Exists("amount") = True Then
                                 objdoc.Bookmarks("amount").Range.Text = Format(frmGLMenu.txtCASLGAmt, "#,##0.00")
                             End If
                         End If
                         End If
                            
                        If objdoc.Bookmarks.Exists("GrpCompany") = True Then
                            objdoc.Bookmarks("GrpCompany").Range.Text = tmpMBMGroupCompany
                        End If
                        
                        If objdoc.Bookmarks.Exists("LGID") = True Then
                            objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
                        End If
                        
                        If txtMBMNumber <> "" Then
                            If objdoc.Bookmarks.Exists("mbmnumber") = True Then
                                objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
                            End If
                        End If
                        
                        If txtHOSCode <> "" Then
                            If objdoc.Bookmarks.Exists("hospital") = True Then
                                objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
                            End If
                        End If
                        
                        If txtDateAdmitted <> "" Then
                            If objdoc.Bookmarks.Exists("dateofadmission") = True Then
                                objdoc.Bookmarks("dateofadmission").Range.Text = txtDateAdmitted
                            End If
                        End If
                        
                        If objdoc.Bookmarks.Exists("Diagnosis") = True Then
                            objdoc.Bookmarks("Diagnosis").Range.Text = txtReason
                        End If
                        
                        If txtDateDischarge <> "" Then
                            If objdoc.Bookmarks.Exists("DOD") = True Then
                                objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
                            End If
                        End If
                        
                        If txtMBMName <> "" Then
                            If objdoc.Bookmarks.Exists("Principal") = True Then
                                objdoc.Bookmarks("Principal").Range.Text = txtMBMName
                            End If
                        End If
                                                
                        If tmpRBType <> "" Then
                            If objdoc.Bookmarks.Exists("RB") = True Then
                                objdoc.Bookmarks("RB").Range.Text = tmpRBType
                            End If
                        ElseIf RB <> "" Then
                            If IsNumeric(ICU) Then
                                If objdoc.Bookmarks.Exists("RB") = True Then
                                    objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - RM" & ICU
                                End If
                            Else
                                If objdoc.Bookmarks.Exists("RB") = True Then
                                    objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - FULL COVER"
                                End If
                            End If
                        End If
                
                        If Len(HOS!HOSFaxNo) Then
                            If objdoc.Bookmarks.Exists("Fax") = True Then
                                objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
                            End If
                        End If
                        
                        If frmGLMenu.txtremarks <> "" Then
                            If objdoc.Bookmarks.Exists("Remark") = True Then
                                objdoc.Bookmarks("Remark").Range.Text = frmGLMenu.txtremarks
                            End If
                        End If
                        
                        If objdoc.Bookmarks.Exists("LGDate") = True Then
                            objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
                        End If
                        bookmark = True
            Else
            
                If Trim(cboPatientList) <> "" Then
                    If objdoc.Bookmarks.Exists("patient") = True Then
                        objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                    End If
                Else
                    If objdoc.Bookmarks.Exists("patient") = True Then
                        objdoc.Bookmarks("patient").Range.Text = " Please key in patient name here!"
                    End If
                End If
                
                'If Len(lstLGListing.SelectedItem.Text) Then
                    If objdoc.Bookmarks.Exists("BTBNo") = True Then
                        objdoc.Bookmarks("BTBNo").Range.Text = Trim(lstLGListing.SelectedItem.Text)
                    End If
                'End If
                    
                If LGNumber <> "" Then
                    If objdoc.Bookmarks.Exists("LGno") = True Then
                        objdoc.Bookmarks("LGno").Range.Text = LGNumber
                    End If
                End If
                    
                LGVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
                    
                If objdoc.Bookmarks.Exists("GrpCompany") = True Then
                    objdoc.Bookmarks("GrpCompany").Range.Text = tmpMBMGroupCompany
                End If
                    
                If objdoc.Bookmarks.Exists("BTBID") = True Then
                    objdoc.Bookmarks("BTBID").Range.Text = Trim(Logon)
                End If
                
                If objdoc.Bookmarks.Exists("folDays") = True Then
                    objdoc.Bookmarks("folDays").Range.Text = Trim(tmpFolDays)
                End If
                                
                If txtMBMNumber <> "" Then
                    If objdoc.Bookmarks.Exists("mbmnumber") = True Then
                        objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
                    End If
                End If
                
                If objdoc.Bookmarks.Exists("Principal") = True Then
                    objdoc.Bookmarks("Principal").Range.Text = IIf(Len(txtMBMName), txtMBMName, "")
                End If
                
                If txtHOSCode <> "" Then
                    If objdoc.Bookmarks.Exists("hospital") = True Then
                        objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
                    End If
                End If
                
                If frmGLMenu.txtremarks <> "" Then
                    If objdoc.Bookmarks.Exists("Remark") = True Then
                        objdoc.Bookmarks("Remark").Range.Text = frmGLMenu.txtremarks
                    End If
                End If
        
                If txtDateAdmitted <> "" Then
                    If objdoc.Bookmarks.Exists("dateofadmission") = True Then
                        objdoc.Bookmarks("dateofadmission").Range.Text = txtDateAdmitted
                    End If
                End If
                
                If tmptxtMBMPolNo <> "" Then
                    If objdoc.Bookmarks.Exists("policyno") = True Then
                        objdoc.Bookmarks("policyno").Range.Text = tmptxtMBMPolNo
                    End If
                End If
                
                    If objdoc.Bookmarks.Exists("Diagnosis") = True Then
                        objdoc.Bookmarks("Diagnosis").Range.Text = txtReason
                    End If
        
                If RB <> "" Then
                    If IsNumeric(ICU) Then
                        If objdoc.Bookmarks.Exists("RB") = True Then
                            objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - RM" & ICU
                        End If
                    Else
                        If objdoc.Bookmarks.Exists("RB") = True Then
                            objdoc.Bookmarks("RB").Range.Text = RB & " (Inclusive of Meal)" & "/" & " ICU - FULL COVER"
                        End If
                    End If
                End If
                
                If Len(HOS!HOSContact) Then
                    If objdoc.Bookmarks.Exists("Attention") = True Then
                        objdoc.Bookmarks("Attention").Range.Text = HOS!HOSContact
                    End If
                End If
            
                If Len(HOS!HOSFaxNo) Then
                    If objdoc.Bookmarks.Exists("Fax") = True Then
                        objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
                    End If
                End If
                
                If objdoc.Bookmarks.Exists("LGDate") = True Then
                    objdoc.Bookmarks("LGDate").Range.Text = Date
                End If
                
                If tmpHltCode = "S" Then
                    strAppendCase = "19"
                    StrAppend1 = " AND PLNCode = '" & Trim(txtPlanCode) & "'"
                Else
                    strAppendCase = "19"
                    StrAppend1 = " AND PLNCode = '" & Trim(txtPlanCode) & "' AND PAYCode = '" & Trim(txtPayCode) & "'"
                End If
                Set cmd2.ActiveConnection = medixmasterconn
                cmd2.CommandText = "Case_Adjustment"
                cmd2.CommandType = adCmdStoredProc
                cmd2.CommandTimeout = 300
                Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
                cmd2.Parameters.Append Prm1
                Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd2.Parameters.Append Prm2
                MBM.CursorLocation = adUseClient
                MBM.CursorType = adOpenForwardOnly
                MBM.CacheSize = 100
                Set MBM = cmd2.Execute
                
                If MBM.EOF = False Then
                    If MBM!PLNMeal = "Y" Then
                        Including = Including + "Meal"
                    End If
                    If MBM!PLNNursing = "Y" Then
                        Including = Including + " " & ",Nursing Charges"
                    End If
                    If MBM!PLNTax = "Y" Then
                        Including = Including + " " & ",Tax"
                    End If
                End If
                
                If objdoc.Bookmarks.Exists("LGDate") = True Then
                    objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
                End If
                If objdoc.Bookmarks.Exists("string") = True Then
                    objdoc.Bookmarks("string").Range.Text = Terms
                End If
                If tmpPAYLGDesc <> "" Then
                    If objdoc.Bookmarks.Exists("PayDesc") = True Then
                        objdoc.Bookmarks("PayDesc").Range.Text = tmpPAYLGDesc
                    End If
                End If
                
                bookmark = True
            End If
        Else
                If txtHOSCode <> "" Then
                    If objdoc.Bookmarks.Exists("hospital") = True Then
                        objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
                    End If
                End If
                
                If HOS!HOSAdd1 <> "" Then
                    If objdoc.Bookmarks.Exists("Add1") = True Then
                        objdoc.Bookmarks("Add1").Range.Text = Trim(HOS!HOSAdd1)
                    End If
                End If
                  
                If HOS!HOSAdd2 <> "" Then
                    If objdoc.Bookmarks.Exists("Add2") = True Then
                        objdoc.Bookmarks("Add2").Range.Text = Trim(HOS!HOSAdd2)
                    End If
                End If
                
                If HOS!HOSAdd3 <> "" Then
                    If objdoc.Bookmarks.Exists("Add3") = True Then
                        objdoc.Bookmarks("Add3").Range.Text = Trim(HOS!HOSAdd3)
                    End If
                End If
                
                If Len(HOS!HOSTelNo) Then
                    If objdoc.Bookmarks.Exists("Tel") = True Then
                        objdoc.Bookmarks("Tel").Range.Text = HOS!HOSTelNo
                    End If
                End If
                
                If Len(HOS!HOSFaxNo) Then
                    If objdoc.Bookmarks.Exists("Fax") = True Then
                        objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
                    End If
                End If
        
                If objdoc.Bookmarks.Exists("LGDate") = True Then
                    objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
                End If
        
                If txtMBMName <> "" Then
                    If objdoc.Bookmarks.Exists("Principal") = True Then
                        objdoc.Bookmarks("Principal").Range.Text = txtMBMName
                    End If
                End If
                                                
                If objdoc.Bookmarks.Exists("StaffNo") = True Then
                    objdoc.Bookmarks("StaffNo").Range.Text = tmpMBMEmpNo
                End If
                
                If objdoc.Bookmarks.Exists("StaffICNo") = True Then
                    objdoc.Bookmarks("StaffICNo").Range.Text = txtMBMIcBcPP
                End If
                
                If objdoc.Bookmarks.Exists("Nationality") = True Then
                    objdoc.Bookmarks("Nationality").Range.Text = "Malaysian"
                End If
                
                If Trim(cboPatientList) <> "" Then
                    If objdoc.Bookmarks.Exists("patient") = True Then
                        objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                    End If
                Else
                    If objdoc.Bookmarks.Exists("patient") = True Then
                        objdoc.Bookmarks("patient").Range.Text = " Please key in patient name here!"
                    End If
                End If
                
                If objdoc.Bookmarks.Exists("Rel") = True Then
                    objdoc.Bookmarks("Rel").Range.Text = tmpRel
                End If
                If objdoc.Bookmarks.Exists("PatientICNo") = True Then
                    objdoc.Bookmarks("PatientICNo").Range.Text = txtPatientIC
                End If
                If objdoc.Bookmarks.Exists("LGno") = True Then
                    objdoc.Bookmarks("LGno").Range.Text = LGNumber
                End If
                
        End If
    ElseIf Letter = "R" Then
        bookmark = False
        strsqlRepRea = ""
        
        If Trim(tmpRepStatus) <> "" And Trim(tmpCashType) = "R" Then
            If objdoc.Bookmarks.Exists("PrincipalName") = True Then
                objdoc.Bookmarks("PrincipalName").Range.Text = txtMBMName
            End If
            
            If objdoc.Bookmarks.Exists("Add1") = True Then
                    objdoc.Bookmarks("Add1").Range.Text = txtAdd1
            End If
            
            If objdoc.Bookmarks.Exists("Add2") = True Then
                objdoc.Bookmarks("Add2").Range.Text = txtAdd2 & txtAdd3
            End If
            
            If objdoc.Bookmarks.Exists("PostCode") = True Then
                objdoc.Bookmarks("PostCode").Range.Text = txtPostCode & "" & txtCity
            End If
            
            If objdoc.Bookmarks.Exists("State") = True Then
                objdoc.Bookmarks("State").Range.Text = txtState
            End If
            
            If Trim(cboPatientList) <> "" Then
                If objdoc.Bookmarks.Exists("patientname") = True Then
                    objdoc.Bookmarks("patientname").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
                End If
            Else
                If objdoc.Bookmarks.Exists("patientname") = True Then
                    objdoc.Bookmarks("patientname").Range.Text = " Please key in patient name here!"
                End If
            End If
            
            If tmpProductName <> "" Then
                If objdoc.Bookmarks.Exists("Product") = True Then
                    objdoc.Bookmarks("Product").Range.Text = tmpProductName
                End If
            End If
            
            If txtDateAdmitted <> "" Then
                If objdoc.Bookmarks.Exists("DOA") = True Then
                    objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
                End If
            End If
            
            If tmptxtMBMPolNo <> "" Then
                If objdoc.Bookmarks.Exists("policynumber") = True Then
                    objdoc.Bookmarks("policynumber").Range.Text = tmptxtMBMPolNo
                End If
            End If
            
            If tmptxtMBMPolNo <> "" Then
                If objdoc.Bookmarks.Exists("MAAPolNo") = True Then
                    objdoc.Bookmarks("MAAPolNo").Range.Text = tmptxtMBMPolNo
                End If
            End If
            
            If txtFinalDiag1(0) <> "" Then
                Dim strsql5 As String
                strsql5 = strsql5 + txtFinalDiag1(0)
                If txtFinalDiag1(1) <> "" Then
                    strsql5 = strsql5 + ", " + txtFinalDiag1(1)
                    If txtFinalDiag1(2) <> "" Then
                        strsql5 = strsql5 + ", " + txtFinalDiag1(2)
                    End If
                End If
                If objdoc.Bookmarks.Exists("Reasons") = True Then
                    objdoc.Bookmarks("Reasons").Range.Text = strsql5
                End If
                
            End If
            
            If Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - 1)) <> "RC6" Then
                If Len(Trim(cboREPCode1)) Then
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
                        strsqlRepRea = Trim(REPCat!RepCatName)
                    End If
                    
                    REP.Close
                    REPCat.Close
                    Set REP = Nothing
                    Set REPCat = Nothing
                End If
                
                If Len(Trim(cboREPCode2)) Then
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
                        strsqlRepRea = strsqlRepRea + ", " & Trim(REPCat!RepCatName)
                    End If
                    
                    REP.Close
                    REPCat.Close
                    Set REP = Nothing
                    Set REPCat = Nothing
                End If
                
                If objdoc.Bookmarks.Exists("RepReason") = True Then
                    objdoc.Bookmarks("RepReason").Range.Text = strsqlRepRea
                End If
            End If
            
            If objdoc.Bookmarks.Exists("ReasonAdm") = True Then
                objdoc.Bookmarks("ReasonAdm").Range.Text = txtReason
            End If
            If objdoc.Bookmarks.Exists("Reasons3") = True Then
                objdoc.Bookmarks("Reasons3").Range.Text = strsql5
            End If
            If objdoc.Bookmarks.Exists("Reason") = True Then
                objdoc.Bookmarks("Reason").Range.Text = strsql5
            End If
            If objdoc.Bookmarks.Exists("RepReason2") = True Then
                objdoc.Bookmarks("RepReason2").Range.Text = strsqlRepRea
            End If
            If objdoc.Bookmarks.Exists("RepdName") = True Then
                objdoc.Bookmarks("RepdName").Range.Text = IIf(Len(PAY!PAYRepdName), PAY!PAYRepdName, "")
            End If
            If objdoc.Bookmarks.Exists("RepdDivision") = True Then
                objdoc.Bookmarks("RepdDivision").Range.Text = IIf(Len(PAY!PAYRepdDivision), PAY!PAYRepdDivision, "")
            End If
            If objdoc.Bookmarks.Exists("RepdTel") = True Then
                objdoc.Bookmarks("RepdTel").Range.Text = IIf(Len(PAY!PAYTELno), PAY!PAYTELno, "")
            End If
            If objdoc.Bookmarks.Exists("RepdEmail") = True Then
                objdoc.Bookmarks("RepdEmail").Range.Text = IIf(Len(PAY!PAYRepdEmail), PAY!PAYRepdEmail, "")
            End If
            If objdoc.Bookmarks.Exists("INSBy") = True Then
                objdoc.Bookmarks("INSBy").Range.Text = PAY!PAYCompanyName
            End If
        Else
        End If
        
    End If
    End If
        HOS.Close
        Set HOS = Nothing
    End Function
Public Sub BindBookmark(objdoc As Word.Document, HOS As ADODB.Recordset)
    If objdoc.Bookmarks.Exists("patient") = True Then
        objdoc.Bookmarks("patient").Range.Text = IIf(Trim(cboPatientList) <> "", Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-"))), " Please key in patient name here!")
    End If
    If objdoc.Bookmarks.Exists("Principal") = True Then
        objdoc.Bookmarks("Principal").Range.Text = IIf(Len(txtMBMName), txtMBMName, "")
    End If
    If objdoc.Bookmarks.Exists("ServiceData") = True Then
        objdoc.Bookmarks("ServiceData").Range.Text = GetServerTime
    End If
    If objdoc.Bookmarks.Exists("mbmnumber") = True Then
        objdoc.Bookmarks("mbmnumber").Range.Text = IIf(Len(txtMBMNumber), txtMBMNumber, "")
    End If
    If objdoc.Bookmarks.Exists("hospital") = True Then
        objdoc.Bookmarks("hospital").Range.Text = IIf(Len(Trim(txtHOSCode)), Trim(txtHOSCode), "")
    End If
    If objdoc.Bookmarks.Exists("Add1") = True Then
        objdoc.Bookmarks("Add1").Range.Text = IIf(Len(Trim(HOS!HOSAdd1)), Trim(HOS!HOSAdd1), "")
    End If
    If objdoc.Bookmarks.Exists("Add2") = True Then
        objdoc.Bookmarks("Add2").Range.Text = IIf(Len(Trim(HOS!HOSAdd2)), Trim(HOS!HOSAdd2), "")
    End If
    If objdoc.Bookmarks.Exists("Add3") = True Then
        objdoc.Bookmarks("Add3").Range.Text = IIf(Len(Trim(HOS!HOSAdd3)), Trim(HOS!HOSAdd3), "")
    End If
    If objdoc.Bookmarks.Exists("Attention") = True Then
        objdoc.Bookmarks("Attention").Range.Text = IIf(Len(HOS!HOSContact), HOS!HOSContact, "")
    End If
    If objdoc.Bookmarks.Exists("Tel") = True Then
        objdoc.Bookmarks("Tel").Range.Text = IIf(Len(HOS!HOSTelNo), HOS!HOSTelNo, "")
    End If
    If objdoc.Bookmarks.Exists("Fax") = True Then
        objdoc.Bookmarks("Fax").Range.Text = IIf(Len(HOS!HOSFaxNo), HOS!HOSFaxNo, "")
    End If
    If objdoc.Bookmarks.Exists("DOA") = True Then
        objdoc.Bookmarks("DOA").Range.Text = IIf(Len(txtDateAdmitted), txtDateAdmitted, "")
    End If
    If objdoc.Bookmarks.Exists("GRPCompany") = True Then
        objdoc.Bookmarks("GRPCompany").Range.Text = IIf(Len(tmpMBMGroupCompany), tmpMBMGroupCompany, "")
    End If
    
    If txtDateDischarge <> "__/__/____" Then
        If objdoc.Bookmarks.Exists("lblDOD") = True Then
            objdoc.Bookmarks("lblDOD").Range.Text = "Discharge Date"
        End If
        If objdoc.Bookmarks.Exists("lblDODDot") = True Then
            objdoc.Bookmarks("lblDODDot").Range.Text = ":"
        End If
        If objdoc.Bookmarks.Exists("DOD") = True Then
            objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
        End If
    End If
    If objdoc.Bookmarks.Exists("ICU") = True Then
        objdoc.Bookmarks("ICU").Range.Text = ICU
    End If
    
   ' If IsNumeric(ICU) Then
   '     If ObjDoc.Bookmarks.Exists("RB") = True Then
   '         ObjDoc.Bookmarks("RB").Range.Text = tmpRMAmt & " (Inclusive of Meal)" & "/" & " ICU - RM" & ICU
   '     End If
   ' Else
   '     If ObjDoc.Bookmarks.Exists("RB") = True Then
   '         ObjDoc.Bookmarks("RB").Range.Text = tmpRMAmt & " (Inclusive of Meal)" & "/" & " ICU - FULL COVER"
   '     End If
   ' End If
    If objdoc.Bookmarks.Exists("LGno") = True Then
        objdoc.Bookmarks("LGno").Range.Text = IIf(Len(LGNumber), LGNumber, "")
    End If
    LGVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
    
        If LGVer = "1" Then
            If objdoc.Bookmarks.Exists("Initial") = True Then
                objdoc.Bookmarks("Initial").Range.Text = "Initial "
            End If
        End If
        If txtCASLGAmt <> "" Then
            If objdoc.Bookmarks.Exists("GLAmt") = True Then
                objdoc.Bookmarks("GLAmt").Range.Text = Format(frmGLMenu.txtCASLGAmt, "#,##0.00")
            End If
        End If
        If frmGLMenu.txtremarks <> "" Then
            If objdoc.Bookmarks.Exists("Remark") = True Then
                objdoc.Bookmarks("Remark").Range.Text = frmGLMenu.txtremarks
            End If
        End If
                            
       ' If Trim(tmpPLNLGPrivate) <> "34" And Trim(tmpPLNLGPrivate) <> "35" Then
       '     If frmGLMenu.txtCASLGAmt <> "" Then
       '         If ObjDoc.Bookmarks.Exists("amount") = True Then
       '             ObjDoc.Bookmarks("amount").Range.Text = Format(frmGLMenu.txtCASLGAmt, "#,##0.00")
       '         End If
       '     End If
       ' End If
        If objdoc.Bookmarks.Exists("LGID") = True Then
            objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
        End If
        If objdoc.Bookmarks.Exists("Diagnosis") = True Then
            objdoc.Bookmarks("Diagnosis").Range.Text = txtReason
        End If
        If txtDateAdmitted <> "" Then
            If objdoc.Bookmarks.Exists("dateofadmission") = True Then
                objdoc.Bookmarks("dateofadmission").Range.Text = txtDateAdmitted
            End If
        End If
                        
        If txtDateDischarge <> "" Then
            If objdoc.Bookmarks.Exists("DOD") = True Then
                objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
            End If
        End If
    If objdoc.Bookmarks.Exists("LGDate") = True Then
        objdoc.Bookmarks("LGDate").Range.Text = GetServerTime
    End If
    
    If tmptxtMBMPolNo <> "" Then
        If objdoc.Bookmarks.Exists("policyno") = True Then
            objdoc.Bookmarks("policyno").Range.Text = tmptxtMBMPolNo
        End If
    End If
                  
                  
        If frmGLMenu.txtCASLGAmt <> "" Then
            If objdoc.Bookmarks.Exists("amount") = True Then
                objdoc.Bookmarks("amount").Range.Text = Format(frmGLMenu.txtCASLGAmt, "#,##0.00")
            End If
        End If
        
        If IsNumeric(ICU) Then
            If objdoc.Bookmarks.Exists("RB") = True Then
                objdoc.Bookmarks("RB").Range.Text = RB & " / " & tmpRBType & " /" & " ICU - RM" & ICU
            End If
        Else
            If objdoc.Bookmarks.Exists("RB") = True Then
                objdoc.Bookmarks("RB").Range.Text = RB & " / " & tmpRBType & " /" & " ICU - FULL COVER"
            End If
        End If
                        
End Sub

    
Private Function BookMarkNote(Optional strAppend As String) As Integer
        If Trim(txtCaseNo) <> "" Then
            objdoc.Bookmarks("CASE").Range.Text = Trim(txtCaseNo)
        End If
        If Trim(txtCASRemark) <> "" Then
          objdoc.Bookmarks("Note").Range.Text = Trim(txtCASRemark) & Trim(txtCASRemarkTwo) & Trim(txtCASRemark3) & Trim(txtCASRemark4) & Trim(txtCASRemark5)
        End If
End Function

Private Sub funcword()
'On Error GoTo openerror
Dim X As Boolean
Dim DocSavePath As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset

Dim TemplateName As String
TemplateName = ""
    X = CreatePath(DocPath & "InternalDocuments\" & Trim(txtCaseNo))
    DocSavePath = DocPath & "InternalDocuments\" & Trim(txtCaseNo) & "\"
    
    strsqlHOS = ""
    strsqlHOS = "Select LGNumber,LGTemplateName,PSInd,PSTemplate from Tbl_LGTemplateMaster where LGNumber ='" + tmpPLNLGPrivate + "'"
    HOS.Open strsqlHOS, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    Do Until HOS.EOF
        If (Mid(frmGLMenu.cboFollowup, 1, 1) = "Y") Then
            TemplateName = HOS!LGTemplateName
        ElseIf (Mid(frmGLMenu.cboFollowup, 1, 1) = "N") Then
            TemplateName = HOS!PSTemplate
        Else
            TemplateName = HOS!LGTemplateName
        End If
    HOS.MoveNext
    Loop
    
    HOS.Close
    Set HOS = Nothing
    
    strsqlHOS = ""
    strsqlHOS = "select HosCode, HosName, HOSStatus from HOS where HosName = '" & Trim(txtHOSCode) & "'"
    HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
          Do Until HOS.EOF
            If HOS.recordCount Then
                HOSType = IIf(Len(HOS!HOSStatus), HOS!HOSStatus, "")
            Else
                HOSType = ""
            End If
          HOS.MoveNext
          Loop
        
        
If Letter = "D" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                If txtPayCode = "EP" Or txtPayCode = "EQ" Or txtPayCode = "ET" Or txtPayCode = "EC" Or txtPayCode = "ES" Or txtPayCode = "EM" Then
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\DocRequestDisETQ.dot")
                Else
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\DocRequestDis.dot")
                End If
        End If
        bookmark ("DocRequestDis.dot")
        objdoc.SaveAs (DocSavePath & "D" & Trim(txtCaseNo) & ".doc")
        sFilePath = DocSavePath & "D" & Trim(txtCaseNo) & ".doc"
        Me.Show
        Exit Sub
ElseIf Letter = "I" Then
    Set objword = Nothing
    If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\Investigation.dot")
    End If
ElseIf Letter = "S" Then
            Set objword = Nothing
            If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\DocReimReq.dot")
            End If
                bookmark ("DocReimReq")
            objdoc.SaveAs (DocSavePath & Trim(txtCaseNo) & ".doc")
            Me.Show
            X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
            objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
            objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(txtCaseNo) & ".pdf", FileFormat:=Word.wdFormatPDF
            sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(txtCaseNo) & ".pdf"
ElseIf Len(TemplateName) And frmGLMenu.cbPurpose <> "Outpatient" And HOSType <> "G" And Trim(Mid(txtPayCode, 1, 1)) <> Trim("E") Then
    Set objword = Nothing
    If objword Is Nothing Then
        Set objword = New Word.Application
        objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & TemplateName)
    End If
        bookmark (TemplateName)
    objdoc.SaveAs (DocSavePath & Trim(txtCaseNo) & ".doc")
    Me.Show
    If Len(LGNumber) Then
        X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
        objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
        objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
        sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
    Else
        X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
        objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
        objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\L" & Trim(txtCaseNo) & "-" & Trim(tmpLGVer) & ".pdf", FileFormat:=Word.wdFormatPDF
        sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\L" & Trim(txtCaseNo) & "-" & Trim(tmpLGVer) & ".pdf"
    End If
Else
    If Letter = "S" Then
            Set objword = Nothing
            If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\DocReimReq.dot")
            End If
                bookmark ("DocReimReq")
            objdoc.SaveAs (DocSavePath & Trim(txtCaseNo) & ".doc")
            Me.Show
            X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
            objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
            objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(txtCaseNo) & ".pdf", FileFormat:=Word.wdFormatPDF
            sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(txtCaseNo) & ".pdf"
    ElseIf Letter = "L" And tmpPLNPerVisit <> "Y" Then
        If Trim(HOSType) = "P" Then
        Dim IsValid As Boolean
        If (Trim(tmpPLNLGPrivate) <> "4" And Trim(tmpPLNLGPrivate) <> "5" And Trim(tmpPLNLGPrivate) <> "6" And Trim(tmpPLNLGPrivate) <> "7" And Trim(tmpPLNLGPrivate) <> "8" And Trim(tmpPLNLGPrivate) <> "26" And Trim(tmpPLNLGPrivate) <> "34") Then
            IsValid = True
        Else
            IsValid = False
        End If
        
            If Mid(frmGLMenu.cbPurpose, 1, 1) = "O" And IsValid Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    
                    If txtPayCode = "SI" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MedixOPGLSI.dot")
                    ElseIf txtPlanCode = "AC62" Or txtPlanCode = "AC63" Or txtPlanCode = "AC64" Or txtPlanCode = "AC65" Or txtPlanCode = "AC66" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\STMBOPSME.dot")
                    Else
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MedixOPGL.dot")
                    End If
                End If
            
                If bookmark("MedixOPStandardGL.dot") = True Then
                    Me.Show
                End If
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo) ''
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            
            ElseIf Trim(tmpPLNLGPrivate) = "9" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TKStandardGL.dot")
                End If
                    bookmark ("TKStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "44" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\STLHDNGL.dot")
                End If
                bookmark ("STLHDNGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "46" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\STMBIPSME.dot")
                End If
                bookmark ("STMBIPSME")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "24" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\IBMStandardGL.dot")
                End If
                bookmark ("IBMStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "5" Or Trim(tmpPLNLGPrivate) = "6" Or Trim(tmpPLNLGPrivate) = "7" Or Trim(tmpPLNLGPrivate) = "8" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    If txtPayCode = "ET" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaGroupInitialGL.dot")
                        bookmark ("EtiqaGroupInitialGL")
                    ElseIf txtPayCode = "EQ" Then
                        If tmpETQIndGL = "EtiqaIndUMRGL" Then
                            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaIndUMRGL.dot")
                            bookmark ("EtiqaIndUMRGL")
                        Else
                            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaIndMedicSaveGL.dot")
                            bookmark ("EtiqaIndMedicSaveGL")
                        End If
                    ElseIf txtPayCode = "EP" Then
                        If tmpETQIndGL = "EtiqaIndUMRGL" Then
                            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaIndUMRGL.dot")
                            bookmark ("EtiqaIndUMRGL")
                        Else
                            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaIndMedicSaveGL.dot")
                            bookmark ("EtiqaIndMedicSaveGL")
                        End If
                    ElseIf txtPayCode = "EC" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaGroupInitialGL.dot")
                        bookmark ("EtiqaGroupInitialGL")
                    ElseIf txtPayCode = "ZN" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaGroupInitialGL.dot")
                        bookmark ("EtiqaGroupInitialGL")
                    End If
                End If
                    objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                    Me.Show
                    X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                    objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                    objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                    sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "4" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    If Letter = "B" Then
                        If Mid(frmGLMenu.cbPurpose, 1, 1) = "O" Then
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankOPGL.dot")
                                bookmark ("MaybankOPGL")
                            Else
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaOPGL.dot")
                                bookmark ("EtiqaOPGL")
                            End If
                        ElseIf Mid(frmGLMenu.cbPurpose, 1, 1) = "E" Then
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankESPGL.dot")
                                bookmark ("MaybankESPGL")
                            Else
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaESPGL.dot")
                                bookmark ("EtiqaESPGL")
                            End If
                        Else
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankInitialGL.dot")
                                 bookmark ("MaybankInitialGL")
                            ElseIf txtPayCode = "ES" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaInitialGL.dot")
                                 bookmark ("EtiqaInitialGL")
                            End If
                        End If
                    Else
                        If Mid(frmGLMenu.cbPurpose, 1, 1) = "O" Then
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankOPGL.dot")
                                bookmark ("MaybankOPGL")
                            Else
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaOPGL.dot")
                                bookmark ("EtiqaOPGL")
                            End If
                        ElseIf Mid(frmGLMenu.cbPurpose, 1, 1) = "E" Then
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankESPGL.dot")
                                bookmark ("MaybankESPGL")
                            Else
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaESPGL.dot")
                                bookmark ("EtiqaESPGL")
                            End If
                        Else
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankInitialGL.dot")
                                 bookmark ("MaybankInitialGL")
                            ElseIf txtPayCode = "ES" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaInitialGL.dot")
                                 bookmark ("EtiqaInitialGL")
                            End If
                        End If
                    End If
                End If
                   objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                    Me.Show
                    X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                    objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                    objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                    sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "34" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    If Letter = "B" Then
                        If Mid(frmGLMenu.cbPurpose, 1, 1) = "O" Then
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankOPGLNew.dot")
                                bookmark ("MaybankOPGL")
                            Else
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaOPGL.dot")
                                bookmark ("EtiqaOPGL")
                            End If
                        ElseIf Mid(frmGLMenu.cbPurpose, 1, 1) = "E" Then
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankESPGL.dot")
                                bookmark ("MaybankESPGL")
                            Else
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaESPGL.dot")
                                bookmark ("EtiqaESPGL")
                            End If
                        Else
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankInitialGL.dot")
                                 bookmark ("MaybankInitialGL")
                                 
                            ElseIf txtPayCode = "ES" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaInitialGL.dot")
                                 bookmark ("EtiqaInitialGL")
                            End If
                        End If
                    Else
                        If Mid(frmGLMenu.cbPurpose, 1, 1) = "O" Then
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankOPGLNew.dot")
                                bookmark ("MaybankOPGL")
                            Else
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaOPGL.dot")
                                bookmark ("EtiqaOPGL")
                            End If
                        ElseIf Mid(frmGLMenu.cbPurpose, 1, 1) = "E" Then
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankESPGL.dot")
                                bookmark ("MaybankESPGL")
                            Else
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaESPGL.dot")
                                bookmark ("EtiqaESPGL")
                            End If
                        Else
                            If txtPayCode = "EM" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MaybankInitialGL.dot")
                                 bookmark ("MaybankInitialGL")
                                 
                            ElseIf txtPayCode = "ES" Then
                                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\EtiqaInitialGL.dot")
                                 bookmark ("EtiqaInitialGL")
                            End If
                        End If
                    End If
                End If
                    objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                    Me.Show
                    X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                    objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                    objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                    sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "3" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    If Letter = "B" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\GuaranteeLetterAXABTBG.dot")
                    Else
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\GuaranteeLetterAXA.dot")
                    End If
                End If
                bookmark ("GuaranteeLetterAXA")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "2" Then
            ElseIf Trim(tmpPLNLGPrivate) = "10" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\XLStandardGL.DOT")
                End If
                bookmark ("XLStandardGL.DOT")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                Me.Show
            ElseIf Trim(tmpPLNLGPrivate) = "30" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\AYStandardGL_2015.DOT")
                End If
                bookmark ("AYStandardGL.DOT")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                Me.Show
            ElseIf Trim(tmpPLNLGPrivate) = "37" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\AYStandardGL_2015NoCoPayment.DOT")
                End If
                bookmark ("AYStandardGL_2015NoCoPayment.DOT")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                Me.Show
            ElseIf Trim(tmpPLNLGPrivate) = "11" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\XLStandardGL210.DOT")
                End If
                bookmark ("XLStandardGL210.DOT")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                Me.Show
            ElseIf Trim(tmpPLNLGPrivate) = "12" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TAStandardGL.dot")
                End If
                bookmark ("TAStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                Me.Show
            ElseIf Trim(tmpPLNLGPrivate) = "28" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TAStandardGLNoCoPmt.dot")
                End If
                bookmark ("TAStandardGLNoCoPmt")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "13" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\BSNStandardGL.dot")
                End If
                bookmark ("BSNStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "14" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\PKStandardGL.dot")
                End If
                bookmark ("PKStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "15" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\XLStandardGL207.dot")
                End If
                bookmark ("XLStandardGL207")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "16" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\XFStandardGL.dot")
                End If
                bookmark ("XFStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "17" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\WMBTBGGL.dot")
                End If
                bookmark ("WMBTBGGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "18" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\WKStandardGL.dot")
                End If
                bookmark ("WKStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "20" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\AMIWKGL.dot")
                End If
                bookmark ("AMIWKGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "21" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\AMGoodYearGL.dot")
                End If
                bookmark ("AMGoodYearGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "22" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\AMStandardGL.dot")
                End If
                bookmark ("AMStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "23" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MMStandardGL.dot")
                End If
                bookmark ("MMStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "25" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TKStandardGLNoPost.dot")
                End If
                bookmark ("TKStandardGLNoPost")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "31" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\SIStandardGL.dot")
                End If
                bookmark ("SIStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "32" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MXStandardGL.dot")
                End If
                bookmark ("MXStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "27" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\XIStandardGL.dot")
                End If
                bookmark ("XIStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "29" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\AMPNSBGL.dot")
                End If
                    bookmark ("AMPNSBGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "33" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TKStandardGLCoIns.dot")
                End If
                bookmark ("TKStandardGLCoIns")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "34" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\STVIP.dot")
                End If
                bookmark ("STVIP")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "35" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\STMBIPVIP.dot")
                End If
                bookmark ("STMBIPVIP")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "36" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TKStandardGL90Days.dot")
                End If
                bookmark ("TKStandardGL90Days")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "38" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\SIStandardGLCoPayment.dot")
                End If
                bookmark ("SIStandardGLCoPayment")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "39" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\SIStandardGLAdm.dot")
                End If
                bookmark ("SIStandardGLAdm")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "40" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\SIStandardGL90.dot")
                End If
                bookmark ("SIStandardGL90")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "41" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\SIStandardGLNoFL.dot")
                End If
                bookmark ("SIStandardGLNoFL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "42" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MZStandardGL.dot")
                End If
                bookmark ("MZStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "47" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\KSStandardGL.dot")
                End If
                bookmark ("KSStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "48" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TEstandardGL.dot")
                End If
                bookmark ("TEstandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "49" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\AMBTBGL.dot")
                End If
                bookmark ("AMBTBGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
             ElseIf Trim(tmpPLNLGPrivate) = "50" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\SHStandardGL.dot")
                End If
                bookmark ("SHStandardGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
             ElseIf Trim(tmpPLNLGPrivate) = "51" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\WKStandardGLNoFL.dot")
                End If
                bookmark ("WKStandardGLNoFL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
             ElseIf Trim(tmpPLNLGPrivate) = "52" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\AMIWKGLNOFL.dot")
                End If
                bookmark ("AMIWKGLNOFL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
             ElseIf Trim(tmpPLNLGPrivate) = "53" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TKSBTBGGL.dot")
                End If
                bookmark ("TKSBTBGGL")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            Else
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\New LG-2.DOT")
                End If
                bookmark ("New LG-2.DOT")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                Me.Show
            End If
            
        ElseIf Trim(HOSType) = "G" Then
            If tmpPLNLGGov = "2" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\GLGovernmentAXA.dot")
                End If
                bookmark ("GLGovernmentAXA")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf tmpPLNLGGov = "1" Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\GLGovernmentMAS.DOT")
                End If
                bookmark ("GLGovernmentMAS")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            Else
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\GLGovernmentMedix.DOT")
                End If
                bookmark ("LGGovernmentMedix")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            End If
        Else
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\New LG-2.DOT")
                End If
                If bookmark("New LG-2.DOT") = True Then
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                End If
        End If
        Exit Sub
    ElseIf Letter = "L" And tmpPLNPerVisit = "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\LGVisit.DOT")
        End If
        If bookmark("LGVisit.DOT") = True Then
            objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
            Me.Show
            X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
            objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
            objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
            sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
        End If
        Exit Sub
        
     ElseIf Letter = "B" And tmpPLNPerVisit <> "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            If Trim(HOSType) = "P" Then
              If Mid(frmGLMenu.cbPurpose, 1, 1) = "O" And (Trim(tmpPLNLGPrivate) <> "5" Or Trim(tmpPLNLGPrivate) <> "6" Or Trim(tmpPLNLGPrivate) <> "7" Or Trim(tmpPLNLGPrivate) Or "8") Then
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\MedixOPStandardGL.dot")
                    If bookmark("MedixOPStandardGL.dot") = True Then
                       Me.Show
                    End If
                    objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                    Me.Show
                    X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                    objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                    objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                    sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                ElseIf Trim(tmpPLNLGPrivate) = "3" Then
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\GuaranteeLetterAXABTBG.dot")
                    bookmark ("GuaranteeLetterAXABTBG")
                ElseIf Trim(tmpPLNLGPrivate) = "19" Then
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TKSBTBGGL.dot")
                    bookmark ("TKSBTBGGL")
                ElseIf Trim(tmpPLNLGPrivate) = "26" Then
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TKFeldaGL.dot")
                    bookmark ("TKFeldaGL")
                ElseIf Trim(tmpPLNLGPrivate) = "17" Then
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\WMBTBGGL.dot")
                    bookmark ("WMBTBGGL")
                ElseIf Trim(tmpPLNLGPrivate) = "35" Then
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\STMBIPVIP.dot")
                    bookmark ("STMBIPVIP")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            Else
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\BTBG.DOT")
                    bookmark ("BTBG")
                End If
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            ElseIf Trim(tmpPLNLGPrivate) = "36" Then
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\TKStandardGL90Days.dot")
                bookmark ("TKStandardGL90Days")
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            Else
                If tmpPLNLGGov = "2" Then
                    Set objword = Nothing
                    If objword Is Nothing Then
                        Set objword = New Word.Application
                        objword.Visible = True
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\GLGovernmentAXA.dot")
                    End If
                    bookmark ("GLGovernmentAXA")
                    objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                ElseIf tmpPLNLGGov = "1" Then
                    Set objword = Nothing
                    If objword Is Nothing Then
                        Set objword = New Word.Application
                        objword.Visible = True
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\GLGovernmentMAS.DOT")
                    End If
                    bookmark ("GLGovernmentMAS")
                    objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Else
                    Set objword = Nothing
                    If objword Is Nothing Then
                        Set objword = New Word.Application
                        objword.Visible = True
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\GLGovernmentMedix.DOT")
                    End If
                    bookmark ("LGGovernmentMedix")
                End If
                objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                Me.Show
                X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
            End If
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
              vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              objdoc.Save
              objdoc.Close
            ElseIf vbCancel Then
              Exit Sub
            End If
      
            If Trim(tmpPLNLGPrivate) = "3" Then
                If bookmark("GuaranteeLetterAXABTBG.dot") = True Then
                    objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                    Me.Show
                    X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                    objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                    objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                    sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                End If
            ElseIf Trim(tmpPLNLGPrivate) = "9" Then
                If bookmark("GLTK.dot") = True Then
                    objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                    Me.Show
                    X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                    objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                    objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                    sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                End If
            Else
                If bookmark("BTBG.DOT") = True Then
                    objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
                    Me.Show
                    X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
                    objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
                    objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
                    sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
                End If
            End If
        End If
        Exit Sub
        
     ElseIf Letter = "B" And tmpPLNPerVisit = "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\BTBGVisit.DOT")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
              vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              objdoc.Save
              objdoc.Close
            ElseIf vbCancel Then
              Exit Sub
            End If
        End If
        
        If bookmark("BTBGVisit.DOT") = True Then
            objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
            Me.Show
            X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
            objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
            objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
            sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
        End If
        Exit Sub
        
    ElseIf Letter = "C" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\New LG-2.DOT")
        End If
        
        If bookmark("New LG-2.DOT") = True Then
            objdoc.SaveAs (DocSavePath & Trim(LGNumber) & ".doc")
            Me.Show
            X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
            objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
            objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
            sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & Trim(LGNumber) & ".pdf"
        End If
        Exit Sub
    
    ElseIf Letter = "R" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                    Set objdoc = Nothing
                    If Trim(tmpRepStatus) = "R" And Mid(CboPayType, 1, 1) = "C" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetterMAA-XR.DOT")
                    ElseIf Trim(tmpRepStatus) = "Y" And Mid(CboPayType, 1, 1) = "C" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetterMAA.DOT")
                    ElseIf Trim(tmpRepStatus) = "Y" And Mid(CboPayType, 1, 1) = "R" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPMAAFormat.DOT")
                    ElseIf Trim(tmpRepStatus) = "R" And Mid(CboPayType, 1, 1) = "R" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPMAAFormat.DOT")
                    Else
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetter.DOT")
                    End If
                'End If
        ElseIf Not objdoc Is Nothing Then
              If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
                vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              ElseIf vbCancel Then
                Exit Sub
              End If
                Set objdoc = Nothing
                    If Trim(tmpRepStatus) = "R" And Mid(CboPayType, 1, 1) = "C" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetterMAA-XR.DOT")
                    ElseIf Trim(tmpRepStatus) = "Y" And Mid(CboPayType, 1, 1) = "C" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetterMAA.DOT")
                    ElseIf Trim(tmpRepStatus) = "Y" And Mid(CboPayType, 1, 1) = "R" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPMAAFormat.DOT")
                    ElseIf Trim(tmpRepStatus) = "R" And Mid(CboPayType, 1, 1) = "R" Then
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPMAAFormat.DOT")
                    Else
                        Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\REPLetter.DOT")
                    End If
        End If
        bookmark ("REPLetter.DOT")
        objdoc.SaveAs (DocSavePath & "R" & Trim(txtCaseNo) & ".doc")
         X = CreatePath(DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo)
        objword.WindowState = Word.WdWindowState.wdWindowStateMinimize
        objword.ActiveDocument.SaveAs filename:=DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & "R" & Trim(LGNumber) & ".pdf", FileFormat:=Word.wdFormatPDF
        sFilePath = DocPath & "ScannedDoc\" & "FinalLgs\" & txtCaseNo & "\" & "R" & Trim(LGNumber) & ".pdf"
        Me.Show
    ElseIf Letter = "D" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                If txtPayCode = "EP" Or txtPayCode = "EQ" Or txtPayCode = "ET" Or txtPayCode = "EC" Or txtPayCode = "ES" Or txtPayCode = "EM" Then
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\DocRequestDisETQ.dot")
                Else
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\DocRequestDis.dot")
                End If
        End If
        bookmark ("DocRequestDis.dot")
        objdoc.SaveAs (DocSavePath & "D" & Trim(txtCaseNo) & ".doc")
      sFilePath = DocSavePath & "D" & Trim(txtCaseNo) & ".doc"
        Me.Show
        Exit Sub
    ElseIf Letter = "I" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\Investigation.dot")
        End If
End If
             
    If bookmark("Investigation.dot") = True Then
            Me.Show
            objdoc.SaveAs (DocSavePath & "I" & Trim(txtCaseNo) & ".doc")
        End If
        Exit Sub
    End If
    Exit Sub
End Sub
Private Function FuncWordTakeHG()
Dim DocSavePath As String
    If Letter = "T" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\Take Home Drugs.doc\")
        End If
        If bookmark("Take Home Drugs.doc") = True Then
            objdoc.SaveAs (DocSavePath & "T" & Trim(txtCaseNo) & ".doc")
            Me.Show
        End If
    End If
End Function

Private Function FuncReqAdm()
Dim DocSavePath As String
    If Letter = "A" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\DocRequestAdm.dot")
        End If
        If bookmark("DocRequestAdm.dot") = True Then
            Me.Show
        End If
    End If
End Function
Private Function FuncPriorLetter()
Dim DocSavePath As String
    If Letter = "P" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\PriorLetter.DOT")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
              vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              objdoc.Save
              objdoc.Close
            ElseIf vbCancel Then
            End If
        End If
        If bookmark("PriorLetterNew.dot") = True Then
            Me.Show
            objdoc.SaveAs (DocSavePath & "P" & Trim(txtCaseNo) & ".doc")
        End If
    End If
End Function


Private Function FuncWordAuditreview()
    If Letter = "M" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "Case\CaseAudit.DOT")
        End If
        
        If Trim(txtCaseNo) <> "" Then
            objdoc.Bookmarks("CASE").Range.Text = Trim(txtCaseNo)
        End If
        
        If Trim(txtPrincipalExc) <> "" Then
          objdoc.Bookmarks("Note").Range.Text = Trim(txtPrincipalExc)
        End If
        
        Me.Show
        
    End If
End Function
Private Function FuncWordRemark()
    If Letter = "M" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "Case\CaseRemark.DOT")
        End If
        If BookMarkNote("CaseRemark.dot") = True Then
            Me.Show
        End If
    End If
End Function

Public Function CreatePath(NewPath) As Boolean
    If Right(NewPath, 1) <> "\" Then
        NewPath = NewPath & "\"
    End If
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        CreatePath = True
    End If
End Function


Private Sub showMBMListing()
    Dim MBM As New ADODB.Recordset
    Dim sqlstr As String
    Dim MBMListing As ListItem
        sqlstr = "Exec SearchMBMCas '" & Trim(txtMBMNumber) & "'"
        Set MBM = New ADODB.Recordset
        MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
        lstMember.ListItems.Clear
        If Not MBM.EOF Then
            With MBM
                Set MBMListing = lstMember.ListItems.Add(, , "00")
                If Len(Trim(!MBMName)) <> 0 Then
                    MBMListing.SubItems(1) = ChkStr(!MBMName)
                End If
                If Len(Trim(!MBMIcBcPp)) Then
                    MBMListing.SubItems(2) = !MBMIcBcPp
                End If
                If Len(Trim(!MBMMemberType)) Then
                    MBMListing.SubItems(3) = !MBMMemberType
                End If
                If Len(Trim(!MBMDOB)) Then
                    MBMListing.SubItems(4) = !MBMDOB
                End If
            End With
        End If
    MBM.Close
    Set MBM = Nothing
End Sub

' * Start
Private Sub showMBMCovListing()
    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "Exec SearchMBMCas '" & Trim(txtMBMNumber) & "'"
    Set MBMCov = New ADODB.Recordset
    MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic

    If Len(MBMCov!MBMCCoverID) Then
        Dim MBMListing As ListItem
        Do While Not MBMCov.EOF
            Set MBMListing = lstMember.ListItems.Add(, , MBMCov!MBMCCoverID)
            With MBMCov
                If Len(Trim(!MBMCName)) <> 0 Then
                    MBMListing.SubItems(1) = ChkStr(!MBMCName)
                End If
                If Len(Trim(!MBMCICBCPPNo)) <> 0 Then
                    MBMListing.SubItems(2) = !MBMCICBCPPNo
                End If
                If Len(Trim(!MBMCRELCode)) Then
                    MBMListing.SubItems(3) = !MBMCRELCode
                End If
                
                If Len(Trim(!MBMCDOB)) Then
                    MBMListing.SubItems(4) = !MBMCDOB
                End If

            End With
            MBMCov.MoveNext
        Loop
    End If
    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub cboSuppExc_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboSuppExc, cboSuppExc.SelStart) + Chr(KeyAscii) + Mid(cboSuppExc, cboSuppExc.SelStart + cboSuppExc.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboSuppExc.ListCount - 1
        
        If NewText = LCase(Left(cboSuppExc.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboSuppExc.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboSuppExc = ValidValue
               cboSuppExc.SelStart = 0
               cboSuppExc.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub UpdateLG()
Dim strsqlRMK As String
Dim LG As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
    strAppend = ""
    strAppendCase = "4"
    strAppend = " AND CASENo = '" & Trim(LGNo) & "' AND LGNo = '" & Trim(LGVer) & "'"
    Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "Case_Adjustment"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        LG.CursorLocation = adUseClient
        LG.CursorType = adOpenForwardOnly
        LG.CacheSize = 100
        Set LG = cmd.Execute
        If LG.EOF = False Then
                LGRemark = Empty
                strsqlRMK = "EXEC Case_AdjUpdateLG '" & Trim(LGNo) & "', '" & Trim(LGVer) & "', '" & LGRemark & "'"
            medixmasterconnMaint.Execute strsqlRMK
        End If
    LG.Close
    Set LG = Nothing
End Sub

Private Sub ExportFeeDetails()
On Error Resume Next
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
    Screen.MousePointer = vbHourglass
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    On Error GoTo HANDLE_ERROR
    objExcel.Interactive = False
    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    objExcel.WindowState = xlMaximized
    Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Case\Fee Details.xlt")
    Set objWorksheet = objWorkbook.Sheets(1)
    'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    objWorksheet.Cells(5, "B") = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
    objWorksheet.Cells(6, "B") = Trim(txtMBMNumber)
    
    Screen.MousePointer = vbDefault
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Exit Sub
ExitFunction:
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Screen.MousePointer = vbDefault
    objExcel.Quit
    Exit Sub
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    GoTo ExitFunction
End Sub

Private Sub UpdateClaimability()
Dim strsql As String
Dim ACPeriod As New ADODB.Recordset
    cboClaimability.Text = "NN - Not Decided"
    strsql = " update CAS set " & _
             " CASClaimability = '" & Mid(cboClaimability, 1, 2) & "' " & _
             " Where CASNumber = '" & Trim(txtCaseNo) & "'"
    medixmasterconn.Execute strsql
End Sub
Private Function SearchMMARate(OPCSCode As String)
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
    MMASurgFees = 0
    MMAAnaeFees = 0
    MMASurgFees13th = 0
    MMAAnaeFees13th = 0
    MMASurgFees13thMin = 0
    strsqlMMA = "Select SurgeonsFees, AnaesthetistsFees,Surgeon13th,Anesthetist13th,Surgeon13thMin  from VIEW_MMA_Schedule where  OPCSCode = '" & Trim(OPCSCode) & "'"
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        Do While Not MMA.EOF
            MMASurgFees = IIf(Len(MMA!SurgeonsFees), MMA!SurgeonsFees, 0)
            MMAAnaeFees = IIf(Len(MMA!AnaesthetistsFees), MMA!AnaesthetistsFees, 0)
            MMASurgFees13th = IIf(Len(MMA!Surgeon13th), MMA!Surgeon13th, 0)
            MMAAnaeFees13th = IIf(Len(MMA!Anesthetist13th), MMA!Anesthetist13th, 0)
            MMASurgFees13thMin = IIf(Len(MMA!Surgeon13thMin), MMA!Surgeon13thMin, 0)
        MMA.MoveNext
        Loop
    MMA.Close
    Set MMA = Nothing
End Function

Private Function Search14thRate(OPCSCode As String)
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
    MMASurgFees = 0
    MMAAnaeFees = 0
    MMASurgFees13th = 0
    MMAAnaeFees13th = 0
    MMASurgFees13thMin = 0
    strsqlMMA = "Select  Surgeon13th,Anesthetist13th,Surgeon13thMin  from MMA14Schedule where  OPCSCode = '" & Trim(OPCSCode) & "'"
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        Do While Not MMA.EOF
            MMASurgFees13th = IIf(Len(MMA!Surgeon13th), MMA!Surgeon13th, 0)
            MMAAnaeFees13th = IIf(Len(MMA!Anesthetist13th), MMA!Anesthetist13th, 0)
            MMASurgFees13thMin = IIf(Len(MMA!Surgeon13thMin), MMA!Surgeon13thMin, 0)
        MMA.MoveNext
        Loop
    MMA.Close
    Set MMA = Nothing
End Function
Private Sub SearchMMARateCal(CPTCode As String)
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
    MMASurgFees = 0
    MMAAnaeFees = 0
    MMASurgFees13th = 0
    MMAAnaeFees13th = 0
    MMASurgFees13thMin = 0
    strsqlMMA = "Select SurgeonsFees, AnaesthetistsFees,Surgeon13th,Anesthetist13th,Surgeon13thMin  from VIEW_MMA_Schedule where  OPCSCode = '" & Trim(CPTCode) & "'"
    MMA.Open strsqlMMA, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        Do While Not MMA.EOF
            MMASurgFees = IIf(Len(MMA!SurgeonsFees), MMA!SurgeonsFees, 0)
            MMAAnaeFees = IIf(Len(MMA!AnaesthetistsFees), MMA!AnaesthetistsFees, 0)
            MMASurgFees13th = IIf(Len(MMA!Surgeon13th), MMA!Surgeon13th, 0)
            MMAAnaeFees13th = IIf(Len(MMA!Anesthetist13th), MMA!Anesthetist13th, 0)
            MMASurgFees13thMin = IIf(Len(MMA!Surgeon13thMin), MMA!Surgeon13thMin, 0)
        MMA.MoveNext
        Loop
    MMA.Close
    Set MMA = Nothing
 End Sub

Private Sub WSListing()
Dim ws As New ADODB.Recordset
Dim WSList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim actb4gst As Double
Dim GST As Double
Dim actualWtGST As Double
actb4gst = 0
GST = 0
actualWtGST = 0
    strAppend = ""
    If lstWorksheetListing.ListItems.Count = 0 Then
        strAppendCase = "2"
        strAppend = " AND CASNumber ='" & Trim(txtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        ws.CursorLocation = adUseClient
        ws.CursorType = adOpenForwardOnly
        ws.CacheSize = 100
        Set ws = cmd.Execute
        If Not ws.EOF Then
        Do Until ws.EOF
            With ws
                Set WSList = lstWorksheetListing.ListItems.Add(, , !CasNumber & "-" & !claimversion)
                    If Len(!CLMTotalActual) Then
                        WSList.SubItems(1) = Format(!CLMTotalActual, "#,##0.00")
                        actb4gst = !CLMTotalActual
                    End If
                      If Len(!TotGST) Then
                        WSList.SubItems(2) = !TotGST
                        GST = !TotGST
                    End If
                    actualWtGST = actb4gst + GST
                     If Len(actualWtGST) Then
                        WSList.SubItems(3) = Format(actualWtGST, "#,##0.00")
                    End If
                    If Len(!CLMTotalApproved) Then
                        WSList.SubItems(4) = Format(!CLMTotalApproved, "#,##0.00")
                    End If
                    If Len(!CLMPayableByMedix2H) Then
                        WSList.SubItems(5) = Format(!CLMPayableByMedix2H, "#,##0.00")
                    End If
                    If Len(!CLMPayableByMedix2M) Then
                        WSList.SubItems(6) = Format(!CLMPayableByMedix2M, "#,##0.00")
                    End If
                    If Len(!BordxRefNo) Then
                        WSList.SubItems(7) = !BordxRefNo
                    End If
                    If Len(!BordxDate) Then
                        WSList.SubItems(8) = !BordxDate
                    End If
                    If Len(!CLMBordxAmt) Then
                        WSList.SubItems(9) = Format(!CLMBordxAmt, "#,##0.00")
                    End If
                    If Len(!CLMPayorAppStatus) Then
                        WSList.SubItems(10) = !CLMPayorAppStatus
                    End If
                    If Len(!CLMPayorAppDate) Then
                        WSList.SubItems(11) = !CLMPayorAppDate
                    End If
                    If Len(!CLMPayorAppCode) Then
                        WSList.SubItems(12) = !CLMPayorAppCode
                    End If
                    If Len(!CLMBordxRevInd) Then
                        WSList.SubItems(13) = !CLMBordxRevInd
                    End If
                    If Len(!CRTCartonNo) Then
                        WSList.SubItems(14) = !CRTCartonNo
                    End If
                    If Len(!DoSend) Then
                        WSList.SubItems(15) = Format(!DoSend, "dd/mm/yyyy")
                    End If
                    If Len(!DebitCreditRefNo) Then
                        WSList.SubItems(16) = Trim(!DebitCreditRefNo)
                    End If
                    If Len(!DebitCreditDate) Then
                        WSList.SubItems(17) = Format(!DebitCreditDate, "dd/mm/yyyy")
                    End If
                    If Len(!CLMRevBordxStatus) Then
                        WSList.SubItems(18) = Trim(!CLMRevBordxStatus)
                    End If
                .MoveNext
                End With
        Loop
        ws.Close
        Set ws = Nothing
    End If
End If
End Sub
Private Sub CheckLimit()
    Dim StrSqlAnnLmt As String
    Dim AnnLmt As New ADODB.Recordset
    Dim strsqlMBMOthers As String
    Dim MBMOthers As New ADODB.Recordset
    Dim strsqlMBM As String
    Dim MBM As New ADODB.Recordset
    Dim srtsqlMBMCref As String
    Dim MBMCRef As New ADODB.Recordset
    Dim PLNSOF  As Boolean
    Dim strsqlmbmlmt As String
    Dim MBMLmt As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim strAppendCase As String
    Dim tmpTotalBordx As Double
    Dim tmpTotalLTBordx As Double
    Dim strsqlCLMBal As String
    Dim tmpMBMAnnLmt As Double
    Dim tmpMBMSuppAnnLmt As Double
    Dim tmpMBMLTAnnLmt As Double
    Dim CLMBal As New ADODB.Recordset
    Dim CAS As New ADODB.Recordset
    Dim AnnInd As Boolean
    Dim LifeTimeStatus As Boolean
    Dim Topup As Boolean
    Dim MBMTopupNumber As String
    Dim TopPlanCode As String
    Dim TopINSCode As String
    Dim TopEffDate As String
    Dim TopNewMBMNumber As String
    Dim TopNewPolNumber As String
    Dim TopNewBalLmt As Double
    Dim LTGuaRenewal As Boolean
    Dim CASTopup As Boolean
    Dim PolEffDate As String
    Dim PolExpDate As String
    Dim tmpMBMNumber As String
    Dim PLNCode As String
    Dim PatientPLN As String
    Dim tmpINSType As String
    Dim TmpMBMType As String
    Dim tmpPatientID As String
    Dim LifetimeAnnBal As Double
    Dim tmpTotalBalLmt As Double
    Dim OldBalLmt As Double
    Dim tmpMBMTopupInd As Boolean
    LifetimeAnnBal = 0
    tmpTotalBalLmt = 0
    
        strAppend = ""
        strAppendCase = "3"
        strAppend = strAppend & " CAS.CASNumber = '" & Trim(txtCaseNo) & "'"
        
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "SearchCasMBMClaim"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
        cmd.Parameters.Append Prm2
        CAS.CursorLocation = adUseClient
        CAS.CursorType = adOpenForwardOnly
        CAS.CacheSize = 100
        Set CAS = cmd.Execute

            If Not CAS.EOF Then
                If CAS!AnnLmtInd = "N" Then
                    AnnInd = False
                Else
                    AnnInd = True
                End If
                If CAS!PLNSOF = "Y" Then
                    PLNSOF = True
                Else
                    PLNSOF = False
                End If
                If CAS!PLNLifeTimeStatus = "Y" Then
                    LifeTimeStatus = True
                    Topup = False
                Else
                    LifeTimeStatus = False
                End If
                If Trim(CAS!MBMTopupInd) = "Y" Then
                    tmpMBMTopupInd = True
                Else
                    tmpMBMTopupInd = False
                End If
                If Len(CAS!MBMTopupNumber) Then
                    srtsqlMBMCref = "Select MBMNumber, INSCode,PLNCode,MBMPayorEffDate,MBMPayorExpDate,MBMPolicyNo,MBMTopupNumber from MBMCrossReference Where MBMTopupNumber = '" & Trim(CAS!MBMTopupNumber) & "'"
                    MBMCRef.Open srtsqlMBMCref, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                    If MBMCRef.recordCount Then
                        MBMTopupNumber = MBMCRef!MBMTopupNumber
                        MBMCRef.Close
                        Set MBMCRef = Nothing
                            strsqlMBM = "Select MBMNumber, INSCode,PLNCode, MBMPayorEffDate,MBMPayorExpDate, MBMPolicyNo, MBMAvailableLimit,MBMTopupNumber from MBM Where MBMNumber = '" & Trim(MBMTopupNumber) & "'"
                            MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
                                If MBM.recordCount Then
                                    TopPlanCode = MBM!PLNCode
                                    TopINSCode = MBM!INSCode
                                    TopEffDate = MBM!MBMPayorEffDate
                                    TopNewMBMNumber = MBM!MBMNumber
                                    TopNewPolNumber = MBM!MBMPolicyNo
                                    TopNewBalLmt = MBM!MBMAvailableLimit
                                    MBM.Close
                                    Set MBM = Nothing
                                End If
                    End If
                        
                        strsqlMBMOthers = "Select MBMNumber,MBMCovID from MBMLifeTime Where MBMNumber = '" & Trim(MBMTopupNumber) & "' And MBMCovID = '" & Trim(CAS!MBMPatientCovID) & "' "
                        MBMOthers.Open strsqlMBMOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If MBMOthers.recordCount Then
                            LTGuaRenewal = True
                        Else
                            LTGuaRenewal = False
                        End If
                        MBMOthers.Close
                        Set MBMOthers = Nothing
                 End If
                If Trim(CAS!CASTopupLmt) = "Y" Then
                    CASTopup = True
                Else
                    CASTopup = False
                End If
                    PolEffDate = CAS!MBMPayorEffDate
                    If Len(CAS!PLNCode) Then
                        PLNCode = CAS!PLNCode
                    End If
                    If Len(CAS!MBMPayorEXPDate) Then
                        PolExpDate = CAS!MBMPayorEXPDate
                    End If
                    If Len(CAS!MBMNumber) Then
                        tmpMBMNumber = CAS!MBMNumber
                    End If
                    tmpPatientID = CAS!MBMPatientCovID
                    If CAS!MBMPatientCovID <> "00" And Len(Trim(MBMCPLNCode)) Then
                        PatientPLN = MBMCPLNCode
                    Else
                        PatientPLN = CAS!PLNCode
                    End If
                    
                    If CAS!INSCode <> "" Then
                        tmpINSType = CAS!INSCode
                    End If
                    
                    If Len(CAS!MBMMemberType) Then
                        TmpMBMType = CAS!MBMMemberType
                    End If
                    
                If AnnInd = True Then
                
                    StrSqlAnnLmt = "Select DisabilityLmt, SuppLimitStatus,AnnLimit,SuppLimit,LifeTimeLimit, PLNCode from AnnualLimit where PLNCode = '" & Trim(PLNCode) & "' And INSCode = '" & Trim(tmpINSType) & "' And AnnualEffDate <= '" & Format(CAS!CASDateAdmitted, "mm/dd/yyyy") & "' order by AnnualEffDate  desc"
                    AnnLmt.Open StrSqlAnnLmt, medixmasterconn, adOpenKeyset, adLockOptimistic
                    
                    tmpMBMAnnLmt = IIf(IsNumeric(AnnLmt!AnnLimit), AnnLmt!AnnLimit, 0)
                    tmpMBMSuppAnnLmt = IIf(IsNumeric(AnnLmt!SuppLimit), AnnLmt!SuppLimit, 0)
                    tmpMBMLTAnnLmt = IIf(IsNumeric(AnnLmt!LifeTimeLimit), AnnLmt!LifeTimeLimit, 0)
                    
                    If AnnLmt!SuppLimitStatus = "D" And tmpPatientID = "00" Then
                        AnnInd = False
                    End If
                    
                    If LifeTimeStatus = False Then
                        If AnnLmt!SuppLimitStatus = "A" Then
                            strsqlCLMBal = "Select CLMBordxAmt  from View_Claim_GetLimit_TypeA " & _
                            " where MBMNUmber   ='" & Trim(tmpMBMNumber) & "'"
                        ElseIf AnnLmt!SuppLimitStatus = "B" Then
                            If Trim(tmpPatientID) = "00" Then
                                strsqlCLMBal = "Select CLMBordxAmt  from View_Claim_GetPrinLimit " & _
                                " where MBMNUmber   ='" & Trim(tmpMBMNumber) & "'"
                            Else
                                strsqlCLMBal = "Select CLMBordxAmt  from View_Claim_GetSuppLimit_TypeB " & _
                                " where MBMNUmber   ='" & Trim(tmpMBMNumber) & "'"
                            End If
                        ElseIf AnnLmt!SuppLimitStatus = "C" Then
                                If Trim(tmpPatientID) = "00" Then
                                strsqlCLMBal = "Select CLMBordxAmt  from View_Claim_GetPrinLimit " & _
                                " where MBMNUmber   ='" & Trim(tmpMBMNumber) & "'"
                            Else
                                strsqlCLMBal = "Select CLMBordxAmt  from View_Claim_GetSuppLimit_TypeC " & _
                                " where MBMNUmber   ='" & Trim(tmpMBMNumber) & "' And MBMPatientCovID   = '" & Trim(tmpPatientID) & "'"
                            End If
                            
                        ElseIf AnnLmt!SuppLimitStatus = "D" Then
                            If Trim(tmpPatientID) = "00" Then
                                strsqlCLMBal = "Select CLMBordxAmt  from View_Claim_GetPrinLimit " & _
                                " where MBMNUmber   ='" & Trim(tmpMBMNumber) & "'"
                            Else
                                strsqlCLMBal = "Select CLMBordxAmt  from View_Claim_GetSuppLimit_TypeC " & _
                                " where MBMNUmber   ='" & Trim(tmpMBMNumber) & "'"
                            End If
                            
                        End If
            
                        CLMBal.Open strsqlCLMBal, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                        
                        Do While Not CLMBal.EOF
                            tmpTotalBordx = IIf(IsNumeric(CLMBal!CLMBordxAmt), CLMBal!CLMBordxAmt, 0)
                            CLMBal.MoveNext
                        Loop
                        CLMBal.Close
                        Set CLMBal = Nothing
                    End If
                    
                    If LifeTimeStatus = True Then
                        'If LTGuaRenewal = True Then
                            strsqlCLMBal = ""
                            strsqlCLMBal = "Select CLMBordxAmt  from View_Claim_GetLifetime_Limit " & _
                            " where MBMTopupNumber   ='" & Trim(TopNewMBMNumber) & "' And MBMPatientCovID  = '" & Trim(tmpPatientID) & "'"
                            CLMBal.Open strsqlCLMBal, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                        
                            Do While Not CLMBal.EOF
                                tmpTotalLTBordx = IIf(IsNumeric(CLMBal!CLMBordxAmt), CLMBal!CLMBordxAmt, 0)
                                CLMBal.MoveNext
                            Loop
                            CLMBal.Close
                            Set CLMBal = Nothing
                        'Else
                            strsqlCLMBal = ""
                            If Trim(tmpPatientID) = "00" Then
                                strsqlCLMBal = "Select CLMBordxAmt  from View_Claim_GetPrinLimit " & _
                                " where MBMNUmber   ='" & Trim(tmpMBMNumber) & "'"
                            Else
                                strsqlCLMBal = "Select CLMBordxAmt  from View_Claim_GetSuppLimit_TypeC " & _
                                " where MBMNUmber   ='" & Trim(tmpMBMNumber) & "' And MBMPatientCovID   = '" & Trim(tmpPatientID) & "'"
                            End If
                            
                            CLMBal.Open strsqlCLMBal, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                        
                            Do While Not CLMBal.EOF
                                tmpTotalBordx = IIf(IsNumeric(CLMBal!CLMBordxAmt), CLMBal!CLMBordxAmt, 0)
                                CLMBal.MoveNext
                            Loop
                            CLMBal.Close
                            Set CLMBal = Nothing
                        'End If
                    End If
                                                                                        
                    If ChkStr(TmpMBMType) = "Q" And AnnLmt!SuppLimitStatus <> "A" Then
                        If AnnLmt!SuppLimitStatus = "B" And CAS!MBMPatientCovID <> "00" Then
                            If LifeTimeStatus = True Then
                                If LTGuaRenewal = True Then
                                    LifetimeAnnBal = Format(CDbl(tmpMBMLTAnnLmt) - CDbl(tmpTotalLTBordx), "#,#.00")
                                    tmpChkLTBalLmt = LifetimeAnnBal
                                    
                                    tmpChkAnnBalLmt = Format(CDbl(tmpMBMSuppAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                                    If CDbl(tmpChkLTBalLmt) < CDbl(tmpChkAnnBalLmt) Then
                                        tmpChkAnnBalLmt = Format(tmpChkLTBalLmt, "#,#.00")
                                    End If
                                Else
                                    tmpChkAnnBalLmt = Format(CDbl(tmpMBMSuppAnnLmt) - CDbl(tmpTotalLTBordx), "#,#.00")
                                End If
                            Else
                                tmpChkAnnBalLmt = Format(CDbl(tmpMBMSuppAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                            End If
                        ElseIf AnnLmt!SuppLimitStatus = "C" And CAS!MBMPatientCovID <> "00" Then
                            
                            If Topup = False Or LifeTimeStatus = True Then
                                If LTGuaRenewal = True Then
                                    LifetimeAnnBal = Format(CDbl(tmpMBMLTAnnLmt) - CDbl(tmpTotalLTBordx), "#,#.00")
                                    tmpChkLTBalLmt = LifetimeAnnBal
                                        tmpChkAnnBalLmt = Format(CDbl(tmpMBMSuppAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                                    If CDbl(tmpChkLTBalLmt) < CDbl(tmpChkAnnBalLmt) Then
                                        tmpChkAnnBalLmt = Format(tmpChkLTBalLmt, "#,#.00")
                                    End If
                                Else
                                    tmpChkAnnBalLmt = Format(CDbl(tmpMBMSuppAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                                End If
                            End If
                        Else
                            If Topup = False Then
                                If LifeTimeStatus = True Then
                                    If LTGuaRenewal = True Then
                                        LifetimeAnnBal = Format(CDbl(tmpMBMLTAnnLmt) - CDbl(tmpTotalLTBordx), "#,#.00")
                                        tmpChkLTBalLmt = LifetimeAnnBal
                                        tmpChkAnnBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                                        If CDbl(tmpChkLTBalLmt) < CDbl(tmpChkAnnBalLmt) Then
                                            tmpChkAnnBalLmt = Format(tmpChkLTBalLmt, "#,#.00")
                                        End If
                                    Else
                                        tmpChkAnnBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalLTBordx), "#,#.00")
                                    End If
                                Else
                                    tmpChkAnnBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                                End If
                            End If
                            If IsNumeric(tmpChkAnnBalLmt) = False Then
                                tmpChkAnnBalLmt = 0
                            End If
                        End If
                    Else
                        If Topup = False Then
                            If LifeTimeStatus = True Then
                                If LTGuaRenewal = True Then
                                    LifetimeAnnBal = Format(CDbl(tmpMBMLTAnnLmt) - CDbl(tmpTotalLTBordx), "#,#.00")
                                    tmpChkLTBalLmt = LifetimeAnnBal
                                    tmpChkAnnBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                                    If CDbl(tmpChkLTBalLmt) < CDbl(tmpChkAnnBalLmt) Then
                                        tmpChkAnnBalLmt = Format(tmpChkLTBalLmt, "#,#.00")
                                    End If
                                Else
                                    tmpChkAnnBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalLTBordx), "#,#.00")
                                End If
                            Else
                                tmpChkAnnBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                            End If
                        ElseIf Topup = True Then
                            If CASTopup = False Then
                                    tmpChkAnnBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                            End If
                        End If
                    End If
                    OldBalLmt = tmpChkAnnBalLmt
                    AnnLmt.Close
                    Set AnnLmt = Nothing
                If IsNumeric(tmpChkAnnBalLmt) = False Then
                    tmpChkAnnBalLmt = 0
                End If
                If AnnInd = True Then
                    If IsNumeric(frmGLMenu.txtCASLGAmt) Then
                        If CDbl(frmGLMenu.txtCASLGAmt) > CDbl(tmpChkAnnBalLmt) Then
                            LGStatus = False
                        Else
                            LGStatus = True
                        End If
                    Else
                        LGStatus = True
                    End If
                Else
                    LGStatus = True
                End If
            Else
                LGStatus = True
            End If
        End If
End Sub

Public Function FindHosName(HOSCode As String) As String
    Dim rst3 As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    'HOSSql = " select HOSCode, HOSNAme from HOS where HOSCOde='" & Trim(HOSCode) & "'"
    'rst3.Open HOSSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "3"
    strAppend = " AND HOSCOde = '" & Trim(HOSCode) & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst3.CursorLocation = adUseClient
        rst3.CursorType = adOpenForwardOnly
        rst3.CacheSize = 100
        Set rst3 = cmd.Execute
        
    If Not rst3.EOF Then
        FindHosName = rst3!HOSName
    Else
        MsgBox "Hospital Name not found!", vbCritical + vbOKOnly
    End If
    rst3.Close
    Set rst3 = Nothing
End Function


Private Sub funcwordMgmt()
'On Error Resume Next
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim i As Integer
Dim ii As Integer
Dim RecordSelected As Integer
Dim CaseRec As New ADODB.Recordset
Dim strsqlCaseRec As String
Dim WSHistoryRec As New ADODB.Recordset
Dim strsqlWSHistoryRec As String
Dim amntb4GST As Double
Dim GST As Double
Dim amntWithGST As Double
Dim ListCount As Integer
Dim ForLoop As Integer
Dim Itemcnt As Integer

       amntb4GST = 0
       GST = 0
       amntWithGST = 0
       
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
    
    On Error GoTo HANDLE_ERROR
    objExcel.Interactive = False
    
        If objExcel.Visible = False Then
            objExcel.Visible = True
        End If
    
        objExcel.WindowState = xlMaximized

        Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Case\MgmtNote_New.xlt")
    
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
                        

            ForLoop = 0
            ListCount = IIf(lstNote1(0).ListItems.Count > 0, lstNote1(0).ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstNote1(0).ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 6, "A") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 6, "B") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 6, "C") = .SubItems(2)
                    objWorksheet.Cells(CDbl(Itemcnt) + 6, "D") = .SubItems(3)
                    objWorksheet.Cells(CDbl(Itemcnt) + 6, "E") = .SubItems(4)
                    objWorksheet.Cells(CDbl(Itemcnt) + 6, "F") = .SubItems(5)
                    objWorksheet.Cells(CDbl(Itemcnt) + 6, "G") = .SubItems(6)
                    objWorksheet.Cells(CDbl(Itemcnt) + 6, "H") = .SubItems(7)
                End With
            Next
        
            ForLoop = 0
            ListCount = IIf(lstNote1(1).ListItems.Count > 0, lstNote1(1).ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstNote1(1).ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 19, "A") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 19, "B") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 19, "C") = .SubItems(2)
                    objWorksheet.Cells(CDbl(Itemcnt) + 19, "D") = .SubItems(3)
                    objWorksheet.Cells(CDbl(Itemcnt) + 19, "E") = .SubItems(4)
                    objWorksheet.Cells(CDbl(Itemcnt) + 19, "F") = .SubItems(5)
                    objWorksheet.Cells(CDbl(Itemcnt) + 19, "G") = .SubItems(6)
                    objWorksheet.Cells(CDbl(Itemcnt) + 19, "H") = .SubItems(7)
                    objWorksheet.Cells(CDbl(Itemcnt) + 19, "I") = .SubItems(8)
                    objWorksheet.Cells(CDbl(Itemcnt) + 19, "J") = .SubItems(9)
                End With
            Next
        
        
            ForLoop = 0
            ListCount = IIf(lstLGListing.ListItems.Count > 0, lstLGListing.ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstLGListing.ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 34, "A") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 34, "B") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 34, "C") = .SubItems(2)
                    objWorksheet.Cells(CDbl(Itemcnt) + 34, "D") = .SubItems(3)
                    objWorksheet.Cells(CDbl(Itemcnt) + 34, "E") = .SubItems(4)
                    objWorksheet.Cells(CDbl(Itemcnt) + 34, "F") = .SubItems(5)
                    objWorksheet.Cells(CDbl(Itemcnt) + 34, "G") = .SubItems(6)
                    objWorksheet.Cells(CDbl(Itemcnt) + 34, "H") = .SubItems(7)
                    objWorksheet.Cells(CDbl(Itemcnt) + 34, "I") = .SubItems(8)
                End With
            Next
        
            ForLoop = 0
            ListCount = IIf(lstNote1(2).ListItems.Count > 0, lstNote1(2).ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstNote1(2).ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 41, "A") = .Text

                End With
            Next
        
            ForLoop = 0
            ListCount = IIf(lstNote1(8).ListItems.Count > 0, lstNote1(2).ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstNote1(8).ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 41, "F") = .Text

                End With
            Next
            
            ForLoop = 0
            ListCount = IIf(lstNote1(5).ListItems.Count > 0, lstNote1(5).ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstNote1(5).ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 61, "A") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 61, "B") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 61, "C") = .SubItems(2)
                    objWorksheet.Cells(CDbl(Itemcnt) + 61, "D") = .SubItems(3)
                    objWorksheet.Cells(CDbl(Itemcnt) + 61, "E") = .SubItems(4)
                    objWorksheet.Cells(CDbl(Itemcnt) + 61, "F") = .SubItems(5)
                    objWorksheet.Cells(CDbl(Itemcnt) + 61, "G") = .SubItems(6)
                    objWorksheet.Cells(CDbl(Itemcnt) + 61, "H") = .SubItems(7)
                    objWorksheet.Cells(CDbl(Itemcnt) + 61, "I") = .SubItems(8)
                    objWorksheet.Cells(CDbl(Itemcnt) + 61, "J") = .SubItems(9)
                End With
            Next
            
            ForLoop = 0
            ListCount = IIf(lstNote1(3).ListItems.Count > 0, lstNote1(3).ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstNote1(3).ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 76, "A") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 76, "B") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 76, "C") = .SubItems(2)
                End With
            Next
            
            ForLoop = 0
            ListCount = IIf(lstNote1(7).ListItems.Count > 0, lstNote1(7).ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstNote1(7).ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 90, "A") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 90, "B") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 90, "C") = .SubItems(2)
                    objWorksheet.Cells(CDbl(Itemcnt) + 90, "D") = .SubItems(3)
                    objWorksheet.Cells(CDbl(Itemcnt) + 90, "E") = .SubItems(4)
                    objWorksheet.Cells(CDbl(Itemcnt) + 90, "F") = .SubItems(5)
                    objWorksheet.Cells(CDbl(Itemcnt) + 90, "G") = .SubItems(6)
                    objWorksheet.Cells(CDbl(Itemcnt) + 90, "H") = .SubItems(7)

                End With
            Next
            
            objWorksheet.Cells(51, "A") = txtProcCode1
            objWorksheet.Cells(51, "B") = txtProcedure1
            objWorksheet.Cells(51, "C") = txtSActFee1
            objWorksheet.Cells(51, "D") = txtS13Rate1
            objWorksheet.Cells(51, "E") = txtAnaesActFees1
            objWorksheet.Cells(51, "F") = txtAnae13Rate1
            
            objWorksheet.Cells(52, "A") = txtProcCode2
            objWorksheet.Cells(52, "B") = txtProcedure2
            objWorksheet.Cells(52, "C") = txtSActFee2
            objWorksheet.Cells(52, "D") = txtS13Rate2
            objWorksheet.Cells(52, "E") = txtAnaesActFees2
            objWorksheet.Cells(52, "F") = txtAnae13Rate2
            
            objWorksheet.Cells(98, "A") = txtSpecialRemark
            
            'objWorksheet.Cells(108, "A") = txtPrincipalExc
            
        objExcel.DisplayAlerts = True
        'objWorksheet.Columns.AutoFit
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        Exit Sub
HANDLE_ERROR:
        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        GoTo ExitFunction
End Sub

