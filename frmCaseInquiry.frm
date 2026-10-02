VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCaseInquiry 
   Caption         =   "Case Inquiry"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame23 
      Height          =   975
      Index           =   18
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   11655
      Begin VB.CommandButton cmdMBMEnq 
         Caption         =   "Enq"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   15
         Top             =   600
         Width           =   570
      End
      Begin VB.ComboBox cboPending 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10680
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox CboClaimability 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10200
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   330
         Left            =   3290
         TabIndex        =   12
         Top             =   240
         Width           =   570
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev"
         Height          =   330
         Left            =   2720
         TabIndex        =   11
         Top             =   240
         Width           =   570
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Height          =   330
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   570
      End
      Begin VB.TextBox txtMBMPolNo 
         Height          =   285
         Left            =   5400
         MaxLength       =   25
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         MaxLength       =   15
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtCaseNo 
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
         Left            =   720
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   720
         MaxLength       =   25
         TabIndex        =   7
         Top             =   600
         Width           =   1365
      End
      Begin VB.CommandButton cmdSearch 
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
         Height          =   330
         Left            =   3840
         TabIndex        =   6
         Top             =   240
         Width           =   690
      End
      Begin VB.ComboBox cboStatus1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8160
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtMBMName 
         Height          =   285
         Left            =   3720
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtINSCode1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtPLNCode 
         Height          =   285
         Left            =   8640
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Pending"
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
         Index           =   47
         Left            =   10080
         TabIndex        =   23
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Claimability"
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
         Index           =   46
         Left            =   9360
         TabIndex        =   22
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "IC Num"
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
         Index           =   45
         Left            =   7320
         TabIndex        =   21
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Cas Status"
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
         Index           =   44
         Left            =   7320
         TabIndex        =   20
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Pol No"
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
         Index           =   43
         Left            =   4800
         TabIndex        =   19
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Mem Name"
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
         Index           =   42
         Left            =   2880
         TabIndex        =   18
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Mem No"
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
         Index           =   41
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Case No"
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
         Index           =   40
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1140
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5850
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   10319
      _Version        =   393216
      Tabs            =   19
      TabsPerRow      =   19
      TabHeight       =   529
      TabCaption(0)   =   "Case Det"
      TabPicture(0)   =   "frmCaseInquiry.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame23(30)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Others"
      TabPicture(1)   =   "frmCaseInquiry.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame23(15)"
      Tab(1).Control(1)=   "Frame23(9)"
      Tab(1).Control(2)=   "Frame23(12)"
      Tab(1).Control(3)=   "Frame23(7)"
      Tab(1).Control(4)=   "Frame23(1)"
      Tab(1).Control(5)=   "Frame23(0)"
      Tab(1).Control(6)=   "Frame23(6)"
      Tab(1).Control(7)=   "Frame23(5)"
      Tab(1).Control(8)=   "Frame23(4)"
      Tab(1).Control(9)=   "Frame23(3)"
      Tab(1).Control(10)=   "Frame23(2)"
      Tab(1).Control(11)=   "Frame2"
      Tab(1).Control(12)=   "Frame23(14)"
      Tab(1).Control(13)=   "Frame23(8)"
      Tab(1).Control(14)=   "FrameS13MinRate"
      Tab(1).Control(15)=   "FrameS13Rate"
      Tab(1).Control(16)=   "Frame13Rate"
      Tab(1).Control(17)=   "FrameNS13RateMin"
      Tab(1).Control(18)=   "FrameNS13Rate"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Take Over/Others"
      TabPicture(2)   =   "frmCaseInquiry.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame23(31)"
      Tab(2).Control(1)=   "Frame1(0)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Mgmt Note"
      TabPicture(3)   =   "frmCaseInquiry.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame23(10)"
      Tab(3).Control(1)=   "Frame23(11)"
      Tab(3).Control(2)=   "txtCASRemark"
      Tab(3).Control(3)=   "txtCASRemarkTwo"
      Tab(3).Control(4)=   "txtCASRemark3"
      Tab(3).Control(5)=   "txtCASRemark4"
      Tab(3).Control(6)=   "txtCASRemark5"
      Tab(3).Control(7)=   "lblPage1"
      Tab(3).Control(8)=   "lblPage2"
      Tab(3).Control(9)=   "lblPage3"
      Tab(3).Control(10)=   "lblPage4"
      Tab(3).Control(11)=   "lblPage5"
      Tab(3).Control(12)=   "lblPage"
      Tab(3).Control(13)=   "Label1(84)"
      Tab(3).Control(14)=   "Label1(0)"
      Tab(3).Control(15)=   "Label1(85)"
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "LG"
      TabPicture(4)   =   "frmCaseInquiry.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstFollowUp"
      Tab(4).Control(1)=   "lstLGListing"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "WS"
      TabPicture(5)   =   "frmCaseInquiry.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame20(12)"
      Tab(5).Control(1)=   "lstWorksheet"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Inv"
      TabPicture(6)   =   "frmCaseInquiry.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lstINVListing"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Pymt && Rcpt"
      TabPicture(7)   =   "frmCaseInquiry.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label1(88)"
      Tab(7).Control(1)=   "Label1(87)"
      Tab(7).Control(2)=   "Label1(86)"
      Tab(7).Control(3)=   "Label1(91)"
      Tab(7).Control(4)=   "Label1(90)"
      Tab(7).Control(5)=   "Label1(89)"
      Tab(7).Control(6)=   "lstRecPayor"
      Tab(7).Control(7)=   "lstRecMNRB"
      Tab(7).Control(8)=   "lstRecFRMBM"
      Tab(7).Control(9)=   "lstReissuedMBM"
      Tab(7).Control(10)=   "lstMBMPayment"
      Tab(7).Control(11)=   "lstHOSPayment"
      Tab(7).Control(12)=   "cmdPrintPymt"
      Tab(7).ControlCount=   13
      TabCaption(8)   =   "Reminder"
      TabPicture(8)   =   "frmCaseInquiry.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lstReminder"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Non-Dis"
      TabPicture(9)   =   "frmCaseInquiry.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "lstPreUndIllness"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Benefit"
      TabPicture(10)  =   "frmCaseInquiry.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lstBenefits"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "Audit Review"
      TabPicture(11)  =   "frmCaseInquiry.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Label1(93)"
      Tab(11).Control(1)=   "txtPrincipalExc"
      Tab(11).Control(2)=   "txtSuppExc"
      Tab(11).Control(3)=   "cboSuppExc"
      Tab(11).ControlCount=   4
      TabCaption(12)  =   "Policy Terms"
      TabPicture(12)  =   "frmCaseInquiry.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Fgrid"
      Tab(12).ControlCount=   1
      TabCaption(13)  =   "Surg Schedule"
      TabPicture(13)  =   "frmCaseInquiry.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "lstSurgProc"
      Tab(13).ControlCount=   1
      TabCaption(14)  =   "LifeTime"
      TabPicture(14)  =   "frmCaseInquiry.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "lstLTLmt"
      Tab(14).ControlCount=   1
      TabCaption(15)  =   "Forms"
      TabPicture(15)  =   "frmCaseInquiry.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "Frame23(17)"
      Tab(15).Control(1)=   "Frame23(16)"
      Tab(15).Control(2)=   "CmnDlg"
      Tab(15).Control(3)=   "List1"
      Tab(15).Control(4)=   "Frame23(13)"
      Tab(15).Control(5)=   "List3"
      Tab(15).ControlCount=   6
      TabCaption(16)  =   "Policy"
      TabPicture(16)  =   "frmCaseInquiry.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "lstPolicyWording"
      Tab(16).ControlCount=   1
      TabCaption(17)  =   "Carton"
      TabPicture(17)  =   "frmCaseInquiry.frx":01DC
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "lstStorage"
      Tab(17).ControlCount=   1
      TabCaption(18)  =   "Emails"
      TabPicture(18)  =   "frmCaseInquiry.frx":01F8
      Tab(18).ControlEnabled=   0   'False
      Tab(18).Control(0)=   "lstEmails"
      Tab(18).ControlCount=   1
      Begin VB.ListBox List3 
         Height          =   2535
         ItemData        =   "frmCaseInquiry.frx":0214
         Left            =   -68040
         List            =   "frmCaseInquiry.frx":0216
         Style           =   1  'Checkbox
         TabIndex        =   357
         ToolTipText     =   "right click to randomize contents"
         Top             =   3120
         Width           =   4455
      End
      Begin VB.CommandButton cmdPrintPymt 
         Caption         =   "Print"
         Height          =   375
         Left            =   -64440
         TabIndex        =   353
         Top             =   5280
         Width           =   975
      End
      Begin VB.ListBox lstEmails 
         Height          =   5325
         ItemData        =   "frmCaseInquiry.frx":0218
         Left            =   -74880
         List            =   "frmCaseInquiry.frx":021F
         TabIndex        =   350
         ToolTipText     =   "right click to randomize contents"
         Top             =   400
         Width           =   4455
      End
      Begin VB.ListBox lstPolicyWording 
         Height          =   5325
         ItemData        =   "frmCaseInquiry.frx":022E
         Left            =   -74880
         List            =   "frmCaseInquiry.frx":0230
         TabIndex        =   348
         ToolTipText     =   "right click to randomize contents"
         Top             =   360
         Width           =   4455
      End
      Begin VB.Frame Frame23 
         Height          =   975
         Index           =   13
         Left            =   -70200
         TabIndex        =   339
         Top             =   480
         Width           =   6615
         Begin VB.TextBox txtINSCode 
            BackColor       =   &H80000001&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   14400
            TabIndex        =   345
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton CmdCopy 
            Caption         =   "Copy"
            Height          =   255
            Left            =   5760
            TabIndex        =   344
            Top             =   630
            Width           =   735
         End
         Begin VB.TextBox TxtDest 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   343
            Top             =   600
            Width           =   4575
         End
         Begin VB.TextBox TxtSrc 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   342
            Top             =   240
            Width           =   4575
         End
         Begin VB.CommandButton CmdBrowse 
            Caption         =   "Browse"
            Height          =   255
            Left            =   5760
            TabIndex        =   341
            Top             =   260
            Width           =   735
         End
         Begin VB.CheckBox chkVerify 
            Caption         =   "Verify Copy"
            Height          =   285
            Left            =   3600
            TabIndex        =   340
            Top             =   2040
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Source :"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   347
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Destination :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   346
            Top             =   600
            Width           =   945
         End
      End
      Begin VB.ListBox List1 
         Height          =   2205
         ItemData        =   "frmCaseInquiry.frx":0232
         Left            =   -74760
         List            =   "frmCaseInquiry.frx":0234
         TabIndex        =   338
         ToolTipText     =   "right click to randomize contents"
         Top             =   600
         Width           =   4455
      End
      Begin VB.ComboBox cboSuppExc 
         Height          =   315
         Left            =   -73320
         TabIndex        =   333
         Top             =   4920
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.TextBox txtSuppExc 
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
         Height          =   2265
         Left            =   -74640
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   332
         Top             =   4680
         Visible         =   0   'False
         Width           =   11415
      End
      Begin VB.TextBox txtPrincipalExc 
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
         Height          =   2580
         Left            =   -74880
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   331
         Top             =   840
         Width           =   11415
      End
      Begin VB.Frame Frame20 
         Height          =   615
         Index           =   12
         Left            =   -74880
         TabIndex        =   286
         Top             =   5100
         Width           =   11415
         Begin VB.ComboBox cboAppStatus 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5040
            TabIndex        =   289
            Top             =   200
            Width           =   1215
         End
         Begin VB.TextBox txtCLMAppCode 
            Height          =   285
            Left            =   7080
            MaxLength       =   15
            TabIndex        =   288
            Top             =   240
            Width           =   2055
         End
         Begin VB.ComboBox cboREVBordxInd 
            Height          =   315
            Left            =   960
            TabIndex        =   287
            Top             =   675
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSMask.MaskEdBox txtAppDate 
            Height          =   300
            Left            =   2880
            TabIndex        =   290
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
            Caption         =   "App Status"
            Height          =   255
            Index           =   112
            Left            =   4200
            TabIndex        =   296
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "App Date"
            Height          =   255
            Index           =   113
            Left            =   2160
            TabIndex        =   295
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label65 
            Caption         =   "Claim No"
            Height          =   255
            Index           =   115
            Left            =   120
            TabIndex        =   294
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblClaimNo 
            Height          =   255
            Left            =   840
            TabIndex        =   293
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label65 
            Caption         =   "App Code"
            Height          =   255
            Index           =   118
            Left            =   6360
            TabIndex        =   292
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label65 
            Caption         =   "Rev Bordx"
            Height          =   255
            Index           =   119
            Left            =   120
            TabIndex        =   291
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.Frame Frame23 
         Height          =   4935
         Index           =   10
         Left            =   -73680
         TabIndex        =   285
         Top             =   720
         Width           =   15
      End
      Begin VB.Frame Frame23 
         Height          =   4935
         Index           =   11
         Left            =   -67320
         TabIndex        =   284
         Top             =   720
         Width           =   15
      End
      Begin VB.TextBox txtCASRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4995
         Left            =   -74760
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   283
         Top             =   720
         Width           =   8895
      End
      Begin VB.TextBox txtCASRemarkTwo 
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
         Left            =   -74760
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   282
         Top             =   720
         Width           =   8895
      End
      Begin VB.TextBox txtCASRemark3 
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
         Left            =   -74760
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   281
         Top             =   720
         Width           =   8895
      End
      Begin VB.TextBox txtCASRemark4 
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
         Left            =   -74760
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   280
         Top             =   720
         Width           =   8895
      End
      Begin VB.TextBox txtCASRemark5 
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
         Left            =   -74760
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   279
         Top             =   720
         Width           =   8895
      End
      Begin VB.Frame Frame23 
         Height          =   675
         Index           =   31
         Left            =   -74880
         TabIndex        =   268
         Top             =   360
         Width           =   11415
         Begin VB.ComboBox cboIntExtTO 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5520
            TabIndex        =   273
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtLinkCASNo 
            Appearance      =   0  'Flat
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
            Left            =   9960
            MaxLength       =   15
            TabIndex        =   272
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtCASTOPlan 
            Appearance      =   0  'Flat
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
            Left            =   7920
            MaxLength       =   4
            TabIndex        =   271
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtCASTakeOverAmt 
            Appearance      =   0  'Flat
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
            MaxLength       =   15
            TabIndex        =   270
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cboCASTakeOveInd 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   720
            TabIndex        =   269
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
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
            Index           =   120
            Left            =   4200
            TabIndex        =   278
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "T/O Ind"
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
            Index           =   68
            Left            =   120
            TabIndex        =   277
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "T/O Amt"
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
            Index           =   69
            Left            =   2160
            TabIndex        =   276
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "T/O Pln"
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
            Index           =   70
            Left            =   7320
            TabIndex        =   275
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "Link Case"
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
            Index           =   75
            Left            =   9240
            TabIndex        =   274
            Top             =   240
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1395
         Index           =   0
         Left            =   -74880
         TabIndex        =   253
         Top             =   960
         Width           =   11415
         Begin VB.TextBox txtREDAmt 
            Appearance      =   0  'Flat
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
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   258
            Top             =   165
            Width           =   1440
         End
         Begin VB.ComboBox cboPayPending 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4815
            Style           =   1  'Simple Combo
            TabIndex        =   254
            Top             =   165
            Width           =   1440
         End
         Begin MSMask.MaskEdBox txtFirstDocDate 
            Height          =   315
            Left            =   4815
            TabIndex        =   259
            Top             =   450
            Width           =   1440
            _ExtentX        =   2540
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
            Left            =   4815
            TabIndex        =   260
            Top             =   735
            Width           =   1440
            _ExtentX        =   2540
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
         Begin VB.ComboBox cboTopUpLmt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   257
            Top             =   420
            Width           =   1440
         End
         Begin VB.ComboBox cboCASInd 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            Style           =   1  'Simple Combo
            TabIndex        =   256
            Top             =   705
            Width           =   1440
         End
         Begin VB.ComboBox cboQuery 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            Style           =   1  'Simple Combo
            TabIndex        =   255
            Top             =   960
            Width           =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "Red Amt"
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
            Index           =   71
            Left            =   120
            TabIndex        =   267
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "Topup Lmt"
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
            Index           =   72
            Left            =   120
            TabIndex        =   266
            Top             =   435
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "Special App"
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
            Index           =   73
            Left            =   120
            TabIndex        =   265
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "Query Status"
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
            Index           =   74
            Left            =   120
            TabIndex        =   264
            Top             =   1005
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "Pymt P'ding App"
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
            Index           =   76
            Left            =   3360
            TabIndex        =   263
            Top             =   195
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "First Doc Date"
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
            Index           =   114
            Left            =   3360
            TabIndex        =   262
            Top             =   465
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "Subseq. Doc Date"
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
            Index           =   115
            Left            =   3360
            TabIndex        =   261
            Top             =   765
            Width           =   1380
         End
      End
      Begin VB.Frame FrameNS13Rate 
         Enabled         =   0   'False
         Height          =   2010
         Left            =   -67920
         TabIndex        =   245
         Top             =   2505
         Width           =   1095
         Begin VB.TextBox txtNS13Rate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   251
            Top             =   400
            Width           =   975
         End
         Begin VB.TextBox txtNS13Rate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   250
            Top             =   660
            Width           =   975
         End
         Begin VB.TextBox txtNS13Rate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   249
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox txtNS13Rate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   248
            Top             =   1140
            Width           =   975
         End
         Begin VB.TextBox txtNS13Rate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   247
            Top             =   1410
            Width           =   975
         End
         Begin VB.TextBox txtNS13Rate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   246
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "13th Sch (Max)"
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
            Left            =   120
            TabIndex        =   252
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame FrameNS13RateMin 
         Enabled         =   0   'False
         Height          =   2010
         Left            =   -69000
         TabIndex        =   237
         Top             =   2505
         Width           =   1095
         Begin VB.TextBox txtNS13MinRate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   243
            Top             =   400
            Width           =   975
         End
         Begin VB.TextBox txtNS13MinRate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   242
            Top             =   660
            Width           =   975
         End
         Begin VB.TextBox txtNS13MinRate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   241
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox txtNS13MinRate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   240
            Top             =   1140
            Width           =   975
         End
         Begin VB.TextBox txtNS13MinRate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   239
            Top             =   1410
            Width           =   975
         End
         Begin VB.TextBox txtNS13MinRate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   238
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
            Left            =   120
            TabIndex        =   244
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame Frame13Rate 
         Enabled         =   0   'False
         Height          =   2025
         Left            =   -64680
         TabIndex        =   229
         Top             =   480
         Width           =   1215
         Begin VB.TextBox txtAnae13Rate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   235
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtAnae13Rate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   234
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtAnae13Rate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   233
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtAnae13Rate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   232
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtAnae13Rate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   231
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtAnae13Rate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   230
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Anaes 13th Sch"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   66
            Left            =   240
            TabIndex        =   236
            Top             =   100
            Width           =   705
         End
      End
      Begin VB.Frame FrameS13Rate 
         Enabled         =   0   'False
         Height          =   2040
         Left            =   -67920
         TabIndex        =   221
         Top             =   480
         Width           =   1095
         Begin VB.TextBox txtS13Rate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   227
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtS13Rate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   226
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtS13Rate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   225
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtS13Rate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   224
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtS13Rate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   223
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtS13Rate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   222
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "13th Sch (Max)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   65
            Left            =   240
            TabIndex        =   228
            Top             =   100
            Width           =   705
         End
      End
      Begin VB.Frame FrameS13MinRate 
         Enabled         =   0   'False
         Height          =   2040
         Left            =   -69000
         TabIndex        =   213
         Top             =   480
         Width           =   1095
         Begin VB.TextBox txtS13MinRate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   219
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtS13MinRate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   218
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtS13MinRate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   217
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtS13MinRate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   216
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtS13MinRate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   215
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtS13MinRate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   60
            MaxLength       =   8
            TabIndex        =   214
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "13th Sch (Min)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   220
            Top             =   100
            Width           =   705
         End
      End
      Begin VB.Frame Frame23 
         Height          =   2010
         Index           =   8
         Left            =   -70080
         TabIndex        =   205
         Top             =   2505
         Width           =   1095
         Begin VB.TextBox txtNSMMARate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   211
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtNSMMARate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   210
            Top             =   1410
            Width           =   975
         End
         Begin VB.TextBox txtNSMMARate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   209
            Top             =   1140
            Width           =   975
         End
         Begin VB.TextBox txtNSMMARate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   208
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox txtNSMMARate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   207
            Top             =   660
            Width           =   975
         End
         Begin VB.TextBox txtNSMMARate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   206
            Top             =   400
            Width           =   975
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   64
            Left            =   225
            TabIndex        =   212
            Top             =   120
            Width           =   465
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Medical Management"
         Height          =   1155
         Index           =   14
         Left            =   -71160
         TabIndex        =   201
         Top             =   4485
         Width           =   3525
         Begin VB.TextBox txtMEDSURDone3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   80
            MaxLength       =   150
            TabIndex        =   204
            Top             =   795
            Width           =   3390
         End
         Begin VB.TextBox txtMEDSURDone2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   80
            MaxLength       =   150
            TabIndex        =   203
            Top             =   540
            Width           =   3390
         End
         Begin VB.TextBox txtMEDSURDone1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   80
            MaxLength       =   150
            TabIndex        =   202
            Top             =   285
            Width           =   3390
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Anaes"
         Height          =   1155
         Left            =   -67635
         TabIndex        =   197
         Top             =   4485
         Width           =   4215
         Begin VB.TextBox txtAnaesDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   100
            TabIndex        =   200
            Top             =   270
            Width           =   4035
         End
         Begin VB.TextBox txtAnaesDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   120
            MaxLength       =   100
            TabIndex        =   199
            Top             =   535
            Width           =   4035
         End
         Begin VB.TextBox txtAnaesDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   120
            MaxLength       =   100
            TabIndex        =   198
            Top             =   785
            Width           =   4035
         End
      End
      Begin VB.Frame Frame23 
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
         Height          =   1800
         Index           =   2
         Left            =   -72240
         TabIndex        =   190
         Top             =   720
         Width           =   975
         Begin VB.TextBox txtProcCode6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   196
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtProcCode5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   195
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtProcCode4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   194
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtProcCode3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   193
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtProcCode2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   192
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtProcCode1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   191
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame23 
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
         Index           =   3
         Left            =   -70080
         TabIndex        =   182
         Top             =   480
         Width           =   1095
         Begin VB.TextBox txtSMMARate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   188
            Top             =   1695
            Width           =   975
         End
         Begin VB.TextBox txtSMMARate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   187
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtSMMARate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   186
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtSMMARate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   185
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtSMMARate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   184
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtSMMARate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   183
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   58
            Left            =   360
            TabIndex        =   189
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame23 
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
         Index           =   4
         Left            =   -65760
         TabIndex        =   174
         Top             =   480
         Width           =   1095
         Begin VB.TextBox txtAnaesMMARate5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   180
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtAnaesMMARate4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   179
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtAnaesMMARate3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   178
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtAnaesMMARate2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   177
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtAnaesMMARate1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   176
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtAnaesMMARate6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   175
            Top             =   1695
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Aneas MMA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   61
            Left            =   240
            TabIndex        =   181
            Top             =   105
            Width           =   465
         End
      End
      Begin VB.Frame Frame23 
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
         Index           =   5
         Left            =   -71160
         TabIndex        =   166
         Top             =   480
         Width           =   1080
         Begin VB.TextBox txtSActFee5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   45
            MaxLength       =   8
            TabIndex        =   172
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtSActFee4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   45
            MaxLength       =   8
            TabIndex        =   171
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtSActFee3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   45
            MaxLength       =   8
            TabIndex        =   170
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtSActFee2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   45
            MaxLength       =   8
            TabIndex        =   169
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtSActFee1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   45
            MaxLength       =   8
            TabIndex        =   168
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtSActFee6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   45
            MaxLength       =   8
            TabIndex        =   167
            Top             =   1695
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Actual"
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
            Index           =   60
            Left            =   360
            TabIndex        =   173
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame23 
         Height          =   2040
         Index           =   6
         Left            =   -66840
         TabIndex        =   158
         Top             =   480
         Width           =   1095
         Begin VB.TextBox txtAnaesActFees5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   164
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox txtAnaesActFees4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   163
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtAnaesActFees3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   162
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtAnaesActFees2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   161
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtAnaesActFees1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   160
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtAnaesActFees6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   159
            Top             =   1695
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Actual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   62
            Left            =   240
            TabIndex        =   165
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Surgical Procedures"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Index           =   0
         Left            =   -74880
         TabIndex        =   151
         Top             =   720
         Width           =   2640
         Begin VB.TextBox txtProcedure6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   157
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   156
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   155
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   154
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   153
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtProcedure1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   150
            TabIndex        =   152
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Non-Surgical Procedures"
         Height          =   1770
         Index           =   1
         Left            =   -74880
         TabIndex        =   144
         Top             =   2730
         Width           =   2640
         Begin VB.TextBox txtNSProcedure6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   200
            TabIndex        =   150
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   200
            TabIndex        =   149
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   200
            TabIndex        =   148
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   200
            TabIndex        =   147
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   200
            TabIndex        =   146
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtNSProcedure1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   200
            TabIndex        =   145
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "OPCS"
         Height          =   1770
         Index           =   7
         Left            =   -72240
         TabIndex        =   137
         Top             =   2730
         Width           =   975
         Begin VB.TextBox txtNSProcCode6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   6
            TabIndex        =   143
            Top             =   1440
            Width           =   885
         End
         Begin VB.TextBox txtNSProcCode5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   6
            TabIndex        =   142
            Top             =   1200
            Width           =   885
         End
         Begin VB.TextBox txtNSProcCode4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   6
            TabIndex        =   141
            Top             =   960
            Width           =   885
         End
         Begin VB.TextBox txtNSProcCode3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   6
            TabIndex        =   140
            Top             =   720
            Width           =   885
         End
         Begin VB.TextBox txtNSProcCode2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   6
            TabIndex        =   139
            Top             =   480
            Width           =   885
         End
         Begin VB.TextBox txtNSProcCode1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   40
            MaxLength       =   6
            TabIndex        =   138
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Att Doc"
         Height          =   1605
         Index           =   12
         Left            =   -66810
         TabIndex        =   131
         Top             =   2505
         Width           =   3375
         Begin VB.TextBox txtATTDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   136
            Top             =   1245
            Width           =   3165
         End
         Begin VB.TextBox txtATTDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   135
            Top             =   1005
            Width           =   3165
         End
         Begin VB.TextBox txtATTDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   134
            Top             =   750
            Width           =   3165
         End
         Begin VB.TextBox txtATTDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   133
            Top             =   495
            Width           =   3165
         End
         Begin VB.TextBox txtATTDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   100
            TabIndex        =   132
            Top             =   240
            Width           =   3165
         End
      End
      Begin VB.Frame Frame23 
         Height          =   2010
         Index           =   9
         Left            =   -71160
         TabIndex        =   123
         Top             =   2505
         Width           =   1095
         Begin VB.TextBox txtNSActFee6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   129
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtNSActFee5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   128
            Top             =   1410
            Width           =   975
         End
         Begin VB.TextBox txtNSActFee4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   127
            Top             =   1140
            Width           =   975
         End
         Begin VB.TextBox txtNSActFee3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   126
            Top             =   900
            Width           =   975
         End
         Begin VB.TextBox txtNSActFee2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   125
            Top             =   660
            Width           =   975
         End
         Begin VB.TextBox txtNSActFee1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   40
            MaxLength       =   8
            TabIndex        =   124
            Top             =   400
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Actual"
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
            Index           =   63
            Left            =   180
            TabIndex        =   130
            Top             =   120
            Width           =   465
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Investigation"
         Height          =   1155
         Index           =   15
         Left            =   -74880
         TabIndex        =   116
         Top             =   4485
         Width           =   3675
         Begin VB.TextBox txtPROCINVDone3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   150
            TabIndex        =   122
            Top             =   795
            Width           =   2535
         End
         Begin VB.TextBox txtPROCINVDone2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   150
            TabIndex        =   121
            Top             =   540
            Width           =   2535
         End
         Begin VB.TextBox txtPROCINVDone1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   150
            TabIndex        =   120
            Top             =   285
            Width           =   2535
         End
         Begin VB.TextBox txtINVCode3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            MaxLength       =   150
            TabIndex        =   119
            Top             =   795
            Width           =   975
         End
         Begin VB.TextBox txtINVCode2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            MaxLength       =   150
            TabIndex        =   118
            Top             =   540
            Width           =   975
         End
         Begin VB.TextBox txtINVCode1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            MaxLength       =   150
            TabIndex        =   117
            Top             =   285
            Width           =   975
         End
      End
      Begin VB.Frame Frame23 
         Enabled         =   0   'False
         Height          =   5445
         Index           =   30
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   11415
         Begin VB.Frame Frame23 
            Height          =   1400
            Index           =   27
            Left            =   5760
            TabIndex        =   74
            Top             =   2600
            Width           =   5535
            Begin VB.TextBox cboFDICDCode1 
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   77
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox cboFDICDCode2 
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   76
               Top             =   620
               Width           =   615
            End
            Begin VB.TextBox cboFDICDCode3 
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   75
               Top             =   870
               Width           =   615
            End
            Begin MSMask.MaskEdBox txtFDDOS1 
               Height          =   285
               Left            =   3600
               TabIndex        =   79
               Top             =   360
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtFDDOS2 
               Height          =   285
               Left            =   3600
               TabIndex        =   80
               Top             =   615
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtFDDOS3 
               Height          =   285
               Left            =   3600
               TabIndex        =   81
               Top             =   870
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtFinalDiag1 
               Height          =   285
               Index           =   0
               Left            =   90
               TabIndex        =   78
               Top             =   360
               Width           =   3525
            End
            Begin VB.TextBox txtFinalDiag1 
               Height          =   285
               Index           =   1
               Left            =   90
               TabIndex        =   358
               Top             =   615
               Width           =   3525
            End
            Begin VB.TextBox txtFinalDiag1 
               Height          =   285
               Index           =   2
               Left            =   90
               TabIndex        =   359
               Top             =   870
               Width           =   3525
            End
            Begin VB.Label Label1 
               Caption         =   "ICD"
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
               Index           =   56
               Left            =   4680
               TabIndex        =   84
               Top             =   120
               Width           =   540
            End
            Begin VB.Label Label1 
               Caption         =   "Date Onset"
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
               Index           =   55
               Left            =   3600
               TabIndex        =   83
               Top             =   120
               Width           =   900
            End
            Begin VB.Label Label1 
               Caption         =   "Final Diag"
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
               Index           =   54
               Left            =   120
               TabIndex        =   82
               Top             =   120
               Width           =   1140
            End
         End
         Begin VB.Frame Frame23 
            Height          =   1400
            Index           =   25
            Left            =   120
            TabIndex        =   63
            Top             =   2600
            Width           =   5415
            Begin VB.TextBox txtPrelimDiag1 
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   67
               Top             =   360
               Width           =   3375
            End
            Begin VB.TextBox cboICDCode1 
               Height          =   285
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   66
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox cboICDCode2 
               Height          =   285
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   65
               Top             =   620
               Width           =   615
            End
            Begin VB.TextBox cboICDCode3 
               Height          =   285
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   64
               Top             =   870
               Width           =   615
            End
            Begin MSMask.MaskEdBox txtPDDOS1 
               Height          =   285
               Left            =   3480
               TabIndex        =   68
               Top             =   360
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtPDDOS2 
               Height          =   285
               Left            =   3480
               TabIndex        =   69
               Top             =   620
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtPDDOS3 
               Height          =   285
               Left            =   3480
               TabIndex        =   70
               Top             =   870
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtPrelimDiag1 
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   360
               Top             =   600
               Width           =   3375
            End
            Begin VB.TextBox txtPrelimDiag1 
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   361
               Top             =   840
               Width           =   3375
            End
            Begin VB.Label Label1 
               Caption         =   "ICD"
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
               Index           =   50
               Left            =   4560
               TabIndex        =   73
               Top             =   120
               Width           =   780
            End
            Begin VB.Label Label1 
               Caption         =   "Date Onset"
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
               Index           =   49
               Left            =   3480
               TabIndex        =   72
               Top             =   120
               Width           =   900
            End
            Begin VB.Label Label1 
               Caption         =   "Prelim. Diag"
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
               Index           =   48
               Left            =   120
               TabIndex        =   71
               Top             =   120
               Width           =   1140
            End
         End
         Begin VB.ComboBox cboPatientList 
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
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   62
            Top             =   255
            Width           =   5000
         End
         Begin VB.ComboBox txtHOSCode 
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
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   61
            Top             =   525
            Width           =   5000
         End
         Begin VB.TextBox txtPatientIC 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   58
            Top             =   210
            Width           =   3495
         End
         Begin VB.TextBox KIVPatientName 
            Height          =   285
            Left            =   1200
            TabIndex        =   57
            Top             =   260
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Frame Frame23 
            Height          =   1335
            Index           =   26
            Left            =   120
            TabIndex        =   44
            Top             =   4000
            Width           =   5415
            Begin VB.TextBox txtUD1 
               Height          =   285
               Left            =   90
               TabIndex        =   50
               Top             =   380
               Width           =   3375
            End
            Begin VB.TextBox txtUD2 
               Height          =   285
               Left            =   90
               TabIndex        =   49
               Top             =   620
               Width           =   3375
            End
            Begin VB.TextBox txtUD3 
               Height          =   285
               Left            =   90
               TabIndex        =   48
               Top             =   870
               Width           =   3375
            End
            Begin VB.TextBox cboUDICDCode1 
               Height          =   285
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   47
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox cboUDICDCode2 
               Height          =   285
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   46
               Top             =   620
               Width           =   615
            End
            Begin VB.TextBox cboUDICDCode3 
               Height          =   285
               Left            =   4560
               MaxLength       =   3
               TabIndex        =   45
               Top             =   870
               Width           =   615
            End
            Begin MSMask.MaskEdBox txtUDDOS1 
               Height          =   285
               Left            =   3480
               TabIndex        =   51
               Top             =   360
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtUDDOS2 
               Height          =   285
               Left            =   3480
               TabIndex        =   52
               Top             =   620
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtUDDOS3 
               Height          =   285
               Left            =   3480
               TabIndex        =   53
               Top             =   870
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "ICD"
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
               Index           =   53
               Left            =   4560
               TabIndex        =   56
               Top             =   120
               Width           =   780
            End
            Begin VB.Label Label1 
               Caption         =   "Date Onset"
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
               Index           =   52
               Left            =   3480
               TabIndex        =   55
               Top             =   120
               Width           =   900
            End
            Begin VB.Label Label1 
               Caption         =   "Underlying Disease"
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
               Index           =   51
               Left            =   120
               TabIndex        =   54
               Top             =   120
               Width           =   1620
            End
         End
         Begin VB.Frame Frame23 
            Height          =   1335
            Index           =   23
            Left            =   5760
            TabIndex        =   39
            Top             =   4000
            Width           =   5535
            Begin VB.ComboBox cboREPCode1 
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
               Left            =   120
               Style           =   1  'Simple Combo
               TabIndex        =   42
               Top             =   360
               Width           =   5370
            End
            Begin VB.ComboBox cboREPCode2 
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
               Left            =   120
               Style           =   1  'Simple Combo
               TabIndex        =   41
               Top             =   645
               Width           =   5370
            End
            Begin VB.ComboBox cboREPCode3 
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
               Left            =   120
               Style           =   1  'Simple Combo
               TabIndex        =   40
               Top             =   915
               Width           =   5370
            End
            Begin VB.Label Label1 
               Caption         =   "Repudiation Reason"
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
               Index           =   57
               Left            =   120
               TabIndex        =   43
               Top             =   120
               Width           =   1740
            End
         End
         Begin MSMask.MaskEdBox txtFirstCall 
            Height          =   285
            Left            =   7800
            TabIndex        =   86
            Top             =   460
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFirstTreated 
            Height          =   285
            Left            =   7800
            TabIndex        =   87
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDateAdmitted 
            Height          =   285
            Left            =   9960
            TabIndex        =   88
            Top             =   460
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDateDischarge 
            Height          =   285
            Left            =   9960
            TabIndex        =   89
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtCASReserveAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   7800
            TabIndex        =   32
            Top             =   980
            Width           =   1215
         End
         Begin VB.TextBox txtCASLGAmt 
            Alignment       =   1  'Right Justify
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
            Left            =   9960
            TabIndex        =   33
            Top             =   980
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtLGTime 
            Height          =   315
            Left            =   9960
            TabIndex        =   90
            Top             =   1210
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
         Begin VB.ComboBox cboCondwillOcc 
            Height          =   315
            Left            =   7800
            Style           =   1  'Simple Combo
            TabIndex        =   31
            Top             =   1215
            Width           =   1215
         End
         Begin VB.ComboBox cboFLWTreatment 
            Height          =   315
            Left            =   7800
            Style           =   1  'Simple Combo
            TabIndex        =   30
            Top             =   1500
            Width           =   3495
         End
         Begin MSMask.MaskEdBox txtREGDate 
            Height          =   285
            Left            =   7800
            TabIndex        =   91
            Top             =   1785
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboFolUp 
            Height          =   315
            Left            =   9960
            Style           =   1  'Simple Combo
            TabIndex        =   28
            Top             =   1790
            Width           =   1335
         End
         Begin VB.ComboBox cboYearDay 
            Height          =   315
            Left            =   10440
            TabIndex        =   26
            Top             =   2070
            Width           =   855
         End
         Begin VB.TextBox txtPatientAge 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9960
            TabIndex        =   27
            Top             =   2070
            Width           =   500
         End
         Begin VB.TextBox txtMBMPolYr 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7800
            TabIndex        =   29
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtFOLDate 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   305
            Left            =   7800
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   2315
            Width           =   1215
         End
         Begin VB.TextBox txtHOSName 
            Height          =   285
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   351
            Top             =   810
            Width           =   5000
         End
         Begin VB.TextBox txtReason 
            Height          =   540
            Left            =   1200
            ScrollBars      =   3  'Both
            TabIndex        =   60
            Top             =   1050
            Width           =   5000
         End
         Begin VB.TextBox txtREFBy 
            Height          =   285
            Left            =   1200
            TabIndex        =   59
            Top             =   1560
            Width           =   5000
         End
         Begin VB.ComboBox cboNature 
            Height          =   315
            Left            =   1200
            Style           =   1  'Simple Combo
            TabIndex        =   38
            Top             =   1800
            Width           =   1800
         End
         Begin VB.ComboBox cboModality 
            Height          =   315
            Left            =   1200
            Style           =   1  'Simple Combo
            TabIndex        =   37
            Top             =   2040
            Width           =   1800
         End
         Begin VB.ComboBox CboStayType 
            Height          =   315
            ItemData        =   "frmCaseInquiry.frx":0236
            Left            =   4400
            List            =   "frmCaseInquiry.frx":0238
            Style           =   1  'Simple Combo
            TabIndex        =   36
            Top             =   1800
            Width           =   1800
         End
         Begin VB.ComboBox cboPayType 
            Height          =   315
            Left            =   4400
            Style           =   1  'Simple Combo
            TabIndex        =   35
            Top             =   2040
            Width           =   1800
         End
         Begin MSMask.MaskEdBox MaskEdBox12 
            Height          =   315
            Left            =   4800
            TabIndex        =   85
            Top             =   1920
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label26 
            Caption         =   "Hospital Name"
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
            Left            =   120
            TabIndex        =   352
            Top             =   820
            Width           =   1140
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
            Left            =   3960
            TabIndex        =   114
            Top             =   2400
            Visible         =   0   'False
            Width           =   2655
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
            Left            =   9120
            TabIndex        =   115
            Top             =   2100
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Policy Year"
            Height          =   255
            Index           =   1
            Left            =   6600
            TabIndex        =   113
            Top             =   2040
            Width           =   1065
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
            Left            =   9120
            TabIndex        =   112
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "LG Amt"
            Height          =   255
            Index           =   118
            Left            =   9120
            TabIndex        =   111
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "DOD"
            Height          =   255
            Index           =   117
            Left            =   9120
            TabIndex        =   110
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "DOA"
            Height          =   255
            Index           =   116
            Left            =   9120
            TabIndex        =   109
            Top             =   520
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Fol-Up"
            Height          =   255
            Index           =   39
            Left            =   9120
            TabIndex        =   108
            Top             =   1830
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "Cas Reg Date"
            Height          =   255
            Index           =   38
            Left            =   6600
            TabIndex        =   107
            Top             =   1800
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Fol. Req"
            Height          =   255
            Index           =   37
            Left            =   6600
            TabIndex        =   106
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Recur"
            Height          =   255
            Index           =   36
            Left            =   6600
            TabIndex        =   105
            Top             =   1250
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "Reserve Amt"
            Height          =   255
            Index           =   35
            Left            =   6600
            TabIndex        =   104
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "F. Treated"
            Height          =   255
            Index           =   34
            Left            =   6600
            TabIndex        =   103
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "F. Call"
            Height          =   255
            Index           =   33
            Left            =   6600
            TabIndex        =   102
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "IC Number"
            Height          =   255
            Index           =   32
            Left            =   6600
            TabIndex        =   101
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Cash Type"
            Height          =   255
            Index           =   31
            Left            =   3360
            TabIndex        =   100
            Top             =   2120
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Stay Type"
            Height          =   255
            Index           =   30
            Left            =   3360
            TabIndex        =   99
            Top             =   1875
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Fol. date"
            Height          =   255
            Index           =   29
            Left            =   6600
            TabIndex        =   98
            Top             =   2350
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Treat Mode"
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   97
            Top             =   2120
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Nat of Clm"
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   96
            Top             =   1850
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Refered By"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   95
            Top             =   1560
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Reason"
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   94
            Top             =   1080
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Hospital"
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   93
            Top             =   520
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Patient"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   825
         End
      End
      Begin MSComctlLib.ListView lstBenefitDetails 
         Height          =   5115
         Left            =   -74880
         TabIndex        =   297
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
            Charset         =   204
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
      Begin MSComctlLib.ListView lstFollowUp 
         Height          =   2640
         Left            =   -74880
         TabIndex        =   298
         Top             =   3150
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FollowUp No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "LG Amt"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Issued By"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Issued Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstLGListing 
         Height          =   2760
         Left            =   -74880
         TabIndex        =   299
         Top             =   350
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4868
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
         NumItems        =   12
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
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "LG Amt"
            Object.Width           =   2646
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
            Text            =   "LG User Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "LG System Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Purpose"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorksheet 
         Height          =   4725
         Left            =   -74880
         TabIndex        =   300
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8334
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
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Claim No"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "WS Reg Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amnt B4 GST"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "GST"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amnt wt GST"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "App Amt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Pay2HOS"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PaytoMBM"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Bordx No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Bordx Date"
            Object.Width           =   2117
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
      Begin MSComctlLib.ListView lstINVListing 
         Height          =   5400
         Left            =   -74880
         TabIndex        =   301
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9525
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Claim No"
            Object.Width           =   1764
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
      Begin MSComctlLib.ListView lstHOSPayment 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   302
         Top             =   645
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2355
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Version"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PV No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PV Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cheq Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cheq No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cheque Amt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Pay Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PV Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Remarks"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBMPayment 
         Height          =   1320
         Left            =   -74880
         TabIndex        =   303
         Top             =   2295
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2328
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Version"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PV No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PV Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cheq Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cheq No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cheq Amt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "PAY Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PV Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Remarks"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstReissuedMBM 
         Height          =   1320
         Left            =   -74880
         TabIndex        =   304
         Top             =   3855
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2328
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Version"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PV No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PV Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cheq Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cheq No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cheq Amt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "PAY Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PV Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Remarks"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstRecFRMBM 
         Height          =   1335
         Left            =   -69120
         TabIndex        =   322
         Top             =   645
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2355
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Version"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Rec No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rec Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cheq No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Rec Amt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Remarks"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstRecMNRB 
         Height          =   1320
         Left            =   -69120
         TabIndex        =   323
         Top             =   3855
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2328
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Version"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Rec No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rec Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cheq No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Rec Amt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Remarks"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstRecPayor 
         Height          =   1320
         Left            =   -69120
         TabIndex        =   324
         Top             =   2295
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2328
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Version"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Rec No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rec Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cheq No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Rec Amt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Remarks"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lstReminder 
         Height          =   5400
         Left            =   -74880
         TabIndex        =   328
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9525
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "REMNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CASNumber"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Version"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "RMF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Orig Inv."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Medical Form"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Proposal Form"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Ref Letter"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Orig Receipt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Student ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Police Report"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Subrogation From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Others 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Others 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "REMDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "REMRegUser"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "REMRegDate"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstPreUndIllness 
         Height          =   5400
         Left            =   -74880
         TabIndex        =   329
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9525
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MBMNumber"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CASNumber"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Admission Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pre. Illness Date"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Pre. Illness "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Pre. Illness Inf. By"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Und. Illness Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Und. Illness"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Und. Illness Inf. By"
            Object.Width           =   2734
         EndProperty
      End
      Begin MSComctlLib.ListView lstBenefits 
         Height          =   5400
         Left            =   -74880
         TabIndex        =   330
         Top             =   360
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   9525
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
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Dis"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid Fgrid 
         Height          =   2835
         Left            =   -74880
         TabIndex        =   335
         Top             =   360
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   5001
         _Version        =   393216
         Rows            =   11
         ScrollBars      =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lstSurgProc 
         Height          =   5400
         Left            =   -74880
         TabIndex        =   336
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9525
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
            Text            =   "Payor Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Plan Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Surgical Main Category"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Surgical Sub Cat"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Perc (% ) of Max. Benefit"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView lstLTLmt 
         Height          =   5385
         Left            =   -74880
         TabIndex        =   337
         Top             =   360
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   9499
         SortKey         =   1
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
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mem No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cover ID"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "MBM Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Bal Ann Lmt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Bal LTL"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Current Claims"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total Claims"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Current Claim Amt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Total Claim Amt"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComDlg.CommonDialog CmnDlg 
         Left            =   -64080
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lstStorage 
         Height          =   5445
         Left            =   -74880
         TabIndex        =   349
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9604
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Claim No"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Carton No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File Location"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Carton Location"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Request By"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Request Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Return Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame23 
         Caption         =   "Documents"
         Height          =   2655
         Index           =   16
         Left            =   -74880
         TabIndex        =   354
         Top             =   360
         Width           =   11415
      End
      Begin VB.Frame Frame23 
         Caption         =   "Internal Documents"
         Height          =   2775
         Index           =   17
         Left            =   -74880
         TabIndex        =   355
         Top             =   3000
         Width           =   4695
         Begin VB.ListBox List2 
            Height          =   2400
            ItemData        =   "frmCaseInquiry.frx":023A
            Left            =   120
            List            =   "frmCaseInquiry.frx":023C
            TabIndex        =   356
            ToolTipText     =   "right click to randomize contents"
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   93
         Left            =   -74760
         TabIndex        =   334
         Top             =   480
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Receipt from Member"
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
         Index           =   89
         Left            =   -69240
         TabIndex        =   327
         Top             =   420
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Receipt from Payor"
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
         Index           =   90
         Left            =   -69090
         TabIndex        =   326
         Top             =   2055
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Receipt from MNRB"
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
         Index           =   91
         Left            =   -69090
         TabIndex        =   325
         Top             =   3615
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Payment to Hospital"
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
         Index           =   86
         Left            =   -74850
         TabIndex        =   316
         Top             =   420
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Payment to Member"
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
         Index           =   87
         Left            =   -74850
         TabIndex        =   315
         Top             =   2055
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Reissued to Member"
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
         Index           =   88
         Left            =   -74850
         TabIndex        =   314
         Top             =   3615
         Width           =   2220
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
         TabIndex        =   313
         Top             =   2640
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
         TabIndex        =   312
         Top             =   2640
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
         TabIndex        =   311
         Top             =   2640
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
         TabIndex        =   310
         Top             =   2640
         Width           =   135
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
         TabIndex        =   309
         Top             =   2640
         Width           =   135
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
         Left            =   -65040
         TabIndex        =   308
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
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
         Index           =   84
         Left            =   -74760
         TabIndex        =   307
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks"
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
         Index           =   0
         Left            =   -73680
         TabIndex        =   306
         Top             =   480
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Person Incharge"
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
         Index           =   85
         Left            =   -67320
         TabIndex        =   305
         Top             =   480
         Width           =   1620
      End
   End
   Begin Crystal.CrystalReport crptSummary 
      Left            =   360
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame23 
      Height          =   555
      Index           =   24
      Left            =   120
      TabIndex        =   317
      Top             =   7120
      Width           =   11655
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10620
         TabIndex        =   320
         Top             =   160
         Width           =   885
      End
      Begin VB.CommandButton cmdPrint2 
         Caption         =   "Print &Mgmt Note"
         Height          =   300
         Left            =   9310
         TabIndex        =   319
         Top             =   160
         Width           =   1305
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print Case"
         Height          =   300
         Left            =   8335
         TabIndex        =   318
         Top             =   160
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Case Inquiry"
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
      Index           =   123
      Left            =   0
      TabIndex        =   321
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmCaseInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intLastPoz2 As Integer
Dim ScanCode As String
Public CaseInquiryCriteria As String
Dim strAppend As String
Dim PatientName As String
Dim txtSRVPath As String
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim rng As Range
Dim sqlstrCAS As String
Dim CASSearch As New ADODB.Recordset
Dim txtHLTCode As String
Dim Suplncode As String
Dim RB As String
Dim MBMCPLNCode As String
Dim MMASurgFees As Double
Dim MMAAnaeFees As Double
Dim MMASurgFees13th As Double
Dim MMAAnaeFees13th As Double
Dim MMASurgFees13thMin As Double
Dim PrintNotesStatus As String
Dim txtPayCode As String
Dim MMARateCal As Integer
Dim PictureFileName As String
Dim tmpPolicyWording As String
Dim tmpPrincipalDOB As String
Dim TmpExclussion As String
Private WithEvents cFileCopy As clsFileCopy
Attribute cFileCopy.VB_VarHelpID = -1
Private Type PointAPI
    X  As Long
    Y  As Long
End Type

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

Private Sub cboPatientList_Change()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cnn As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

   If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
    
        If Mid(cboPatientList, 1, 2) = "00" Then
            'strsql = "Select MBMICBcPp from MBMCrossReference where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Inquiry"
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
                
                If MBM.EOF Then
                    MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    txtPatientIC = MBM!MBMIcBcPp
                End If
            MBM.Close
            Set MBM = Nothing
            
        ElseIf Mid(cboPatientList, 1, 2) <> "00" Then
            'strsqlCOV = "Select MBMCICBCPPNo,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            'MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            Set cnn.ActiveConnection = medixmasterconn
            cnn.CommandText = "Case_Inquiry"
            cnn.CommandType = adCmdStoredProc
            cnn.CommandTimeout = 300
            Set Prm1 = cnn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cnn.Parameters.Append Prm1
            Set Prm2 = cnn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cnn.Parameters.Append Prm2
            MBMCov.CursorLocation = adUseClient
            MBMCov.CursorType = adOpenForwardOnly
            MBMCov.CacheSize = 100
            Set MBMCov = cnn.Execute
            
                If Not MBMCov.EOF Then
                    txtPatientIC = MBMCov!MBMCICBCPPNo
                    cboSuppExc = cboPatientList
                    If Len(MBMCov!MBMCPLNCode) Then
                        MBMCPLNCode = MBMCov!MBMCPLNCode
                    End If
                End If
            MBMCov.Close
            Set MBMCov = Nothing
        End If
    End If
End Sub

Private Sub cboPatientList_GotFocus()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cnn As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

   If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
    
        If Mid(cboPatientList, 1, 2) = "00" Then
            'strsql = "Select MBMICBcPp from MBMCrossReference where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            'MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Inquiry"
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
                
                If MBM.EOF Then
                    MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    txtPatientIC = MBM!MBMIcBcPp
                End If
            MBM.Close
            Set MBM = Nothing
            
        ElseIf Mid(cboPatientList, 1, 2) <> "00" Then
            'strsqlCOV = "Select MBMCICBCPPNo,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            'MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
            Set cnn.ActiveConnection = medixmasterconn
            cnn.CommandText = "Case_Inquiry"
            cnn.CommandType = adCmdStoredProc
            cnn.CommandTimeout = 300
            Set Prm1 = cnn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cnn.Parameters.Append Prm1
            Set Prm2 = cnn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cnn.Parameters.Append Prm2
            MBMCov.CursorLocation = adUseClient
            MBMCov.CursorType = adOpenForwardOnly
            MBMCov.CacheSize = 100
            Set MBMCov = cnn.Execute
            
                If Not MBMCov.EOF Then
                    txtPatientIC = MBMCov!MBMCICBCPPNo
                    cboSuppExc = cboPatientList
                End If
            MBMCov.Close
            Set MBMCov = Nothing
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
    CmnDlg.Filename = ""
    CmnDlg.CancelError = True
    On Error GoTo ErrHandler
    CmnDlg.InitDir = DocPath & "FaxDocuments"
    CmnDlg.Flags = cdlOFNHideReadOnly
    CmnDlg.Filter = "All files|*.*"
    'Show the open file window
    CmnDlg.ShowOpen
    'Put the file bath in the source text
    TxtSrc.Text = CmnDlg.Filename
    tmpString = DocPath & "FaxDocuments\"
    cnt = Len(tmpString)
    TxtDest.Text = DocPath & "ScannedDoc\" & txtCaseNo & "" & Mid(TxtSrc.Text, cnt, 100) & ""
ErrHandler:
  Exit Sub

End Sub

Public Function CreatePath(NewPath) As Boolean
    'Add a trailing slash if none
    If Right(NewPath, 1) <> "\" Then
        NewPath = NewPath & "\"
    End If
    
    'Call API
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        'No errors, return True
        CreatePath = True
    End If

End Function

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
    cFileCopy.Verify = (chkVerify.Value = vbChecked)
        
    'Show percentage
    'LblPercent.Caption = "0"
    
    If Not cFileCopy.copyfile(ErrCode, ErrMsg) Then
        GoTo CopyErr
    End If
    
    'Hide progress bar
    'ProgressBar1.Visible = False
    
    'Skip the error message and exit sub
    Call DestroyFile(TxtSrc)
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
End Sub
Private Sub copyfile()
On Error GoTo CopyErr

Dim ErrCode As Long
Dim ErrMsg As String

Set cFileCopy = New clsFileCopy

'Disable the copy button
CmdCopy.Enabled = False

'The source file the you want to copy
cFileCopy.SourceFile = TxtSrc
'The destination file name
cFileCopy.TargetFile = TxtDest


cFileCopy.ByteSize = 4096 '4kb

'Check for verify
cFileCopy.Verify = (chkVerify.Value = vbChecked)
    
'Show percentage

If Not cFileCopy.copyfile(ErrCode, ErrMsg) Then
    GoTo CopyErr
End If

'Hide progress bar

'Skip the error message and exit sub
MsgBox "File copy successful..."
GoTo Ex

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

Private Sub cmdMBMEnq_Click()
    If txtMBMNumber <> "" Then
        frmNewMBMEnquiry.txtMBMNumber = txtMBMNumber
        frmNewMBMEnquiry.cmdShow_Click
    End If
End Sub



Private Sub cmdPrev_Click()
    Dim CAS As New ADODB.Recordset
    Dim CASStatus As Boolean
    Dim CasNumber  As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    Dim strOrderby As String
    

    If Trim(txtCaseNo) <> "" Then
      '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        'sql = " select CASNumber from CASCrossReference where CASnumber < '" & RemStr(txtCaseNo) & "' " & _
            "Order by CASNumber DESC "
        CAS.MaxRecords = 1
        'CAS.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        strAppendCase = "2"
        strAppend = " AND CASnumber < '" & RemStr(txtCaseNo) & "'"
        strOrderby = "Order by CASNumber DESC"
        strAppend = strAppend + strOrderby
        Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "Case_Inquiry"
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
        
        If Not CAS.EOF Then
            CASStatus = True
            CasNumber = CAS!CasNumber
        Else
            CASStatus = False
        End If
        
        CAS.Close
        Set CAS = Nothing
     '   medixmasterconnMaint.Close
     '   Set medixmasterconnMaint = Nothing

        If CASStatus = True Then
            txtCaseNo.Text = CasNumber
            cmdShow_Click
        Else
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Case Number Format! ", vbCritical + vbOKOnly
        End If
    End If
    
End Sub

Private Sub CmdNext_Click()
    Dim CAS As New ADODB.Recordset
    Dim CASStatus As Boolean
    Dim CasNumber As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    Dim strOrderby As String
    

'    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    'sql = " select CASNumber from CASCrossReference where CASNumber > '" & RemStr(txtCaseNo) & "' " & _
        " Order by CASnumber "
    'CAS.MaxRecords = 1
    'CAS.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    strAppendCase = "2"
    strAppend = " AND CASNumber > '" & RemStr(txtCaseNo) & "'"
    strOrderby = "Order by CASnumber"
    strAppend = strAppend + strOrderby
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "Case_Inquiry"
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
        
    If Not CAS.EOF Then
        CASStatus = True
        CasNumber = CAS!CasNumber
    Else
        CASStatus = False
    End If
    
    CAS.Close
    Set CAS = Nothing
'    medixmasterconnMaint.Close
 '   Set medixmasterconnMaint = Nothing

    If CASStatus = True Then
        txtCaseNo.Text = CasNumber
        cmdShow_Click
    Else
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
    End If
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    If txtCaseNo <> "" Then
        Call funcword2
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = Default
End Sub

Private Sub cmdPrint2_Click()
    Screen.MousePointer = vbHourglass
    If txtCaseNo <> "" Then
        PrintNotesStatus = "M"
        Call funcword
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = Default
End Sub

Private Sub cboQuery_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboQuery, cboQuery.SelStart) + Chr(KeyAscii) + Mid(cboQuery, cboQuery.SelStart + cboQuery.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboQuery.ListCount - 1
        
        If NewText = LCase(Left(cboQuery.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboQuery.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboQuery = ValidValue
               cboQuery.SelStart = 0
               cboQuery.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cmdCASPrintIntNote_Click()
    Screen.MousePointer = vbHourglass
        
    If txtCaseNo <> "" Then
        PrintNotesStatus = "I"
        Call funcword
        
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
              
    Screen.MousePointer = Default
    
End Sub




Private Sub cmdPrintPymt_Click()
    Screen.MousePointer = vbHourglass
        
    If txtCaseNo <> "" Then
        Call funcwordPymt
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
              
    Screen.MousePointer = Default
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
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

Private Sub CmdExit_Click()
    Unload Me
    Set frmCaseInquiry = Nothing
End Sub

Private Sub CmdSearch_Click()
    frmCASList.Source = 4
    frmCASList.Show
End Sub

Private Sub ComboListing()
    cboStatus1.AddItem "O" & " - " & "Open"
    cboStatus1.AddItem "S" & " - " & "Suspend"
    cboStatus1.AddItem "C" & " - " & "Close"
    cboStatus1.AddItem "D" & " - " & "Delete"
    cboStatus1.AddItem "R" & " - " & "Repudiation"
    
    'CboClaimability.AddItem "A" & " - " & "Accepted"
    'CboClaimability.AddItem "R" & " - " & "Repudiation"
    'CboClaimability.AddItem "C" & " - " & "Cancelled"
    'CboClaimability.AddItem "S" & " - " & "Special Approval"
    'CboClaimability.AddItem "P" & " - " & "Pay & Claim"
    'CboClaimability.AddItem "N" & " - " & "Not Decided"
   
    cboCondwillOcc.AddItem "Y - YES"
    cboCondwillOcc.AddItem "N - NO"
    
    cboFLWTreatment.AddItem "Y - YES"
    cboFLWTreatment.AddItem "N - NO"

    cboPending.AddItem "Y - Yes"
    cboPending.AddItem "N - No"

End Sub
                 
Public Sub cmdShow_Click()
'On Error GoTo ErrMsg
Dim getRecord As Integer
Dim CaseNo As String
    
    SSTab1.Tab = 0
    strAppend = ""
    Screen.MousePointer = vbHourglass
    lblFaxDocStatus.Visible = False
    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
    Else
        'txtCaseNo = Trim(txtCaseNo)
        ScanCode = Trim(txtCaseNo)
        If Mid(ScanCode, 1, 2) = "LO" Then
            txtCaseNo = ""
            Call SearchForm(ScanCode)
        Else
           ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
          '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
           ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
            CaseNo = txtCaseNo
            
            '=============================
            strAppend = strAppend + " AND CASNumber = '" & Trim(txtCaseNo) & "'"
            CaseInquiryCriteria = strAppend
            '=============================
            
            Call ClearEntries
            Call ClearForm(Me)
            txtCaseNo = Trim(CaseNo)
            cboPatientList.Clear
            MMARateCal = 0
            getRecord = showCAS
            Call CheckFiles
            MMARateCal = 1
          '  medixmasterconn.Close
         '   Set medixmasterconn = Nothing
          '  medixmasterconnWS.Close
          ''  Set medixmasterconnWS = Nothing
          '  medixmasterconnMaint.Close
          '  Set medixmasterconnMaint = Nothing
        End If
        'End If
    End If
    Screen.MousePointer = Default
    Exit Sub

'ErrMsg:
'    MsgBox Err.Description
'    Screen.MousePointer = vbDefault

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
Dim sqlstr As String
Dim CAS As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim strsqlHOS As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim StrSqlMsgProm As String
Dim MsgProm As New ADODB.Recordset

    sqlstrCAS = "Exec SearchCasAdjustment '" & Trim(txtCaseNo) & "'"
    
    Set CASSearch = New ADODB.Recordset
    CASSearch.Open sqlstrCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If Not CASSearch.EOF Then
    
        showCAS = 1
        With CASSearch
            txtCaseNo = !CasNumber
            txtMBMNumber = Trim(!MBMNumber)
            If !MBMPatientCovID = "00" Then
                cboPatientList = Trim(!MBMPatientCovID) & " - " & Trim(!MBMName)
                PatientName = Trim(!MBMName)
            Else
                cboPatientList = Trim(!MBMPatientCovID) & " - " & Trim(!MBMCName)
                cboSuppExc = Trim(!MBMPatientCovID) & " - " & Trim(!MBMCName)
                PatientName = Trim(!MBMCName)
            End If
        
            txtDateAdmitted.mask = ""
            txtDateAdmitted = Format(!CASDateAdmitted, "dd/mm/yyyy")
            If Len(!ReferedBy) Then
                txtREFBy = !ReferedBy
            End If
            If Len(!CASLGAmt) Then
                txtCASLGAmt = !CASLGAmt
            End If
            If Len(!CASReserveAmount) Then
                txtCASReserveAmount = !CASReserveAmount
            End If
            If Len(!CASDischargeDate) Then
                txtDateDischarge.mask = ""
                txtDateDischarge.mask = "##/##/####"
                txtDateDischarge = Format(!CASDischargeDate, "dd/mm/yyyy")
            End If
            If !CASQuery <> "" Then
                cboQuery = !CASQuery
            End If
                        
            
            'strsqlHOS = "Select HOSCode,HOSName from HOS where HOSCode = '" & Trim(!CASHOSCode) & "'"
            'HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppendCase = "3"
            strAppend = " AND HOSCode = '" & Trim(!CASHOSCode) & "'"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_Inquiry"
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
            
            If Not HOS.EOF Then
                 If Len(!CASHOSCode) Then
                     txtHOSCode = HOS!HosName
                 End If
             End If
            
            If Len(!CASStayType) Then
                CboStayType = !CASStayType
            End If
            If Len(!CASPAYType) Then
                cboPayType = !CASPAYType
            End If
            If Len(!ModalityCode) Then
                cboModality = !ModalityCode
            End If
            If Len(!ClaimNatureCode) Then
                cboNature = !ClaimNatureCode
            End If
            If Len(!CASAdmissionReason) Then
                txtReason = !CASAdmissionReason
            End If
            
            cboStatus1.AddItem "O" & " - " & "Open"

            If Trim(!CASStatus) = "O" Then
                cboStatus1.Text = "O" & " - " & "Open"
            ElseIf Trim(!CASStatus) = "S" Then
                cboStatus1.Text = "S" & " - " & "Suspend"
            ElseIf Trim(!CASStatus) = "C" Then
                cboStatus1.Text = "C" & " - " & "Close"
            ElseIf Trim(!CASStatus) = "D" Then
                cboStatus1.Text = "D" & " - " & "Delete"
            ElseIf Trim(!CASStatus) = "R" Then
               cboStatus1.Text = "R" & " - " & "Repudiation"
            End If
            
            If Trim(!CASClaimability) = "AA" Then
                CboClaimability.Text = "AA - Accepted"
             ElseIf Trim(!CASClaimability) = "RR" Then
                CboClaimability.Text = "RR - Rejected"
             ElseIf Trim(!CASClaimability) = "CC" Then
                CboClaimability.Text = "CC - Cancelled"
             ElseIf Trim(!CASClaimability) = "SS" Then
                CboClaimability.Text = "SA - Special Approval"
             ElseIf Trim(!CASClaimability) = "PP" Then
                CboClaimability.Text = "PP - Pay & Claim"
             ElseIf Trim(!CASClaimability) = "NN" Then
                CboClaimability.Text = "NN - Not Decided"
             ElseIf Trim(!CASClaimability) = "NA" Then
                CboClaimability.Text = "NA - Not Decided/Accepted"
             ElseIf Trim(!CASClaimability) = "NR" Then
                CboClaimability.Text = "NR - Not Decided/Rejected"
            ElseIf Trim(!CASClaimability) = "PA" Then
                CboClaimability.Text = "PA - Pay&Claim/Accepted"
            ElseIf Trim(!CASClaimability) = "PR" Then
                CboClaimability.Text = "PR - Pay&Claim/Rejected"
            ElseIf Trim(!CASClaimability) = "AR" Then
                CboClaimability.Text = "AR - Accepted/Rejected"
            ElseIf Trim(!CASClaimability) = "RA" Then
                CboClaimability.Text = "RA - Rejected/Accepted"
            ElseIf Trim(!CASClaimability) = "ER" Then
                CboClaimability.Text = "ER - Expired/Rejected"
            ElseIf Trim(!CASClaimability) = "AI" Then
                CboClaimability.Text = "AI - Accepted/Insurance"
            End If
            
            If Trim(!CASPending) = "Y" Then
                cboPending.Text = "Y" & " - " & "Yes"
            ElseIf Trim(!CASPending) = "N" Then
                cboPending.Text = "N" & " - " & "No"
            End If
            If Len(!CASInd) Then
                cboCASInd = !CASInd
            End If
            If Len(!LINKCASNo) Then
                TxtLinkCASNo = !LINKCASNo
            End If
            If Len(!CASTopupLmt) Then
                cboTopUpLmt = !CASTopupLmt
            End If
            If Len(!CaseRegDate) Then
                txtREGDate = Format(!CaseRegDate, "dd/mm/yyyy")
            End If
            
            txtCASTakeOverAmt = IIf(Len(!CASTakeOverAmt), !CASTakeOverAmt, "")
            
            cboCASTakeOveInd = IIf(Len(!CASTakeOverInd), !CASTakeOverInd, "")
            
            txtCASTOPlan = IIf(Len(!CASTakeOverPlan), !CASTakeOverPlan, "")
            
            If Len(!CASLGTime) Then
                txtLGTime = !CASLGTime
            End If
            txtMBMPolYr = IIf(Len(!CASPolYear), !CASPolYear, "")

        End With
                Call showMBM
                Call CallLimit
                Call showCASRemarks
                Call showOther
                Call showOther2
                Call ShowFirstDiag
                Call ShowFinalDiag
                Call ShowUD
                'Call GetAccountSummary
                Call ShowRepudiation
                'Call ShowExclusion
                If Trim(txtHLTCode) = "S" Then
                    StrSqlMsgProm = "Select VIPInd,PLNSpecialNote,PLNPolicyWording from PLN where PLNCode = '" & Trim(txtPLNCode) & "'"
                Else
                    StrSqlMsgProm = "Select VIPInd,PLNSpecialNote,PLNPolicyWording from PLN where PLNCode = '" & Trim(txtPLNCode) & "' And PAYCode = '" & Trim(txtPayCode) & "' "
                End If
                MsgProm.Open StrSqlMsgProm, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                    Do While Not MsgProm.EOF
                        tmpPolicyWording = IIf(Len(MsgProm!PLNPolicyWording), MsgProm!PLNPolicyWording, "")
                        If Mid(MsgProm!VIPInd, 1, 1) = "Y" Then
                            MsgBox MsgProm!PLNSpecialNote, vbOKOnly + vbInformation
                        End If
                    MsgProm.MoveNext
                    Loop
                MsgProm.Close
                Set MsgProm = Nothing
            'If txtUpgPLNCode <> "" Or txtDOC <> "__/__/____" Then
            '    Call NewBNFListing
            'End If
    Else
        MsgBox "No record found!", vbCritical + vbOKOnly
    End If
    CASSearch.Close
    Set CASSearch = Nothing
End Function

Public Sub CallLimit()
Dim AnnualLimit As String
Dim PLN As String
Dim INS As String
Dim HLT As String
Dim MBM As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
    
    'sqlstr = "Select PLNCode, INSCode, HLTCode from MBM Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "4"
    strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Inquiry"
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
        
        PLN = Mid(MBM!PLNCode, 1, 2)
        INS = Mid(MBM!INSCode, 1, 1)
        HLT = Mid(MBM!HLTCode, 1, 1)
            
        MBM.Close
        Set MBM = Nothing
        'txtAnnualLimit = ""
        'txtAnnualLimit = Format(GetAnnualLimit(INS, PLN, HLT, txtMBMEffDate), "#,##0.00")

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
Dim MMASurgFees As Double
Dim MMAAnaeFees As Double
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
        'Do Until Other.EOF
        With CASSearch
            If Len(Trim(!CASOCondWillRecur)) Then
                'If Len(Trim(!CASOCondWillRecur)) = "Y" Then
                '    optCONwilloccY.value = True
                'Else
                '    optCONwilloccN.value = False
                'End If
                 cboCondwillOcc = !CASOCondWillRecur
            End If
            If Len(Trim(!CASOFollowUpRequired)) Then
                'If Len(Trim(!CASOFollowUpRequired)) = "Y" Then
                '    optFOLreqY.value = True
                'Else
                '    optFOLreqN.value = False
                'End If
                 cboFLWTreatment = !CASOFollowUpRequired
            End If
            If Len(Trim(!CASFolUp)) Then
                 cboFolUp = !CASFolUp
            End If
            
            If Len(!CASFirstCall) Then
                txtFirstCall = Format(!CASFirstCall, "dd/mm/yyyy")
            End If
            
            If Len(!CASFirstTreated) Then
                txtFirstTreated = Format(!CASFirstTreated, "dd/mm/yyyy")
            End If
            'If Len(!CASFCallTime) Then
            '    txtFCTime = Format(!CASFCallTime, "hh:mm")
            'End If
            
            'If Len(!CASFTreatedTime) Then
            '     txtFTTime = Format(!CASFTreatedTime, "hh:mm")
            'End If
            
            'If Len(!CASDOATime) Then
            '     txtDOATime = Format(!CASDOATime, "hh:mm")
            'End If
            
            'If Len(!CASDODTime) Then
            '     txtDODTime = Format(!CASDODTime, "hh:mm")
            'End If

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
            
            'If Len(!CASOMedOnDis1) Then
            '     txtMOD1 = !CASOMedOnDis1
            'End If

            'If Len(!CASOMedOnDis2) Then
            '     txtMOD2 = !CASOMedOnDis2
            'End If
            
            'If Len(!CASOMedOnDis3) Then
            '     txtMOD3 = !CASOMedOnDis3
            'End If
            
            'If Len(!CASOReccurence) Then
            '    cboRecurrence = !CASOReccurence
            'End If
            
            'If Len(!CASDRD) Then
            '    txtDateRECDoc = Format(!CASDRD, "dd/mm/yyyy")
            'End If
            If Len(!CASRedAmt) Then
                txtREDAmt = !CASRedAmt
            End If
            
            
            'If Len(!CASOPROCRequired1) Then
            '    txtProcedure1 = !CASOPROCRequired1
            'End If
            
            'If Len(!CASOPROCRequired2) Then
            '   txtProcedure1 = !CASOPROCRequired2
            'End If

            'If Len(!CASOPROCRequired3) Then
            '     txtProcedure1 = !CASOPROCRequired3
            'End If
            
            'If Len(!CASOPROCRequired4) Then
            '   txtProcedure1 = !CASOPROCRequired4
            'End If

            'If Len(!CASOPROCRequired5) Then
            '     txtProcedure1 = !CASOPROCRequired5
            'End If
            
            If Len(!CASCptCode1) Then
                txtProcCode1 = !CASCptCode1
            End If
            
            If Len(!CASCptCode2) Then
                txtProcCode2 = !CASCptCode2
            End If
            
            If Len(!CASCptCode3) Then
                txtProcCode3 = !CASCptCode3
            End If
            
            If Len(!CASCptCode4) Then
                txtProcCode4 = !CASCptCode4
            End If
            
            If Len(!CASCptCode5) Then
                txtProcCode5 = !CASCptCode5
            End If
            
            If Len(!CASCptCode6) Then
                txtProcCode6 = !CASCptCode6
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
                txtNSProcedure1 = !CASONonSurgPROCRequired1
            End If
            
            If Len(!CASONonSurgPROCRequired2) Then
               txtNSProcedure2 = !CASONonSurgPROCRequired2
            End If

            If Len(!CASONonSurgPROCRequired3) Then
                 txtNSProcedure3 = !CASONonSurgPROCRequired3
            End If

            If Len(!CASONonSurgPROCRequired4) Then
               txtNSProcedure4 = !CASONonSurgPROCRequired4
            End If

            If Len(!CASONonSurgPROCRequired5) Then
                 txtNSProcedure5 = !CASONonSurgPROCRequired5
            End If
            
            If Len(!CASONonSurgPROCRequired6) Then
                 txtNSProcedure6 = !CASONonSurgPROCRequired6
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
                txtATTDoc(0) = !CASAttendDoctor1
            End If
            If Len(!CASAttendDoctor2) Then
                txtATTDoc(1) = !CASAttendDoctor2
            End If
            If Len(!CASAttendDoctor3) Then
                txtATTDoc(2) = !CASAttendDoctor3
            End If
            If Len(!CASAttendDoctor4) Then
                txtATTDoc(3) = !CASAttendDoctor4
            End If
            If Len(!CASAttendDoctor5) Then
                txtATTDoc(4) = !CASAttendDoctor5
            End If
            If Len(!CASAnaesName1) Then
                txtAnaesDoc(0) = !CASAnaesName1
            End If
            If Len(!CASAnaesName2) Then
                txtAnaesDoc(1) = !CASAnaesName2
            End If
            If Len(!CASAnaesName3) Then
                txtAnaesDoc(2) = !CASAnaesName3
            End If
            
            If Len(!CASFirstDocDate) Then
                txtFirstDocDate = !CASFirstDocDate
            End If
            If Len(!CASSubSeqDocDate) Then
                txtSubDocDate = !CASSubSeqDocDate
            End If
            cboIntExtTO = IIf(Len(!CASIntExtTO), !CASIntExtTO, "")
            
            txtHOSName = IIf(Len(!CASHOSName), !CASHOSName, "")
            'If Len(!CASDOATime) Then
            '    txtDOATime = Format(!CASDOATime, "HH:MM")
            'End If
            'If Len(!CASDODTime) Then
            '    txtDODTime = Format(!CASDODTime, "HH:MM")
            'End If
            'If Len(!CASDOALGTime) Then
            '    txtDOALGTime = Format(!CASDOALGTime, "HH:MM")
            'End If
            'If Len(!CASDODLGTime) Then
            '    txtDODLGTime = Format(!CASDODLGTime, "HH:MM")
            'End If
            'If Len(!CASDOAComment) Then
            '    txtDOAComment = !CASDOAComment
            'End If
            'If Len(!CASDODComment) Then
            '    txtDODComment = !CASDODComment
            'End If
            'If Len(!CASComplaintDate) Then
            '    'if format(!CASComplaintDate, "HH:MM") = "00:00"
            '    txtComplaintDate = !CASComplaintDate
            'End If
            'If Len(!CASComplaintIssue) Then
            '    txtComplaintIssue = !CASComplaintIssue
            'End If
            'If Len(!CASComplainant) Then
            '    txtComplain = !CASComplainant
            'End If
            'If Len(!CASCompAction) Then
            '    txtCompAction = !CASCompAction
            'End If
            
            If Len(!CASAuditReview) Then
                txtPrincipalExc = Replace(!CASAuditReview, "~", vbNewLine)
            End If
            
            
        End With
        
    End If
End Function

Private Sub cboStatus1_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboStatus1.Text, cboStatus1.SelStart) + Chr(KeyAscii) + Mid(cboStatus1.Text, cboStatus1.SelStart + cboStatus1.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboStatus1.ListCount - 1
        
        If NewText = LCase(Left(cboStatus1.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboStatus1.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
            If ValidCount = 1 Then
                cboStatus1.Text = ValidValue
                cboStatus1.SelStart = 0
                cboStatus1.SelLength = Len(ValidValue)
            End If
        
        End If
End Sub

Private Sub ShowFinalDiag()
    Dim diafinal As New ADODB.Recordset
    Dim strsql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
    'strsql = "select * from DIAFinal where CASNumber = '" & txtCaseNo & "'"
    'diafinal.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "5"
    strAppend = " AND CASNumber = '" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Inquiry"
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
        
        If Not diafinal.EOF Then
            Do Until diafinal.EOF
                With diafinal
                    If Len(!ICDCode1) Then
                        cboFDICDCode1 = !ICDCode1
                    End If
                    If Len(!DIADOS1) Then
                        txtFDDOS1 = Format(!DIADOS1, "dd/mm/yyyy")
                    End If
                    If Len(!DIAFinalDiag1) Then
                        txtFinalDiag1(0) = !DIAFinalDiag1
                    End If
                    If Len(!ICDCode2) Then
                        cboFDICDCode2 = !ICDCode2
                    End If
                    If Len(!DIADOS2) Then
                        txtFDDOS2 = Format(!DIADOS2, "dd/mm/yyyy")
                    End If
                    If Len(!DIAFinalDiag2) Then
                        txtFinalDiag1(1) = !DIAFinalDiag2
                    End If
                    If Len(!ICDCode3) Then
                        cboFDICDCode3 = !ICDCode3
                    End If
                    If Len(!DIADOS3) Then
                        txtFDDOS3 = Format(!DIADOS3, "dd/mm/yyyy")
                    End If
                    If Len(!DIAFinalDiag3) Then
                        txtFinalDiag1(2) = !DIAFinalDiag3
                    End If
        
                End With
                diafinal.MoveNext
            Loop
            End If
    diafinal.Close
    Set diafinal = Nothing
End Sub

Private Sub ShowFirstDiag()
    Dim diafirst As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String

    'strsql = "Select casnumber , * from DIAFirst where CasNumber = '" & txtCaseNo & "'"
    'diafirst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "2"
    strAppend = " AND CasNumber = '" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Inquiry"
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
        
            Do Until diafirst.EOF
            With diafirst
                If Len(!ICDCode1) Then
                    cboICDCode1 = !ICDCode1
                End If
                If Len(!DIADOS1) Then
                    txtPDDOS1 = Format(!DIADOS1, "dd/mm/yyyy")
                End If
                If Len(!DIAPrelimDiag1) Then
                    txtPrelimDiag1(0) = !DIAPrelimDiag1
                End If
                If Len(!ICDCode2) Then
                    cboICDCode2 = !ICDCode2
                End If
                If Len(!DIADOS2) Then
                    txtPDDOS2 = Format(!DIADOS2, "dd/mm/yyyy")
                End If
                If Len(!DIAPrelimDiag2) Then
                    txtPrelimDiag1(1) = !DIAPrelimDiag2
                End If
                If Len(!ICDCode3) Then
                    cboICDCode3 = !ICDCode3
                End If
                If Len(!DIADOS3) Then
                    txtPDDOS3 = Format(!DIADOS3, "dd/mm/yyyy")
                End If
                If Len(!DIAPrelimDiag3) Then
                    txtPrelimDiag1(2) = !DIAPrelimDiag3
                End If
            End With
            diafirst.MoveNext
            Loop
    diafirst.Close
    Set diafirst = Nothing
End Sub

Private Function showMBM()
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim sqlstr1, sqlPolStatus   As String
Dim PolStatus As New ADODB.Recordset
Dim ValueDate As Date

    'Call SetSupp(False)
    'txtCASReserveAmount = 2000
    'cmdGenerate.Enabled = True
    'txtCaseNo = "0"
    'txtFirstAuthoriseCode = "0"
    Screen.MousePointer = vbHourglass
     txtMBMName.ForeColor = vbBlack
    'sqlstr = " Select MBMNumber, HLTCode, MBMPolicyNo,  MBMName ," & _
    '        " MBMIcBcPp,  MBMMemberType , " & _
    '        " MBMStatus, PAYCode, AGECode , " & _
    '        " INSCode, PLNCode, " & _
    '        " MBMAddress1, MBMAddress2, " & _
    '        " MBMAddress3, MBMPostCode, CTYCode, STACode,  " & _
     '       " MBMDOB , BRNCode, MBMTakeOver ,  MBMAvailableLimit,   " & _
     '       " MBMPayorEffDate, MBMPayorExpDate, MBMEndtEffDate "
    'sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    '
    'sqlstr1 = " select MBMCNumber, MBMCCoverID, MBMCName, " & _
    '         " MBMCRELCode, MBMCStatus, MBMCAnnualLimit " & _
    '         " from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
             
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    'MBMCov.Open sqlstr1, medixmasterconn, adOpenKeyset, adLockOptimistic
    sqlstr = "Exec SearchMBMCas '" & Trim(txtMBMNumber) & "'"
    
    Set MBM = New ADODB.Recordset
    MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    If Not MBM.EOF Then
    
    'If MBM.recordCount = 0 Then
       'txtTotalRecord = 0
        With MBM
            If Len(!MBMName) Then
                txtMBMName = !MBMName
            End If
            If Len(!MBMPolicyNo) Then
                txtMBMPolNo = !MBMPolicyNo
            End If
            If Len(!MBMIcBcPp) Then
                txtMBMIcBcPP = !MBMIcBcPp
            End If
            If Len(!INSCode) Then
                txtINSCode1 = !INSCode
            End If
            If Len(!PLNCode) Then
                txtPLNCode = !PLNCode
            End If
            If Len(!HLTCode) Then
                txtHLTCode = !HLTCode
            End If
            If Len(!MBMCPLNCode) Then
                Suplncode = !MBMCPLNCode
            Else
                Suplncode = ""
            End If
            'If Len(Trim(!MBMTakeover)) Then
            '    txtTakeOver = !MBMTakeover
            'End If
            If Len(!MBMExclusion) Then
                TmpExclussion = !MBMExclusion
            End If
            
            If Len(Format(!MBMDOB, "dd/mm/yyyy")) Then
            '    txtDOB.mask = ""
                tmpPrincipalDOB = Trim(Format(!MBMDOB, "dd/mm/yyyy"))
            Else
                tmpPrincipalDOB = ""
            End If
            
            If Len(!PAYCode) Then
                txtPayCode = !PAYCode
            End If
                
            'txtPAYCode = !txtPAYCode
            'txtAGECode = !AgeCode
            'If Len(!MBMPayorEffDate) Then
            '    txtMBMEffDate.mask = ""
            '    txtMBMEffDate = Format(!MBMPayorEffDate, "dd/mm/yyyy")
            'End If
            'If Len(!MBMPayorEXPDate) Then
            '    txtMBMExpiryDate.mask = ""
            '    txtMBMExpiryDate = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
            'End If
            
                
            'txtAnnualLimit = Format(GetAnnualLimit(txtINSCode1, txtPLNCode, !HLTCode, CDate(!MBMPayorEffDate)), "#,##0.00")
            'txtAvailableAnnualLimit = Format(!MBMAvailableLimit, "#,##0.00")
            'txtHLTCode = !HLTCode
            'txtPolStatus = !MBMStatus
            
            'sqlPolStatus = " select pstatuscode ,pstatusDesc from PolicyStatus where PStatusCode ='" & Trim(txtPolStatus) & "'"
            'PolStatus.Open sqlPolStatus, medixmasterconn, adOpenKeyset, adLockOptimistic
            
           ' If PolStatus.recordCount Then
                'MsgBox "Membership Status = " & PolStatus!PStatusDesc & "..! ", vbOKOnly + vbCritical, DialogBoxTitle
            '    txtMBMName = txtMBMName & PolStatus!PStatusDesc
            '    txtPolStatus = PolStatus!PStatusDesc
                'txtMBMName.ForeColor = vbRed
                'cmdGenerate.Enabled = False
           ' Else
           '     MsgBox "Error in Policy Status..! ", vbOKOnly + vbCritical, DialogBoxTitle
           ' End If
            
'            PolStatus.Close
'            Set PolStatus = Nothing
    
'            If Len(Trim(!MBMCancelDate)) = 0 Then
'                txtCancellationDate = !MBMCancelDate
'            End If
            'If Len(Trim(!MBMLastAdjustmentDate)) = 0 Then
            '    txtLastAdjustment = !MBMLastAdjustmentDate
            'End If
            'cboPatientList.Clear
            'cboPatientList.AddItem "0 -" & Me.txtMBMName
            'txtHLTCOde = !HLTCode
                'If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
                '    Do While Not MBM.EOF
                '        cboPatientList.AddItem MBM!MBMCCoverID & "-" & MBM!MBMCName
                '        MBM.MoveNext
                '    Loop
                'End If
        End With
        '-----------------------------------------------------------------------------------------------
        ' To check prompt message box when member policy have expired
        'If Len(txtMBMExpiryDate) Then
        '    If DateDiff("ww", txtMBMExpiryDate, DateAdd("d", 14, Date)) >= 0 Then
        '        MsgBox "Membership has been expired (+14 days grace period..!) ", vbOKOnly + vbCritical, DialogBoxTitle
        '    End If
        'End If
        '-----------------------------------------------------------------------------------------------
               
'        ElseIf txtPolStatus = "S" Then
 '           MsgBox "Membership has been  SUSPENDED..! ", vbOKOnly + vbCritical, DialogBoxTitle
  '          txtMBMName = txtMBMName & " (SUSPENDED)"
   '         txtMBMName.ForeColor = vbRed
    '        cmdGenerate.Enabled = False
     '   ElseIf txtPolStatus = "R" Then
      '      MsgBox "Membership has been  REJECTED..! ", vbOKOnly + vbCritical, DialogBoxTitle
       '     txtMBMName = txtMBMName & " (REJECTED)"
        '    txtMBMName.ForeColor = vbRed
         '   cmdGenerate.Enabled = False
        'ElseIf txtPolStatus = "D" Then
'         '   MsgBox "Policy has been  DELETED..! ", vbOKOnly + vbCritical, DialogBoxTitle
 '           txtMBMName = txtMBMName & " (DELETED)"
  '          txtMBMName.ForeColor = vbRed
   '         cmdGenerate.Enabled = False
    '    End If
    
    ' * Start
    'If Date - CDate(txtMBMEffDate) <= 30 Then
    '    lblWaitingPeriod.Visible = True
    'Else
    '    lblWaitingPeriod.Visible = False
    'End If
    
    'If CDate(txtMBMExpiryDate) < Date Then
    '    lblExpired.Visible = True
    '    lblExpired.Caption = "Expired"
    'ElseIf CDate(txtMBMExpiryDate) - Date <= 30 Then
    '    lblExpired.Visible = True
    '    lblExpired.Caption = "< 30 Days to Expiry"
    'Else
    '    lblExpired.Visible = False
    'End If

    
    'If Len(txtMBMExpiryDate) Then
    '    If CDate(txtMBMExpiryDate) < Date Then
    '        lblExpired.Visible = True
    '    Else
    '        lblExpired.Visible = False
    '    End If
    'End If
    
    ' * End
    If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
        Do While Not MBM.EOF
            cboPatientList.AddItem MBM!MBMCCoverID & "-" & MBM!MBMCName
            cboSuppExc.AddItem MBM!MBMCCoverID & "-" & MBM!MBMCName
            MBM.MoveNext
        Loop
    End If

    Else
        MsgBox "No record found!", vbCritical

    End If
    
    Screen.MousePointer = vbDefault
   
    MBM.Close
    Set MBM = Nothing
    'MBMCov.Close
    'Set MBMCov = Nothing
End Function

Private Sub GetAccountSummary()
    Dim Acc As New ADODB.Recordset
    Dim strsql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
    'strsql = " select casnumber, totalApprove, totalPayable2Hos, " & _
            " totalPaid2Hos, totalPayable2Mbm, totalPaidByMbm , CLMFinalApprove, " & _
            " totalPaid2MBM , totalExcess , totalReimburse , CLMFinalApprove from view_account_summary where CASNumber = '" & txtCaseNo & "'"
    'Acc.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    strAppendCase = "1"
    strAppend = " AND CASNumber = '" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_Inquiry"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        Acc.CursorLocation = adUseClient
        Acc.CursorType = adOpenForwardOnly
        Acc.CacheSize = 100
        Set Acc = cmd.Execute
        
    If Not Acc.EOF Then
        With Acc
            'txtBordxAmt = Format(!totalApprove, "#,#.00")
            'txtApproveAmt = Format(!totalApprove, "#,#.00")
            'txtPayable2Hos = Format(!totalPayable2Hos, "#,#.00")
            'If IsNumeric(txtPayable2Hos) = False Then
            '    txtPayable2Hos = 0
            'End If
            'If Trim(!totalPaid2Hos) <> "" Then
            '    txtPaid2Hos = Format(!totalPaid2Hos, "#,#.00")
            'End If
            'If IsNumeric(txtPaid2Hos) = False Then
            '    txtPaid2Hos = 0
            'End If
            
            ' Out standing hos
            'txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,#.00;(#,#.00)")
            
            'txtPayableMem = Format(!totalPayable2Mbm, "#,#.00;(#,#.00)")
            'If IsNumeric(txtPayableMem) = False Then
            '    txtPayableMem = 0
            'End If
            'If Trim(!totalPaid2MBM) <> "" Then
            '    txtPaid2MBM = Format(!totalPaid2MBM, "#,#.00")
            'End If
            'If IsNumeric(txtPaid2MBM) = False Then
            '    txtPaid2MBM = 0
            'End If
            'If Trim(!totalPaidByMbm) <> "" Then
            '    txtPaidByMem = Format(!totalPaidByMbm, "#,#.00")
            'End If
            'If IsNumeric(txtPaidByMem) = False Then
            '    txtPaidByMem = 0
            'End If
            'txtMemOS = Format(CDbl(txtPayableMem) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
            'txtReimbursement = Format(!TotalReimburse, "#,#.00")
            'txtExcess = Format(!totalExcess, "#,#.00;(#,#.00)")
            'txtBordxAmt = Format(!totalApprove, "#,#.00")
        End With
    End If
    Acc.Close
    Set Acc = Nothing

End Sub

Private Sub ShowUD()
    Dim CASUD As New ADODB.Recordset
    Dim strsql As String
    
    'StrSql = "select * from CAS_UD where CASNumber = '" & txtCaseNo & "'"
    'CASUD.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
        If Not CASSearch.EOF Then
            With CASSearch
                If Len(!CASUDCode1) Then
                    cboUDICDCode1 = !CASUDCode1
                End If
                If Len(!CASUDDOS1) Then
                    txtUDDOS1 = Format(!CASUDDOS1, "dd/mm/yyyy")
                End If
                If Len(!CASUDReasons1) Then
                    txtUD1 = !CASUDReasons1
                End If
                If Len(!CASUDCode2) Then
                    cboUDICDCode2 = !CASUDCode2
                End If
                If Len(!CASUDDOS2) Then
                    txtUDDOS2 = Format(!CASUDDOS2, "dd/mm/yyyy")
                End If
                If Len(!CASUDReasons2) Then
                    txtUD2 = !CASUDReasons2
                End If
                If Len(!CASUDCode3) Then
                    cboUDICDCode3 = !CASUDCode3
                End If
                If Len(!CASUDDOS3) Then
                    txtUDDOS3 = Format(!CASUDDOS3, "dd/mm/yyyy")
                End If
                If Len(!CASUDReasons3) Then
                    txtUD3 = !CASUDReasons3
                End If
               
            End With
        End If
End Sub

Private Sub LGListing()
Dim LG As New ADODB.Recordset
Dim LGList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
    
    If Len(Trim(txtMBMNumber)) <> 0 Then
        lstLGListing.ListItems.Clear
        If lstLGListing.ListItems.Count = 0 Then
                
            'sql = "Select CASeNo, LGNo, LGDate, " & _
                  " LGAmt, LGIssuedBy, LGRemark From CASLG " & _
                  " Where CASENo = '" & Trim(txtCaseNo) & "'"

            'LG.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppendCase = "3"
            strAppend = " AND CASENo = '" & Trim(txtCaseNo) & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Inquiry"
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
                        If Len(!LGAmt) Then
                            LGList.SubItems(2) = Format(!LGAmt, "#,##0.00")
                        End If
                        If Len(cboPatientList) Then
                            LGList.SubItems(3) = Trim(Mid(ChkStr(cboPatientList), 1, InStr(1, ChkStr(cboPatientList), "-") - 1))
                        End If
                        If Len(PatientName) Then
                            LGList.SubItems(4) = Trim(PatientName)
                        End If
                        If Len(txtDateAdmitted) Then
                           LGList.SubItems(5) = Format(txtDateAdmitted, "dd/mm/yyyy")
                        End If
                        If Len(txtHOSCode) Then
                            LGList.SubItems(6) = Trim(txtHOSCode)
                        End If
                        If Len(!LGIssuedBy) Then
                            LGList.SubItems(7) = Trim(!LGIssuedBy)
                        End If
                        If Len(!LGRemark) Then
                            LGList.SubItems(8) = Trim(!LGRemark)
                        End If
                        If Len(!LGUserTime) Then
                            LGList.SubItems(9) = Trim(!LGUserTime)
                        End If
                        If Len(!LGSystemTime) Then
                            LGList.SubItems(10) = Trim(!LGSystemTime)
                        End If
                        If Len(!LGPurpose) Then
                            LGList.SubItems(11) = Trim(!LGPurpose)
                        End If
                        
                    .MoveNext
                    End With
                Loop
                LG.Close
                Set LG = Nothing
            End If
        End If
    End If
End Sub

Private Sub ClearEntries()
Dim ctl As Control
Dim s As Integer

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        'ElseIf TypeOf ctl Is ComboBox Then
         '   cboSRVCode.Text = "EA - EUROPE ASSISTANCE"
        '    txtHOSCode.Text = ""
        '    cboPatientList.Clear
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        'ElseIf TypeOf ctl Is MaskEdBox Then
        '    ctl.mask = ""
        '    ctl.Text = ""
        '    ctl.mask = "##/##/####"
        End If
    Next ctl
    '  Call ComboListing
      'cmdSaveKIV.Enabled = True
        'txtFCTime.mask = ""
        'txtFCTime.Text = ""
        'txtFCTime.mask = "##:##"
        'txtFTTime.mask = ""
        'txtFTTime.Text = ""
        'txtFTTime.mask = "##:##"
        'txtDOATime.mask = ""
        'txtDOATime.Text = ""
        'txtDOATime.mask = "##:##"
        'txtDODTime.mask = ""
        'txtDODTime.Text = ""
        'txtDODTime.mask = "##:##"
        CboClaimability = ""
        cboStatus1 = ""
        cboPending = ""
        txtCASReserveAmount = 2500
        cboCASInd = ""
        cboTopUpLmt = ""
        cboREPCode1 = ""
        cboREPCode2 = ""
        cboREPCode3 = ""
        cboCondwillOcc = ""
        cboFLWTreatment = ""
        cboFolUp = ""
        cboSuppExc.Clear
        txtCASRemark.Visible = True
        lblPage.Caption = "Page 1"
        lstLTLmt.ListItems.Clear
        lblPage1.FontBold = True
        lblPage2.FontBold = False
        lblPage3.FontBold = False
        lblPage4.FontBold = False
        lblPage5.FontBold = False
        
        Fgrid.col = 1
        For s = 1 To 10
            Fgrid.row = (s)
        Fgrid = ""
        Next
        txtLGTime.mask = ""
        txtLGTime.Text = ""
        txtLGTime.mask = "##:##"
        List1.Clear
        lblClaimNo = ""
        lstPolicyWording.Clear
End Sub

Private Sub ShowRepudiation()
Dim REP As New ADODB.Recordset
Dim REPCode As New ADODB.Recordset
Dim strsqlRep As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    ' INSERT INTO CASRepudiation TABLE
     
    'strsqlREP = "select * from CASRepudiation where CASNumber = '" & txtCaseNo & "'"
    'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "6"
    strAppend = " AND CASNumber = '" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Inquiry"
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
        
    'strsqlCAS = "Select *  From CAS where CASNumber = '" & Trim(txtCaseNo) & "'"
    'CAS.Open strsqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'If Mid(cboStatus, 1, 1) = "R" Then
    'If REP.recordCount = 0 Then
    '    REP.AddNew
    '    REP!CASNumber = txtCaseNo
    'Else
    'End If
    If Not REP.EOF Then
        With REP
            If Len(Trim(!CASREPCode1)) Then
                strsqlRep = "EXEC Case_Inquiry '" & Trim(!CASREPCode1) & "',7"
                REPCode.Open strsqlRep, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                If REPCode.EOF = False Then
                With REPCode
                    cboREPCode1 = Trim(!REPCode) & " - " & Trim(!RepDescription)
                End With
                End If
                REPCode.Close
                Set REPCode = Nothing
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
        End With
        'REP.Update
        
            'With CAS
            '    !CASStatus = "R"
            'End With
            'CAS.Update
        End If
    REP.Close
    'CAS.Close
    Set REP = Nothing
    'Set CAS = Nothing
End Sub

Private Sub WSListing()
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
    If lstWorksheet.ListItems.Count = 0 Then
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
                Set WSList = lstWorksheet.ListItems.Add(, , !CasNumber & "-" & !claimversion)
                    If Len(!ClaimCloseDate) Then
                        WSList.SubItems(1) = !ClaimCloseDate
                    End If
                    If Len(!CLMTotalActual) Then
                        WSList.SubItems(2) = Format(!CLMTotalActual, "#,##0.00")
                        amntb4GST = !CLMTotalActual
                    End If
                    
                    'GST AMount
                    
                    If Len(!TotGST) Then
                        WSList.SubItems(3) = Format(!TotGST, "#,##0.00")
                        GST = !TotGST
                    End If
                    
                   'Total Actual With GST
                    amntWithGST = amntb4GST + GST
                   If Len(amntWithGST) Then
                      WSList.SubItems(4) = Format(amntWithGST, "#,##0.00")
                    End If
                    
                    If Len(!CLMTotalApproved) Then
                        WSList.SubItems(5) = !CLMTotalApproved
                    End If
                    If Len(!CLMPayableByMedix2H) Then
                        WSList.SubItems(6) = !CLMPayableByMedix2H
                    End If
                    If Len(!CLMPayableByMedix2M) Then
                        WSList.SubItems(7) = !CLMPayableByMedix2M
                    End If
                    If Len(!BordxRefNo) Then
                        WSList.SubItems(8) = !BordxRefNo
                    End If
                    If Len(!BordxDate) Then
                        WSList.SubItems(9) = !BordxDate
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

Private Sub PaymentHOSListing()
Dim PH As New ADODB.Recordset
Dim PHList As ListItem
Dim sql As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

If Len(Trim(txtCaseNo)) <> 0 Then
    If lstHOSPayment.ListItems.Count = 0 Then
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
            Set PHList = lstHOSPayment.ListItems.Add(, , !CasNumber & "-" & !claimversion)
                     If Len(!PVHOSNo) Then
                         PHList.SubItems(1) = !PVHOSNo 'Format(!CLMTotalActual, "#,##0.00")
                     End If
                     If Len(!PVHOSDate) Then
                         PHList.SubItems(2) = !PVHOSDate
                     End If
                     If Len(!PVHOSChequeDate) Then
                         PHList.SubItems(3) = !PVHOSChequeDate
                     End If
                     If Len(!PVHOSChequeNo) Then
                         PHList.SubItems(4) = !PVHOSChequeNo
                     End If
                     If Len(!HOSPaidAmt) Then
                         PHList.SubItems(5) = Format(!HOSPaidAmt, "#,##0.00")
                     End If
                     
                    If Len(!PVStatus) Then
                        PHList.SubItems(6) = Format(!PVPayHOSStatus, "#,##0.00")
                     End If

                     If Len(!PVPayHOSStatus) Then
                        PHList.SubItems(7) = Format(!PVStatus, "#,##0.00")
                     End If
                     If Len(!HOSRemark) Then
                        PHList.SubItems(8) = Format(!HOSRemark, "#,##0.00")
                     End If
                .MoveNext
                End With
        Loop
        PH.Close
        Set PH = Nothing
    End If
End If
End If
End Sub

Private Sub PaymentMBMListing()
Dim PM As New ADODB.Recordset
Dim PMList As ListItem
Dim sql As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

If Len(Trim(txtCaseNo)) <> 0 Then
    If lstMBMPayment.ListItems.Count = 0 Then
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
            Set PMList = lstMBMPayment.ListItems.Add(, , !CasNumber & "-" & !claimversion)
                     If Len(!PVNo) Then
                         PMList.SubItems(1) = !PVNo 'Format(!CLMTotalActual, "#,##0.00")
                     End If
                     If Len(!PVDate) Then
                         PMList.SubItems(2) = !PVDate
                     End If
                     If Len(!PVChequeDate) Then
                         PMList.SubItems(3) = !PVChequeDate
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
                    If Len(!PVStatus) Then
                        PMList.SubItems(7) = !PVStatus
                    End If
                    
                    If Len(!Remark) Then
                        PMList.SubItems(8) = !Remark
                    End If

                .MoveNext
                End With
        Loop
        PM.Close
        Set PM = Nothing
    End If
End If
End If
End Sub

Private Sub PaymentReissuedListing()
Dim rst As New ADODB.Recordset
Dim masterlist As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

If Len(Trim(txtCaseNo)) <> 0 Then
    If lstReissuedMBM.ListItems.Count = 0 Then
            'sql = " Select CASNumber, ClaimVersion, PVNo, PVDate, " & _
            " PVChequeNo,PVChequeDate,MBMPaidAmt,Remark " & _
            " From Payment2MBMReissue " & _
            " where CASNumber ='" & Trim(txtCaseNo) & "'"
                
        'rst.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
        strAppendCase = "3"
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
            Set masterlist = lstReissuedMBM.ListItems.Add(, , !CasNumber & "-" & !claimversion)
                     If Len(!PVNo) Then
                         masterlist.SubItems(1) = !PVNo
                     End If
                     If Len(!PVDate) Then
                         masterlist.SubItems(2) = !PVDate
                     End If
                     If Len(!PVChequeDate) Then
                         masterlist.SubItems(3) = !PVChequeDate
                     End If
                     If Len(!PVChequeNo) Then
                         masterlist.SubItems(4) = !PVChequeNo
                     End If
                     If Len(!MBMPaidAmt) Then
                         masterlist.SubItems(5) = Format(!MBMPaidAmt, "#,##0.00")
                     End If
                     If Len(!Remark) Then
                         masterlist.SubItems(6) = !Remark
                     End If
                     
                     
                .MoveNext
                End With
        Loop
       
        rst.Close
        Set rst = Nothing
    End If
End If
End If
End Sub

Private Sub ReceiptFRMNRBListing()
Dim rst As New ADODB.Recordset
Dim masterlist As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

If Len(Trim(txtCaseNo)) <> 0 Then
    If lstRecMNRB.ListItems.Count = 0 Then
        'sql = " Select CASNumber, ClaimVersion, ReceiptNo, ReceiptDate, " & _
              " ReceiptChequeNo,ReceiptAmt,MBMRemarks " & _
              " From ReceiptFrMBM" & _
              " where CASNumber ='" & Trim(txtCaseNo) & "'"
                
        'rst.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
        strAppendCase = "6"
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
            
                    Set masterlist = lstRecMNRB.ListItems.Add(, , !CasNumber & "-" & !claimversion)
                    If Len(!ReceiptNo) Then
                        masterlist.SubItems(1) = !ReceiptNo
                    End If
                    If Len(!ReceiptDate) Then
                        masterlist.SubItems(2) = !ReceiptDate
                    End If
                    If Len(!ChequeNo) Then
                        masterlist.SubItems(3) = !ChequeNo
                    End If
                    If Len(!CheckAmt) Then
                        masterlist.SubItems(4) = !CheckAmt
                    End If
                    
                    If Len(!Remark) Then
                        masterlist.SubItems(5) = !Remark
                    End If
                         
                .MoveNext
                End With
            Loop
       
        rst.Close
        Set rst = Nothing
        End If
    End If
End If

End Sub

Private Sub ReceiptFRMPYRListing()
Dim rst As New ADODB.Recordset
Dim masterlist As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

If Len(Trim(txtCaseNo)) <> 0 Then
    If lstRecPayor.ListItems.Count = 0 Then
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
            
                    Set masterlist = lstRecPayor.ListItems.Add(, , !CasNumber & "-" & !claimversion)
                    If Len(!ReceiptNo) Then
                        masterlist.SubItems(1) = !ReceiptNo
                    End If
                    If Len(!ReceiptDate) Then
                        masterlist.SubItems(2) = !ReceiptDate
                    End If
                    If Len(!ChequeNo) Then
                        masterlist.SubItems(3) = !ChequeNo
                    End If
                    If Len(!CheckAmt) Then
                        masterlist.SubItems(4) = !CheckAmt
                    End If
                    
                    If Len(!Remark) Then
                        masterlist.SubItems(5) = !Remark
                    End If
                         
                .MoveNext
                End With
            Loop
       
        rst.Close
        Set rst = Nothing
        End If
    End If
End If

End Sub


Private Sub ReceiptFRMBMListing()
Dim rst As New ADODB.Recordset
Dim masterlist As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

If Len(Trim(txtCaseNo)) <> 0 Then
    If lstRecFRMBM.ListItems.Count = 0 Then
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
            
                    Set masterlist = lstRecFRMBM.ListItems.Add(, , !CasNumber & "-" & !claimversion)
                    If Len(!ReceiptNo) Then
                        masterlist.SubItems(1) = !ReceiptNo
                    End If
                    If Len(!ReceiptDate) Then
                        masterlist.SubItems(2) = !ReceiptDate
                    End If
                    If Len(!ReceiptChequeNo) Then
                        masterlist.SubItems(3) = !ReceiptChequeNo
                    End If
                    If Len(!ReceiptAmt) Then
                        masterlist.SubItems(4) = !ReceiptAmt
                    End If
                    
                    If Len(!MBMRemarks) Then
                        masterlist.SubItems(5) = !MBMRemarks
                    End If
                         
                .MoveNext
                End With
            Loop
       
        rst.Close
        Set rst = Nothing
        End If
    End If
End If

End Sub


Private Sub Form_Load()
    Call DoInitialSettings
End Sub

Sub DoInitialSettings()
    
    Fgrid.row = 0
  
    Fgrid.ColWidth(0) = 2500     ' Set column's width
    Fgrid.ColWidth(1) = 1000
    
    Fgrid.col = 0
    Fgrid = ""
    
    Fgrid.col = 1
    Fgrid = ""
    
    Fgrid.row = 1
    Fgrid.col = 0
    Fgrid = "Ref. Req. Adm"
    Fgrid.row = 2
    Fgrid = "Ref . Req for Pre-Specialist"
    Fgrid.row = 3
    Fgrid = "Min. Dur. Hosp"
    Fgrid.row = 4
    Fgrid = "No. Phy Vst/Day"
    Fgrid.row = 5
    Fgrid = "Spec. Per. Prin (Days)"
    Fgrid.row = 6
    Fgrid = "Spec. Per. Supp (Days)"
    Fgrid.row = 7
    Fgrid = "Lower Age Lmt"
    Fgrid.row = 8
    Fgrid = "Upper Age Lmt"
    Fgrid.row = 9
    Fgrid = "Student Age"
    Fgrid.row = 10
    Fgrid = "Child Under Age"
    
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

Private Sub FollowUpListing()
Dim rst As New ADODB.Recordset
Dim strsql As String
Dim Follow As ListItem
    
    lstFollowUp.ListItems.Clear
    strsql = "EXEC SearchFollowUpLG '" & ChkStr(txtCaseNo) & "'"
    rst.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
    
    Do Until rst.EOF
        With rst
            Set Follow = lstFollowUp.ListItems.Add(, , !FollowUpNo)
            Follow.SubItems(1) = !LGAmt
            Follow.SubItems(2) = !FollowUpIssuedBy
            Follow.SubItems(3) = !FollowUpIssuedDate
            .MoveNext
        End With
    Loop
    
End Sub

Private Sub lstMember_BeforeLabelEdit(Cancel As Integer)

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

Private Sub lstEmails_Click()
On Error Resume Next
Dim lRet    As Long
Dim sFile   As String
    
    PictureFileName = DocPath & "Emails\" & txtCaseNo & ""
    '\Medical Rpt"
    crt PictureFileName & "\" & lstEmails.List(lstEmails.ListIndex)

    'If Index = 0 Then
        sFile = PictureFileName & "\" & lstEmails.List(lstEmails.ListIndex)
        lRet = ShellExecute(Me.hwnd, "Open", sFile, &H0&, &H0&, SW_RESTORE)
       ' Picture2.Picture = LoadPicture()
       ' Picture2.Picture = LoadPicture(sFile)
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

Private Sub lstWorksheet_Click()
Dim lstWS As ListItem


    lblClaimNo = ""
    txtAppDate = "__/__/____"
    cboAppStatus = ""
    
    If lstWorksheet.ListItems.Count > 0 Then
        Set lstWS = lstWorksheet.SelectedItem
        lblClaimNo = IIf(Len(Trim(lstWorksheet.SelectedItem)), lstWorksheet.SelectedItem, "")
        cboAppStatus = IIf(Len(Trim(lstWS.SubItems(11))), lstWS.SubItems(11), "")
        txtAppDate = IIf(Len(lstWS.SubItems(12)), Format(lstWS.SubItems(12), "dd/mm/yyyy"), "__/__/____")
        txtCLMAppCode = IIf(Len(Trim(lstWS.SubItems(13))), lstWS.SubItems(13), "")
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Screen.MousePointer = vbHourglass
Dim frmTermsStatus As Integer
'Dim PolList As ListItem
Dim PolList As MSFlexGrid
Dim frmPanelHospStatus As Integer
Dim tmpHOSCode As String

    
    If SSTab1.Tab = 4 Then
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
            Call LGListing
            Call FollowUpListing
       ' medixmasterconnMaint.Close
        'Set medixmasterconnMaint = Nothing
    End If
   
    If SSTab1.Tab = 5 Then
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            Call WSListing
       ' medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
    End If
    
    If SSTab1.Tab = 6 Then
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call INVListing
     '   medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
    
    If SSTab1.Tab = 7 Then
      '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
            Call PaymentHOSListing
            Call PaymentMBMListing
            Call PaymentReissuedListing
            Call ReceiptFRMBMListing
            Call ReceiptFRMPYRListing
            Call ReceiptFRMNRBListing
        
      '  medixmasterconnPayment.Close
     '   Set medixmasterconnPayment = Nothing
    End If
    
    'If SSTab1.Tab = 8 Then
    '    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    '    medixmasterconnPayment.Close
    '    Set medixmasterconnPayment = Nothing
    'End If
    
    If SSTab1.Tab = 8 Then
        'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            Call showReminder
       ' medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
    End If
        
    If SSTab1.Tab = 9 And lstPreUndIllness.ListItems.Count = 0 Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call SearchPrincipalIll
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
    
        If SSTab1.Tab = 10 Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call BNFListing
            'Call NewBNFListing
       ' medixmasterconn.Close
        'Set medixmasterconn = Nothing
    End If
        
    'If SSTab1.Tab = 11 Then
    '    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            'Call BNFListing
            'Call NewBNFListing
    '    medixmasterconn.Close
    '    Set medixmasterconn = Nothing
    'End If
        
    
    If SSTab1.Tab = 12 Then
        If txtPLNCode <> "" Then
            'If lstPolicyTerm.listitems.Count = 0 Then
                frmTermsStatus = 3
              '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                Call PolicyTermsListing(txtPLNCode, frmTermsStatus, txtPayCode, txtHLTCode)
              '  medixmasterconn.Close
            'Set medixmasterconn = Nothing
            'End If
        End If
    End If
        
    If SSTab1.Tab = 13 Then
        If txtPLNCode <> "" Then
            If lstSurgProc.ListItems.Count = 0 Then
            frmTermsStatus = 3
           ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call ScheduleSurg(txtPayCode, txtPLNCode, frmTermsStatus)
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
            End If
        End If
    End If
    
    If SSTab1.Tab = 14 Then
        'If cboPLNCode <> "" Then
            If lstLTLmt.ListItems.Count = 0 Then
              '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                Call LTLmtListing
              '  medixmasterconn.Close
              '  Set medixmasterconn = Nothing
            End If
        'End If
    End If
    
    If SSTab1.Tab = 15 Then
        'Frame1(1).Visible = True
        Call LoadDoc
        Call LoadIntDoc
        Call LoadWebDoc
    End If
    
    If SSTab1.Tab = 16 Then
        If Trim(tmpPolicyWording) <> "" Then
            Call LoadPolicy
        End If
    End If

    If SSTab1.Tab = 17 Then
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            If Trim(txtCaseNo) <> "" Then
                Call CartonListing
            End If
       ' medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
    End If

    If SSTab1.Tab = 18 Then
        'Frame1(1).Visible = True
        Call LoadEmails
    End If

    'If SSTab1.Tab = 15 Then
    '    If txtPLNCode <> "" Then
    '        If lstPanelHosp.listitems.Count = 0 Then
    '            frmTermsStatus = 1
    '            frmPanelHospStatus = 4
    '            medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    '            Call PanelHospListing(txtPLNCode, frmTermsStatus, frmPanelHospStatus, txtPAYCode, tmpHOSCode, txtHOSCode, txtHLTCode)
                
    '        If medixmasterconn Is Nothing Then
                
    '            medixmasterconn.Close
    '        End If
            
    '        Set medixmasterconn = Nothing
    '        End If
    '    End If
    'End If
        
Screen.MousePointer = Default
End Sub

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

Sub crt(pth As String)
Dim f
f = FreeFile
Open App.path + "\temp.html" For Output As #f
    Print #f, "<html><body><center><img src=""" & pth & """></center></body></html>"
Close #f
DoEvents
'WebBrowser1.Navigate App.Path & "\temp.html"
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
    
    If Right(B, 4) = ".doc" Then
        List1.AddItem B
    End If

    If Right(B, 4) = ".pdf" Then
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
        
        If Right(B, 4) = ".doc" Then
            List1.AddItem B
        End If
        
        If Right(B, 4) = ".pdf" Then
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
        
    Else
        Exit Do
    End If
Loop
'List1.Height = Me.Height - 2520 - 350
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
    g = IIf(IsNumeric(txtAnaesMMARate6), txtAnaesMMARate6, 0)
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

Private Sub txtCaseNo_GotFocus()
    cmdShow.Default = True
End Sub

Private Sub TxtCaseNo_LostFocus()
    txtCaseNo = ChkStr(txtCaseNo)
    cmdShow.Default = False
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
        FindHosName = rst3!HosName
    Else
        MsgBox "Hospital Name not found!", vbCritical + vbOKOnly
    End If
    rst3.Close
    Set rst3 = Nothing
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
        
        If Len(CASSearch!CasRemarks) Then
            txtCASRemark3 = txtCASRemark3 + Replace(CASSearch!CasRemarks, "~", Chr(13))
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

Private Sub funcword()
On Error GoTo openerror
   
    ' make sure only one word application launched
    txtSRVPath = crdPath
    Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
                
            'Set objdoc = objword.Documents.Add
            Set objdoc = objword.Documents.Add(LocCardPath & "Case\CaseRemark.DOT")
            'Set objdoc = objword.Documents.Add(txtSRVPath & "\template\" & "\CaseRemark.DOT\")
        End If
        
        With objword.ActiveDocument
                        
            '.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = "CASE MANAGEMENT NOTE"
            '.Sections(1).Headers(wdHeaderFooterPrimary).Range.Bold = True
            '.Sections(1).Headers(wdHeaderFooterPrimary).Range.Font.Size = "18"
            '.Sections(1).Headers(wdHeaderFooterPrimary).Range.Font.Name = "Century Gothic"
            
            '.Sections(1).Range.Text = "Case Number: " & Trim(txtCaseNo)
            '.Sections(1).Range.Bold = True
            '.Sections(1).Range.Font.Size = "12"
            '.Sections(1).Range.Font.Name = "Century Gothic"
            
            'objword.Selection.EndKey Unit:=wdLine
            'objword.Selection.TypeParagraph
            'objword.Selection.TypeParagraph

        End With
        If bookmark("CaseRemark.dot") = True Then
            Me.Show
        End If
        Exit Sub
    
openerror:
    Me.Show
    If Err.Number = "5356" Then
        MsgBox Err.Description & ". Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If

End Sub

Private Sub funcword2()
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

        Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Case\CaseEnq2.xlt")
    
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
                        
            strsqlCaseRec = " SELECT CASNumber, MBMNumber, MBMPolicyNo, MBMName, MBMIcBcPp, " & _
                            " MBMDOB, MBMPayorEffDate, MBMPayorExpDate, CASStatus, CASClaimability, " & _
                            " MBMTakeover, MBMRenewFlag, PatientName, HOSName, AttendDoctor, ReferedBy, " & _
                            " CASDateAdmitted, CASDischargeDate, CASAdmissionReason, CASStayType, ClaimNatureCode, " & _
                            " ModalityCode, CASPAYType, FirstICDCode1, DIAPrelimDiag1, FirstICDCode2, DIAPrelimDiag2, " & _
                            " FirstICDCode3, DIAPrelimDiag3, FinalICDCode1, DIAFinalDiag1, FinalICDCode2, DIAFinalDiag2, " & _
                            " FinalICDCode3, DIAFinalDiag3, CASUDReasons1, CASUDReasons2 , CASUDReasons3, " & _
                            " CASREPCode1, CASREPCode2, CASREPCode3, CASOProcInvestigationDone1, CASOProcInvestigationDone2, " & _
                            " CASOProcInvestigationDone3, CASOPROCRequired1, CASOPROCRequired2, CASOPROCRequired3, " & _
                            " CASOPROCRequired4, CASOPROCRequired5 FROM VIEW_NEW_CAS_REPORT Where CASNumber = '" & Trim(txtCaseNo) & "'"
          
            CaseRec.Open strsqlCaseRec, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
            If CaseRec.recordCount Then
            
                objWorksheet.Cells(6, "B") = Trim(CaseRec!CasNumber)
                objWorksheet.Cells(7, "B") = Trim(CaseRec!MBMNumber)
                objWorksheet.Cells(8, "B") = Trim(CaseRec!MBMPolicyNo)
                objWorksheet.Cells(9, "B") = Trim(CaseRec!MBMName)
                objWorksheet.Cells(10, "B") = Trim(CaseRec!MBMIcBcPp)
                objWorksheet.Cells(11, "B") = Format(CaseRec!MBMDOB, "mm/dd/yyyy")
                objWorksheet.Cells(12, "B") = Format(CaseRec!MBMPayorEffDate, "mm/dd/yyyy")
                objWorksheet.Cells(13, "B") = Format(CaseRec!MBMPayorEXPDate, "mm/dd/yyyy")
                objWorksheet.Cells(6, "G") = Trim(CaseRec!CASStatus)
                objWorksheet.Cells(7, "G") = Trim(CaseRec!CASClaimability)
                objWorksheet.Cells(12, "G") = Trim(CaseRec!MBMTakeover)
                objWorksheet.Cells(13, "G") = Trim(CaseRec!MBMRenewFlag)
                objWorksheet.Cells(18, "B") = Trim(CaseRec!PatientName)
                objWorksheet.Cells(19, "B") = Trim(CaseRec!HosName)
                objWorksheet.Cells(20, "B") = Trim(CaseRec!AttendDoctor)
                objWorksheet.Cells(21, "B") = Trim(CaseRec!ReferedBy)
                objWorksheet.Cells(22, "B") = Format(CaseRec!CASDateAdmitted, "mm/dd/yyyy")
                objWorksheet.Cells(23, "B") = Format(CaseRec!CASDischargeDate, "mm/dd/yyyy")
                objWorksheet.Cells(24, "B") = Trim(CaseRec!CASStayType)
                objWorksheet.Cells(25, "B") = Trim(CaseRec!CASAdmissionReason)
                objWorksheet.Cells(22, "G") = Trim(CaseRec!ClaimNatureCode)
                objWorksheet.Cells(23, "G") = Trim(CaseRec!ModalityCode)
                objWorksheet.Cells(24, "G") = Trim(CaseRec!CASPAYType)
                objWorksheet.Cells(29, "B") = Trim(CaseRec!FirstICDCode1) & " - " & Trim(CaseRec!DIAPrelimDiag1)
                objWorksheet.Cells(30, "B") = Trim(CaseRec!FirstICDCode2) & " - " & Trim(CaseRec!DIAPrelimDiag2)
                objWorksheet.Cells(31, "B") = Trim(CaseRec!FirstICDCode3) & " - " & Trim(CaseRec!DIAPrelimDiag3)
                objWorksheet.Cells(32, "B") = Trim(CaseRec!FinalICDCode1) & " - " & Trim(CaseRec!DIAFinalDiag1)
                objWorksheet.Cells(33, "B") = Trim(CaseRec!FinalICDCode2) & " - " & Trim(CaseRec!DIAFinalDiag2)
                objWorksheet.Cells(34, "B") = Trim(CaseRec!FinalICDCode3) & " - " & Trim(CaseRec!DIAFinalDiag3)
                objWorksheet.Cells(35, "B") = Trim(CaseRec!CASUDReasons1) & ", " & Trim(CaseRec!CASUDReasons2) & ", " & Trim(CaseRec!CASUDReasons3)
                objWorksheet.Cells(37, "B") = Trim(CaseRec!CASREPCode1) & ", " & Trim(CaseRec!CASREPCode2) & ", " & Trim(CaseRec!CASREPCode3)
                objWorksheet.Cells(38, "B") = Trim(CaseRec!CASOProcInvestigationDone1)
                objWorksheet.Cells(39, "B") = Trim(CaseRec!CASOProcInvestigationDone2)
                objWorksheet.Cells(40, "B") = Trim(CaseRec!CASOProcInvestigationDone3)
                objWorksheet.Cells(41, "B") = Trim(CaseRec!CASOPROCRequired1)
                objWorksheet.Cells(42, "B") = Trim(CaseRec!CASOPROCRequired2)
                objWorksheet.Cells(43, "B") = Trim(CaseRec!CASOPROCRequired3)
                objWorksheet.Cells(44, "B") = Trim(CaseRec!CASOPROCRequired4)
                objWorksheet.Cells(45, "B") = Trim(CaseRec!CASOPROCRequired5)
                
                Dim cmd As New ADODB.Command
                Dim Prm1 As New ADODB.Parameter
                Dim WSRec As New ADODB.Recordset
                Dim cnt As Integer
                
                Set cmd.ActiveConnection = medixmasterconnWS
                cmd.CommandText = "CaseEnquiryNotes"
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 300
                Set Prm1 = cmd.CreateParameter("CasNumber", adVarChar, adParamInput, 20, Trim(CaseRec!CasNumber))
                cmd.Parameters.Append Prm1
                WSRec.CursorLocation = adUseClient
                WSRec.CursorType = adOpenForwardOnly
                WSRec.CacheSize = 100
                Set WSRec = cmd.Execute
                
                
        
            
                'strsqlWSHistoryRec = " SELECT CASNumber, ClaimVersion, CLMTotalApproved, CLMTotalActual, " & _
                '                     " CLMPayableByMedix2H, CLMPayableByMedix2M, BordxRefNo, ClaimCloseDate, " & _
                '                     " BordxDate FROM VIEW_WS_Listing WHERE CASNumber = '" & Trim(CaseRec!CasNumber) & "'"
                '
                'WSHistoryRec.Open strsqlWSHistoryRec, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                
                cnt = 1
                 Do Until WSRec.EOF
                    With WSRec
                        objWorksheet.Cells(50 + CDbl(cnt), "A") = Trim(!CasNumber)
                        objWorksheet.Cells(50 + CDbl(cnt), "B") = Trim(!claimversion)
                        objWorksheet.Cells(50 + CDbl(cnt), "C") = Format(!ClaimCloseDate, "mm/dd/yyyy")
                        objWorksheet.Cells(50 + CDbl(cnt), "D") = Format(!CLMTotalActual, "#,##0.00")
                   
                       If Len(!CLMTotalActual) Then
                          amntb4GST = !CLMTotalActual
                        End If
                        
                        If Len(!TotGST) Then
                            objWorksheet.Cells(50 + CDbl(cnt), "E") = Format(!TotGST, "#,##0.00")
                            GST = !TotGST
                        End If
                         
                         amntWithGST = amntb4GST + GST
                         
                         If Len(amntWithGST) Then
                         objWorksheet.Cells(50 + CDbl(cnt), "F") = Format(amntWithGST, "#,##0.00")
                         End If
                         
                        objWorksheet.Cells(50 + CDbl(cnt), "G") = Format(!CLMTotalApproved, "#,##0.00")
                        objWorksheet.Cells(50 + CDbl(cnt), "H") = Format(!CLMPayableByMedix2H, "#,##0.00")
                        objWorksheet.Cells(50 + CDbl(cnt), "I") = Format(!CLMPayableByMedix2M, "#,##0.00")
                        objWorksheet.Cells(50 + CDbl(cnt), "J") = Trim(!BordxRefNo)
                        objWorksheet.Cells(50 + CDbl(cnt), "K") = Format(!BordxDate, "mm/dd/yyyy")
                    End With
                    cnt = cnt + 1
                    WSRec.MoveNext
                 Loop
                 
                 WSRec.Close
                 
                        
                'If WSHistoryRec.recordCount Then
                '    RecordSelected = WSHistoryRec.recordCount
               '
               '     For i = 1 To RecordSelected
               '
               '         objWorksheet.Cells(50 + CDbl(i), "A") = Trim(WSHistoryRec!CasNumber)
               '         objWorksheet.Cells(50 + CDbl(i), "B") = Trim(WSHistoryRec!claimversion)
               '         objWorksheet.Cells(50 + CDbl(i), "C") = Format(WSHistoryRec!ClaimCloseDate, "mm/dd/yyyy")
               '         objWorksheet.Cells(50 + CDbl(i), "D") = Format(WSHistoryRec!CLMTotalActual, "#,##0.00")
               '         objWorksheet.Cells(50 + CDbl(i), "E") = Format(WSHistoryRec!CLMTotalApproved, "#,##0.00")
               '         objWorksheet.Cells(50 + CDbl(i), "F") = Format(WSHistoryRec!CLMPayableByMedix2H, "#,##0.00")
               '         objWorksheet.Cells(50 + CDbl(i), "G") = Format(WSHistoryRec!CLMPayableByMedix2M, "#,##0.00")
               '         objWorksheet.Cells(50 + CDbl(i), "H") = Trim(WSHistoryRec!BordxRefNo)
               '         objWorksheet.Cells(50 + CDbl(i), "I") = Format(WSHistoryRec!BordxDate, "mm/dd/yyyy")
               '
               '     WSHistoryRec.MoveNext
               '
               '    Next i
                    
               ' End If
                
                'WSHistoryRec.Close
                'Set WSHistoryRec = Nothing
            
        End If
        
        CaseRec.Close
        Set CaseRec = Nothing
        
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

Private Sub funcwordPymt()
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
Dim ListCount As Integer
Dim ForLoop As Integer
Dim Itemcnt As Integer
    
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
    
    On Error GoTo HANDLE_ERROR
    objExcel.Interactive = False
    
    
        If objExcel.Visible = False Then
            objExcel.Visible = True
        End If
    
        objExcel.WindowState = xlMaximized

        Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Case\CasePayment.xlt")
    
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
            ForLoop = 0
            ListCount = IIf(lstHOSPayment.ListItems.Count > 0, lstHOSPayment.ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstHOSPayment.ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 5, "A") = Itemcnt
                    objWorksheet.Cells(CDbl(Itemcnt) + 5, "B") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 5, "C") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 5, "D") = Format(.SubItems(2), "DD/MM/YYYY")
                    objWorksheet.Cells(CDbl(Itemcnt) + 5, "E") = .SubItems(4)
                    objWorksheet.Cells(CDbl(Itemcnt) + 5, "F") = Format(.SubItems(5), "#,##0.00")
                    objWorksheet.Cells(CDbl(Itemcnt) + 5, "G") = .SubItems(8)
                End With
            Next
        
            ForLoop = 0
            ListCount = IIf(lstMBMPayment.ListItems.Count > 0, lstMBMPayment.ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstMBMPayment.ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 17, "A") = Itemcnt
                    objWorksheet.Cells(CDbl(Itemcnt) + 17, "B") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 17, "C") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 17, "D") = Format(.SubItems(2), "DD/MM/YYYY")
                    objWorksheet.Cells(CDbl(Itemcnt) + 17, "E") = .SubItems(4)
                    objWorksheet.Cells(CDbl(Itemcnt) + 17, "F") = Format(.SubItems(5), "#,##0.00")
                    objWorksheet.Cells(CDbl(Itemcnt) + 17, "G") = .SubItems(8)
                End With
            Next
        
            ForLoop = 0
            ListCount = IIf(lstReissuedMBM.ListItems.Count > 0, lstReissuedMBM.ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstReissuedMBM.ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 29, "A") = Itemcnt
                    objWorksheet.Cells(CDbl(Itemcnt) + 29, "B") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 29, "C") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 29, "D") = Format(.SubItems(2), "DD/MM/YYYY")
                    objWorksheet.Cells(CDbl(Itemcnt) + 29, "E") = .SubItems(4)
                    objWorksheet.Cells(CDbl(Itemcnt) + 29, "F") = Format(.SubItems(5), "#,##0.00")
                    objWorksheet.Cells(CDbl(Itemcnt) + 29, "G") = .SubItems(8)
                End With
            Next
        
            ForLoop = 0
            ListCount = IIf(lstRecFRMBM.ListItems.Count > 0, lstRecFRMBM.ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstRecFRMBM.ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 35, "A") = Itemcnt
                    objWorksheet.Cells(CDbl(Itemcnt) + 35, "B") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 35, "C") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 35, "D") = Format(.SubItems(2), "DD/MM/YYYY")
                    objWorksheet.Cells(CDbl(Itemcnt) + 35, "E") = .SubItems(3)
                    objWorksheet.Cells(CDbl(Itemcnt) + 35, "F") = Format(.SubItems(4), "#,##0.00")
                    objWorksheet.Cells(CDbl(Itemcnt) + 35, "G") = .SubItems(5)
                End With
            Next
        
            ForLoop = 0
            ListCount = IIf(lstRecPayor.ListItems.Count > 0, lstRecPayor.ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstRecPayor.ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 41, "A") = Itemcnt
                    objWorksheet.Cells(CDbl(Itemcnt) + 41, "B") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 41, "C") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 41, "D") = Format(.SubItems(2), "DD/MM/YYYY")
                    objWorksheet.Cells(CDbl(Itemcnt) + 41, "E") = .SubItems(3)
                    objWorksheet.Cells(CDbl(Itemcnt) + 41, "F") = Format(.SubItems(4), "#,##0.00")
                    objWorksheet.Cells(CDbl(Itemcnt) + 41, "G") = .SubItems(5)
                End With
            Next
        
            ForLoop = 0
            ListCount = IIf(lstRecMNRB.ListItems.Count > 0, lstRecMNRB.ListItems.Count, 0)
            Itemcnt = 0
            
            For ForLoop = 1 To ListCount
                Itemcnt = Itemcnt + 1
                With lstRecMNRB.ListItems(ForLoop)
                    objWorksheet.Cells(CDbl(Itemcnt) + 49, "A") = Itemcnt
                    objWorksheet.Cells(CDbl(Itemcnt) + 49, "B") = .Text
                    objWorksheet.Cells(CDbl(Itemcnt) + 49, "C") = .SubItems(1)
                    objWorksheet.Cells(CDbl(Itemcnt) + 49, "D") = Format(.SubItems(2), "DD/MM/YYYY")
                    objWorksheet.Cells(CDbl(Itemcnt) + 49, "E") = .SubItems(3)
                    objWorksheet.Cells(CDbl(Itemcnt) + 49, "F") = Format(.SubItems(4), "#,##0.00")
                    objWorksheet.Cells(CDbl(Itemcnt) + 49, "G") = .SubItems(5)
                End With
            Next
        
        objExcel.DisplayAlerts = True
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

Private Function bookmark(Optional strAppend As String) As Integer
'Dim CASERemark As New ADODB.Recordset
'Dim strsqlCASERemark As String

    'strsqlCASERemark = "Select CasNumber,CASORemarks from CASRemarks where CASNumber = '" & Trim(txtCaseNo) & "'"
    'CASERemark.Open strsqlCASERemark, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'With CASERemark
    If Trim(txtCaseNo) <> "" Then
        If Trim(PrintNotesStatus) = "M" Then
         'If Trim(!CasNumber) <> "" Then
            objdoc.Bookmarks("CASE").Range.Text = Trim(txtCaseNo) '(!CasNumber)
            
            If Trim(txtCASRemark) <> "" Then 'If Trim(!CASORemarks) <> "" Then
                objdoc.Bookmarks("Note").Range.Text = Trim(txtCASRemark) & Trim(txtCASRemarkTwo) & Trim(txtCASRemark3) & Trim(txtCASRemark4) & Trim(txtCASRemark5)
            End If


        ElseIf Trim(PrintNotesStatus) = "I" Then
            objdoc.Bookmarks("CASE").Range.Text = Trim(txtCaseNo) & " (Internal Notes)"
        
        ''    If Trim(txtIntCASRemark) <> "" Then 'If Trim(!CASORemarks) <> "" Then
        '        objdoc.Bookmarks("Note").Range.Text = Trim(txtIntCASRemark)
        '    End If
        End If
    End If
    'End With
    'CASERemark.Close
    'Set CASERemark = Nothing

End Function

Private Sub showReminder()
Dim REMD As New ADODB.Recordset
Dim strsql As String
Dim REMListing As ListItem

    strsql = "Exec SearchCLMReminder '" & Trim(txtCaseNo) & "'"
    REMD.Open strsql, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
    
    lstReminder.ListItems.Clear
    If Not REMD.EOF Then
        Do Until REMD.EOF
            With REMD
            Set REMListing = lstReminder.ListItems.Add(, , !REMNumber)
            
            If Len(!CasNumber) Then
                REMListing.SubItems(1) = !CasNumber
            End If
            If Len(!REMVersion) Then
                REMListing.SubItems(2) = !REMVersion
            End If
            If Len(!REMRMF) Then
                REMListing.SubItems(3) = !REMRMF
            End If
            If Len(!REMOrinInv) Then
                REMListing.SubItems(4) = !REMOrinInv
            End If
            If Len(!REMRpt) Then
                REMListing.SubItems(5) = !REMRpt
            End If
            If Len(!REMProposalForm) Then
                REMListing.SubItems(6) = !REMProposalForm
            End If
            If Len(!REMRefLetter) Then
                REMListing.SubItems(7) = !REMRefLetter
            End If
            If Len(!REMReceipt) Then
                REMListing.SubItems(8) = !REMReceipt
            End If
            If Len(!REMStudent) Then
                REMListing.SubItems(9) = !REMStudent
            End If
            If Len(!REMPolice) Then
                REMListing.SubItems(10) = !REMPolice
            End If
            If Len(!REMSubroForm) Then
                REMListing.SubItems(11) = !REMSubroForm
            End If
            If Len(!REMOthers1) Then
                REMListing.SubItems(12) = !REMOthers1
            End If
            If Len(!REMOthers2) Then
                REMListing.SubItems(13) = !REMOthers2
            End If
            If Len(!REMDate) Then
                REMListing.SubItems(14) = !REMDate
            End If
            If Len(!REMRegUser) Then
                REMListing.SubItems(15) = !REMRegUser
            End If
            If Len(!REMRegDate) Then
                REMListing.SubItems(16) = !REMRegDate
            End If
            .MoveNext
            End With
        Loop
        REMD.Close
        Set REMD = Nothing
    End If
End Sub

Private Sub SearchPrincipalIll()
Dim rst As New ADODB.Recordset
Dim strsql As String
Dim cmdIll As New ADODB.Command
Dim PrmIll As ADODB.Parameter
Dim strAppend As String
Dim CASIllListing As ListItem


    strAppend = " AND MBMIcBcPp = '" & Trim(txtMBMIcBcPP) & "'"
    Set cmdIll.ActiveConnection = medixmasterconn
    cmdIll.CommandText = "Case_SearchIllness"
    cmdIll.CommandType = adCmdStoredProc
    cmdIll.CommandTimeout = 300
    Set PrmIll = cmdIll.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmdIll.Parameters.Append PrmIll
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.CacheSize = 100
    Set rst = cmdIll.Execute
    
    If Not rst.EOF Then

        If SSTab1.Tab = 9 And lstPreUndIllness.ListItems.Count = 0 Then
            Do Until rst.EOF
            
            With rst
                
            Set CASIllListing = lstPreUndIllness.ListItems.Add(, , !MBMNumber)
                If Len(Trim(!CasNumber)) <> 0 Then
                    CASIllListing.SubItems(1) = ChkStr(!CasNumber)
                End If
                If Len(Trim(!MBMPatientCovID)) <> 0 Then
                    CASIllListing.SubItems(2) = !MBMPatientCovID
                End If
                If Len(Trim(!CASDateAdmitted)) Then
                    CASIllListing.SubItems(3) = !CASDateAdmitted
                End If
                
                If Len(Trim(!CASPreIllDate)) Then
                    CASIllListing.SubItems(4) = !CASPreIllDate
                End If
                
                If Len(Trim(!CASPreIllCAT)) Then
                    If !CASPreIllCAT = "1" Then
                        CASIllListing.SubItems(5) = "Illness"
                    ElseIf !CASPreIllCAT = "2" Then
                        CASIllListing.SubItems(5) = "Injury"
                    ElseIf !CASPreIllCAT = "3" Then
                        CASIllListing.SubItems(5) = "Hospitalisation"
                    ElseIf !CASPreIllCAT = "4" Then
                        CASIllListing.SubItems(5) = IIf(Len(!CASPreIllOthers), !CASPreIllOthers, "")
                    End If
                End If
                'If Len(Trim(!CASPreIllOthers)) Then
                '    CASIllListing.SubItems(5) = !CASPreIllOthers
                'End If
                If Len(Trim(!CASPreIllInfBy)) Then
                    CASIllListing.SubItems(6) = !CASPreIllInfBy
                End If
                If Len(Trim(!CASUndIllDate)) Then
                    CASIllListing.SubItems(7) = !CASUndIllDate
                End If
                If Len(Trim(!CASUndIllCAT)) Then
                    If !CASUndIllCAT = "1" Then
                        CASIllListing.SubItems(8) = "Diabetes"
                    ElseIf !CASUndIllCAT = "2" Then
                        CASIllListing.SubItems(8) = "Hypertension"
                    ElseIf !CASUndIllCAT = "3" Then
                        CASIllListing.SubItems(8) = !CASUndIllOthers
                    End If
                End If
                'If Len(Trim(!CASUndIllOthers)) Then
                '    CASIllListing.SubItems(9) = !CASUndIllOthers
                'End If
                If Len(Trim(!CASUndIllInfBy)) Then
                    CASIllListing.SubItems(9) = !CASUndIllInfBy
                End If
            .MoveNext
            End With
            Loop
        'Else
        End If
        
    
    End If
        
    rst.Close
    Set rst = Nothing
End Sub
Private Sub BNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim bnf As ListItem
    Dim sql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
    If Len(Trim(txtHLTCode)) <> 0 _
     And Len(Trim(txtINSCode1)) <> 0 And Len(Trim(txtPLNCode)) <> 0 Then
        lstBenefits.ListItems.Clear
        If Len(Suplncode) And Mid(cboPatientList, 1, 2) <> "00" Then
            strAppendCase = "17"
            strAppend = " AND PLNCOde ='" & Trim(Suplncode) & "'" & _
              " AND HLTCode ='" & Trim(txtHLTCode) & "'" & _
              " AND INSCODe ='" & Trim(txtINSCode1) & "'"
        Else
            strAppendCase = "17"
            strAppend = " AND PLNCOde ='" & Trim(txtPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(txtHLTCode) & "'" & _
              " AND INSCODe ='" & Trim(txtINSCode1) & "'"
        End If
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
                'If rst2!BNFCode = "101" Then
                '    RB = BNF.SubItems(3)
                'End If
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

'Private Sub NewBNFListing()
'    Dim rst2 As New ADODB.Recordset
'    Dim BNF As ListItem
'    Dim sql As String
'    Dim cmd As New ADODB.Command
'    Dim Prm1 As ADODB.Parameter
'    Dim Prm2 As ADODB.Parameter
'    Dim strAppendCase As String
'    Dim strAppend As String

'    If lstNewBenefitDetails.listitems.Count = 0 And Len(Trim(txtHLTCode)) <> 0 _
'     And Len(Trim(txtINSCode1)) <> 0 And Len(Trim(txtUpgPLNCode)) <> 0 Then
'        strAppendCase = "17"
'        strAppend = " AND PLNCOde ='" & Trim(txtUpgPLNCode) & "'" & _
'          " AND HLTCode ='" & Trim(txtHLTCode) & "'" & _
'          " AND INSCODe ='" & Trim(txtINSCode1) & "'"
'        Set cmd.ActiveConnection = medixmasterconn
'        cmd.CommandText = "Case_Adjustment"
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandTimeout = 300
'        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
'        cmd.Parameters.Append Prm1
'        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
'        cmd.Parameters.Append Prm2
'        rst2.CursorLocation = adUseClient
'        rst2.CursorType = adOpenForwardOnly
'        rst2.CacheSize = 100
'        Set rst2 = cmd.Execute
'
'        Do Until rst2.EOF
'            Set BNF = lstNewBenefitDetails.listitems.Add(, , rst2!BNFCode)
'                If Trim(rst2!BNFdescription) <> "" Then
'                    BNF.SubItems(1) = rst2!BNFdescription
'                End If
'                BNF.SubItems(2) = IIf(IsNull(rst2!BNFNoOfDays), 0, rst2!BNFNoOfDays)
'                 If Trim(rst2!BNFClaimAbleAmount) <> "" Then
'                    BNF.SubItems(3) = Format(rst2!BNFClaimAbleAmount, "###,###,###.00")
'                End If
'                If Trim(rst2!BNFRMperDis) <> "" Then
'                    BNF.SubItems(4) = Format(rst2!BNFRMperDis, "###,###,###.00")
'                End If
'                If rst2!BNFCode = "101" Then
'                    RB = BNF.SubItems(3)
'                End If
'                rst2.MoveNext
'        Loop
'
'        rst2.Close
'        Set rst2 = Nothing
'    End If
'End Sub

Private Sub cboSuppExc_Change()
    Call ShowSuppExclusion(False)
End Sub

Private Sub cboSuppExc_Click()
    Call ShowSuppExclusion(True)
End Sub


Private Sub cboSuppExc_GotFocus()
    Call ShowSuppExclusion(True)
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


Private Sub ShowSuppExclusion(tmpInd As Boolean)
Dim strsql As String
Dim MBMCov As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim conn As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
    
    If tmpInd = True Then
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    End If
    txtSuppExc = ""
    If txtMBMNumber <> "" And IsNumeric(Mid(cboSuppExc, 1, 1)) = True Then
        'strsql = "Select MBMCnumber , MBMCCoverID , MBMCExclusion from MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Trim(Mid(cboSuppExc, 1, 2)) & "'"
        'MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockPessimistic
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
            If Trim(MBMCov!MBMCExclusion) <> "" Then
                txtSuppExc = MBMCov!MBMCExclusion
            End If
        End If
        MBMCov.Close
        Set MBMCov = Nothing
    End If
    If tmpInd = True Then
     '   medixmasterconn.Close
     '   Set medixmasterconn = Nothing
    End If
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
    If Len(Trim(txtNSProcCode1)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode1)
        Else
            Call SearchMMARateCal(txtNSProcCode1)
        End If
        txtNSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtNS13Rate1 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        txtNS13MinRate1 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
    
    Else
        txtNSMMARate1 = 0
        txtNS13Rate1 = 0
        txtNS13MinRate1 = 0
        
    End If
End Sub

Private Sub txtNSProcCode2_Change()
    If Len(Trim(txtNSProcCode2)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode2)
        Else
            Call SearchMMARateCal(txtNSProcCode2)
        End If
        txtNSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtNS13Rate2 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        txtNS13MinRate2 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
    Else
        txtNSMMARate2 = 0
        txtNS13Rate2 = 0
        txtNS13MinRate2 = 0
    End If

End Sub

Private Sub txtNSProcCode3_Change()
    If Len(Trim(txtNSProcCode3)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode3)
        Else
            Call SearchMMARateCal(txtNSProcCode3)
        End If
        txtNSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtNS13Rate3 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        txtNS13MinRate3 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
    Else
        txtNSMMARate3 = 0
        txtNS13Rate3 = 0
        txtNS13MinRate3 = 0
    End If
End Sub

Private Sub txtNSProcCode4_Change()
    If Len(Trim(txtNSProcCode4)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode4)
        Else
            Call SearchMMARateCal(txtNSProcCode4)
        End If
        txtNSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtNS13Rate4 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        txtNS13MinRate4 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
    Else
        txtNSMMARate4 = 0
        txtNS13Rate4 = 0
        txtNS13MinRate4 = 0
        
    End If
End Sub

Private Sub txtNSProcCode5_Change()
    If Len(Trim(txtNSProcCode5)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode5)
        Else
            Call SearchMMARateCal(txtNSProcCode5)
        End If
        txtNSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtNS13Rate5 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        txtNS13MinRate5 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
    Else
        txtNSMMARate5 = 0
        txtNS13Rate5 = 0
        txtNS13MinRate5 = 0
        
    End If

End Sub

Private Sub txtNSProcCode6_Change()
    If Len(Trim(txtNSProcCode6)) Then
        If MMARateCal = 0 Then
            Call SearchMMARate(txtNSProcCode6)
        Else
            Call SearchMMARateCal(txtNSProcCode6)
        End If
        txtNSMMARate6 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtNS13Rate6 = IIf(Len(MMASurgFees13th), MMASurgFees13th, 0)
        txtNS13MinRate6 = IIf(Len(MMASurgFees13thMin), MMASurgFees13thMin, 0)
        
    Else
        txtNSMMARate6 = 0
        txtNS13Rate6 = 0
        txtNS13MinRate6 = 0
        
    End If

End Sub

Private Sub txtProcCode1_Change()
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
    Else
        txtSMMARate1 = 0
        txtAnaesMMARate1 = 0
        txtS13Rate1 = 0
        txtAnae13Rate1 = 0
        txtS13MinRate1 = 0
    End If
End Sub

Private Sub txtProcCode2_Change()
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
    Else
        txtSMMARate2 = 0
        txtAnaesMMARate2 = 0
        txtS13Rate2 = 0
        txtAnae13Rate2 = 0
        txtS13MinRate2 = 0
         
    End If

End Sub

Private Sub txtProcCode3_Change()
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
        
    Else
        txtSMMARate3 = 0
        txtAnaesMMARate3 = 0
        txtS13Rate3 = 0
        txtAnae13Rate3 = 0
        txtS13MinRate3 = 0
         
    End If
End Sub

Private Sub txtProcCode4_Change()
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
        
    Else
        txtSMMARate4 = 0
        txtAnaesMMARate4 = 0
        txtS13Rate4 = 0
        txtAnae13Rate4 = 0
        txtS13MinRate4 = 0
         
    End If
End Sub

Private Sub txtProcCode5_Change()
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
        
    Else
        txtSMMARate5 = 0
        txtAnaesMMARate5 = 0
        txtS13Rate5 = 0
        txtAnae13Rate5 = 0
        txtS13MinRate5 = 0
         
    End If
End Sub



Private Sub txtProcCode6_Change()
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
        
    Else
        txtSMMARate6 = 0
        txtAnaesMMARate6 = 0
        txtS13Rate6 = 0
        txtAnae13Rate6 = 0
        txtS13MinRate6 = 0
    End If
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

Private Function SearchMMARate(OPCSCode As String)
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
    MMASurgFees = 0
    MMAAnaeFees = 0
    MMASurgFees13th = 0
    MMAAnaeFees13th = 0
    MMASurgFees13thMin = 0
    
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
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
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
End Function
Private Sub SearchMMARateCal(CPTCode As String)
Dim strsqlMMA As String
Dim MMA As New ADODB.Recordset
    
    MMASurgFees = 0
    MMAAnaeFees = 0
    MMASurgFees13th = 0
    MMAAnaeFees13th = 0
    MMASurgFees13thMin = 0
    
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

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
   ' medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing

End Sub

Private Sub LTLmtListing()
Dim LTList As ListItem
Dim strsql As String
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim strAppend As String
Dim cmd2 As New ADODB.Command
Dim Prm3 As ADODB.Parameter
Dim prm4 As ADODB.Parameter
Dim cmd3 As New ADODB.Command
Dim Prm5 As ADODB.Parameter
Dim prm6 As ADODB.Parameter
Dim rst4 As New ADODB.Recordset
Dim cmd4 As New ADODB.Command
Dim Prm7 As ADODB.Parameter
Dim prm8 As ADODB.Parameter

Dim rst5 As New ADODB.Recordset
Dim cmd5 As New ADODB.Command
Dim Prm9 As ADODB.Parameter
Dim prm10 As ADODB.Parameter

Dim rst6 As New ADODB.Recordset
Dim cmd6 As New ADODB.Command
Dim prm11 As ADODB.Parameter
Dim prm12 As ADODB.Parameter

Dim rst7 As New ADODB.Recordset
Dim cmd7 As New ADODB.Command
Dim prm13 As ADODB.Parameter
Dim prm14 As ADODB.Parameter

Dim rst8 As New ADODB.Recordset
Dim rst9 As New ADODB.Recordset
Dim tmpAnnBalLmt As Double
Dim tmpLTBalLmt As Double
    'lstLTLmt.listitems.Clear
            
    Screen.MousePointer = vbHourglass
    
    strAppend = " AND (MBMNumber = '" & Trim(txtMBMNumber) & "')"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "SP_SearchLTLmt"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, "1")
        cmd.Parameters.Append Prm2
        rst1.CursorLocation = adUseClient
        rst1.CursorType = adOpenForwardOnly
        rst1.CacheSize = 100
        Set rst1 = cmd.Execute

    Do While Not rst1.EOF
        With rst1
        
            Set LTList = lstLTLmt.ListItems.Add(, , !MBMNumber)
                LTList.SubItems(1) = !MBMCOVID
                LTList.SubItems(2) = !MBMName
                LTList.SubItems(3) = IIf(IsNumeric(!MBMAvailableLimit), Format(!MBMAvailableLimit, "#,##0.00"), "0")
                
                tmpAnnBalLmt = IIf(IsNumeric(!MBMAvailableLimit), Format(!MBMAvailableLimit, "#,##0.00"), "0")
                
                If Trim(!MBMGuaRenewal) = "3" Then
                    strAppend = ""
                    strAppend = " Select MBMNumber,MBMLifeTimeLmt from View_Search_LTPrinLmt where MBMNumber = '" & Trim(!MBMTopupNumber) & "'"
                    rst8.Open strAppend, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
                         tmpLTBalLmt = IIf(IsNumeric(rst8!MBMLifeTimeLmt), Format(rst8!MBMLifeTimeLmt, "#,##0.00"), "0")
                    rst8.Close
                    Set rst8 = Nothing
                    
                    'If CDbl(tmpAnnBalLmt) > CDbl(tmpLTBalLmt) Then
                    '    LTList.SubItems(4) = IIf(IsNumeric(!tmpLTBalLmt), Format(!tmpLTBalLmt, "#,##0.00"), "0")
                    'Else
                        LTList.SubItems(4) = IIf(IsNumeric(tmpLTBalLmt), Format(tmpLTBalLmt, "#,##0.00"), "0")
                    'End If
                Else
                    LTList.SubItems(4) = IIf(IsNumeric(tmpAnnBalLmt), Format(tmpAnnBalLmt, "#,##0.00"), "0")
                End If
            strAppend = ""
            strAppend = " AND (MBMNumber = '" & Trim(rst1!MBMNumber) & "')"
            Set cmd2.ActiveConnection = medixmasterconn
                cmd2.CommandText = "SP_SearchLTLmt"
                cmd2.CommandType = adCmdStoredProc
                cmd2.CommandTimeout = 300
                Set Prm3 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd2.Parameters.Append Prm3
                Set prm4 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, "2")
                cmd2.Parameters.Append prm4
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd2.Execute
                
                Do While Not rst2.EOF
                    LTList.SubItems(5) = IIf(IsNumeric(rst2!Cases), rst2!Cases, "0")
                    LTList.SubItems(7) = IIf(IsNumeric(rst2!BordxAmt), Format(rst2!BordxAmt, "#,##0.00"), "0")
                rst2.MoveNext
                Loop
            rst2.Close
            Set rst2 = Nothing
            
            strAppend = ""
            strAppend = " AND (MBMTopupNumber = '" & Trim(rst1!MBMTopupNumber) & "')"
            Set cmd3.ActiveConnection = medixmasterconn
                cmd3.CommandText = "SP_SearchLTLmt"
                cmd3.CommandType = adCmdStoredProc
                cmd3.CommandTimeout = 300
                Set Prm5 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd3.Parameters.Append Prm5
                Set prm6 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, "3")
                cmd3.Parameters.Append prm6
                rst3.CursorLocation = adUseClient
                rst3.CursorType = adOpenForwardOnly
                rst3.CacheSize = 100
                Set rst3 = cmd3.Execute
                Do While Not rst3.EOF
                    LTList.SubItems(6) = IIf(IsNumeric(rst3!CasHistory), rst3!CasHistory, "0")
                    LTList.SubItems(8) = IIf(IsNumeric(rst3!BordxAmt), Format(rst3!BordxAmt, "#,##0.00"), "0")
                rst3.MoveNext
                Loop
            rst3.Close
            Set rst3 = Nothing
            
                
                If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                
                    strAppend = ""
                    strAppend = " AND (MBMCNumber = '" & Trim(!MBMNumber) & "')"
                    Set cmd5.ActiveConnection = medixmasterconn
                        cmd5.CommandText = "SP_SearchLTLmt"
                        cmd5.CommandType = adCmdStoredProc
                        cmd5.CommandTimeout = 300
                        Set Prm9 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                        cmd5.Parameters.Append Prm9
                        Set prm10 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, "4")
                        cmd5.Parameters.Append prm10
                        rst5.CursorLocation = adUseClient
                        rst5.CursorType = adOpenForwardOnly
                        rst5.CacheSize = 100
                        Set rst5 = cmd5.Execute
                    
                    Do While Not rst5.EOF
                        tmpAnnBalLmt = IIf(IsNumeric(rst5!MBMCAvailableLimit), Format(rst5!MBMCAvailableLimit, "#,##0.00"), "0")
                        
                        'tmpLTBalLmt = IIf(IsNumeric(!MBMLifeTimeLmt), Format(!MBMLifeTimeLmt, "#,##0.00"), "0")

                        Set LTList = lstLTLmt.ListItems.Add(, , rst5!MBMCNumber)
                        LTList.SubItems(1) = rst5!MBMCCoverID
                        LTList.SubItems(2) = rst5!MBMCName
                        LTList.SubItems(3) = IIf(IsNumeric(tmpAnnBalLmt), Format(tmpAnnBalLmt, "#,##0.00"), "0")
                        If Trim(!MBMGuaRenewal) = "3" Then

                            strAppend = ""
                            strAppend = " Select MBMCNumber,MBMCCoverID,MBMLifeTimeLmt from View_Search_LTSuppLmt where MBMCNumber = '" & Trim(!MBMTopupNumber) & "' And MBMCCoverID = '" & Trim(rst5!MBMCCoverID) & "'"
                            rst9.Open strAppend, medixmasterconn, adOpenKeyset, adLockOptimistic
                                 tmpLTBalLmt = IIf(IsNumeric(rst9!MBMLifeTimeLmt), Format(rst9!MBMLifeTimeLmt, "#,##0.00"), "0")
                            rst9.Close
                            Set rst9 = Nothing
    
                            'If CDbl(tmpAnnBalLmt) > CDbl(tmpLTBalLmt) Then
                            '    LTList.SubItems(3) = IIf(IsNumeric(tmpLTBalLmt), Format(tmpLTBalLmt, "#,##0.00"), "0")
                            'Else
                                LTList.SubItems(4) = IIf(IsNumeric(tmpLTBalLmt), Format(tmpLTBalLmt, "#,##0.00"), "0")
                            'End If
                        Else
                            LTList.SubItems(4) = IIf(IsNumeric(tmpAnnBalLmt), Format(tmpAnnBalLmt, "#,##0.00"), "0")
                        End If
                        'LTList.SubItems(4) = IIf(IsNumeric(rst5!MBMLifeTimeLmt), Format(rst5!MBMLifeTimeLmt, "#,##0.00"), "0")
                         
                        'cmd6.ActiveConnection.Close
                        'prm11.
                        'prm12.value = ""
                        strAppend = ""
                        'strAppend = " AND (MBMCNumber = '" & Trim(rst5!MBMCNumber) & "' And MBMCCoverID = '" & Trim(rst5!MBMCCoverID) & "')"
                        
                        strAppend = " Select MBMCNumber,MBMCCoverID,Cases,BordxAmt from View_Search_LTSuppCase where MBMCNumber = '" & Trim(rst5!MBMCNumber) & "' And MBMCCoverID = '" & Trim(rst5!MBMCCoverID) & "'"
                        rst6.Open strAppend, medixmasterconn, adOpenKeyset, adLockOptimistic
                        'Set cmd6.ActiveConnection = medixmasterconn
                        '    cmd6.CommandText = "SP_SearchLTLmt"
                        '    cmd6.CommandType = adCmdStoredProc
                        '    cmd6.CommandTimeout = 300
                        '    Set prm11 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                        '    cmd6.Parameters.Append prm11
                        '    Set prm12 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, "5")
                        '    cmd6.Parameters.Append prm12
                        '    rst6.CursorLocation = adUseClient
                        '    rst6.CursorType = adOpenForwardOnly
                        '    rst6.CacheSize = 100
                        '    Set rst6 = cmd6.Execute
                            LTList.SubItems(5) = "0"
                            LTList.SubItems(7) = "0"
                            Do While Not rst6.EOF
                                LTList.SubItems(5) = IIf(IsNumeric(rst6!Cases), rst6!Cases, "0")
                                LTList.SubItems(7) = IIf(IsNumeric(rst6!BordxAmt), Format(rst6!BordxAmt, "#,##0.00"), "0")

                            rst6.MoveNext
                            Loop
                        rst6.Close
                        Set rst6 = Nothing
                        
                        strAppend = ""
                        'strAppend = " AND (MBMTopupNumber = '" & Trim(rst5!MBMTopupNumber) & "' And MBMCCoverID = '" & Trim(rst5!MBMCCoverID) & "')"
                        'Set cmd7.ActiveConnection = medixmasterconn
                        '    cmd7.CommandText = "SP_SearchLTLmt"
                        '    cmd7.CommandType = adCmdStoredProc
                        '    cmd7.CommandTimeout = 300
                        '    Set prm13 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                        '    cmd7.Parameters.Append prm13
                        '    Set prm14 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, "6")
                        '    cmd7.Parameters.Append prm14
                        '    rst7.CursorLocation = adUseClient
                        '    rst7.CursorType = adOpenForwardOnly
                        '    rst7.CacheSize = 100
                        '    Set rst7 = cmd7.Execute
                        strAppend = " Select MBMTopupNumber,MBMCCoverID,CasHistory,BordxAmt from VIEW_Search_LTSuppCaseHistory where MBMTopupNumber = '" & Trim(rst5!MBMTopupNumber) & "' And MBMCCoverID = '" & Trim(rst5!MBMCCoverID) & "'"
                        rst7.Open strAppend, medixmasterconn, adOpenKeyset, adLockOptimistic
                            LTList.SubItems(6) = "0"
                            LTList.SubItems(8) = "0"
                            Do While Not rst7.EOF
                                LTList.SubItems(6) = IIf(IsNumeric(rst7!CasHistory), rst7!CasHistory, "0")
                                LTList.SubItems(8) = IIf(IsNumeric(rst7!BordxAmt), Format(rst7!BordxAmt, "#,##0.00"), "0")
                            rst7.MoveNext
                            Loop
                    
                        rst7.Close
                        Set rst7 = Nothing

                    rst5.MoveNext
                    Loop
                    rst5.Close
                    Set rst5 = Nothing
            End If
        
        End With
        rst1.MoveNext
    Loop
    rst1.Close
    Set rst1 = Nothing
    Screen.MousePointer = vbDefault
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

Private Sub CartonListing()
Dim ws As New ADODB.Recordset
Dim WSList As ListItem
Dim strsql As String
    If Len(Trim(txtCaseNo)) <> "" Then
        If lstStorage.ListItems.Count = 0 Then
            strsql = ""
            lstStorage.ListItems.Clear
            strsql = ""
            strsql = " select * from VIEW_Storage_CLM_Search where CASNumber = '" & Trim(txtCaseNo) & "'"
            ws.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            
            If Not ws.EOF Then
                Do Until ws.EOF
                    With ws
                        Set WSList = lstStorage.ListItems.Add(, , !CasNumber & "-" & !claimversion)
                            If Len(!STRCartonNo) Then
                                WSList.SubItems(1) = !STRCartonNo
                            End If
                            If Len(!STRFileStatus) Then
                                WSList.SubItems(2) = Format(!STRFileStatus, "#,##0.00")
                            End If
                            If Len(!CRTCartonStatus) Then
                                WSList.SubItems(3) = Format(!CRTCartonStatus, "#,##0.00")
                            End If
                            
                            If Len(!STRRequestBy) Then
                                WSList.SubItems(4) = !STRRequestBy
                            End If
                            If Len(!DORequest) Then
                                WSList.SubItems(5) = !DORequest
                            End If
                            If Len(!DOReturn) Then
                                WSList.SubItems(6) = !DOReturn
                            End If
                            If Len(!STRRemarks) Then
                                WSList.SubItems(7) = !STRRemarks
                            End If
                        .MoveNext
                    End With
                Loop
            End If
            ws.Close
            Set ws = Nothing
                
            strsql = ""
            strsql = " select * from VIEW_Storage_CLMRep_Search where CASNumber = '" & Trim(txtCaseNo) & "'"
            ws.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            
            If Not ws.EOF Then
                Do Until ws.EOF
                    With ws
                        Set WSList = lstStorage.ListItems.Add(, , !CasNumber)
                            If Len(!STRCartonNo) Then
                                WSList.SubItems(1) = !STRCartonNo
                            End If
                            If Len(!STRFileStatus) Then
                                WSList.SubItems(2) = Format(!STRFileStatus, "#,##0.00")
                            End If
                            If Len(!CRTCartonStatus) Then
                                WSList.SubItems(3) = Format(!CRTCartonStatus, "#,##0.00")
                            End If
                            
                            If Len(!STRRequestBy) Then
                                WSList.SubItems(4) = !STRRequestBy
                            End If
                            If Len(!DORequest) Then
                                WSList.SubItems(5) = !DORequest
                            End If
                            If Len(!DOReturn) Then
                                WSList.SubItems(6) = !DOReturn
                            End If
                            If Len(!STRRemarks) Then
                                WSList.SubItems(7) = !STRRemarks
                            End If
                        .MoveNext
                    End With
                Loop
               
            End If
            ws.Close
            Set ws = Nothing

        End If
    End If
End Sub


Private Sub ShowExclusion()
    Dim rst As New ADODB.Recordset
    Dim strsql As String
    
        strsql = " SELECT * from View_Get_MemberExclusion " & _
                " where PAYCode = '" & Trim(txtPayCode) & "' " & _
                " And MBMDOB = '" & Format(tmpPrincipalDOB, "yyyy/mm/dd") & "' " & _
                " And INSCode = '" & Trim(txtINSCode1) & "'" & _
                " And MBMName = '" & Trim(txtMBMName) & "'"
        rst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            Do While Not rst.EOF
                If Trim(TmpExclussion) = "" Then
                    TmpExclussion = IIf(Len(Trim(rst!MBMExclusion)), Trim(rst!MBMExclusion), "")
                Else
                    TmpExclussion = TmpExclussion & vbNewLine & IIf(Len(Trim(rst!MBMExclusion)), Trim(rst!MBMExclusion), "")
                End If
            rst.MoveNext
            Loop
        rst.Close
        Set rst = Nothing
    'End If
End Sub


Private Sub LoadEmails()
On Error Resume Next
Dim B As String, f As Integer, d As Integer, e As Integer
Call addemails



'f = FreeFile
'Open App.Path & "\temp.html" For Output As #f
'    Print #f, "<html><body><div align=center><table border=0><tr><td valign=center><center><a href=" & Chr(34) & "Please Click File to view" & Chr(34) & " target=" & Chr(34) & "_blank" & Chr(34) & ">visit xeek.net</a></td></tr></table></body>"
'Close #f
'WebBrowser1.Navigate App.Path & "\temp.html"
Exit Sub
End Sub

Sub addemails()
Dim B As String
lstEmails.Clear

PictureFileName = DocPath & "Emails\" & txtCaseNo & ""
'\Medical Rpt"
B = Dir(PictureFileName & "\*.*")
'b = Dir(Dir1.Path & "\*.*")
If B <> "" Then
    If Right(B, 4) = ".eml" Then
        lstEmails.AddItem B
    End If
    
Else
    Exit Sub
End If
Do: DoEvents
    B = Dir
    If B <> "" Then
        If Right(B, 4) = ".eml" Then
            lstEmails.AddItem B
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



