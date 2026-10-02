VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseAdmission 
   Caption         =   "Case Admission"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6240
      Left            =   120
      TabIndex        =   146
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11007
      _Version        =   393216
      Tabs            =   12
      Tab             =   1
      TabsPerRow      =   12
      TabHeight       =   529
      TabCaption(0)   =   "General Info"
      TabPicture(0)   =   "frmCaseAdmission.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstMember"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Case Details"
      TabPicture(1)   =   "frmCaseAdmission.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Case Others"
      TabPicture(2)   =   "frmCaseAdmission.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Take Over/Others"
      TabPicture(3)   =   "frmCaseAdmission.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtAuditReview"
      Tab(3).Control(1)=   "Frame10"
      Tab(3).Control(2)=   "Frame9"
      Tab(3).Control(3)=   "Frame5"
      Tab(3).Control(4)=   "Frame8"
      Tab(3).Control(5)=   "Label26(9)"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Benefit"
      TabPicture(4)   =   "frmCaseAdmission.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label26(54)"
      Tab(4).Control(1)=   "Label46"
      Tab(4).Control(2)=   "txtDOC"
      Tab(4).Control(3)=   "lstBenefitDetails"
      Tab(4).Control(4)=   "lstNewBenefitDetails"
      Tab(4).Control(5)=   "txtUpgPLNCode"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "MBM'Ship History"
      TabPicture(5)   =   "frmCaseAdmission.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstMBMHis"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Warranty"
      TabPicture(6)   =   "frmCaseAdmission.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cboSuppExc"
      Tab(6).Control(1)=   "txtSuppExc"
      Tab(6).Control(2)=   "txtPrincipalExc"
      Tab(6).Control(3)=   "Label34"
      Tab(6).Control(4)=   "Label24"
      Tab(6).ControlCount=   5
      TabCaption(7)   =   "Non-Disclosure"
      TabPicture(7)   =   "frmCaseAdmission.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lstPreUndIllness"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Case Mgmt Notes"
      TabPicture(8)   =   "frmCaseAdmission.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label26(57)"
      Tab(8).Control(1)=   "Label26(56)"
      Tab(8).Control(2)=   "Label26(55)"
      Tab(8).Control(3)=   "lblPage"
      Tab(8).Control(4)=   "lblPage5"
      Tab(8).Control(5)=   "lblPage4"
      Tab(8).Control(6)=   "lblPage3"
      Tab(8).Control(7)=   "lblPage2"
      Tab(8).Control(8)=   "lblPage1"
      Tab(8).Control(9)=   "txtCASRemark5"
      Tab(8).Control(10)=   "txtCASRemark4"
      Tab(8).Control(11)=   "txtCASRemark3"
      Tab(8).Control(12)=   "txtCASRemarkTwo"
      Tab(8).Control(13)=   "txtCASRemark"
      Tab(8).Control(14)=   "Frame11"
      Tab(8).Control(15)=   "Frame12"
      Tab(8).ControlCount=   16
      TabCaption(9)   =   "Policy Terms"
      TabPicture(9)   =   "frmCaseAdmission.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Fgrid"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Surgical Schedule"
      TabPicture(10)  =   "frmCaseAdmission.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lstSurgProc"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "Policy Wording"
      TabPicture(11)  =   "frmCaseAdmission.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "lstPolicyWording"
      Tab(11).Control(1)=   "CmnDlg"
      Tab(11).ControlCount=   2
      Begin VB.TextBox txtAuditReview 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Left            =   -74880
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   320
         Top             =   3600
         Width           =   11295
      End
      Begin VB.ListBox lstPolicyWording 
         Height          =   5520
         ItemData        =   "frmCaseAdmission.frx":0150
         Left            =   -74880
         List            =   "frmCaseAdmission.frx":0152
         TabIndex        =   311
         ToolTipText     =   "right click to randomize contents"
         Top             =   400
         Width           =   4455
      End
      Begin VB.Frame Frame10 
         Caption         =   "Others Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74880
         TabIndex        =   246
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
         Begin VB.TextBox txtCompAction 
            Height          =   1215
            Left            =   5760
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   1920
            Width           =   5415
         End
         Begin MSMask.MaskEdBox txtDOATime 
            Height          =   285
            Left            =   1440
            TabIndex        =   126
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDOALGTime 
            Height          =   285
            Left            =   1440
            TabIndex        =   127
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDODTime 
            Height          =   285
            Left            =   4080
            TabIndex        =   128
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDODLGTime 
            Height          =   285
            Left            =   4080
            TabIndex        =   129
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDODComment 
            Height          =   285
            Left            =   4080
            TabIndex        =   130
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtComplaintIssue 
            Height          =   1215
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   131
            Top             =   1920
            Width           =   5535
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame12"
         Height          =   5295
         Left            =   -66960
         TabIndex        =   218
         Top             =   720
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame11"
         Height          =   5295
         Left            =   -73200
         TabIndex        =   217
         Top             =   720
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "Take Over Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         TabIndex        =   237
         Top             =   360
         Width           =   11415
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
            Left            =   8160
            MaxLength       =   4
            TabIndex        =   116
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cboCASTakeOveInd 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   720
            TabIndex        =   113
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox TxtLinkCASNo 
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
            Left            =   10200
            MaxLength       =   15
            TabIndex        =   117
            Top             =   240
            Visible         =   0   'False
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
            Left            =   3000
            TabIndex        =   114
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cboIntExtTO 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5760
            TabIndex        =   115
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label26 
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
            Index           =   69
            Left            =   4440
            TabIndex        =   242
            Top             =   240
            Width           =   1380
         End
         Begin VB.Label Label26 
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
            Index           =   45
            Left            =   120
            TabIndex        =   241
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label26 
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
            Index           =   46
            Left            =   2350
            TabIndex        =   240
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "T/O PLN"
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
            Left            =   7440
            TabIndex        =   239
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label26 
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
            Index           =   52
            Left            =   9240
            TabIndex        =   238
            Top             =   240
            Visible         =   0   'False
            Width           =   1140
         End
      End
      Begin VB.TextBox txtUpgPLNCode 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67440
         TabIndex        =   232
         Top             =   345
         Width           =   975
      End
      Begin VB.ComboBox cboSuppExc 
         Height          =   315
         Left            =   -73680
         TabIndex        =   227
         Top             =   3180
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
         Height          =   2460
         Left            =   -74880
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   226
         Top             =   3540
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
         Height          =   2520
         Left            =   -74880
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   225
         Top             =   600
         Width           =   11415
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
         Height          =   5265
         Left            =   -74160
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   219
         Top             =   720
         Width           =   8895
      End
      Begin VB.Frame Frame7 
         Height          =   5625
         Left            =   160
         TabIndex        =   165
         Top             =   360
         Width           =   11415
         Begin VB.TextBox txtPatientList 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   23
            Top             =   120
            Width           =   5000
         End
         Begin VB.ComboBox cboPatientList 
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
            ItemData        =   "frmCaseAdmission.frx":0154
            Left            =   1300
            List            =   "frmCaseAdmission.frx":0156
            Sorted          =   -1  'True
            TabIndex        =   171
            Top             =   120
            Width           =   5000
         End
         Begin VB.TextBox txtPatientIC 
            Height          =   285
            Left            =   7680
            MaxLength       =   15
            TabIndex        =   32
            Top             =   140
            Width           =   3375
         End
         Begin VB.Frame Frame15 
            Height          =   1395
            Left            =   120
            TabIndex        =   170
            Top             =   2700
            Width           =   5655
            Begin VB.CommandButton CmdDiagPre 
               Caption         =   ">"
               Height          =   255
               Index           =   0
               Left            =   5280
               TabIndex        =   213
               Top             =   180
               Width           =   225
            End
            Begin MSMask.MaskEdBox txtPDDOS1 
               Height          =   285
               Left            =   3960
               TabIndex        =   48
               Top             =   465
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox txtPDDOS2 
               Height          =   285
               Left            =   3960
               TabIndex        =   51
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
            Begin MSMask.MaskEdBox txtPDDOS3 
               Height          =   285
               Left            =   3960
               TabIndex        =   54
               Top             =   975
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
            Begin VB.TextBox txtPrelimDiag1 
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   47
               Top             =   465
               Width           =   3860
            End
            Begin VB.TextBox txtPrelimDiag2 
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   50
               Top             =   720
               Width           =   3860
            End
            Begin VB.TextBox txtPrelimDiag3 
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   53
               Top             =   975
               Width           =   3860
            End
            Begin VB.TextBox cboICDCode1 
               Height          =   285
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   49
               Top             =   465
               Width           =   615
            End
            Begin VB.TextBox cboICDCode2 
               Height          =   285
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   52
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox cboICDCode3 
               Height          =   285
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   55
               Top             =   975
               Width           =   615
            End
            Begin VB.Label Label26 
               Caption         =   "ICD"
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
               Left            =   4920
               TabIndex        =   196
               Top             =   180
               Width           =   540
            End
            Begin VB.Label Label26 
               Caption         =   "Date OnSet"
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
               Left            =   3960
               TabIndex        =   195
               Top             =   180
               Width           =   900
            End
            Begin VB.Label Label26 
               Caption         =   "Prem. Diag"
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
               Left            =   120
               TabIndex        =   194
               Top             =   180
               Width           =   1140
            End
         End
         Begin VB.CheckBox ChkDOA 
            Caption         =   "Check1"
            Height          =   255
            Left            =   11100
            TabIndex        =   168
            Top             =   400
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox ChkDOD 
            Caption         =   "Check1"
            Height          =   255
            Left            =   11100
            TabIndex        =   167
            Top             =   660
            Value           =   1  'Checked
            Width           =   255
         End
         Begin MSMask.MaskEdBox txtFirstCall 
            Height          =   315
            Left            =   7680
            TabIndex        =   33
            Top             =   375
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
         Begin MSMask.MaskEdBox txtFirstTreated 
            Height          =   315
            Left            =   7680
            TabIndex        =   35
            Top             =   620
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
         Begin MSMask.MaskEdBox txtDateAdmitted 
            Height          =   300
            Left            =   9720
            TabIndex        =   34
            Top             =   360
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
            Height          =   285
            Left            =   9720
            TabIndex        =   36
            Top             =   620
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
            Height          =   285
            Left            =   7680
            TabIndex        =   37
            Top             =   900
            Width           =   1335
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
            Height          =   305
            Left            =   9720
            TabIndex        =   38
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboCondwillOcc 
            Height          =   315
            Left            =   7680
            TabIndex        =   39
            Top             =   1120
            Width           =   1335
         End
         Begin VB.Frame Frame16 
            Height          =   1365
            Left            =   120
            TabIndex        =   172
            Top             =   4140
            Width           =   5655
            Begin VB.CommandButton CmdDiagUD 
               Caption         =   ">"
               Height          =   255
               Index           =   2
               Left            =   5280
               TabIndex        =   215
               Top             =   180
               Width           =   225
            End
            Begin MSMask.MaskEdBox txtUDDOS1 
               Height          =   285
               Left            =   3960
               TabIndex        =   66
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
               Height          =   285
               Left            =   3960
               TabIndex        =   69
               Top             =   705
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
               Height          =   285
               Left            =   3960
               TabIndex        =   72
               Top             =   960
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
            Begin VB.TextBox txtUD1 
               Height          =   285
               Left            =   90
               MaxLength       =   150
               TabIndex        =   65
               Top             =   480
               Width           =   3880
            End
            Begin VB.TextBox txtUD2 
               Height          =   285
               Left            =   90
               MaxLength       =   150
               TabIndex        =   68
               Top             =   720
               Width           =   3880
            End
            Begin VB.TextBox txtUD3 
               Height          =   285
               Left            =   90
               MaxLength       =   150
               TabIndex        =   71
               Top             =   960
               Width           =   3880
            End
            Begin VB.TextBox cboUDICDCode1 
               Height          =   285
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   67
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox cboUDICDCode2 
               Height          =   285
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   70
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox cboUDICDCode3 
               Height          =   285
               Left            =   4920
               MaxLength       =   3
               TabIndex        =   73
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label26 
               Caption         =   "Underlying Disease"
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
               Left            =   120
               TabIndex        =   197
               Top             =   180
               Width           =   1620
            End
            Begin VB.Label Label58 
               Caption         =   "Date Onset"
               Height          =   255
               Left            =   3960
               TabIndex        =   174
               Top             =   180
               Width           =   855
            End
            Begin VB.Label Label56 
               Caption         =   "ICD "
               Height          =   255
               Left            =   4920
               TabIndex        =   173
               Top             =   180
               Width           =   495
            End
         End
         Begin VB.Frame Frame17 
            Height          =   1395
            Left            =   5880
            TabIndex        =   169
            Top             =   2700
            Width           =   5400
            Begin VB.CommandButton CmdDiagFinal 
               Caption         =   ">"
               Height          =   255
               Index           =   1
               Left            =   5040
               TabIndex        =   214
               Top             =   180
               Width           =   225
            End
            Begin VB.TextBox txtFinalDiag1 
               Height          =   285
               Left            =   90
               MaxLength       =   150
               TabIndex        =   56
               Top             =   465
               Width           =   3650
            End
            Begin VB.TextBox txtFinalDiag2 
               Height          =   285
               Left            =   90
               MaxLength       =   150
               TabIndex        =   59
               Top             =   720
               Width           =   3650
            End
            Begin VB.TextBox txtFinalDiag3 
               Height          =   285
               Left            =   90
               MaxLength       =   150
               TabIndex        =   62
               Top             =   975
               Width           =   3650
            End
            Begin MSMask.MaskEdBox txtFDDOS1 
               Height          =   285
               Left            =   3720
               TabIndex        =   57
               Top             =   465
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
               Height          =   285
               Left            =   3720
               TabIndex        =   60
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
               Height          =   285
               Left            =   3720
               TabIndex        =   63
               Top             =   975
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
            Begin VB.TextBox cboFDICDCode1 
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   58
               Top             =   465
               Width           =   615
            End
            Begin VB.TextBox cboFDICDCode2 
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   61
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox cboFDICDCode3 
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   64
               Top             =   975
               Width           =   615
            End
            Begin VB.Label Label26 
               Caption         =   "ICD"
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
               Index           =   33
               Left            =   4680
               TabIndex        =   200
               Top             =   180
               Width           =   540
            End
            Begin VB.Label Label26 
               Caption         =   "Date OnSet"
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
               Left            =   3720
               TabIndex        =   199
               Top             =   180
               Width           =   900
            End
            Begin VB.Label Label26 
               Caption         =   "Final Diag"
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
               Left            =   120
               TabIndex        =   198
               Top             =   180
               Width           =   1620
            End
         End
         Begin VB.Frame Frame18 
            Enabled         =   0   'False
            Height          =   1365
            Left            =   5880
            TabIndex        =   175
            Top             =   4140
            Width           =   5400
            Begin VB.CommandButton CmdDiagRepudiation 
               Caption         =   ">"
               Height          =   255
               Index           =   3
               Left            =   1800
               TabIndex        =   216
               Top             =   180
               Width           =   225
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
               Left            =   75
               Style           =   1  'Simple Combo
               TabIndex        =   74
               Top             =   480
               Width           =   5250
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
               Left            =   75
               Style           =   1  'Simple Combo
               TabIndex        =   75
               Top             =   720
               Width           =   5250
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
               Left            =   75
               Style           =   1  'Simple Combo
               TabIndex        =   76
               Top             =   975
               Width           =   5250
            End
            Begin VB.Label Label26 
               Caption         =   "Repudiation Resons"
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
               Left            =   120
               TabIndex        =   201
               Top             =   180
               Width           =   1620
            End
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
            Left            =   1305
            Sorted          =   -1  'True
            TabIndex        =   24
            Top             =   360
            Width           =   5000
         End
         Begin MSMask.MaskEdBox txtLGTime 
            Height          =   285
            Left            =   9720
            TabIndex        =   40
            Top             =   1110
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
            Left            =   7680
            TabIndex        =   41
            Top             =   1360
            Width           =   3375
         End
         Begin VB.ComboBox cboFolUp 
            Height          =   315
            Left            =   7680
            TabIndex        =   42
            Top             =   1650
            Width           =   840
         End
         Begin VB.TextBox txtRoomNBoardDays 
            Height          =   300
            Left            =   9480
            TabIndex        =   43
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtMBMPolYr 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7680
            TabIndex        =   44
            Top             =   1920
            Width           =   840
         End
         Begin VB.CheckBox chkDiag 
            Height          =   255
            Left            =   7320
            TabIndex        =   166
            Top             =   3480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtPatientAge 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9480
            TabIndex        =   45
            Top             =   1920
            Width           =   615
         End
         Begin VB.ComboBox cboYearDay 
            Height          =   315
            Left            =   10065
            TabIndex        =   46
            Top             =   1920
            Width           =   990
         End
         Begin VB.TextBox txtHOSName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   25
            Top             =   650
            Width           =   5000
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
            Height          =   540
            Left            =   1305
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   900
            Width           =   5000
         End
         Begin VB.TextBox txtREFBy 
            Height          =   285
            Left            =   1305
            MaxLength       =   100
            TabIndex        =   27
            Top             =   1395
            Width           =   5000
         End
         Begin VB.ComboBox cboNature 
            Height          =   315
            Left            =   1300
            TabIndex        =   28
            Text            =   "Ill - Illness"
            Top             =   1650
            Width           =   2000
         End
         Begin VB.ComboBox CboStayType 
            Height          =   315
            Left            =   4320
            TabIndex        =   29
            Top             =   1650
            Width           =   2000
         End
         Begin VB.ComboBox cboModality 
            Height          =   315
            Left            =   1300
            TabIndex        =   30
            Top             =   1920
            Width           =   2000
         End
         Begin VB.ComboBox CboPayType 
            Height          =   315
            Left            =   4320
            TabIndex        =   31
            Top             =   1920
            Width           =   2000
         End
         Begin VB.ComboBox cboMaternityStatus 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmCaseAdmission.frx":0158
            Left            =   1320
            List            =   "frmCaseAdmission.frx":015A
            TabIndex        =   313
            Top             =   2210
            Width           =   2000
         End
         Begin VB.ComboBox cboOPChemo 
            Height          =   315
            ItemData        =   "frmCaseAdmission.frx":015C
            Left            =   4320
            List            =   "frmCaseAdmission.frx":015E
            TabIndex        =   318
            Top             =   2210
            Width           =   2000
         End
         Begin VB.Label Label2 
            Caption         =   "OP Chemo"
            Height          =   375
            Left            =   3480
            TabIndex        =   319
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lblAgent 
            BackColor       =   &H0000FFFF&
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   6480
            TabIndex        =   317
            Top             =   2280
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.Label Label19 
            Caption         =   "Maternity"
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
            Left            =   120
            TabIndex        =   314
            Top             =   2250
            Width           =   1005
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
            TabIndex        =   312
            Top             =   680
            Width           =   1140
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
            Index           =   0
            Left            =   8625
            TabIndex        =   310
            Top             =   1965
            Width           =   1095
         End
         Begin VB.Label Label26 
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
            Index           =   38
            Left            =   6600
            TabIndex        =   306
            Top             =   1960
            Width           =   900
         End
         Begin VB.Label Label26 
            Caption         =   "Copy Diag."
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
            Index           =   37
            Left            =   7920
            TabIndex        =   305
            Top             =   3480
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label26 
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
            Index           =   35
            Left            =   9120
            TabIndex        =   296
            Top             =   1125
            Width           =   1140
         End
         Begin VB.Label Label65 
            Caption         =   "R&&B Days"
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
            Index           =   120
            Left            =   8625
            TabIndex        =   249
            Top             =   1695
            Width           =   1095
         End
         Begin VB.Label Label26 
            Caption         =   "Fol. Up"
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
            Left            =   6600
            TabIndex        =   193
            Top             =   1660
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Fol. Req"
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
            Left            =   6600
            TabIndex        =   192
            Top             =   1400
            Width           =   1140
         End
         Begin VB.Label Label29 
            Caption         =   "1st. Treated"
            Height          =   255
            Index           =   0
            Left            =   6600
            TabIndex        =   191
            Top             =   660
            Width           =   1230
         End
         Begin VB.Label Label26 
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
            Index           =   10
            Left            =   120
            TabIndex        =   190
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label Label26 
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
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   189
            Top             =   420
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Reasons"
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
            Left            =   120
            TabIndex        =   188
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Ref By"
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
            Left            =   120
            TabIndex        =   187
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Nat of Clm"
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
            Left            =   120
            TabIndex        =   186
            Top             =   1700
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Stay Type"
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
            Left            =   3480
            TabIndex        =   185
            Top             =   1700
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Treat. Mode"
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
            Left            =   120
            TabIndex        =   184
            Top             =   1980
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Cash Type"
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
            Left            =   3480
            TabIndex        =   183
            Top             =   1980
            Width           =   1140
         End
         Begin VB.Label Label26 
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
            Index           =   18
            Left            =   6600
            TabIndex        =   182
            Top             =   150
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "1st. Call"
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
            Left            =   6600
            TabIndex        =   181
            Top             =   420
            Width           =   1140
         End
         Begin VB.Label Label26 
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
            Index           =   20
            Left            =   9120
            TabIndex        =   180
            Top             =   420
            Width           =   1140
         End
         Begin VB.Label Label26 
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
            Index           =   21
            Left            =   9120
            TabIndex        =   179
            Top             =   660
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Rsvr Amt"
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
            Index           =   22
            Left            =   6600
            TabIndex        =   178
            Top             =   920
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "LG Amt "
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
            Left            =   9120
            TabIndex        =   177
            Top             =   920
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Rcur"
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
            Left            =   6600
            TabIndex        =   176
            Top             =   1150
            Width           =   1140
         End
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5750
         Left            =   -74880
         TabIndex        =   150
         Top             =   330
         Width           =   11415
         Begin VB.Frame FrameSOPCS 
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
            Height          =   2000
            Left            =   2805
            TabIndex        =   293
            Top             =   120
            Width           =   1095
            Begin VB.TextBox txtProcCode1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   78
               Top             =   420
               Width           =   975
            End
            Begin VB.CommandButton cmdOPCS 
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
               Height          =   210
               Index           =   1
               Left            =   360
               TabIndex        =   294
               Top             =   180
               Width           =   615
            End
            Begin VB.TextBox txtProcCode2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   80
               Top             =   680
               Width           =   975
            End
            Begin VB.TextBox txtProcCode3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   82
               Top             =   920
               Width           =   975
            End
            Begin VB.TextBox txtProcCode4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   84
               Top             =   1160
               Width           =   975
            End
            Begin VB.TextBox txtProcCode5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   86
               Top             =   1400
               Width           =   975
            End
            Begin VB.TextBox txtProcCode6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   88
               Top             =   1670
               Width           =   975
            End
         End
         Begin VB.Frame FrameSMMARate 
            Caption         =   "Surg MMA"
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
            Height          =   2000
            Left            =   3945
            TabIndex        =   285
            Top             =   120
            Width           =   1335
            Begin VB.TextBox txtSTotalMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   292
               Top             =   2120
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox txtSMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   291
               Top             =   440
               Width           =   1215
            End
            Begin VB.TextBox txtSMMARate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   290
               Top             =   680
               Width           =   1215
            End
            Begin VB.TextBox txtSMMARate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   289
               Top             =   920
               Width           =   1215
            End
            Begin VB.TextBox txtSMMARate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   288
               Top             =   1160
               Width           =   1215
            End
            Begin VB.TextBox txtSMMARate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   287
               Top             =   1400
               Width           =   1215
            End
            Begin VB.TextBox txtSMMARate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   286
               Top             =   1670
               Width           =   1215
            End
         End
         Begin VB.Frame FrameAnaesMMA 
            Caption         =   "Anaes MMA"
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
            Height          =   2000
            Left            =   5280
            TabIndex        =   277
            Top             =   120
            Width           =   1335
            Begin VB.TextBox txtTotalAnaesMMARate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   120
               MaxLength       =   8
               TabIndex        =   284
               Top             =   2120
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox txtAnaesMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   283
               Top             =   440
               Width           =   1215
            End
            Begin VB.TextBox txtAnaesMMARate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   282
               Top             =   680
               Width           =   1215
            End
            Begin VB.TextBox txtAnaesMMARate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   281
               Top             =   920
               Width           =   1215
            End
            Begin VB.TextBox txtAnaesMMARate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   280
               Top             =   1160
               Width           =   1215
            End
            Begin VB.TextBox txtAnaesMMARate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   279
               Top             =   1400
               Width           =   1215
            End
            Begin VB.TextBox txtAnaesMMARate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   278
               Top             =   1670
               Width           =   1215
            End
         End
         Begin VB.Frame FrameNSMMARate 
            Caption         =   "N-Surg MMA"
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
            Height          =   2000
            Left            =   5280
            TabIndex        =   269
            Top             =   2100
            Width           =   1335
            Begin VB.TextBox txtTotalNSMMARate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   276
               Top             =   2070
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox txtNSMMARate1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   275
               Top             =   440
               Width           =   1215
            End
            Begin VB.TextBox txtNSMMARate2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   274
               Top             =   680
               Width           =   1215
            End
            Begin VB.TextBox txtNSMMARate3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   273
               Top             =   920
               Width           =   1215
            End
            Begin VB.TextBox txtNSMMARate4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   272
               Top             =   1160
               Width           =   1215
            End
            Begin VB.TextBox txtNSMMARate5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   271
               Top             =   1400
               Width           =   1215
            End
            Begin VB.TextBox txtNSMMARate6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   40
               MaxLength       =   8
               TabIndex        =   270
               Top             =   1640
               Width           =   1215
            End
         End
         Begin VB.Frame FrameNSOPCS 
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
            Height          =   2000
            Left            =   3960
            TabIndex        =   267
            Top             =   2100
            Width           =   1335
            Begin VB.TextBox txtNSProcCode1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   90
               Top             =   440
               Width           =   1215
            End
            Begin VB.CommandButton cmdOPCS 
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
               Height          =   210
               Index           =   0
               Left            =   360
               TabIndex        =   268
               Top             =   200
               Width           =   615
            End
            Begin VB.TextBox txtNSProcCode2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   92
               Top             =   680
               Width           =   1215
            End
            Begin VB.TextBox txtNSProcCode3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   94
               Top             =   920
               Width           =   1215
            End
            Begin VB.TextBox txtNSProcCode4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   96
               Top             =   1160
               Width           =   1215
            End
            Begin VB.TextBox txtNSProcCode5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   98
               Top             =   1400
               Width           =   1215
            End
            Begin VB.TextBox txtNSProcCode6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   6
               TabIndex        =   100
               Top             =   1670
               Width           =   1215
            End
         End
         Begin VB.Frame Frame21 
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
            Height          =   1050
            Left            =   6620
            TabIndex        =   152
            Top             =   1440
            Width           =   4695
            Begin VB.TextBox txtMEDSURDone1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   107
               Top             =   215
               Width           =   4455
            End
            Begin VB.TextBox txtMEDSURDone2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   108
               Top             =   480
               Width           =   4455
            End
            Begin VB.TextBox txtMEDSURDone3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   109
               Top             =   710
               Width           =   4455
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
            Height          =   1290
            Index           =   0
            Left            =   6620
            TabIndex        =   151
            Top             =   120
            Width           =   4695
            Begin VB.CommandButton cmdInvList 
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
               Height          =   255
               Left            =   3840
               TabIndex        =   153
               Top             =   150
               Width           =   735
            End
            Begin VB.TextBox txtPROCINVDone1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   101
               Top             =   435
               Width           =   3375
            End
            Begin VB.TextBox txtPROCINVDone2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   103
               Top             =   675
               Width           =   3375
            End
            Begin VB.TextBox txtPROCINVDone3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   150
               TabIndex        =   105
               Top             =   930
               Width           =   3375
            End
            Begin VB.TextBox txtINVCode1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3480
               MaxLength       =   150
               TabIndex        =   102
               Top             =   435
               Width           =   1095
            End
            Begin VB.TextBox txtINVCode2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3480
               MaxLength       =   150
               TabIndex        =   104
               Top             =   675
               Width           =   1095
            End
            Begin VB.TextBox txtINVCode3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3480
               MaxLength       =   150
               TabIndex        =   106
               Top             =   930
               Width           =   1095
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
            Height          =   1120
            Index           =   1
            Left            =   6620
            TabIndex        =   265
            Top             =   2520
            Width           =   4695
            Begin VB.TextBox txtAnaesDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   120
               MaxLength       =   30
               TabIndex        =   112
               Top             =   750
               Width           =   4485
            End
            Begin VB.TextBox txtAnaesDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   120
               MaxLength       =   30
               TabIndex        =   111
               Top             =   495
               Width           =   4485
            End
            Begin VB.TextBox txtAnaesDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   120
               MaxLength       =   30
               TabIndex        =   110
               Top             =   240
               Width           =   4485
            End
         End
         Begin VB.Frame FrameSProc 
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
            Height          =   2000
            Left            =   120
            TabIndex        =   295
            Top             =   120
            Width           =   2655
            Begin VB.TextBox txtProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   77
               Top             =   440
               Width           =   2535
            End
            Begin VB.TextBox txtProcedure2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   79
               Top             =   680
               Width           =   2535
            End
            Begin VB.TextBox txtProcedure3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   81
               Top             =   920
               Width           =   2535
            End
            Begin VB.TextBox txtProcedure4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   83
               Top             =   1160
               Width           =   2535
            End
            Begin VB.TextBox txtProcedure5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   85
               Top             =   1400
               Width           =   2535
            End
            Begin VB.TextBox txtProcedure6 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   87
               Top             =   1670
               Width           =   2535
            End
         End
         Begin VB.Frame FrameNSProc 
            Caption         =   "Non-Surgical Charges"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2000
            Left            =   120
            TabIndex        =   266
            Top             =   2100
            Width           =   3825
            Begin VB.TextBox txtNSProcedure1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   89
               Top             =   440
               Width           =   3735
            End
            Begin VB.TextBox txtNSProcedure2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   91
               Top             =   680
               Width           =   3735
            End
            Begin VB.TextBox txtNSProcedure3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   93
               Top             =   920
               Width           =   3735
            End
            Begin VB.TextBox txtNSProcedure4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   95
               Top             =   1160
               Width           =   3735
            End
            Begin VB.TextBox txtNSProcedure5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   97
               Top             =   1400
               Width           =   3735
            End
            Begin VB.TextBox txtNSProcedure6 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   40
               MaxLength       =   200
               TabIndex        =   99
               Top             =   1670
               Width           =   3735
            End
         End
         Begin VB.Frame Frame3 
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
            Height          =   1635
            Left            =   120
            TabIndex        =   158
            Top             =   4070
            Width           =   8175
            Begin VB.CommandButton cmdloadDoc 
               Caption         =   "Load Doc"
               Height          =   255
               Left            =   6720
               TabIndex        =   327
               Top             =   240
               Width           =   975
            End
            Begin VB.ComboBox txtATTDoc 
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   326
               Top             =   1200
               Width           =   6255
            End
            Begin VB.ComboBox txtATTDoc 
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   325
               Top             =   960
               Width           =   6255
            End
            Begin VB.ComboBox txtATTDoc 
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   324
               Top             =   720
               Width           =   6255
            End
            Begin VB.ComboBox txtATTDoc 
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   323
               Top             =   480
               Width           =   6255
            End
            Begin VB.ComboBox txtATTDoc 
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   322
               Top             =   240
               Width           =   6255
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1890
         Left            =   -74880
         TabIndex        =   136
         Top             =   360
         Width           =   11415
         Begin VB.TextBox txtINSCode 
            BackColor       =   &H80000001&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   14400
            TabIndex        =   138
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtAdd1 
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   13
            Top             =   135
            Width           =   4815
         End
         Begin VB.TextBox txtAdd2 
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   14
            Top             =   390
            Width           =   4815
         End
         Begin VB.TextBox txtAdd3 
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   15
            Top             =   645
            Width           =   4815
         End
         Begin VB.TextBox txtPostCode 
            Height          =   285
            Left            =   840
            MaxLength       =   5
            TabIndex        =   16
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox txtCity 
            Height          =   285
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   17
            Top             =   900
            Width           =   2895
         End
         Begin VB.TextBox txtState 
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   18
            Top             =   1155
            Width           =   4815
         End
         Begin MSMask.MaskEdBox txtDOB 
            Height          =   285
            Left            =   7200
            TabIndex        =   20
            Top             =   1560
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMPolNo 
            Height          =   285
            Left            =   840
            MaxLength       =   25
            TabIndex        =   19
            Top             =   1410
            Visible         =   0   'False
            Width           =   4815
         End
         Begin VB.TextBox txtAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7080
            TabIndex        =   21
            Top             =   165
            Width           =   1935
         End
         Begin VB.TextBox txtAvailableAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7080
            TabIndex        =   22
            Top             =   420
            Width           =   1935
         End
         Begin VB.Label Label26 
            Caption         =   "Avail. Lmt"
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
            Left            =   6120
            TabIndex        =   164
            Top             =   435
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Ann Lmt"
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
            Left            =   6120
            TabIndex        =   163
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Pol Num"
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
            Index           =   4
            Left            =   120
            TabIndex        =   162
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Label Label26 
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
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   161
            Top             =   1200
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Town"
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
            Left            =   2280
            TabIndex        =   160
            Top             =   930
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "PCode"
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
            TabIndex        =   159
            Top             =   880
            Width           =   1140
         End
         Begin VB.Label Label10 
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
            Height          =   255
            Left            =   13200
            TabIndex        =   139
            Top             =   1245
            Width           =   1275
         End
         Begin VB.Label Label26 
            Caption         =   "Add"
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
            TabIndex        =   137
            Top             =   195
            Width           =   1140
         End
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
         Height          =   5265
         Left            =   -74160
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   223
         Top             =   720
         Width           =   8895
      End
      Begin MSComctlLib.ListView lstPreUndIllness 
         Height          =   5595
         Left            =   -74880
         TabIndex        =   224
         Top             =   420
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9869
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
      Begin MSComctlLib.ListView lstMBMHis 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   230
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9763
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mem No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pol No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Exp Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Insured Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Plan Code"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstNewBenefitDetails 
         Height          =   5355
         Left            =   -69120
         TabIndex        =   231
         Top             =   660
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   9446
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
      Begin MSComctlLib.ListView lstBenefitDetails 
         Height          =   5355
         Left            =   -74880
         TabIndex        =   233
         Top             =   660
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   9446
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
      Begin MSMask.MaskEdBox txtDOC 
         Height          =   300
         Left            =   -65640
         TabIndex        =   234
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin MSComctlLib.ListView lstMember 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   245
         Top             =   2280
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6588
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
      Begin MSComctlLib.ListView lstSurgProc 
         Height          =   5475
         Left            =   -74880
         TabIndex        =   247
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9657
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
      Begin MSFlexGridLib.MSFlexGrid Fgrid 
         Height          =   2835
         Left            =   -74820
         TabIndex        =   248
         Top             =   420
         Width           =   3315
         _ExtentX        =   5847
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
      Begin VB.Frame Frame5 
         Height          =   2265
         Left            =   -74880
         TabIndex        =   253
         Top             =   960
         Width           =   3015
         Begin VB.TextBox txtREDAmt 
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
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   118
            Top             =   195
            Width           =   1300
         End
         Begin MSMask.MaskEdBox txtDateRECDoc 
            Height          =   300
            Left            =   3840
            TabIndex        =   254
            Top             =   1920
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
         Begin VB.ComboBox cboTopUpLmt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   119
            Text            =   "N - No"
            Top             =   450
            Width           =   1335
         End
         Begin VB.ComboBox cboCASInd 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   120
            Top             =   705
            Width           =   1335
         End
         Begin VB.ComboBox cboQuery 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   121
            Top             =   990
            Width           =   1335
         End
         Begin VB.ComboBox cboPayPending 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   122
            Text            =   "N - No"
            Top             =   1280
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtFirstDocDate 
            Height          =   315
            Left            =   1560
            TabIndex        =   123
            Top             =   1560
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
            Left            =   1560
            TabIndex        =   124
            Top             =   1850
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
         Begin VB.Label Label98 
            Caption         =   "Date Rec. Doc."
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
            Left            =   3600
            TabIndex        =   262
            Top             =   1920
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label26 
            Caption         =   "Red Amt"
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
            Left            =   120
            TabIndex        =   261
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Topup Lmt"
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
            Index           =   49
            Left            =   120
            TabIndex        =   260
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Special App"
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
            Index           =   50
            Left            =   120
            TabIndex        =   259
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label Label26 
            Caption         =   "Query Status"
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
            Index           =   51
            Left            =   120
            TabIndex        =   258
            Top             =   1005
            Width           =   1140
         End
         Begin VB.Label Label26 
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
            Index           =   53
            Left            =   120
            TabIndex        =   257
            Top             =   1320
            Width           =   1380
         End
         Begin VB.Label Label29 
            Caption         =   "Subseq. Doc Date"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   256
            Top             =   1880
            Width           =   1335
         End
         Begin VB.Label Label26 
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
            Index           =   70
            Left            =   120
            TabIndex        =   255
            Top             =   1600
            Width           =   1260
         End
      End
      Begin VB.Frame Frame8 
         Height          =   2265
         Left            =   -71760
         TabIndex        =   263
         Top             =   960
         Width           =   8295
         Begin VB.TextBox TxtLGRemark 
            Appearance      =   0  'Flat
            Height          =   1845
            Left            =   120
            MaxLength       =   100
            ScrollBars      =   1  'Horizontal
            TabIndex        =   125
            Top             =   360
            Width           =   8055
         End
         Begin VB.Label Label26 
            Caption         =   "LG Remark"
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
            Index           =   36
            Left            =   120
            TabIndex        =   264
            Top             =   120
            Width           =   780
         End
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
         Height          =   5265
         Left            =   -74160
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   297
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
         Height          =   5265
         Left            =   -74160
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   303
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
         Height          =   5265
         Left            =   -74160
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   304
         Top             =   720
         Width           =   8895
      End
      Begin MSComDlg.CommonDialog CmnDlg 
         Left            =   -70200
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label26 
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
         Index           =   9
         Left            =   -74880
         TabIndex        =   321
         Top             =   3240
         Width           =   2340
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
         Left            =   -64920
         TabIndex        =   302
         Top             =   3480
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
         Left            =   -64680
         TabIndex        =   301
         Top             =   3480
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
         Left            =   -64440
         TabIndex        =   300
         Top             =   3480
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
         Left            =   -64200
         TabIndex        =   299
         Top             =   3480
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
         Left            =   -63960
         TabIndex        =   298
         Top             =   3480
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
         Left            =   -64800
         TabIndex        =   244
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label46 
         Caption         =   "D.O.C"
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
         Left            =   -66240
         TabIndex        =   236
         ToolTipText     =   "Date of Admission"
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label26 
         Caption         =   "New Benefit"
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
         Index           =   54
         Left            =   -69120
         TabIndex        =   235
         Top             =   420
         Width           =   1140
      End
      Begin VB.Label Label34 
         Caption         =   "Principal"
         Height          =   255
         Left            =   -74880
         TabIndex        =   229
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Supplementary"
         Height          =   255
         Left            =   -74880
         TabIndex        =   228
         Top             =   3135
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "Date"
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
         Index           =   55
         Left            =   -74160
         TabIndex        =   222
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label26 
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
         Height          =   255
         Index           =   56
         Left            =   -73080
         TabIndex        =   221
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label26 
         Caption         =   "Person Incharge"
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
         Index           =   57
         Left            =   -66840
         TabIndex        =   220
         Top             =   480
         Width           =   1500
      End
   End
   Begin VB.Frame Frame14 
      Height          =   480
      Left            =   120
      TabIndex        =   140
      Top             =   7920
      Width           =   11655
      Begin VB.CommandButton cmdDocReqDis 
         Caption         =   "Doc Req Discharge"
         Height          =   285
         Left            =   4200
         TabIndex        =   309
         Top             =   150
         Width           =   1575
      End
      Begin VB.CommandButton cmdDocReqAdm 
         Caption         =   "Doc Req Adm"
         Height          =   285
         Left            =   3000
         TabIndex        =   308
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdTHDrugs 
         Caption         =   "&T.Home Drugs"
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
         Left            =   1800
         TabIndex        =   243
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrior 
         Caption         =   "&Notification"
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
         Left            =   840
         TabIndex        =   156
         Top             =   150
         Width           =   975
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
         TabIndex        =   145
         Top             =   150
         Width           =   735
      End
      Begin VB.TextBox txtINSCode1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   144
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtHLTCode 
         Height          =   285
         Left            =   2040
         TabIndex        =   143
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
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
         Left            =   10800
         TabIndex        =   142
         Top             =   150
         Width           =   660
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
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
         Left            =   10155
         TabIndex        =   141
         Top             =   150
         Width           =   660
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "&Register"
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
         TabIndex        =   133
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label26 
         Caption         =   "Case No"
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
         Left            =   6000
         TabIndex        =   307
         Top             =   195
         Width           =   660
      End
      Begin VB.Label txtCaseNo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6600
         TabIndex        =   154
         Top             =   180
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   490
      Left            =   120
      TabIndex        =   134
      Top             =   220
      Width           =   11655
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   0
         Top             =   150
         Width           =   2055
      End
      Begin VB.ComboBox cboPending 
         Height          =   315
         Left            =   10800
         TabIndex        =   5
         Text            =   "N - No"
         Top             =   120
         Width           =   735
      End
      Begin VB.ComboBox cboStatus1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5640
         TabIndex        =   3
         Top             =   135
         Width           =   975
      End
      Begin VB.ComboBox cboClaimability 
         Height          =   315
         Left            =   7560
         TabIndex        =   4
         Text            =   "AA - Accepted"
         Top             =   120
         Width           =   2535
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Height          =   315
         Left            =   3120
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   315
         Left            =   3840
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label26 
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
         Index           =   68
         Left            =   10200
         TabIndex        =   212
         Top             =   165
         Width           =   1140
      End
      Begin VB.Label Label26 
         Caption         =   "Claimability"
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
         Left            =   6720
         TabIndex        =   211
         Top             =   165
         Width           =   1140
      End
      Begin VB.Label Label26 
         Caption         =   "Cas Status"
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
         Left            =   4680
         TabIndex        =   210
         Top             =   165
         Width           =   1140
      End
      Begin VB.Label Label26 
         Caption         =   "MBM Num"
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
         Index           =   65
         Left            =   250
         TabIndex        =   209
         Top             =   150
         Width           =   1140
      End
   End
   Begin VB.Frame Frame4 
      Height          =   795
      Left            =   120
      TabIndex        =   135
      Top             =   720
      Width           =   11655
      Begin VB.TextBox txtMBMName 
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   6
         Top             =   160
         Width           =   2535
      End
      Begin MSMask.MaskEdBox txtMBMEffDate 
         Height          =   315
         Left            =   6720
         TabIndex        =   11
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtRenewal 
         Height          =   285
         Left            =   4365
         MaxLength       =   1
         TabIndex        =   9
         Top             =   165
         Width           =   375
      End
      Begin MSMask.MaskEdBox txtMBMExpiryDate 
         Height          =   315
         Left            =   8640
         TabIndex        =   12
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   8
         Top             =   420
         Width           =   2535
      End
      Begin VB.TextBox txtTakeOver 
         Height          =   285
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   10
         Top             =   165
         Width           =   375
      End
      Begin VB.TextBox txtPAYCode 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   780
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSMask.MaskEdBox txtMBMTopUpJoinDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   8520
         TabIndex        =   250
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPolEffDate 
         Height          =   315
         Left            =   4680
         TabIndex        =   315
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label26 
         Caption         =   "Policy Eff Date"
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
         Left            =   3600
         TabIndex        =   316
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label lblTopupStatus 
         BackColor       =   &H00C0C0C0&
         Caption         =   "This MBMship has a Topup Policy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7680
         TabIndex        =   252
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "TopUp Join Date"
         Height          =   195
         Index           =   1
         Left            =   6240
         TabIndex        =   251
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label Label26 
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
         Index           =   64
         Left            =   7920
         TabIndex        =   208
         Top             =   420
         Width           =   1140
      End
      Begin VB.Label Label26 
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
         Index           =   63
         Left            =   5880
         TabIndex        =   207
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label26 
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
         Index           =   62
         Left            =   5880
         TabIndex        =   206
         Top             =   195
         Width           =   780
      End
      Begin VB.Label Label26 
         Caption         =   "Renewal"
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
         Left            =   3600
         TabIndex        =   205
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label Label26 
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
         Index           =   60
         Left            =   240
         TabIndex        =   204
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label26 
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
         Index           =   59
         Left            =   240
         TabIndex        =   203
         Top             =   810
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label26 
         Caption         =   "Name"
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
         Index           =   58
         Left            =   240
         TabIndex        =   202
         Top             =   200
         Width           =   1140
      End
      Begin VB.Label txtPolStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   155
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label63 
         Caption         =   "T/Over"
         Height          =   255
         Left            =   4800
         TabIndex        =   147
         Top             =   195
         Width           =   825
      End
      Begin VB.Label lblWaitingPeriod 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Waiting Period"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   9855
         TabIndex        =   149
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblExpired 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Expired"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   9870
         TabIndex        =   148
         Top             =   420
         Visible         =   0   'False
         Width           =   1680
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Case Registration"
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
      TabIndex        =   157
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmCaseAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim LG As String
Dim intLastPoz2 As Integer
Dim REPReasons As Integer
Dim FirstDiag As Integer
Dim FinalDiag As Integer
Dim UDI As Integer
Dim SaveOthers As Integer
Dim HOS As Integer
Dim LGNumber As String
Dim Letter As String
Dim EXCNumber As String
Dim RB As String
Dim txtPath As String
Dim Check As Boolean
Dim ScanCode As String
Dim PrevDOA As String
Dim HOspitalName As String
Dim Relationship As String
Dim AgeBand   As String
Dim txtPLNCode As String
Dim MBMCPLNCode As String
Dim TotalFinalDiag As String
Dim TotalProc As String
Public Document, Personnel, Attention As String
Public Underlying As Boolean
Public CASUnderlying As String
Dim PreIllCAT As String
Dim PreIllCATOthers As String
Dim UndIllCAT As String
Dim UndIllCATOthers As String
Dim medixmasterconnPat As New ADODB.Connection
Dim PatientCovID As String
Dim MMASurgFees As Double
Dim MMAAnaeFees As Double
Dim tmpPLNPerVisit As String
Dim tmpPlnCode As String
Dim HOSArr As Collection
Dim tmpRBType As String
Dim tmpMBMTermination As Boolean
Dim tmpMBMTopupInd As Boolean
Dim tmpLGType As String
Dim tmpGRPCompany As String
Dim tempLGTime As String
Dim tempLGSysTime As String
Dim tmpEmails As String
Dim tmpSenderEmail As String
Dim tmpSender As String
Dim tmpBDay As Boolean
Dim tmpMBMDOB As String
Dim tmpMBMCDOB As String
Dim tmpStartAge As String
Dim tmpEndAge As String
Dim StartAgeStatus As Boolean
Dim EndAgeStatus As Boolean
Dim AgeLenght As String
Dim AgeCat As String
Dim AgeOnly As Integer
Dim StartPatientAge As Boolean
Dim EndPatientAge As Boolean
Dim tmpChkAnnBalLmt As Double
Dim tmpChkLTBalLmt As Double
Dim LGStatus As Boolean
Dim tmpPolicyWording As String
Dim PictureFileName As String
Dim tmpHltCode As String
Dim tmpHospChk As String
Dim tmpMBMCostCenter As String
Dim tmpSuspendDate As String
Dim tmpAgentStatus As Boolean
Dim tmpAgentCode As String
Dim tmpAgentname As String
Dim tmpINSType As String
Dim ASOValid As Boolean

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

Private Sub cboCASInd_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCASInd, cboCASInd.SelStart) + Chr(KeyAscii) + Mid(cboCASInd, cboCASInd.SelStart + cboCASInd.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboCASInd.ListCount - 1
        
        If NewText = LCase(Left(cboCASInd.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCASInd.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboCASInd = ValidValue
               cboCASInd.SelStart = 0
               cboCASInd.SelLength = Len(ValidValue)
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

Private Sub cboMaternityStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboMaternityStatus.Text, cboMaternityStatus.SelStart) + Chr(KeyAscii) + Mid(cboMaternityStatus.Text, cboMaternityStatus.SelStart + cboMaternityStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboMaternityStatus.ListCount - 1
 
        If NewText = LCase(Left(cboMaternityStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboMaternityStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
               cboMaternityStatus.Text = ValidValue
               cboMaternityStatus.SelStart = 0
               cboMaternityStatus.SelLength = Len(ValidValue)
             End If
        End If

End Sub

Private Sub cboPatientList_Change()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim strsql As String
Dim strsqlCOV As String
Dim cmd As New ADODB.Command
Dim cnn As New ADODB.Command
Dim cov As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim tmpPreEIllness As String
Dim tmpSpecificIll As String
Dim tmpWaitingPeriod As String
Dim SelectAge As String

    'medixmasterconnPat.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    AgeBand = ""
    StartAgeStatus = False
    EndAgeStatus = False

    If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
        
        'StrSql = "Select MBMNumber, MBMICBcPp,MBMDOB from MBM where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        'MBM.Open StrSql, medixmasterconnPat, adOpenKeyset, adLockOptimistic
        strAppendCase = "1"
        strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Set cmd.ActiveConnection = medixmasterconnPat
        cmd.CommandText = "Case_Registration"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        MBM.CursorLocation = adUseClient
        MBM.CursorType = adOpenForwardOnly
        MBM.CacheSize = 100
        Set MBM = cmd.Execute
        
        'strsqlCOV = "Select MBMCDOB,MBMCICBCPPNo,MBMCRELCode,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
        'MBMCov.Open strsqlCOV, medixmasterconnPat, adOpenKeyset, adLockOptimistic
        strAppendCase = "2"
        strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
        Set cnn.ActiveConnection = medixmasterconnPat
        cnn.CommandText = "Case_Registration"
        cnn.CommandType = adCmdStoredProc
        cnn.CommandTimeout = 300
        Set Prm1 = cnn.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
        cnn.Parameters.Append Prm1
        Set Prm2 = cnn.CreateParameter("strAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cnn.Parameters.Append Prm2
        MBMCov.CursorLocation = adUseClient
        MBMCov.CursorType = adOpenForwardOnly
        MBMCov.CacheSize = 100
        Set MBMCov = cnn.Execute
        
        If MBM.EOF = True Then
            MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            txtPatientIC = MBM!MBMIcBcPp
            Relationship = "P"
            If Len(MBM!MBMDOB) Then
                AgeBand = (Format(txtMBMEffDate, "yyyy") - Format(MBM!MBMDOB, "yyyy"))
                tmpMBMDOB = MBM!MBMDOB
            End If
            
        End If
        
        If MBMCov.EOF = False Then
            txtPatientIC = MBMCov!MBMCICBCPPNo
            Relationship = MBMCov!MBMCRELCode
            
            If Len(MBMCov!MBMCDOB) Then
                AgeBand = (Format(txtMBMEffDate, "yyyy") - Format(MBMCov!MBMCDOB, "yyyy"))
                tmpMBMDOB = MBM!MBMDOB
            End If
            If Len(MBMCov!MBMCPLNCode) Then
                MBMCPLNCode = MBMCov!MBMCPLNCode
            End If

        End If
       
        'If Mid(cboPatientList, 1, 2) <> "00" Then
        '    cboSuppExc = cboPatientList
        'End If
       
        MBM.Close
        MBMCov.Close
        Set MBM = Nothing
        Set MBMCov = Nothing
        
        If Mid(cboPatientList, 1, 2) = "00" Then
            strAppendCase = "13"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Else
            strAppendCase = "3"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
        End If
        Set cov.ActiveConnection = medixmasterconnPat
        cov.CommandText = "Case_Registration"
        cov.CommandType = adCmdStoredProc
        cov.CommandTimeout = 300
        Set Prm1 = cov.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
        cov.Parameters.Append Prm1
        Set Prm2 = cov.CreateParameter("strAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cov.Parameters.Append Prm2
        MBMCov.CursorLocation = adUseClient
        MBMCov.CursorType = adOpenForwardOnly
        MBMCov.CacheSize = 100
        Set MBMCov = cov.Execute
        
        If MBMCov.EOF = False Then
            If Mid(cboPatientList, 1, 2) = "00" Then
                tmpPreEIllness = IIf(Len(MBMCov!MBMPreExistIllDate), MBMCov!MBMPreExistIllDate, "")
                tmpSpecificIll = IIf(Len(MBMCov!MBMSpecificIllDate), MBMCov!MBMSpecificIllDate, "")
                tmpWaitingPeriod = IIf(Len(MBMCov!MBMWaitingPeriodDate), MBMCov!MBMWaitingPeriodDate, "")
                If Trim(MBMCov!MBMBDayStatus) = "Y" Then
                    tmpBDay = True
                Else
                    tmpBDay = False
                End If
            Else
                tmpPreEIllness = IIf(Len(MBMCov!MBMCPreExistIllDate), MBMCov!MBMCPreExistIllDate, "")
                tmpSpecificIll = IIf(Len(MBMCov!MBMCSpecificIllDate), MBMCov!MBMCSpecificIllDate, "")
                tmpWaitingPeriod = IIf(Len(MBMCov!MBMCWaitingPeriodDate), MBMCov!MBMCWaitingPeriodDate, "")
                If Trim(MBMCov!MBMCBDayStatus) = "Y" Then
                    tmpBDay = True
                Else
                    tmpBDay = False
                End If
                
            End If
        End If
        MBMCov.Close
        Set MBMCov = Nothing
     End If
     
    If Mid(cboPatientList, 1, 2) = "00" Then
        If txtDateAdmitted <> "__/__/____" And tmpMBMDOB <> "__/__/____" Then
            SelectAge = GetAgeDays(CDate(txtDateAdmitted), CDate(tmpMBMDOB))
            Call GetAgeCat(SelectAge)
            If IsNumeric(tmpStartAge) And Trim(AgeCat) <> "Y" Then
                If tmpStartAge >= AgeOnly Then
                    StartAgeStatus = True
                Else
                    StartAgeStatus = False
                End If
            Else
                StartAgeStatus = False
            End If
            
            If IsNumeric(EndAgeStatus) And Trim(AgeCat) <> "D" Then
                If Trim(tmpEndAge) <> "" Then
                    If tmpEndAge < AgeOnly Then
                        EndAgeStatus = True
                    Else
                        EndAgeStatus = False
                    End If
                Else
                    EndAgeStatus = False
                End If
            Else
                EndAgeStatus = False
            End If

        End If
    Else
        If txtDateAdmitted <> "__/__/____" And tmpMBMCDOB <> "__/__/____" Then
            SelectAge = GetAgeDays(CDate(txtDateAdmitted), CDate(tmpMBMCDOB))
            Call GetAgeCat(SelectAge)
            
            If IsNumeric(tmpStartAge) And Trim(AgeCat) <> "Y" Then
                If tmpStartAge >= AgeOnly Then
                    StartAgeStatus = True
                Else
                    StartAgeStatus = False
                End If
            Else
                StartAgeStatus = False
            End If
            
            If IsNumeric(EndAgeStatus) And Trim(AgeCat) <> "D" Then
                If Trim(tmpEndAge) <> "" Then
                    If tmpEndAge < AgeOnly Then
                        EndAgeStatus = True
                    Else
                        EndAgeStatus = False
                    End If
                Else
                    EndAgeStatus = False
                End If
            Else
                EndAgeStatus = False
            End If
            
        End If
    End If
     
    If tmpBDay = True Then
        MsgBox "Please verify birthdate!", vbOKOnly + vbInformation
    End If

    If StartAgeStatus = True Then
        MsgBox "Please verify Patient's Age!", vbOKOnly + vbInformation
    End If

    If EndAgeStatus = True Then
        MsgBox "Please verify Patient's Age!", vbOKOnly + vbInformation
    End If

    '========== Add by CK =============================================
        If (Relationship = "S" And CDbl(AgeBand) >= "19") Or (Relationship = "D" And CDbl(AgeBand) >= "19") Then
            MsgBox "Please check employment status!", vbOKOnly + vbInformation
        End If
    '========== Add by CK =============================================
    If Len(tmpPreEIllness) Then
        If CDate(tmpPreEIllness) >= CDate(GetServerTime) Then
            MsgBox "Subject to Pre-Existing Illness until " & tmpPreEIllness & "!", vbOKOnly + vbInformation
        End If
    End If
    If Len(tmpSpecificIll) Then
        If CDate(tmpSpecificIll) >= CDate(GetServerTime) Then
            MsgBox "Subject to Specific Illness until " & tmpSpecificIll & "!", vbOKOnly + vbInformation
        End If
    End If
    If Len(tmpWaitingPeriod) Then
        If CDate(tmpWaitingPeriod) >= CDate(GetServerTime) Then
            MsgBox "Subject to Waiting Period until " & tmpWaitingPeriod & "!", vbOKOnly + vbInformation
        End If
    End If

    'tmpSpecificIll
    'tmpWaitingPeriod
    Call BNFListing
   ' medixmasterconnPat.Close
   ' Set medixmasterconnPat = Nothing
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

Private Sub cboTopUpLmt_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboTopUpLmt, cboTopUpLmt.SelStart) + Chr(KeyAscii) + Mid(cboTopUpLmt, cboTopUpLmt.SelStart + cboTopUpLmt.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboTopUpLmt.ListCount - 1
        
        If NewText = LCase(Left(cboTopUpLmt.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboTopUpLmt.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               cboTopUpLmt = ValidValue
               cboTopUpLmt.SelStart = 0
               cboTopUpLmt.SelLength = Len(ValidValue)
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

Private Sub cboYearDay_LostFocus()
Dim tmpPatientAge As String
    StartPatientAge = False
    EndPatientAge = False
    
    If IsNumeric(txtPatientAge) And Mid(cboYearDay, 1, 1) <> "Y" And cboYearDay <> "" Then
        If Mid(cboYearDay, 1, 1) = "W" Then
            tmpPatientAge = txtPatientAge * 7
        Else
            tmpPatientAge = txtPatientAge * 30
        End If
        
        If CDbl(tmpStartAge) >= CDbl(tmpPatientAge) Then
            StartPatientAge = True
        Else
            StartPatientAge = False
        End If
        
    Else
        StartPatientAge = False
    End If

    If IsNumeric(txtPatientAge) And Mid(cboYearDay, 1, 1) = "Y" And cboYearDay <> "" Then
        If Trim(tmpEndAge) <> "" Then
            If CDbl(tmpEndAge) < CDbl(txtPatientAge) Then
                EndPatientAge = True
            Else
                EndPatientAge = False
            End If
        Else
            EndPatientAge = False
        End If
    Else
        EndPatientAge = False
    End If
    
    If StartPatientAge = True Then
        MsgBox "Please verify Patient's Age!", vbOKOnly + vbInformation
    End If

    If EndPatientAge = True Then
        MsgBox "Please verify Patient's Age!", vbOKOnly + vbInformation
    End If

End Sub

Private Sub chkDiag_Click()

    If chkDiag.Value = 1 Then
        txtFinalDiag1 = txtPrelimDiag1
        txtFinalDiag2 = txtPrelimDiag2
        txtFinalDiag3 = txtPrelimDiag3
        txtFDDOS1 = txtPDDOS1
        txtFDDOS2 = txtPDDOS2
        txtFDDOS3 = txtPDDOS3
        cboFDICDCode1 = cboICDCode1
        cboFDICDCode2 = cboICDCode2
        cboFDICDCode3 = cboICDCode3
    Else
        txtFinalDiag1 = ""
        txtFinalDiag2 = ""
        txtFinalDiag3 = ""
        txtFDDOS1 = "__/__/____"
        txtFDDOS2 = "__/__/____"
        txtFDDOS3 = "__/__/____"
        cboFDICDCode1 = ""
        cboFDICDCode2 = ""
        cboFDICDCode3 = ""
    End If

End Sub

Private Sub ChkDOA_Click()
    If ChkDOA.Value = 0 And ChkDOD.Value = 0 Then
        cmdRegister.Enabled = True
    End If
End Sub

Private Sub ChkDOD_Click()
    If ChkDOA.Value = 0 And ChkDOD.Value = 0 Then
        cmdRegister.Enabled = True
    End If
End Sub

Private Sub CmdDiagFinal_Click(a As Integer)
        frmSearchICD.Options = 2
        frmSearchICD.optRep.Visible = False
        frmSearchICD.Source = 2
        frmSearchICD.Show vbModal
End Sub

Private Sub CmdDiagPre_Click(a As Integer)
    'If a = "3" Then
    '    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
    '        FrmSearchRepudiation.Source = 2
    '        FrmSearchRepudiation.show vbModal
    '    End If
    'Else
        frmSearchICD.Options = 1
        frmSearchICD.optRep.Visible = False
        frmSearchICD.Source = 2
        frmSearchICD.Show vbModal
    'End If
End Sub

Private Sub CmdDiagUD_Click(a As Integer)
    'If a = "3" Then
    '    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
    '        FrmSearchRepudiation.Source = 2
    '        FrmSearchRepudiation.show vbModal
    '    End If
    'Else
        frmSearchICD.Options = 3
        frmSearchICD.optRep.Visible = False
        frmSearchICD.Source = 2
        frmSearchICD.Show vbModal
    'End If
End Sub

Private Sub CmdDiagRepudiation_Click(a As Integer)
    'If a = "3" Then
    '    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
    '        FrmSearchRepudiation.Source = 2
    '        FrmSearchRepudiation.show vbModal
    '    End If
    'Else
    If a = "3" Then
            If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
            FrmSearchRepudiation.Source = 1
            FrmSearchRepudiation.Show vbModal
            End If
    End If
    'End If
End Sub


Private Sub cmdDocReqAdm_Click()
Dim RecordSaved
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
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
    
            Letter = "A"
            Call funcword
        End If
    End If
  '  medixmasterconn.Close
  ''  Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
End Sub

Private Sub cmdDocReqDis_Click()
Dim RecordSaved
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
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
        End If
    End If
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End Sub

Private Sub cmdInvList_Click()
    frmSearchInv.Source = 2
    frmSearchInv.Show vbModal
End Sub

Private Sub cmdLoadDoc_Click()
Call BindDoc
End Sub

Private Sub cmdOPCS_Click(Index As Integer)
    If Index = 1 Then
        frmSearchCPT.Source = 4
    Else
        frmSearchCPT.Source = 5
    End If
        frmSearchCPT.Show vbModal
End Sub

Private Sub cmdPrior_Click()

    cmdPrior.Enabled = False
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If IsDate(txtMBMExpiryDate) Then
        If Month(txtMBMExpiryDate) = Month(GetServerTime) And Year(txtMBMExpiryDate) = Year(GetServerTime) And CDate(txtMBMExpiryDate) >= CDate(GetServerTime) Then
            Letter = "P"
            Call funcword
        ElseIf CDate(txtMBMExpiryDate) < CDate(GetServerTime) Then
            Letter = "E"
            Call funcword
        End If
    End If
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    cmdPrior.Enabled = True
End Sub

Private Sub cmdTHDrugs_Click()
Dim RecordSaved
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
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
            Letter = "T"
            Call funcword
        End If
    End If
 '   medixmasterconn.Close
   ' Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    'End If
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

    Document = ""
    Personnel = ""
    Attention = ""
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
        If cboREPCode1 = "" Then
            MsgBox "Please Select Repudiation reason", vbOKOnly + vbCritical, DialogBoxTitle
            Frame18.Enabled = True
            SSTab1.Tab = 1
            cboREPCode1.SetFocus
        Else
            Frame18.Enabled = True
        End If
        frmRepudiationSelection.Document = Document
        frmRepudiationSelection.Personnel = Personnel
        frmRepudiationSelection.Attention = Attention
        frmRepudiationSelection.Source = 2
        frmRepudiationSelection.Show vbModal
    Else
        Frame18.Enabled = False
        cmdLG.Enabled = True
    End If
    'medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End Sub

Private Sub cboClaimability_Click()

    Document = ""
    Personnel = ""
    Attention = ""
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
        If cboREPCode1 = "" Then
            MsgBox "Please Select Repudiation reason", vbOKOnly + vbCritical, DialogBoxTitle
            Frame18.Enabled = True
            SSTab1.Tab = 1
            cboREPCode1.SetFocus
        Else
            Frame18.Enabled = True
        End If
        frmRepudiationSelection.Document = Document
        frmRepudiationSelection.Personnel = Personnel
        frmRepudiationSelection.Attention = Attention
        frmRepudiationSelection.Source = 2
        frmRepudiationSelection.Show vbModal
    Else
        Frame18.Enabled = False
        cmdLG.Enabled = True
    End If
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
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
        MsgBox "Treament Mode cannot exceed 30 characters", vbOKOnly + vbCritical, DialogBoxTitle
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
        MsgBox "Claim Nature cannot exceed 30 characters", vbOKOnly + vbCritical, DialogBoxTitle
        cboNature.SetFocus
    End If
End Sub

Private Sub cboPatientList_Click()
Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim strsql As String
Dim strsqlCOV As String
Dim cmd As New ADODB.Command
Dim cnn As New ADODB.Command
Dim cov As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim tmpPreEIllness As String
Dim tmpSpecificIll As String
Dim tmpWaitingPeriod As String
Dim SelectAge As String

  '  medixmasterconnPat.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    AgeBand = ""
    StartAgeStatus = False
    EndAgeStatus = False

    If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 2))) = True Then
        
        'StrSql = "Select MBMNumber, MBMICBcPp,MBMDOB from MBM where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        'MBM.Open StrSql, medixmasterconnPat, adOpenKeyset, adLockOptimistic
        strAppendCase = "1"
        strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Registration"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        MBM.CursorLocation = adUseClient
        MBM.CursorType = adOpenForwardOnly
        MBM.CacheSize = 100
        Set MBM = cmd.Execute
        
        'strsqlCOV = "Select MBMCDOB,MBMCICBCPPNo,MBMCRELCode,MBMCPLNCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
        'MBMCov.Open strsqlCOV, medixmasterconnPat, adOpenKeyset, adLockOptimistic
        strAppendCase = "2"
        strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
        Set cnn.ActiveConnection = medixmasterconn
        cnn.CommandText = "Case_Registration"
        cnn.CommandType = adCmdStoredProc
        cnn.CommandTimeout = 300
        Set Prm1 = cnn.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
        cnn.Parameters.Append Prm1
        Set Prm2 = cnn.CreateParameter("strAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cnn.Parameters.Append Prm2
        MBMCov.CursorLocation = adUseClient
        MBMCov.CursorType = adOpenForwardOnly
        MBMCov.CacheSize = 100
        Set MBMCov = cnn.Execute
        
        If MBM.EOF = True Then
            MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            txtPatientIC = MBM!MBMIcBcPp
            Relationship = "P"
            If Len(MBM!MBMDOB) Then
                AgeBand = (Format(txtMBMEffDate, "yyyy") - Format(MBM!MBMDOB, "yyyy"))
                tmpMBMDOB = MBM!MBMDOB
            End If
            
        End If
        
        If MBMCov.EOF = False Then
            txtPatientIC = MBMCov!MBMCICBCPPNo
            Relationship = MBMCov!MBMCRELCode
            
            If Len(MBMCov!MBMCDOB) Then
                AgeBand = (Format(txtMBMEffDate, "yyyy") - Format(MBMCov!MBMCDOB, "yyyy"))
                tmpMBMCDOB = MBMCov!MBMCDOB
            End If
            If Len(MBMCov!MBMCPLNCode) Then
                MBMCPLNCode = MBMCov!MBMCPLNCode
            End If

        End If
       
        'If Mid(cboPatientList, 1, 2) <> "00" Then
        '    cboSuppExc = cboPatientList
        'End If
       
        MBM.Close
        MBMCov.Close
        Set MBM = Nothing
        Set MBMCov = Nothing
        
        If Mid(cboPatientList, 1, 2) = "00" Then
            strAppendCase = "13"
            strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Else
            strAppendCase = "3"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
        End If
        Set cov.ActiveConnection = medixmasterconn
        cov.CommandText = "Case_Registration"
        cov.CommandType = adCmdStoredProc
        cov.CommandTimeout = 300
        Set Prm1 = cov.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
        cov.Parameters.Append Prm1
        Set Prm2 = cov.CreateParameter("strAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cov.Parameters.Append Prm2
        MBMCov.CursorLocation = adUseClient
        MBMCov.CursorType = adOpenForwardOnly
        MBMCov.CacheSize = 100
        Set MBMCov = cov.Execute
        
        If MBMCov.EOF = False Then
            If Mid(cboPatientList, 1, 2) = "00" Then
                tmpPreEIllness = IIf(Len(MBMCov!MBMPreExistIllDate), MBMCov!MBMPreExistIllDate, "")
                tmpSpecificIll = IIf(Len(MBMCov!MBMSpecificIllDate), MBMCov!MBMSpecificIllDate, "")
                tmpWaitingPeriod = IIf(Len(MBMCov!MBMWaitingPeriodDate), MBMCov!MBMWaitingPeriodDate, "")
                If Trim(MBMCov!MBMBDayStatus) = "Y" Then
                    tmpBDay = True
                Else
                    tmpBDay = False
                End If
            Else
                tmpPreEIllness = IIf(Len(MBMCov!MBMCPreExistIllDate), MBMCov!MBMCPreExistIllDate, "")
                tmpSpecificIll = IIf(Len(MBMCov!MBMCSpecificIllDate), MBMCov!MBMCSpecificIllDate, "")
                tmpWaitingPeriod = IIf(Len(MBMCov!MBMCWaitingPeriodDate), MBMCov!MBMCWaitingPeriodDate, "")
                If Trim(MBMCov!MBMCBDayStatus) = "Y" Then
                    tmpBDay = True
                Else
                    tmpBDay = False
                End If
                
            End If
        End If
        MBMCov.Close
        Set MBMCov = Nothing
     End If
     
    If Mid(cboPatientList, 1, 2) = "00" Then
        If txtDateAdmitted <> "__/__/____" And tmpMBMDOB <> "__/__/____" Then
            SelectAge = GetAgeDays(CDate(txtDateAdmitted), CDate(tmpMBMDOB))
            Call GetAgeCat(SelectAge)
            If IsNumeric(tmpStartAge) And Trim(AgeCat) <> "Y" Then
                If tmpStartAge >= AgeOnly Then
                    StartAgeStatus = True
                Else
                    StartAgeStatus = False
                End If
            Else
                StartAgeStatus = False
            End If
            
            If IsNumeric(EndAgeStatus) And Trim(AgeCat) <> "D" Then
                If Trim(tmpEndAge) <> "" Then
                    If tmpEndAge < AgeOnly Then
                        EndAgeStatus = True
                    Else
                        EndAgeStatus = False
                    End If
                Else
                    EndAgeStatus = False
                End If
            Else
                EndAgeStatus = False
            End If

        End If
    Else
        If txtDateAdmitted <> "__/__/____" And tmpMBMCDOB <> "__/__/____" Then
            
            SelectAge = GetAgeDays(CDate(txtDateAdmitted), CDate(tmpMBMCDOB))
            Call GetAgeCat(SelectAge)
            If IsNumeric(tmpStartAge) And Trim(AgeCat) <> "Y" Then
                If tmpStartAge >= AgeOnly Then
                    StartAgeStatus = True
                Else
                    StartAgeStatus = False
                End If
            Else
                StartAgeStatus = False
            End If
            
            If IsNumeric(EndAgeStatus) And Trim(AgeCat) <> "D" Then
                If Trim(tmpEndAge) <> "" Then
                    If tmpEndAge < AgeOnly Then
                        EndAgeStatus = True
                    Else
                        EndAgeStatus = False
                    End If
                Else
                    EndAgeStatus = False
                End If
            Else
                EndAgeStatus = False
            End If
            
        End If
    End If
     
    If tmpBDay = True Then
        MsgBox "Please verify birthdate!", vbOKOnly + vbInformation
    End If

    If StartAgeStatus = True Then
        MsgBox "Please verify Patient's Age!", vbOKOnly + vbInformation
    End If

    If EndAgeStatus = True Then
        MsgBox "Please verify Patient's Age!", vbOKOnly + vbInformation
    End If

    '========== Add by CK =============================================
        If (Relationship = "S" And CDbl(AgeBand) >= "19") Or (Relationship = "D" And CDbl(AgeBand) >= "19") Then
            MsgBox "Please check employment status!", vbOKOnly + vbInformation
        End If
    '========== Add by CK =============================================
    If Len(tmpPreEIllness) Then
        If CDate(tmpPreEIllness) >= CDate(GetServerTime) Then
            MsgBox "Subject to Pre-Existing Illness until " & tmpPreEIllness & "!", vbOKOnly + vbInformation
        End If
    End If
    If Len(tmpSpecificIll) Then
        If CDate(tmpSpecificIll) >= CDate(GetServerTime) Then
            MsgBox "Subject to Specific Illness until " & tmpSpecificIll & "!", vbOKOnly + vbInformation
        End If
    End If
    If Len(tmpWaitingPeriod) Then
        If CDate(tmpWaitingPeriod) >= CDate(GetServerTime) Then
            MsgBox "Subject to Waiting Period until " & tmpWaitingPeriod & "!", vbOKOnly + vbInformation
        End If
    End If

    'tmpSpecificIll
    'tmpWaitingPeriod
    Call BNFListing
 '   medixmasterconnPat.Close
 '   Set medixmasterconnPat = Nothing
End Sub

Private Sub cboPatientList_GotFocus()
    SSTab1.TabIndex = 1
End Sub

Private Sub cboPatientList_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    'KeyAscii = KeyCharacter(KeyAscii)
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
Private Sub CboStayType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboStayType, CboStayType.SelStart) + Chr(KeyAscii) + Mid(CboStayType, CboStayType.SelStart + CboStayType.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboStayType.ListCount - 1
        
        If NewText = LCase(Left(CboStayType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboStayType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboStayType = ValidValue
               CboStayType.SelStart = 0
               CboStayType.SelLength = Len(ValidValue)
            End If
        End If
End Sub
Private Sub CboPayType_LostFocus()
    Dim Delimiter As Integer
    Dim ClaimCode As String
    Delimiter = InStr(1, CboPayType, "-")
    If Delimiter = 0 Then
        If Len(Trim(CboPayType)) > 2 Then
            MsgBox "Claim Type cannot exceed 2 character!", vbCritical
            CboPayType.SetFocus
        End If
    Else
        ClaimCode = Mid(Trim(CboPayType), 1, Delimiter - 1)
        If Len(Trim(ClaimCode)) > 2 Then
            MsgBox "Claim Type cannot exceed 2 character!", vbCritical
            CboPayType.SetFocus
        End If
    End If
End Sub

Private Sub cboREPCode1_Change()
    If cboREPCode1 <> "" Then
        cmdLG.Enabled = False
    End If
End Sub

Private Sub cboREPCode1_GotFocus()
    If cboREPCode1 <> "" Then
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
    
    For i = 0 To cboREPCode2.ListCount - 1
        
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


Private Sub cboSuppExc_Click()
Dim MBMCov As New ADODB.Recordset
Dim cnn As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    If txtMBMNumber <> "" And IsNumeric(Trim(Mid(cboSuppExc, 1, 2))) Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'StrSql = "Select MBMCNumber , MBMCCoverID , MBMCExclusion  from MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Trim(Mid(cboSuppExc, 1, 2)) & "'"
        'MBMCov.Open StrSql, medixmasterconn, adOpenKeyset, adLockPessimistic
        strAppendCase = "3"
        strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Trim(Mid(cboSuppExc, 1, 2)) & "'"
        Set cnn.ActiveConnection = medixmasterconn
        cnn.CommandText = "Case_Registration"
        cnn.CommandType = adCmdStoredProc
        cnn.CommandTimeout = 300
        Set Prm1 = cnn.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
        cnn.Parameters.Append Prm1
        Set Prm2 = cnn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cnn.Parameters.Append Prm2
        MBMCov.CursorLocation = adUseClient
        MBMCov.CursorType = adOpenForwardOnly
        MBMCov.CacheSize = 100
        Set MBMCov = cnn.Execute
        
        If MBMCov.EOF = False Then
            If Trim(MBMCov!MBMCExclusion) <> "" Then
                txtSuppExc = MBMCov!MBMCExclusion
            End If
        End If
        MBMCov.Close
        Set MBMCov = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
End Sub

Private Sub cmdLG_Click()
Dim RecordSaved
Dim CASLG As New ADODB.Recordset
Dim PLN As New ADODB.Recordset
Dim strsqlPLN As String
Dim strsqlBTBG As String
Dim startTrans
Dim ChkVer As String
Dim TempSysDate As String
Dim TempLGAmt As String
Dim strsqlCASLG As String
Dim strsql As String
Dim rstPLN As New ADODB.Recordset
Dim RecordSave

    tmpChkAnnBalLmt = 0
    tmpChkLTBalLmt = 0
    tmpLGType = ""
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
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
    ElseIf Trim(txtCASLGAmt) = "" Then
        MsgBox "Please Enter Reserve Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASLGAmt.SetFocus
    ElseIf txtLGTime = "__:__" Then
        MsgBox "Please Enter LG Time", vbOKOnly + vbCritical, DialogBoxTitle
        txtLGTime.SetFocus
    Else
    
        If txtPAYCode = "AX" Then
            If txtPolEffDate <> txtMBMEffDate Then
                RecordSave = MsgBox("Have you checked Take Over clause?", vbYesNo + vbExclamation)
                If RecordSave = vbNo Then
                    medixmasterconn.Close
                    Set medixmasterconn = Nothing
                    medixmasterconnMaint.Close
                    Set medixmasterconnMaint = Nothing
                    medixmasterconnWS.Close
                    Set medixmasterconnWS = Nothing
                    Exit Sub
                End If
            End If
        End If
        
        Call CheckLimit
        If LGStatus = True Then
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
            
                Letter = ""
                Letter = "L"
                LGNumber = GetLGCode(ChkStr(txtCaseNo))
                TempSysDate = Format(GetServerTime, "yyyy/mm/dd")
                TempLGAmt = CDbl(IIf(IsNumeric(txtCASLGAmt) = True, txtCASLGAmt, 0))
                If txtLGTime <> "__:__" Then
                    tempLGTime = txtLGTime
                Else
                    tempLGTime = Empty
                End If
                tempLGSysTime = Format(Time, "hh:mm")
                strsql = "Select PLNCode,PLNLGType from PLN where PLNCode = '" & Trim(txtPLNCode) & "' And PayCode = '" & Trim(txtPAYCode) & "'"
                rstPLN.Open strsql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                
                Letter = rstPLN!PLNLGType
                rstPLN.Close
                Set rstPLN = Nothing
                
                If Trim(Letter) <> "C" Then
                    strsqlPLN = "EXEC Case_Registration '" & Trim(txtMBMNumber) & "',4"
                    PLN.Open strsqlPLN, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                    If Not PLN.EOF Then
                        Letter = "B"
                        LGNumber = "B" & Mid(LGNumber, 2, Len(LGNumber))
                    End If
                    PLN.Close
                    Set PLN = Nothing
                End If
                
                ChkVer = ""
                ChkVer = Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-")))
                strsqlCASLG = "EXEC Case_AdjInsertCASLG '" & ChkStr(txtCaseNo) & "', '" & Trim(Mid(LGNumber, InStr(1, LGNumber, "-") + 1, Len(LGNumber) - InStr(1, LGNumber, "-"))) & "', '" & TempSysDate & "', " & TempLGAmt & ", '" & Logon & "', '" & ChkStr(Trim(TxtLGRemark)) & "', '" & Trim(Letter) & "', '" & Trim(tmpPLNPerVisit) & "', '" & Trim(tmpRBType) & "', '" & tempLGTime & "', '" & tempLGSysTime & "'"
                medixmasterconnMaint.Execute strsqlCASLG
                
                Call funcword
                 Call MobileNotification(txtMBMNumber, LGNumber)
                cmdLG.Enabled = False
            End If
        Else
            MsgBox "Please Check Balance Limit", vbOKOnly + vbCritical, DialogBoxTitle
            cmdLG.Enabled = False
        End If
    End If
   
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    cmdLG.Enabled = False
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
    rst2.MoveNext
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
    objDom.async = False
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub

Private Sub Form_Load()
Dim LGRemarks As String

Call DoInitialSettings
LGRemarks = "2)  Co-Payment - In the event the Room and Board limit is exceeded, the Insurance will pay only 80% of other eligible costs." & _
            "3)  Insurance will pay only 80% of MRI, CT Scan and Heart Scan charges." & _
            "4)  Medical Information relating to the patient and treatment must be truthfully disclosed by the attending physician." & _
            "5)  The following items are not covered: -" & _
                    "(a) Personal nature such as telephone calls, newspaper, laundry, television etc." & _
                    "(b) Non-Medical items." & _
                    "(c) Medical Report Charges." & _
                    "(d) Registration Fees, Admission Fees, Billing Fees, Dispensing Fees And Admission Kit." & _
                    "(e) Appliances, Prosthesis - Clutches, Lenses, Pacemaker, Braces, Corset, Arm sling and Cervical Collar." & _
                    "(f) Lodger" & _
                    "(g) Supplies & Services not related to illness injury." & _
            "6)  All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees (4th Edition)."
   
    Call ComboListing
    cboPatientList.Visible = False
    txtPatientList.Visible = True
    txtCASTakeOverAmt.Enabled = False
    Check = False
    ChkDOA.Value = 1
    ChkDOD.Value = 1
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


Public Sub BindDoc()
Dim sqlQry As String
Dim Doc As New ADODB.Recordset

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
Dim ICD As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim SRV As New ADODB.Recordset
Dim Nature As New ADODB.Recordset
Dim Modality As New ADODB.Recordset
Dim ICDCAT As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim PAYMode As New ADODB.Recordset
Dim PAYTYPE As New ADODB.Recordset
Dim STAYTYPE As New ADODB.Recordset
Dim CASRep As New ADODB.Recordset
Dim SPEAPP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim tmpHosName As String
Dim tmpHOSCode As String
Dim HOSIndex As Integer






    
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchSpecApp"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    SPEAPP.CursorLocation = adUseClient
    SPEAPP.CursorType = adOpenForwardOnly
    SPEAPP.CacheSize = 100
    Set SPEAPP = cmd.Execute
 
    Do Until SPEAPP.EOF
        With SPEAPP
            cboCASInd.AddItem !SpecAppCode & " - " & !SpecAppDesc
            SPEAPP.MoveNext
        End With
    Loop
    SPEAPP.Close
    Set SPEAPP = Nothing
    
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
            CboStayType.AddItem !PayTypeCode & " - " & Trim(!PayTypeName)
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
            tmpHosName = HOS!HosName
            txtHOSCode.AddItem Trim(!HosName)
            HOSArr.Add tmpHOSCode, tmpHosName
            HOS.MoveNext
        End With
    Loop
    HOS.Close
    Set HOS = Nothing
    
    
   ' medixmasterconn.Close
   '' Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing

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
    
    cboPending.AddItem "Y - Yes"
    cboPending.AddItem "N - No"
    
    cboCondwillOcc.AddItem "Y - YES"
    cboCondwillOcc.AddItem "N - NO"
    
    cboFLWTreatment.AddItem "Y - YES"
    cboFLWTreatment.AddItem "N - NO"
    
    cboTopUpLmt.AddItem "N - No"
    cboTopUpLmt.AddItem "Y - Yes"
    
    cboQuery.AddItem "N - No"
    cboQuery.AddItem "Y - Yes"
    
    cboFolUp.AddItem "Y - Yes"
    cboFolUp.AddItem "N - No"
          
    With cboPayPending
        .AddItem "Y - Yes"
        .AddItem "N - No"
    End With
    
    cboCASTakeOveInd.AddItem "Y - Yes"
    cboCASTakeOveInd.AddItem "N - No"
    
    cboIntExtTO.AddItem "I - Internal"
    cboIntExtTO.AddItem "E - External"
    
    cboYearDay.AddItem "W - Weeks"
    cboYearDay.AddItem "M - Months"
    cboYearDay.AddItem "Y - Year"
    
    cboMaternityStatus.AddItem "N - No"
    cboMaternityStatus.AddItem "D - Normal Delivery"
    cboMaternityStatus.AddItem "C - Caesarean"

    cboOPChemo.AddItem "Y - Yes"
    cboOPChemo.AddItem "N - No"
End Sub

Private Sub BNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim bnf As ListItem
    Dim sql As String
    Dim medixmasterconnBNF As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
    lstBenefitDetails.ListItems.Clear
    
   ' medixmasterconnBNF.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    If Len(Trim(txtHLTCode)) <> 0 _
     And Len(Trim(txtINSCode1)) <> 0 And Len(Trim(txtPLNCode)) <> 0 Then
        
        
        If Len(MBMCPLNCode) And Mid(cboPatientList, 1, 2) <> "00" Then
            strAppendCase = "5"
            strAppend = " AND PLNCOde ='" & Trim(MBMCPLNCode) & "'" & _
                    " AND HLTCode ='" & Trim(txtHLTCode) & "'" & _
                    " AND INSCODe ='" & Trim(txtINSCode1) & "'"
            'sql = " select BNFCode, BNFNoOFDays , " & _
              " BNFclaimableAmount, BNFdescription, BNFRMPerDis " & _
              " from View_BNF where PLNCOde ='" & Trim(MBMCPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(txtHLTCode) & "'" & _
              " AND INSCODe ='" & Trim(txtINSCode1) & "'"
        Else
            strAppendCase = "5"
            strAppend = " AND PLNCOde ='" & Trim(txtPLNCode) & "'" & _
                    " AND HLTCode ='" & Trim(txtHLTCode) & "'" & _
                    " AND INSCODe ='" & Trim(txtINSCode1) & "'"
            'sql = " select BNFCode, BNFNoOFDays , " & _
              " BNFclaimableAmount, BNFdescription, BNFRMPerDis " & _
              " from View_BNF where PLNCOde ='" & Trim(txtPLNCode) & "'" & _
              " AND HLTCode ='" & Trim(txtHLTCode) & "'" & _
              " AND INSCODe ='" & Trim(txtINSCode1) & "'"
        End If
        'rst2.Open sql, medixmasterconnBNF, adOpenKeyset, adLockOptimistic
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Registration"
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
            Set bnf = lstBenefitDetails.ListItems.Add(, , rst2!BNFCode)
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
                    If IsNumeric(rst2!BNFClaimAbleAmount) = False Then
                        RB = 0
                    Else
                        RB = rst2!BNFClaimAbleAmount
                    End If
                End If
                rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
    End If
   ' medixmasterconnBNF.Close
   ' Set medixmasterconnBNF = Nothing
End Sub

Private Sub CmdClear_Click()
    txtMBMNumber.Enabled = True
    cmdRegister.Enabled = False
    cmdClear.Caption = "&Clear"
    SSTab1.Tab = 0
    Call ClearEntries
    Check = False
    cboPatientList.Visible = False
    txtPatientList.Visible = True
    cmdLG.Enabled = False
    cboSuppExc.Clear
    Frame4.Enabled = True
    ChkDOA.Value = 1
    ChkDOD.Value = 1
    txtMBMNumber.SetFocus
    Document = ""
    Personnel = ""
    Attention = ""
End Sub

'  ############### Search Button ################
Private Sub CmdSearch_Click()
    On Error Resume Next
    'lblTopupStatus.Visible = False
    Call ClearEntries
    'Call SetSupp(False)
    frmSearchMBM.Source = 4
    frmSearchMBM.Show vbModal
End Sub

' ####################################
Private Sub EnableEntries(status As Boolean)
   txtMBMName.ForeColor = vbBlack
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = status
        End If
    Next ctl
End Sub

Private Sub ClearEntries()
Dim ctl As Control
Dim s As Integer

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        End If
    Next ctl
        txtHOSCode = ""
        txtCASReserveAmount = 1600
        txtCASLGAmt = 2500
        cboNature = "ILL - ILLNess"
        cboModality = "S - Surgical"
        cboPatientList.Clear
        cboICDCode1 = ""
        cboICDCode2 = ""
        cboICDCode3 = ""
        cboUDICDCode1 = ""
        cboUDICDCode2 = ""
        cboUDICDCode3 = ""
        cboFDICDCode1 = ""
        cboFDICDCode2 = ""
        cboFDICDCode3 = ""
        cboREPCode1 = ""
        cboREPCode2 = ""
        cboREPCode3 = ""
        cboStatus1 = "O - Open"
        cboClaimability = "AA - Accepted"
        cboTopUpLmt = "N - No"
        cboCASTakeOveInd = ""
        lstSurgProc.ListItems.Clear
        Fgrid.col = 1
        For s = 1 To 10
            Fgrid.row = (s)
        Fgrid = ""
        Next
        txtLGTime.mask = ""
        txtLGTime.Text = ""
        txtLGTime.mask = "##:##"
        lstPolicyWording.Clear
End Sub

' ############################ Buttons

Private Sub cmdRegister_Click()
'On Error GoTo errSave

Dim RecordSaved
Dim X As Boolean
Dim Y As Boolean
    X = False
    Y = False
    cmdRegister.Enabled = False
         Call checkASOPaln(ASOValid)
         
         If ASOValid = False Then
            RecordSaved = MsgBox("ASO Limit already Exceed. Still you want to Register !!!", vbYesNo + vbExclamation)
            
            Else
            RecordSaved = MsgBox("You Want register case", vbYesNo + vbExclamation)
         End If
         
            If RecordSaved = vbYes Then
                Call NewCaseAddmission
            End If
         cmdLG.Visible = True
        If tmpMBMTermination = True Then
            cmdLG.Enabled = False
        Else
            cmdLG.Enabled = True
        End If
    cmdRegister.Enabled = True
    Exit Sub
End Sub
Public Sub checkASOPaln(IsValid As Boolean)

Dim MBMNumber As String
Dim CoverId As String
Dim ASO As String

Dim sqlQry As String
Dim Plnrs As New ADODB.Recordset

Dim ASOLimit As Double
Dim IPUtilized As Double
Dim OPUtilized As Double
Dim IsCombined As String

sqlQry = ""
MBMNumber = txtMBMNumber
CoverId = Mid(cboPatientList, 1, 2)

If CoverId = "00" Then
    sqlQry = "Select PLN.ASOInd from MBM inner join PLN on PLN.PLNCode = MBM.PLNCode where MBM.MBMNumber ='" + MBMNumber + "'"
Else
    sqlQry = "select PLN.ASOInd  from MBMCoveredPersons inner join PLN on PLN.PLNCode = MBMCoveredPersons.MBMCPLNCode where MBMCNumber ='" + MBMNumber + "' and MBMCCoverID ='" + CoverId + "'"
End If

Plnrs.Open sqlQry, medixmasterconn, adOpenKeyset, adLockOptimistic
Do While Not Plnrs.EOF
        ASO = IIf(Len(Plnrs!ASOInd), Plnrs!ASOInd, "")
    Plnrs.MoveNext
Loop
Plnrs.Close
Set Plnrs = Nothing

If ASO = "Y" Then

    sqlQry = ""
    sqlQry = "Select ISNULL(MBMASOLimit,0) as ASOLimit,ISNULL (MBMIPLimit,0) as IPLimit,ISNULL(MBMOPLimit,0) as OPLimit ,ISNULL (MBMIPUtilized,0) as IPUtilized ,ISNULL (MBMOPUtilized,0) as OPUtlized , IsIPOPCombined from Tbl_ASOMemberDetails where MBMNumber ='" + MBMNumber + "' and MBMCoverID ='" + CoverId + "'"
    Plnrs.Open sqlQry, medixmasterconn, adOpenKeyset, adLockPessimistic
    Do While Not Plnrs.EOF
            ASOLimit = CDbl(Plnrs!ASOLimit)
            IPUtilized = CDbl(Plnrs!IPUtilized)
            OPUtilized = CDbl(Plnrs!OPUtlized)
            IsCombined = IIf(Len(Plnrs!IsIPOPCombined), (Plnrs!IsIPOPCombined), "")
        Plnrs.MoveNext
    Loop
    Plnrs.Close
    Set Plnrs = Nothing
    If IsCombined = "Y" Then
        If IPUtilized + OPUtilized < ASOLimit Then
            IsValid = True
        Else
            IsValid = False
        End If
    Else
        If IPUtilized < ASOLimit Then
            IsValid = True
        Else
            IsValid = False
        End If
    End If
Else
    IsValid = True
End If
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

Private Sub CmdExit_Click()
    Unload Me
    Set frmCaseAdmission = Nothing
End Sub

Public Sub MailCDOMsg(msg As String)
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object

Dim strTableBeg As String
Dim strTableEnd As String
Dim strFntNormal As String
Dim strFntHeader As String
Dim strFntEnd As String
Dim strMessage As String
Dim Body As String

                 Body = ""
                 strTableBeg = "<table border=1>"
                 strTableEnd = "</table>"
                 strFntHeader = "<font size=2 face=" & Chr(34) & "Arial" & Chr(34) & "><b> <tr bgcolor=lightblue><td>Agent</td> <td>Case No.</td> <td>DOA</td> <td>Patient Name</td> <td>Policy No</td> <td>Eff Date</td>  <td>Hospital</td> </tr></b></font>"
                 strFntNormal = "<font color=black face=" & Chr(34) & "Arial" & Chr(34) & " size=1>"
                 strFntEnd = "</font>"
                 strMessage = ""
                 strMessage = strTableBeg & strFntNormal & strFntHeader
                 strMessage = strMessage & "<tr></tr> <tr><td> " & msg & "</td><td> " & txtCaseNo & "</td><td> " & txtDateAdmitted & "</td><td> " & cboPatientList & "</td><td> " & txtMBMPolNo & "</td><td> " & txtMBMEffDate & "</td><td> " & HOspitalName & "</td></tr>"
                 strMessage = strMessage & strFntEnd & strTableEnd
                
                    Body = "Dear Sir/Madam,</p></p>" & vbNewLine & vbNewLine & _
                            "</p></p>" & strMessage & vbNewLine & "</p></p>" & _
                            "</p> <br> Thank you.</p></p>" & vbNewLine & _
                            "</p> <br> Regards, </p></p> " & vbNewLine & _
                            "MediExpress (Malaysia) Sdn Bhd"
                            
On Error GoTo debugs

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "Agent Alert Details"
'objMessage.From = "adm@medix.com.my"
objMessage.From = "notifications@medix.com.my"
objMessage.To = "powling.sin@zurich.com.my;zulina.anuar@zurich.com.my"
objMessage.BCC = "trinath@medix.com.my"
objMessage.TextBody = Body
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
'objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "adm@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "notifications@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "medix123"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update
objMessage.Send
debugs:
If Err.Description <> "" Then MsgBox Err.Description

End Sub


Private Sub SentEmail(msg As String)
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object
On Error GoTo debugs


Dim strTableBeg As String
Dim strTableEnd As String
Dim strFntNormal As String
Dim strFntHeader As String
Dim strFntEnd As String
Dim strMessage As String
Dim Body As String

                 Body = ""
                 strTableBeg = "<table border=1>"
                 strTableEnd = "</table>"
                 strFntHeader = "<font size=2 face=" & Chr(34) & "Arial" & Chr(34) & "><b> <tr bgcolor=lightblue><td>Agent</td> <td>Case No.</td> <td>DOA</td> <td>Patient Name</td> <td>Policy No</td> <td>Eff Date</td>  <td>Hospital</td> </tr></b></font>"
                 strFntNormal = "<font color=black face=" & Chr(34) & "Arial" & Chr(34) & " size=1>"
                 strFntEnd = "</font>"
                 strMessage = ""
                 strMessage = strTableBeg & strFntNormal & strFntHeader
                 strMessage = strMessage & "<tr></tr> <tr><td> " & msg & "</td><td> " & txtCaseNo & "</td><td> " & txtDateAdmitted & "</td><td> " & cboPatientList & "</td><td> " & txtMBMPolNo & "</td><td> " & txtMBMEffDate & "</td><td> " & HOspitalName & "</td></tr>"
                 strMessage = strMessage & strFntEnd & strTableEnd
                
                    Body = "Dear Sir/Madam,</p></p>" & vbNewLine & vbNewLine & _
                            "</p></p>" & strMessage & vbNewLine & "</p></p>" & _
                            "</p> <br> Thank you.</p></p>" & vbNewLine & _
                            "</p> <br> Regards, </p></p> " & vbNewLine & _
                            "MediExpress (Malaysia) Sdn Bhd"
                            

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "Alert... Agent Details "
objMessage.From = "notifications@medix.com.my"
objMessage.To = "powling.sin@zurich.com.my;zulina.anuar@zurich.com.my"
objMessage.BCC = "trinath@medix.com.my; bryan@medix.com.my"
objMessage.HTMLBody = Body
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



'Private Sub SentEmail(msg As String)
'On Err GoTo ErrorSave
'Screen.MousePointer = vbHourglass
'Dim oOApp As Outlook.Application
'Dim oOMail As Outlook.MailItem ''
'Dim strTableBeg As String
'Dim strTableEnd As String
'Dim strFntNormal As String
'Dim strFntHeader As String
'Dim strFntEnd As String
'Dim strMessage As String
'Dim Body As String'

'                 Body = ""
'                 strTableBeg = "<table border=1>"
'                strTableEnd = "</table>"
'                 strFntHeader = "<font size=2 face=" & Chr(34) & "Arial" & Chr(34) & "><b> <tr bgcolor=lightblue><td>Agent</td> <td>Case No.</td> <td>DOA</td> <td>Patient Name</td> <td>Policy No</td> <td>Eff Date</td>  <td>Hospital</td> </tr></b></font>"
'                 strFntNormal = "<font color=black face=" & Chr(34) & "Arial" & Chr(34) & " size=1>"
'                 strFntEnd = "</font>"
'                 strMessage = ""
'                 strMessage = strTableBeg & strFntNormal & strFntHeader
'                 strMessage = strMessage & "<tr></tr> <tr><td> " & msg & "</td><td> " & txtCaseNo & "</td><td> " & txtDateAdmitted & "</td><td> " & cboPatientList & "</td><td> " & txtMBMPolNo & "</td><td> " & txtMBMEffDate & "</td><td> " & HOspitalName & "</td></tr>"
'                 strMessage = strMessage & strFntEnd & strTableEnd
'
'                    Body = "Dear Sir/Madam,</p></p>" & vbNewLine & vbNewLine & _
'                            "</p></p>" & strMessage & vbNewLine & "</p></p>" & _
'                            "</p> <br> Thank you.</p></p>" & vbNewLine & _
'                            "</p> <br> Regards, </p></p> " & vbNewLine & _
'                            "MediExpress (Malaysia) Sdn Bhd"
'
'
'                        Set oOApp = CreateObject("Outlook.Application")
'                        Set oOMail = oOApp.CreateItem(olMailItem)
'                        With oOMail
'                            '.To = "powling.sin@zurich.com.my;zulina.anuar@zurich.com.my"
'                            .To = "trinath@medix.com.my;bryan@medix.com.my"
'                            .Subject = "Alert... Agent Details "
'                            .BodyFormat = olFormatHTML
'                            .HTMLBody = "<HTML><BODY>" & strFntNormal & Body & " </BODY></HTML>"
'                            .Send
'                       End With
'    Screen.MousePointer = vbDefault
'    'MsgBox "Emails have been sent out successfully!", vbOKOnly + vbCritical
'    Exit Sub
'ErrorSave:
'        Screen.MousePointer = vbDefault
''        MediConnCISPayment.Close
''        Set MediConnCISPayment = Nothing
'        medixmasterconn.Close
'        Set medixmasterconn = Nothing
'        MsgBox Err.Description
'End Sub



Public Sub cmdShow_Click()
On Error GoTo ErrMsg
    Dim MBM As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String
    Dim sqlstr1, sqlPolStatus   As String
    Dim PolStatus As New ADODB.Recordset
    Dim ValueDate As Date
    Dim tempMBMNo As String
    Dim strsqlAgent As String
    Dim Agent As New ADODB.Recordset
    Dim SpeGracePeriod As Boolean
    Dim SpeGracePeriodDays As String
    Dim strsqlPLN As String
    Dim PLN  As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim cmd1 As New ADODB.Command
    Dim cmd2 As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    Dim StrSqlMsgProm As String
    Dim MsgProm As New ADODB.Recordset
    Dim tmpMsgprom As String
    Dim tmpMBMGLIssuance As String
    tmpSuspendDate = ""
    SSTab1.Tab = 1
    tmpEmails = ""
    tempMBMNo = txtMBMNumber
    Call ClearEntries
    
    txtATTDoc(0).Clear
    txtATTDoc(1).Clear
    txtATTDoc(2).Clear
    txtATTDoc(3).Clear
    txtATTDoc(4).Clear

    tmpMBMTopupInd = False
    txtMBMNumber = tempMBMNo
    cboPatientList.Clear
    cboPatientList.Visible = True
    txtPatientList.Visible = False
    txtPatientIC.Enabled = False
    MBMCPLNCode = ""
    Frame4.Enabled = False
    ScanCode = Trim(txtMBMNumber)
    txtCASRemark.Visible = True
    lblPage.Caption = "Page 1"
    tmpGRPCompany = ""
    tmpStartAge = ""
    tmpEndAge = ""
    txtHOSName.Enabled = False
    If Mid(ScanCode, 1, 2) = "LO" Then
        txtMBMNumber = ""
        Call SearchForm(ScanCode)
    Else
    If Trim(txtMBMNumber) <> "" Then
        
            'If Len(txtMBMNumber) = 8 Or Len(txtMBMNumber) = 9 Then
            If Len(txtMBMNumber) = 8 Or Len(txtMBMNumber) = 9 Or (Len(Trim(txtMBMNumber)) = 11 And Mid(Trim(txtMBMNumber), 9, 1) <> "*") Then
            
                frmSearchMBM05.Source = 4
                frmSearchMBM05.Show
                If frmSearchMBM05.SearchRecCnt < "1" Then
                    Unload frmSearchMBM05
                    MsgBox "No record found!", vbCritical
                ElseIf frmSearchMBM05.SearchRecCnt = "1" Then
                    txtMBMNumber = frmSearchMBM05.lstMBMInquiryList.SelectedItem.Text
                    Unload frmSearchMBM05
                    cmdShow_Click
                End If

            Else
              '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                sqlstr = "Exec SearchMBMCas '" & Trim(txtMBMNumber) & "'"
            
                Set MBM = New ADODB.Recordset
                MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
            If Not MBM.EOF Then
                Screen.MousePointer = vbHourglass
                cmdRegister.Enabled = True
        
                With MBM
                    If Len(!MBMNewPlanCode) Then
                        txtUpgPLNCode = !MBMNewPlanCode
                    End If
                    If Len(!MBMNewPolEffDate) Then
                        txtDOC = CDate(!MBMNewPolEffDate)
                    End If
                    
                    txtMBMName = !MBMName
                    txtMBMPolNo = !MBMPolicyNo
                    txtMBMIcBcPP = !MBMIcBcPp
                    txtDOB.mask = ""
                    txtDOB = Format(!MBMDOB, "dd/mm/yyyy")
                    tmpINSType = !INSCode
                    txtPLNCode = !PLNCode
                    tmpPlnCode = !PLNCode
                    If Len(!MBMAddress1) Then
                        txtAdd1 = !MBMAddress1
                    End If
                    If Len(!MBMAddress2) Then
                        txtAdd2 = !MBMAddress2
                    End If
                    If Len(!MBMAddress3) Then
                        txtAdd3 = !MBMAddress3
                    End If
                    If Len(!MBMPostCode) Then
                        txtPostCode = !MBMPostCode
                    End If
                    If Len(!CTYCode) Then
                        txtCity = !CTYCode
                    End If
                    If Len(!STACode) Then
                        txtState = !STACode
                    End If
                    If Len(Trim(!MBMTakeover)) Then
                        txtTakeOver = !MBMTakeover
                    End If
                    If Len(!MBMRenewFlag) Then
                        txtRenewal = !MBMRenewFlag
                    End If
                    If Len(!MBMGroupCompany) Then
                        tmpGRPCompany = !MBMGroupCompany
                    End If
                    txtHLTCode = !HLTCode

                    '=======================Message Prom==================================
                    txtPAYCode = !PAYCode
                    Dim AnnLmtInd As String

                    txtPolEffDate = IIf(Len(!MBMPolEffDate), Format(!MBMPolEffDate, "dd/mm/yyyy"), "__/__/____")
                    
                    If Trim(txtHLTCode) = "S" Then
                        StrSqlMsgProm = "Select  VIPInd,PLNSpecialNote , AnnLmtIND,MessagePromt,PLNStartAge,PLNEndAge,PLNPolicyWording from PLN where PLNCode = '" & Trim(txtPLNCode) & "'"
                    Else
                        StrSqlMsgProm = "Select   VIPInd,PLNSpecialNote  ,AnnLmtIND,MessagePromt,PLNStartAge,PLNEndAge,PLNPolicyWording from PLN where PLNCode = '" & Trim(txtPLNCode) & "' And PAYCode = '" & Trim(txtPAYCode) & "' "
                    End If
                    MsgProm.Open StrSqlMsgProm, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                        Do While Not MsgProm.EOF
                            If MsgProm!MessagePromt <> "" Then
                                tmpMsgprom = MsgProm!MessagePromt
                                MsgBox tmpMsgprom, vbOKOnly + vbInformation
                            End If
                            AnnLmtInd = MsgProm!AnnLmtInd
                            tmpStartAge = IIf(Len(MsgProm!PLNStartAge), MsgProm!PLNStartAge, "")
                            tmpEndAge = IIf(Len(MsgProm!PLNEndAge), MsgProm!PLNEndAge, "")
                            tmpPolicyWording = IIf(Len(MsgProm!PLNPolicyWording), MsgProm!PLNPolicyWording, "")
                            If Mid(MsgProm!VIPInd, 1, 1) = "Y" Then
                                MsgBox MsgProm!PLNSpecialNote, vbOKOnly + vbInformation
                            End If
                            
                            MsgProm.MoveNext
                        Loop
                    MsgProm.Close
                    Set MsgProm = Nothing
                    
                    txtMBMEffDate.mask = ""
                    txtMBMEffDate = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    'txtMBMEffDate.mask = "##/##/####"
                    txtMBMExpiryDate.mask = ""
                    txtMBMExpiryDate = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    'txtMBMExpiryDate.mask = "##/##/####"
                                        
                    If AnnLmtInd = "Y" Then
                    txtAnnualLimit = Format(GetAnnualLimit(txtINSCode1, txtPLNCode, !HLTCode, CDate(!MBMPayorEffDate)), "#,##0.00")
                    Else
                    txtAnnualLimit = "FULL"
                    End If
                    
                    txtAvailableAnnualLimit = Format(!MBMAvailableLimit, "#,##0.00")
                    
                    txtPolStatus = !MBMStatus
                    
                    If !MBMSuspendStatus = "Y" Then
                        txtPolStatus = "S"
                    End If
                    
                    If !MBMSuspendDate <> "" Then
                        tmpSuspendDate = Format(!MBMSuspendDate, "dd/mm/yyyy")
                    Else
                        tmpSuspendDate = ""
                    End If
                    
                    If Trim(!MBMTopupInd) = "Y" Then
                        tmpMBMTopupInd = True
                    Else
                        tmpMBMTopupInd = False
                    End If
                    
                    If Trim(!MBMTerminateStatus) = "Y" Then
                        tmpMBMTermination = True
                    Else
                        tmpMBMTermination = False
                    End If
                    
                    tmpMBMCostCenter = IIf(Len(!MBMCostCenter), !MBMCostCenter, "")
                    cboPatientList.AddItem "00 -" & Me.txtMBMName
                    
                    If Trim(!PAYCode) = "XL" Then
                        If CDate(GetServerTime) - CDate(txtMBMEffDate) <= 90 Then
                            MsgBox "90 days Non-Cashless!", vbOKOnly + vbCritical, DialogBoxTitle
                        End If
                    End If
        
                    tmpMBMGLIssuance = IIf(Len(!MBMGLIssuance), !MBMGLIssuance, "")
                    tmpAgentStatus = False
                    If !MBMAgentCode <> "" Then
                        strAppendCase = "9"
                        strAppend = " AND AgentCode = '" & Trim(!MBMAgentCode) & "'"
                        Set cmd1.ActiveConnection = medixmasterconn
                        cmd1.CommandText = "Case_Registration"
                        cmd1.CommandType = adCmdStoredProc
                        cmd1.CommandTimeout = 300
                        Set Prm1 = cmd1.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                        cmd1.Parameters.Append Prm1
                        Set Prm2 = cmd1.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                        cmd1.Parameters.Append Prm2
                        Agent.CursorLocation = adUseClient
                        Agent.CursorType = adOpenForwardOnly
                        Agent.CacheSize = 100
                        Set Agent = cmd1.Execute
                        
                        If Agent.EOF = False Then
                            tmpAgentCode = !MBMAgentCode
                            tmpAgentname = Agent!AgentName
                            lblAgent.Visible = True
                            lblAgent.Caption = "Alert... Agent Details : " & Agent!AgentName & "(" & MBM!MBMAgentCode & ")"
                            MsgBox "Alert... Agent Details : " & Agent!AgentName & "(" & MBM!MBMAgentCode & ")", vbOKOnly + vbCritical, DialogBoxTitle
                            tmpAgentStatus = True
                            Call SentEmail(lblAgent.Caption)
                            MsgBox "This is a Suspicious Agent : " & MBM!MBMAgentCode, vbOKOnly + vbCritical, DialogBoxTitle
                        Else
                            lblAgent = ""
                            lblAgent.Visible = False
                        End If
                    Else
                        lblAgent = ""
                        lblAgent.Visible = False
                    End If
                    
                    If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
                        'If MBM.recordCount Then
                            Do While Not MBM.EOF
                                cboPatientList.AddItem MBM!MBMCCoverID & "-" & MBM!MBMCName
                                MBM.MoveNext
                            Loop
                        'End If
                End If
                End With
                '-----------------------------------------------------------------------------------------------
                ' To check prompt message box when member policy have expired
                'strsqlPLN = "Select SpeGPeriodStatus, SpeGPeriodDays from PLN where PLNCode = '" & Trim(txtPLNCode) & "'"
                'PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                strAppendCase = "10"
                strAppend = " AND PLNCode = '" & Trim(txtPLNCode) & "'"
                Set cmd2.ActiveConnection = medixmasterconn
                cmd2.CommandText = "Case_Registration"
                cmd2.CommandType = adCmdStoredProc
                cmd2.CommandTimeout = 300
                Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd2.Parameters.Append Prm1
                Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd2.Parameters.Append Prm2
                PLN.CursorLocation = adUseClient
                PLN.CursorType = adOpenForwardOnly
                PLN.CacheSize = 100
                Set PLN = cmd2.Execute
            
                    If PLN.EOF = False Then
                        If PLN!SpeGPeriodStatus = "Y" Then
                            SpeGracePeriod = True
                            SpeGracePeriodDays = PLN!SpeGPeriodDays
                        Else
                            SpeGracePeriod = False
                        End If
                    Else
                        SpeGracePeriod = False
                    End If
                    tmpPLNPerVisit = IIf(Trim(PLN!PLNPerVisit) = "Y", PLN!PLNPerVisit, "")
                    tmpRBType = IIf(Len(PLN!PLNRBType), PLN!PLNRBType, "")
                PLN.Close
                Set PLN = Nothing
                
                'If SpeGracePeriod = False Then
                '    If DateDiff("ww", txtMBMExpiryDate, DateAdd("d", 14, GetServerTime)) >= 0 Then
                '        MsgBox "Membership has been expired!", vbOKOnly + vbCritical, DialogBoxTitle
                '
                '    End If
                'Else
                '    If DateDiff("ww", txtMBMExpiryDate, DateAdd("d", CDbl(SpeGracePeriodDays), GetServerTime)) >= 0 Then
                '        MsgBox "Membership to be expired (" & SpeGracePeriodDays & " days grace period..!) ", vbOKOnly + vbCritical, DialogBoxTitle
                '    End If
                'End If
                            
        'Call SearchPrincipalIll
        
       ' If Len(PreIllCAT) Then
       '     If Trim(PreIllCAT) = "1" Then
       '         MsgBox "Pre-Existing Illness(Illness!)  ", vbOKOnly + vbCritical, DialogBoxTitle
       '     ElseIf Trim(PreIllCAT) = "2" Then
       '         MsgBox "Pre-Existing Illness(Injury!) ", vbOKOnly + vbCritical, DialogBoxTitle
       '     ElseIf Trim(PreIllCAT) = "3" Then
       '         MsgBox "Pre-Existing Illness(Hospitalisation!) ", vbOKOnly + vbCritical, DialogBoxTitle
       '     ElseIf Trim(PreIllCAT) = "4" Then
       '         MsgBox "Pre-Existing Illness(" & PreIllCATOthers & "!) ", vbOKOnly + vbCritical, DialogBoxTitle
       '     End If
       ' End If
       '
       ' If Len(UndIllCAT) Then
       '     If Trim(UndIllCAT) = "1" Then
       '         MsgBox "Underlying Illness(Diabetes!) ", vbOKOnly + vbCritical, DialogBoxTitle
       '     ElseIf Trim(UndIllCAT) = "2" Then
       '         MsgBox "Underlying Illness(Hypertension!) ", vbOKOnly + vbCritical, DialogBoxTitle
       '     ElseIf Trim(UndIllCAT) = "3" Then
       '         MsgBox "Underlying Illness(" & UndIllCATOthers & "!) ", vbOKOnly + vbCritical, DialogBoxTitle
       '     End If
       ' End If
                                        
                                        
            If CDate(GetServerTime) - CDate(txtMBMEffDate) <= 30 Then
                lblWaitingPeriod.Visible = True
            Else
                lblWaitingPeriod.Visible = False
            End If
            
            If CDate(txtMBMExpiryDate) < CDate(GetServerTime) Then
                MsgBox "Membership has been expired!", vbOKOnly + vbCritical, DialogBoxTitle
                lblExpired.Visible = True
                lblExpired.Caption = "Expired"
                cmdRegister.Enabled = False
            ElseIf CDate(txtMBMExpiryDate) - CDate(GetServerTime) <= 30 Then
                lblExpired.Visible = True
                lblExpired.Caption = "< 30 Days to Expiry"
                cmdRegister.Enabled = True
            Else
                lblExpired.Visible = False
                cmdRegister.Enabled = True
            End If
            
            'If Len(txtAvailableAnnualLimit) Then
            '    If txtAvailableAnnualLimit <= 2000 Then
            '        lblBelow2000.Visible = True
            '    Else
            '        lblBelow2000.Visible = False
            '    End If
            'End If
            
            'txtMorW = Year(GetServerTime) - Year(txtDOB)
            'txtMorW = txtMorW & " Years Old"
            
            Call showMBM
            
            If Trim(txtINSCode1) = "F" Or Trim(txtINSCode1) = "H" Then
                Call showMBMCov
            End If
            Call BNFListing
            
            If tmpMBMTermination = True Then
                cmdLG.Enabled = False
            Else
                cmdLG.Enabled = True
            End If
            
            'If CDbl(txtAvailableAnnualLimit) <= 0 And tmpMBMTopupInd = True Then
            '    MsgBox "Limit Exceeded. Case Registration is not allowed", vbOKOnly + vbCritical, DialogBoxTitle
            '    cmdRegister.Enabled = False
            'Else
            '    cmdRegister.Enabled = True
            'End If
            MBM.Close
            Set MBM = Nothing
            
            'Call ShowExclusion
            
            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
        
            If ChkStr(txtPolStatus) = "S" Then
                MsgBox "Membership has been suspended since " & tmpSuspendDate & ".", vbCritical + vbOKOnly
                cmdRegister.Enabled = False

            End If
        
            If ChkStr(txtPolStatus) = "C" Then
                MsgBox "Membership has been cancelled.", vbCritical + vbOKOnly
                cmdRegister.Enabled = False
            End If
        
            If txtPAYCode = "AX" Then
                If txtPolEffDate <> txtMBMEffDate Then
                    MsgBox "Please check Take Over clause", vbOKOnly + vbCritical, DialogBoxTitle
                End If
            End If
            
            If tmpMBMGLIssuance = "N" Then
                MsgBox "GL is NOT allowed for this member", vbOKOnly + vbInformation, DialogBoxTitle
                cmdLG.Enabled = False
            Else
            
                cmdLG.Enabled = True
            End If
            
        Else
            MsgBox "No record found!", vbCritical
               cmdRegister.Enabled = False
            MBM.Close
            Set MBM = Nothing
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
            End If
        End If

        Else
            MsgBox "Please key in Membership No.!", vbCritical
            txtMBMNumber.SetFocus
            Frame4.Enabled = True
        End If
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrMsg:
   MsgBox Err.Description
   Screen.MousePointer = vbDefault
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


Private Sub lstBenefitDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstBenefitDetails.SortKey = ColumnHeader.Index - 1
    lstBenefitDetails.Sorted = True
    If lstBenefitDetails.SortOrder = lvwAscending Then
        lstBenefitDetails.SortOrder = lvwDescending
    Else
        lstBenefitDetails.SortOrder = lvwAscending
    End If
End Sub

Private Sub NewCaseAddmission()
Dim str As String
Dim RecordSaved, RecordSaved2, Allow30Days, Effdate
Dim CAS As New ADODB.Recordset
Dim MBMChecking As New ADODB.Recordset
Dim MBMCheckEffDate As New ADODB.Recordset
Dim CASChecking As New ADODB.Recordset
Dim CAH As New ADODB.Recordset
Dim MBMCheckStatus As New ADODB.Recordset
Dim strsqlCheckStatus As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim CaseNumber As String
Dim strsqlCheck, strsqlCheck30days, checkEffDate As String
Dim startTrans As Boolean
Dim CASCRef As New ADODB.Recordset
Dim CASRMK As New ADODB.Recordset
Dim strsqlCAS As String
Dim strsqlREM As String
Dim strsqlCASRef As String
Dim HOSCode  As String
Dim DateREC As String
Dim TempResAmt As String
Dim TempLGAmt As String
Dim TempDOA As String
Dim TempDOD As String
Dim TempCASInd As String
Dim TempCLMAdmin As String
Dim TempTokenDate As String
Dim TempUserToken As String
Dim TempDOR As String
Dim Days As Integer
Dim cmd As New ADODB.Command
Dim cmd1 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim cmd3 As New ADODB.Command
Dim cmd4 As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim Reason, RefBy, CASRemark, CASRemarkTwo, CASRemarkThree, CASRemarkFour, CASRemarkFive
Dim Test As String
Dim TmpTakeOverAmt As Double
Dim CASRemarksInt As String
Dim TempDate As Date
Dim strdate As String
Dim RecordSave
Dim tempdays, tempday1, tempday2 As Integer
Dim X As Boolean
Dim Y As Boolean
Dim Z As Boolean


    Check = False
    HOSCode = ""
    tempLGTime = ""
    X = False
    Y = False
    Z = False
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
    
    
    If txtPAYCode = "AX" Then
        If txtPolEffDate <> txtMBMEffDate Then
            RecordSave = MsgBox("Have you checked Take Over clause?", vbYesNo + vbExclamation)
            If RecordSave = vbNo Then
                Exit Sub
            End If
        End If
    End If
        
    If Trim(txtMBMNumber) = "" Then
        MsgBox "Please Enter Membership Information", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf Trim(cboPatientList) = "" Then
        MsgBox "Please Enter Patient Information", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf Trim(cboClaimability) = "" Then
        MsgBox "Please Enter Claimability Status", vbOKOnly + vbCritical, DialogBoxTitle
        cmdRegister.Visible = True
        cmdLG.Enabled = False
        cboClaimability.SetFocus
    ElseIf Trim(txtHOSCode) = "" Then
        MsgBox "Please Enter Hospital Information", vbOKOnly + vbCritical, DialogBoxTitle
        txtHOSCode.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf txtDateAdmitted = "__/__/____" Then
        MsgBox "Please Enter Admission Date", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 1
        txtDateAdmitted.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(Trim(txtDateAdmitted)) = False Then
        MsgBox "Please Enter a Valid Admission Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtDateAdmitted.SetFocus
        SSTab1.Tab = 1
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf cboPatientList = "" Then
        MsgBox "Please Enter Patient's ID", vbOKOnly + vbCritical, DialogBoxTitle
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf Trim(txtCASReserveAmount) = "" Then
        MsgBox "Please Enter Reserve Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASReserveAmount.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf txtCASLGAmt = "" Then
        MsgBox "Please Enter LG Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASLGAmt.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf Mid(CboPayType, 1, 1) = "C" And cboICDCode1 = "" Then
        MsgBox "Please Enter ICD Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboICDCode1.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtPDDOS1) = False And txtPDDOS1 <> "__/__/____" Then
        MsgBox "Please Enter a Valid Diagnosis Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtPDDOS1.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtPDDOS2) = False And txtPDDOS2 <> "__/__/____" Then
        MsgBox "Please Enter a Valid Diagnosis Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtPDDOS2.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtPDDOS3) = False And txtPDDOS3 <> "__/__/____" Then
        MsgBox "Please Enter a Valid Diagnosis Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtPDDOS3.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtUDDOS1) = False And txtUDDOS1 <> "__/__/____" Then
        MsgBox "Please Enter a Valid Diagnosis Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtUDDOS1.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtUDDOS2) = False And txtUDDOS2 <> "__/__/____" Then
        MsgBox "Please Enter a Valid Diagnosis Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtUDDOS2.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtUDDOS3) = False And txtUDDOS3 <> "__/__/____" Then
        MsgBox "Please Enter a Valid Diagnosis Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtUDDOS3.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtFDDOS1) = False And txtFDDOS1 <> "__/__/____" Then
        MsgBox "Please Enter a Valid Diagnosis Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtFDDOS1.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtFDDOS2) = False And txtFDDOS2 <> "__/__/____" Then
        MsgBox "Please Enter a Valid Diagnosis Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtFDDOS2.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtFDDOS3) = False And txtFDDOS3 <> "__/__/____" Then
        MsgBox "Please Enter a Valid Diagnosis Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtFDDOS3.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtFirstCall) = False And txtFirstCall <> "__/__/____" Then
        MsgBox "Please Enter a Valid Fisrt Call Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtFirstCall.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtFirstTreated) = False And txtFirstTreated <> "__/__/____" Then
        MsgBox "Please Enter a Valid Fisrt Treated Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtFirstTreated.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf IsDate(txtDateDischarge) = False And txtDateDischarge <> "__/__/____" Then
        MsgBox "Please Enter a Valid Discharge Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtFirstTreated.SetFocus
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    ElseIf cboTopUpLmt = "" Then
        MsgBox "Please Enter Topup Limit Indicator", vbOKOnly + vbCritical, DialogBoxTitle
        cboTopUpLmt.SetFocus
    ElseIf Mid(cboIntExtTO, 1, 1) = "I" And Trim(txtCASTOPlan) = "" Then
        MsgBox "Please Enter Takeover Plan", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASTOPlan.SetFocus
    ElseIf Mid(CboPayType, 1, 1) = "C" And Trim(txtMBMPolYr) = "" Then
        MsgBox "Please Enter Policy Year", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMPolYr.SetFocus
    ElseIf Trim(txtPatientAge) = "" Then
        MsgBox "Please Enter Patient's Age", vbOKOnly + vbCritical, DialogBoxTitle
        txtPatientAge.SetFocus
    ElseIf Trim(cboYearDay) = "" Then
        MsgBox "Please Select Week/Month/Year", vbOKOnly + vbCritical, DialogBoxTitle
        cboYearDay.SetFocus
    ElseIf (txtHOSCode = "DENTAL CLINIC" Or txtHOSCode = "DIALYSIS CENTRE" Or txtHOSCode = "FOREIGN HOSPITAL" Or txtHOSCode = "GENERAL PRACTITIONER" Or txtHOSCode = "GOVERNMENT HOSPITAL" Or txtHOSCode = "NON-PANEL HOSPITAL" Or txtHOSCode = "SPECIALIST CLINIC") And Trim(txtHOSName) = "" Then
        MsgBox "Please Enter Hospital Name or Clinic Name!", vbOKOnly + vbCritical, DialogBoxTitle
        txtHOSName.Enabled = True
        txtHOSName.SetFocus
    ElseIf (txtHOSCode = "DENTAL CLINIC" Or txtHOSCode = "DIALYSIS CENTRE" Or txtHOSCode = "FOREIGN HOSPITAL" Or txtHOSCode = "GENERAL PRACTITIONER" Or txtHOSCode = "GOVERNMENT HOSPITAL" Or txtHOSCode = "NON-PANEL HOSPITAL" Or txtHOSCode = "SPECIALIST CLINIC") And Trim(txtHOSName) = "" Then
        MsgBox "Please Select Week/Month/Year", vbOKOnly + vbCritical, DialogBoxTitle
        cboYearDay.SetFocus
    ElseIf (txtPAYCode = "XL" Or txtPAYCode = "XB" Or txtPAYCode = "XF") And cboOPChemo = "" And Mid(CboPayType, 1, 1) = "R" Then
        MsgBox "Please Select Outpatient Chemotheraphy Status", vbOKOnly + vbCritical, DialogBoxTitle
        cboOPChemo.SetFocus
    ElseIf Trim(CboStayType) = "" Then
        MsgBox "Please Select Stay Type", vbOKOnly + vbCritical, DialogBoxTitle
        CboStayType.SetFocus
    ElseIf Trim(cboModality) = "" Then
        MsgBox "Please Treatment Mode", vbOKOnly + vbCritical, DialogBoxTitle
        cboModality.SetFocus
    ElseIf Trim(CboPayType) = "" Then
        MsgBox "Please Cash Type", vbOKOnly + vbCritical, DialogBoxTitle
        CboPayType.SetFocus
        
    Else
        If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
            If cboREPCode1 = "" Then
                MsgBox "Please Enter Repudiate Code", vbOKOnly + vbCritical, DialogBoxTitle
                cboREPCode1.SetFocus
                cmdRegister.Visible = True
                cmdLG.Enabled = False
                Exit Sub
            ElseIf Document = "" Or Personnel = "" Or Attention = "" Then
                MsgBox "Please Enter Repudiation Selection Reason", vbOKOnly + vbCritical, DialogBoxTitle
                frmRepudiationSelection.Source = 2
                frmRepudiationSelection.Show vbModal
                cmdRegister.Visible = True
                cmdLG.Enabled = False
                Exit Sub
            End If
        End If
                
    If ChkStr(txtPolStatus) = "C" Then
       MsgBox "Policy has been cancelled!", vbCritical + vbOKOnly
       Exit Sub
    End If
   
     strAppendCase = "1"
     strAppend = " AND MBMNumber = '" & txtMBMNumber & "' AND mbmpayorEffdate >= getdate()"
     Set cmd.ActiveConnection = medixmasterconnMaint
        cmd.CommandText = "Case_Registration"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        MBMCheckEffDate.CursorLocation = adUseClient
        MBMCheckEffDate.CursorType = adOpenForwardOnly
        MBMCheckEffDate.CacheSize = 100
        Set MBMCheckEffDate = cmd.Execute
      
    If ChkDOA.Value = 1 Then
      If MBMCheckEffDate.EOF = False Then
           MsgBox "Policy not yet activated!", vbCritical + vbOKOnly
            MBMCheckEffDate.Close
            Set MBMCheckEffDate = Nothing
            cmdRegister.Visible = True
            cmdLG.Enabled = False
           Exit Sub
      End If
    End If
    
    'Check 30 days
      'strsqlCheck30days = "select MBMNumber from MBMCrossReference where " & _
                    " MBMNumber ='" & txtMBMNumber & "' and    " & _
                    " mbmpayorEffdate >= DateAdd(D, -30, getdate()) " & _
                    " and mbmpayorEffdate <= getdate() "
       'MBMChecking.Open strsqlCheck30days, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    strAppendCase = "1"
    strAppend = " AND MBMNumber ='" & txtMBMNumber & "' AND mbmpayorEffdate >= DateAdd(D, -30, getdate()) AND mbmpayorEffdate <= getdate()"
    Set cmd1.ActiveConnection = medixmasterconnMaint
        cmd1.CommandText = "Case_Registration"
        cmd1.CommandType = adCmdStoredProc
        cmd1.CommandTimeout = 300
        Set Prm1 = cmd1.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd1.Parameters.Append Prm1
        Set Prm2 = cmd1.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd1.Parameters.Append Prm2
        MBMChecking.CursorLocation = adUseClient
        MBMChecking.CursorType = adOpenForwardOnly
        MBMChecking.CacheSize = 100
        Set MBMChecking = cmd1.Execute
        
      If MBMChecking.EOF = False Then
        Allow30Days = MsgBox("Policy has just been activated less than 30 days !" & vbNewLine & "do you want to proceed?", vbOKCancel + vbCritical)
      End If
      
      If Allow30Days = vbCancel Then
        MBMChecking.Close
        Set MBMChecking = Nothing
        cmdRegister.Visible = True
        cmdLG.Enabled = False
        Exit Sub
      End If
      
        ' check discharge
      'strsqlCheck = "select casNumber, CASDateAdmitted  from CASCrossReference " & _
                    " where MBMNumber ='" & Trim(txtMBMNumber) & "' " & _
                    " AND CASDateAdmitted = '" & Format(txtDateAdmitted, "MM/DD/YYYY") & "'"
        'strsqlCheck = strsqlCheck & " and  MBMPatientCovID ='" & Mid(cboPatientList, 1, 2) & "'"
    strAppendCase = "2"
    strAppend = " AND MBMNumber ='" & Trim(txtMBMNumber) & "' AND CASDateAdmitted = '" & Format(txtDateAdmitted, "MM/DD/YYYY") & "' AND  MBMPatientCovID ='" & Mid(cboPatientList, 1, 2) & "'"
    Set cmd2.ActiveConnection = medixmasterconnMaint
        cmd2.CommandText = "Case_Registration"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set Prm1 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd2.Parameters.Append Prm1
        Set Prm2 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd2.Parameters.Append Prm2
        CASChecking.CursorLocation = adUseClient
        CASChecking.CursorType = adOpenForwardOnly
        CASChecking.CacheSize = 100
        Set CASChecking = cmd2.Execute
      
      'CASChecking.Open strsqlCheck, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
      
      If CASChecking.EOF = False Then
     
            RecordSaved = MsgBox("Patient has not been discharged, do you want to continue", vbYesNo + vbExclamation)
            CASChecking.Close
            Set CASChecking = Nothing

      Else
        
            RecordSaved = MsgBox("Are you sure", vbYesNo + vbExclamation)
        
            If RecordSaved = vbNo Then
                CASChecking.Close
                Set CASChecking = Nothing
                cmdRegister.Visible = True
                cmdLG.Enabled = False
            Else
              CASChecking.Close
              Set CASChecking = Nothing
            strAppendCase = "2"
            strAppend = " AND MBMNumber ='" & Trim(txtMBMNumber) & "' AND CASDateAdmitted = '" & Format(txtDateAdmitted, "MM/DD/YYYY") & "' AND  MBMPatientCovID ='" & Mid(cboPatientList, 1, 2) & "'"
            Set cmd3.ActiveConnection = medixmasterconnMaint
                cmd3.CommandText = "Case_Registration"
                cmd3.CommandType = adCmdStoredProc
                cmd3.CommandTimeout = 300
                Set Prm1 = cmd3.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                cmd3.Parameters.Append Prm1
                Set Prm2 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                cmd3.Parameters.Append Prm2
                CASChecking.CursorLocation = adUseClient
                CASChecking.CursorType = adOpenForwardOnly
                CASChecking.CacheSize = 100
                Set CASChecking = cmd3.Execute
                
              If CASChecking.EOF = False Then
                  RecordSaved = MsgBox("Patient has not been discharged, do you want to continue", vbYesNo + vbExclamation)
                  If RecordSaved = vbNo Then
                    Check = True
                    CASChecking.Close
                    Set CASChecking = Nothing
                  End If
              Else
                  CASChecking.Close
                  Set CASChecking = Nothing
                  Check = False
                  cmdRegister.Visible = True
                  cmdLG.Enabled = False
              End If
            End If
        End If
    If RecordSaved = vbYes And Check = False Then
        HOspitalName = Replace(txtHOSCode, "'", "''")
        strAppendCase = "21"
        strAppend = " AND HosName = '" & Trim(HOspitalName) & "'"
        Set cmd4.ActiveConnection = medixmasterconn
            cmd4.CommandText = "Case_Registration"
            cmd4.CommandType = adCmdStoredProc
            cmd4.CommandTimeout = 300
            Set Prm1 = cmd4.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd4.Parameters.Append Prm1
            Set Prm2 = cmd4.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd4.Parameters.Append Prm2
            HOS.CursorLocation = adUseClient
            HOS.CursorType = adOpenForwardOnly
            HOS.CacheSize = 100
            Set HOS = cmd4.Execute
            If HOS.EOF = False Then
                HOSCode = HOS!HOSCode
            End If
        HOS.Close
        Set HOS = Nothing
        
        TempCASInd = IIf(Len(cboCASInd), Mid(cboCASInd, 1, 1), "")
        TempResAmt = CDbl(IIf(IsNumeric(txtCASReserveAmount) = True, txtCASReserveAmount, 0))
        TempLGAmt = CDbl(IIf(IsNumeric(txtCASLGAmt) = True, txtCASLGAmt, 0))
        If txtDateAdmitted <> "__/__/____" Then
            TempDOA = Format(txtDateAdmitted, "yyyy/mm/dd")
        Else
            TempDOA = Empty
        End If
        If txtDateDischarge <> "__/__/____" Then
            TempDOD = Format(txtDateDischarge, "yyyy/mm/dd")
        Else
            TempDOD = Empty
        End If
        If txtDateRECDoc <> "__/__/____" Then
            DateREC = Format(txtDateRECDoc, "mm/dd/yyyy")
        Else
            DateREC = Empty
        End If
        If txtLGTime <> "__:__" Then
            tempLGTime = txtLGTime
        Else
            tempLGTime = Empty
        End If
        
        TempCASInd = IIf(Len(cboCASInd), Mid(cboCASInd, 1, 1), "")
        If Mid(CboPayType, 1, 1) = "R" Then
            TempCLMAdmin = "1"
            TempTokenDate = Format(GetServerTime, "yyyy/mm/dd")
            TempUserToken = Logon
        Else
            TempCLMAdmin = "0"
            TempTokenDate = ""
            TempUserToken = ""
        End If
        
        TempDOR = Format(GetServerTime, "yyyy/mm/dd")
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
        TmpTakeOverAmt = IIf(IsNumeric(txtCASTakeOverAmt) = True, txtCASTakeOverAmt, 0)
        
        
        'Begin Transaction
        medixmasterconn.beginTrans
        startTrans = True
          
        CaseNumber = GenerateCaseNo(txtPAYCode)
        
    Dim sClaimability As String
    sClaimability = ""
    If ASOValid = False Then
        sClaimability = "RR"
    Else
        sClaimability = Mid(cboClaimability, 1, 2)
    End If
    
    
    
    If txtDateDischarge <> "__/__/____" Then
        If Mid(CboPayType, 1, 1) = "R" Then
            strsqlCAS = "EXEC Case_RegInsertCAS2015 '" & ChkStr(CaseNumber) & "', '" & ChkStr(txtMBMNumber) & "', '" & Mid(cboPatientList, 1, 2) & "', " & _
                        "'O', '" & Mid(cboPending, 1, 1) & "', " & TempResAmt & ", " & TempLGAmt & ", " & _
                        "'" & TempDOA & "', '" & TempDOD & "', '" & sClaimability & "', " & _
                        "'" & TempCASInd & "', '" & ChkStr(Mid(cboModality, 1, 2)) & "', '" & Mid(CboStayType, 1, 1) & "' , '" & Mid(CboPayType, 1, 1) & "', '" & ChkStr(TempCLMAdmin) & "', '" & TempTokenDate & "', '" & TempUserToken & "', " & _
                        "'" & ChkStr(Trim(TxtLinkCASNo)) & "' , '" & Mid(cboNature, 1, 3) & "', '" & ChkStr(Logon) & "' ,'" & TempDOR & "', '" & Mid(cboTopUpLmt, 1, 1) & "', '" & ChkStr(HOSCode) & "', " & _
                        "'" & DateREC & "', '" & Mid(cboQuery, 1, 1) & "',  " & TmpTakeOverAmt & ", '" & Mid(cboCASTakeOveInd, 1, 1) & "', '" & txtCASTOPlan & "', '" & tempLGTime & "' , '" & txtMBMPolYr & "', '" & tmpMBMCostCenter & "', '" & Mid(cboOPChemo, 1, 1) & "'"
        Else
            strsqlCAS = "EXEC Case_RegInsertCAS12015 '" & ChkStr(CaseNumber) & "', '" & ChkStr(txtMBMNumber) & "', '" & Mid(cboPatientList, 1, 2) & "', " & _
                "'O', '" & Mid(cboPending, 1, 1) & "', " & TempResAmt & ", " & TempLGAmt & ", " & _
                "'" & TempDOA & "', '" & TempDOD & "', '" & sClaimability & "', " & _
                "'" & TempCASInd & "', '" & ChkStr(Mid(cboModality, 1, 2)) & "', '" & Mid(CboStayType, 1, 1) & "' , '" & Mid(CboPayType, 1, 1) & "', '" & ChkStr(TempCLMAdmin) & "', " & _
                "'" & ChkStr(Trim(TxtLinkCASNo)) & "' , '" & Mid(cboNature, 1, 3) & "', '" & ChkStr(Logon) & "' ,'" & TempDOR & "', '" & Mid(cboTopUpLmt, 1, 1) & "', '" & ChkStr(HOSCode) & "', " & _
                "'" & DateREC & "', '" & Mid(cboQuery, 1, 1) & "',  " & TmpTakeOverAmt & ", '" & Mid(cboCASTakeOveInd, 1, 1) & "', '" & txtCASTOPlan & "', '" & tempLGTime & "', '" & txtMBMPolYr & "', '" & tmpMBMCostCenter & "', '" & Mid(cboOPChemo, 1, 1) & "'"
        End If
    Else
        If Mid(CboPayType, 1, 1) = "R" Then
            strsqlCAS = "EXEC Case_RegInsertCAS22015 '" & ChkStr(CaseNumber) & "', '" & ChkStr(txtMBMNumber) & "', '" & Mid(cboPatientList, 1, 2) & "', " & _
                        "'O', '" & Mid(cboPending, 1, 1) & "', " & TempResAmt & ", " & TempLGAmt & ", " & _
                        "'" & TempDOA & "','" & TempDOD & "','" & sClaimability & "', " & _
                        "'" & TempCASInd & "', '" & ChkStr(Mid(cboModality, 1, 2)) & "', '" & Mid(CboStayType, 1, 1) & "' , '" & Mid(CboPayType, 1, 1) & "', '" & ChkStr(TempCLMAdmin) & "', '" & TempTokenDate & "', '" & TempUserToken & "', " & _
                        "'" & ChkStr(Trim(TxtLinkCASNo)) & "' , '" & Mid(cboNature, 1, 3) & "', '" & ChkStr(Logon) & "' ,'" & TempDOR & "', '" & Mid(cboTopUpLmt, 1, 1) & "', '" & ChkStr(HOSCode) & "', " & _
                        "'" & DateREC & "', '" & Mid(cboQuery, 1, 1) & "',  " & TmpTakeOverAmt & ", '" & Mid(cboCASTakeOveInd, 1, 1) & "', '" & txtCASTOPlan & "', '" & tempLGTime & "' , '" & txtMBMPolYr & "', '" & tmpMBMCostCenter & "', '" & Mid(cboOPChemo, 1, 1) & "'"
        Else
            strsqlCAS = "EXEC Case_RegInsertCAS32015 '" & ChkStr(CaseNumber) & "', '" & ChkStr(txtMBMNumber) & "', '" & Mid(cboPatientList, 1, 2) & "', " & _
                "'O', '" & Mid(cboPending, 1, 1) & "', " & TempResAmt & ", " & TempLGAmt & ", " & _
                "'" & TempDOA & "','" & TempDOD & "', '" & sClaimability & "', " & _
                "'" & TempCASInd & "', '" & ChkStr(Mid(cboModality, 1, 2)) & "', '" & Mid(CboStayType, 1, 1) & "' , '" & Mid(CboPayType, 1, 1) & "', '" & ChkStr(TempCLMAdmin) & "'," & _
                "'" & ChkStr(Trim(TxtLinkCASNo)) & "' , '" & Mid(cboNature, 1, 3) & "','" & ChkStr(Logon) & "' ,'" & TempDOR & "', '" & Mid(cboTopUpLmt, 1, 1) & "', '" & ChkStr(HOSCode) & "', " & _
                "'" & DateREC & "', '" & Mid(cboQuery, 1, 1) & "',  " & TmpTakeOverAmt & ", '" & Mid(cboCASTakeOveInd, 1, 1) & "', '" & txtCASTOPlan & "', '" & tempLGTime & "' , '" & txtMBMPolYr & "', '" & tmpMBMCostCenter & "', '" & Mid(cboOPChemo, 1, 1) & "'"
        End If
    End If
    medixmasterconn.Execute strsqlCAS
  
    If Len(CASRemark) <= 1900 Then
        strsqlREM = "EXEC Case_RegInsertCASRemark '" & ChkStr(CaseNumber) & "', '" & CASRemark & "', ''"
    ElseIf Len(CASRemark) > 1900 Then
        strsqlREM = "EXEC Case_RegInsertCASRemark '" & ChkStr(CaseNumber) & "', '" & Mid(CASRemark, 1, 1900) & "', '" & Mid(CASRemark, 1901, Len(CASRemark) - 1900) & "'"
    End If
    medixmasterconn.Execute strsqlREM
            
    strsqlREM = ""
    If Len(CASRemarkTwo) <= 1900 Then
        strsqlREM = "EXEC Case_RegInsertCASRemark2 '" & ChkStr(CaseNumber) & "', '" & CASRemarkTwo & "', ''"
    ElseIf Len(CASRemarkTwo) > 1900 Then
        strsqlREM = "EXEC Case_RegInsertCASRemark2 '" & ChkStr(CaseNumber) & "', '" & Mid(CASRemarkTwo, 1, 1900) & "', '" & Mid(CASRemarkTwo, 1901, Len(CASRemarkTwo) - 1900) & "'"
    End If
    medixmasterconn.Execute strsqlREM
    
    strsqlREM = ""
    strsqlREM = "EXEC Case_RegInsertCASRemark3 '" & ChkStr(CaseNumber) & "', '" & CASRemarkThree & "', '" & CASRemarkFour & "', '" & CASRemarkFive & "'"
    medixmasterconn.Execute strsqlREM

    
    If txtDateDischarge <> "__/__/____" Then
        strsqlCASRef = " EXEC Case_RegInsertCASRef '" & ChkStr(CaseNumber) & "', '" & ChkStr(txtMBMNumber) & "', '" & Mid(cboPatientList, 1, 2) & "', " & _
                    "'O', '" & Mid(cboPending, 1, 1) & "', '" & ChkStr(HOSCode) & "', " & _
                    "'" & TempDOA & "', '" & TempDOD & "', '" & Mid(cboClaimability, 1, 2) & "', " & _
                    "'" & ChkStr(TxtLinkCASNo) & "', '" & Mid(cboTopUpLmt, 1, 1) & "', '" & Mid(cboPayPending, 1, 1) & "'"
    Else
        strsqlCASRef = " EXEC Case_RegInsertCASRef1 '" & ChkStr(CaseNumber) & "', '" & ChkStr(txtMBMNumber) & "', '" & Mid(cboPatientList, 1, 2) & "', " & _
                    "'O', '" & Mid(cboPending, 1, 1) & "', '" & ChkStr(HOSCode) & "', " & _
                    "'" & TempDOA & "', '" & Mid(cboClaimability, 1, 2) & "', " & _
                    "'" & ChkStr(TxtLinkCASNo) & "', '" & Mid(cboTopUpLmt, 1, 1) & "', '" & Mid(cboPayPending, 1, 1) & "'"
    End If
        
        medixmasterconnMaint.Execute strsqlCASRef
            strsqlCASRef = ""
            strsqlCASRef = "EXEC Case_RegInsertCASRecording '" & ChkStr(CaseNumber) & "','" & TempDOR & "' , '" & tempLGTime & "'," & _
                    "'" & Mid(CboStayType, 1, 1) & "' , '" & ChkStr(Logon) & "' , '" & cboPatientList & "', '" & ChkStr(txtHOSCode) & "',  " & _
                " '" & ChkStr(txtMBMNumber) & "', '" & TempDOA & "'"
            medixmasterconnMaint.Execute strsqlCASRef
            
            
          txtCaseNo = ChkStr(CaseNumber)
          
          
          
                Call SaveFirstDiag(CaseNumber)
                Call SaveFinalDiag(CaseNumber)
                Call SaveUD(CaseNumber)
                Call SaveCASOthers(CaseNumber)
                Call SaveCASOthers2(CaseNumber)
                Call UpdateBalLmt
                
                If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
                    Call SaveRepudiation
                End If
      End If
    End If
    
    If startTrans = True Then
        medixmasterconn.commitTrans
        X = CreatePath(DocPath & "ScannedDoc\" & txtCaseNo)
        Y = CreatePath(DocPath & "Emails\" & txtCaseNo)
        Z = CreatePath(DocPath & "InternalDocuments\" & txtCaseNo)
        MsgBox "Record Saved, Your Case No is : " & CaseNumber, vbOKOnly + vbInformation, DialogBoxTitle
        cmdRegister.Enabled = False
        cmdLG.Enabled = True
        cboStatus1.ListIndex = 0
        Days = CDate(txtMBMExpiryDate) - CDate(GetServerTime)
        If Days <= 45 And Days >= 0 Then
            Letter = "P"
            Call funcword
            cmdPrior.Enabled = True
        ElseIf CDate(GetServerTime) > CDate(txtMBMExpiryDate) Then
            Letter = "E"
            Call funcword
            cmdPrior.Enabled = True
        Else
            cmdPrior.Enabled = False
        End If
        
        If Mid(cboClaimability, 1, 2) = "RR" Or Mid(cboClaimability, 1, 2) = "NR" Or Mid(cboClaimability, 1, 2) = "PR" Or Mid(cboClaimability, 1, 2) = "AR" Or Mid(cboClaimability, 1, 2) = "BR" Or Mid(cboClaimability, 1, 2) = "ER" Then
            Letter = "R"
            Call funcword
        End If
        
    End If
        If tmpAgentStatus = True Then
            Call MailCDOMsg("Alert... Agent Details : " & tmpAgentname & "(" & tmpAgentCode & ")")
        End If
    Exit Sub
    
    

ErrorSave:
    MsgBox Err.Description
    If startTrans = True Then
        medixmasterconn.RollbackTrans
        cmdRegister.Visible = True
        cmdLG.Enabled = False
    End If
'    Exit Sub
    
'ErrorCheck:
'    If startTrans = True Then
'        medixmasterconn.RollbackTrans
'        cmdRegister.Visible = True
'        cmdLG.Enabled = False
'    End If
  
End Sub

Private Sub SaveFirstDiag(CaseNumber As String)
On Error Resume Next
    Dim listrecord As Integer
    Dim listloop As Integer
    Dim strsqlDIAFirst As String
    Dim DIADOS1, DIADOS2, DIADOS3 As String
    Dim PrelimDiag1, PrelimDiag2, PrelimDiag3
    
    ' INSERT INTO DIAFirst TABLE
    'diafirst.Open "select casnumber , * from DIAFirst", medixmasterconn, adOpenKeyset, adLockOptimistic
        'diafirst.AddNew
        'With diafirst
            '!CasNumber = ChkStr(CaseNumber)
                'If cboICDCode1 <> "" Then
                '    !ICDCode1 = ChkStr(cboICDCode1)
                'End If
                If txtPDDOS1 <> "__/__/____" Then
                    DIADOS1 = Format(txtPDDOS1, "mm/dd/yyyy")
                Else
                    DIADOS1 = Empty
                End If
                '!DIAPrelimDiag1 = ChkStr(txtPrelimDiag1)
                'If cboICDCode2 <> "" Then
                '    !ICDCode2 = ChkStr(cboICDCode2)
                'End If
                If txtPDDOS2 <> "__/__/____" Then
                    DIADOS2 = Format(txtPDDOS2, "mm/dd/yyyy")
                Else
                    DIADOS2 = Empty
                End If
                '!DIAPrelimDiag2 = ChkStr(txtPrelimDiag2)
                'If cboICDCode3 <> "" Then
                '    !ICDCode3 = ChkStr(cboICDCode3)
                'End If
                If txtPDDOS3 <> "__/__/____" Then
                    DIADOS3 = Format(txtPDDOS3, "mm/dd/yyyy")
                Else
                    DIADOS3 = Empty
                End If
                '!DIAPrelimDiag3 = ChkStr(txtPrelimDiag3)
        'End With
        'diafirst.Update
    'diafirst.Close
    'Set diafirst = Nothing
    PrelimDiag1 = ChkStr(Replace(txtPrelimDiag1, "'", "''"))
    PrelimDiag2 = ChkStr(Replace(txtPrelimDiag2, "'", "''"))
    PrelimDiag3 = ChkStr(Replace(txtPrelimDiag3, "'", "''"))
    
    strsqlDIAFirst = " EXEC Case_RegInsertDIAFirst '" & ChkStr(CaseNumber) & "', " & _
                    "'" & ChkStr(cboICDCode1) & "', '" & ChkStr(cboICDCode2) & "', '" & ChkStr(cboICDCode3) & "', " & _
                    "'" & PrelimDiag1 & "', '" & PrelimDiag2 & "', '" & PrelimDiag3 & "', " & _
                    "'" & DIADOS1 & "', '" & DIADOS2 & "', '" & DIADOS3 & "'"
    medixmasterconn.Execute strsqlDIAFirst
End Sub

Private Sub SaveFinalDiag(CaseNumber As String)
On Error Resume Next
Dim diafinal As New ADODB.Recordset
Dim strsqlDIAFinal As String
Dim DIADOS1, DIADOS2, DIADOS3
Dim FinalDiag1, FinalDiag2, FinalDiag3
    ' INSERT INTO DIAFinal TABLE
    
    TotalFinalDiag = 0
    'diafinal.Open "select  * from DIAFinal", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'diafinal.AddNew
    'With diafinal
        '!CasNumber = ChkStr(CaseNumber)
        If cboFDICDCode1 <> "" Then
        '    !ICDCode1 = ChkStr(cboFDICDCode1)
            TotalFinalDiag = CDbl(TotalFinalDiag) + 1
        End If
        If txtFDDOS1 <> "__/__/____" Then
            DIADOS1 = Format(txtFDDOS1, "mm/dd/yyyy")
        Else
            DIADOS1 = Empty
        End If
        '!DIAFinalDiag1 = ChkStr(txtFinalDiag1)
        If cboFDICDCode2 <> "" Then
        '    !ICDCode2 = ChkStr(cboFDICDCode2)
            TotalFinalDiag = CDbl(TotalFinalDiag) + 1
        End If
        If txtFDDOS2 <> "__/__/____" Then
            DIADOS2 = Format(txtFDDOS2, "mm/dd/yyyy")
        Else
            DIADOS2 = Empty
        End If
        '!DIAFinalDiag2 = ChkStr(txtFinalDiag2)
        If cboFDICDCode3 <> "" Then
        '    !ICDCode3 = ChkStr(cboFDICDCode3)
            TotalFinalDiag = CDbl(TotalFinalDiag) + 1
        End If
        If txtFDDOS3 <> "__/__/____" Then
            DIADOS3 = Format(txtFDDOS3, "mm/dd/yyyy")
        Else
            DIADOS3 = Empty
        End If
        '!DIAFinalDiag3 = ChkStr(txtFinalDiag3)
    'End With
    'diafinal.Update
    'diafinal.Close
    'Set diafinal = Nothing
    FinalDiag1 = ChkStr(Replace(txtFinalDiag1, "'", "''"))
    FinalDiag2 = ChkStr(Replace(txtFinalDiag2, "'", "''"))
    FinalDiag3 = ChkStr(Replace(txtFinalDiag3, "'", "''"))
    
    strsqlDIAFinal = "EXEC Case_RegInsertDIAFinal '" & ChkStr(CaseNumber) & "', " & _
                    "'" & ChkStr(cboFDICDCode1) & "', '" & ChkStr(cboFDICDCode2) & "', '" & ChkStr(cboFDICDCode3) & "', " & _
                    "'" & FinalDiag1 & "', '" & FinalDiag2 & "', '" & FinalDiag3 & "', " & _
                    "'" & DIADOS1 & "', '" & DIADOS2 & "', '" & DIADOS3 & "'"
    medixmasterconn.Execute strsqlDIAFinal
    
End Sub

Private Sub UpdateBalLmt()
Dim strsqlMBM As String
Dim MBMLimit
Dim MBM As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    If IsNumeric(txtREDAmt) = True Then
        'strsqlMBM = "select MBMAvailableLimit from MBM where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "11"
        strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Registration"
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
        
            If MBM.EOF = False Then
                If CDbl(MBM!MBMAvailableLimit) > CDbl(txtREDAmt) Then
                    MBMLimit = CDbl(MBM!MBMAvailableLimit) - CDbl(txtREDAmt)
                    strsqlMBM = "EXEC Case_RegUpdateBalLmt " & MBMLimit & ", '" & Trim(txtMBMNumber) & "'"
                    
                    'MBM!MBMAvailableLimit = CDbl(MBM!MBMAvailableLimit) - CDbl(txtREDAmt)
                Else
                    strsqlMBM = "EXEC Case_RegUpdateBalLmt " & 0 & ", '" & Trim(txtMBMNumber) & "'"
                    'MBM!MBMAvailableLimit = 0
                End If
                medixmasterconn.Execute strsqlMBM
            End If
        MBM.Close
        Set MBM = Nothing
    End If
End Sub

Private Sub SaveUD(CaseNumber As String)
On Error Resume Next
    Dim UD As New ADODB.Recordset
    Dim strsqlSaveUD As String
    Dim CASUDDOS1, CASUDDOS2, CASUDDOS3 As String
    Dim UD1, UD2, UD3
    
    'UD.Open "CAS_UD", medixmasterconn, adOpenKeyset, adLockOptimistic
    ' INSERT INTO UD TABLE
            'UD.AddNew
            'With UD
                '!CasNumber = ChkStr(CaseNumber)
                    '!CASUDCode1 = Mid(ChkStr(cboUDICDCode1), 1, 5)
                    If txtUDDOS1 <> "__/__/____" Then
                        CASUDDOS1 = Format(txtUDDOS1, "mm/dd/yyyy")
                    Else
                        CASUDDOS1 = Empty
                    End If
                    '!CASUDReasons1 = ChkStr(txtUD1)
                    '!CASUDCode2 = Mid(ChkStr(cboUDICDCode2), 1, 5)
                    If txtUDDOS2 <> "__/__/____" Then
                        CASUDDOS2 = Format(txtUDDOS2, "mm/dd/yyyy")
                    Else
                        CASUDDOS2 = Empty
                    End If
                    '!CASUDReasons2 = ChkStr(txtUD2)
                    '!CASUDCode3 = Mid(ChkStr(cboUDICDCode3), 1, 5)
                    If txtUDDOS3 <> "__/__/____" Then
                        CASUDDOS3 = Format(txtUDDOS3, "mm/dd/yyyy")
                    Else
                        CASUDDOS3 = Empty
                    End If
                    '!CASUDReasons3 = ChkStr(txtUD3)
                'End If
            'End With
            'UD.Update
    'UD.Close
    'Set UD = Nothing
    UD1 = ChkStr(Replace(txtUD1, "'", "''"))
    UD2 = ChkStr(Replace(txtUD2, "'", "''"))
    UD3 = ChkStr(Replace(txtUD3, "'", "''"))
    
    strsqlSaveUD = " EXEC Case_RegInsertUD '" & ChkStr(CaseNumber) & "', " & _
                    "'" & Mid(ChkStr(cboUDICDCode1), 1, 5) & "','" & Mid(ChkStr(cboUDICDCode2), 1, 5) & "','" & Mid(ChkStr(cboUDICDCode3), 1, 5) & "', " & _
                    "'" & UD1 & "', '" & UD2 & "', '" & UD3 & "'," & _
                    "'" & CASUDDOS1 & "', '" & CASUDDOS2 & "', '" & CASUDDOS3 & "'"
    medixmasterconn.Execute strsqlSaveUD
    
End Sub

Private Sub showMBM()
    Dim MBM As New ADODB.Recordset
    Dim sqlstr As String
    Dim strsqlMBMExc As String
    Dim MBMExc As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim cnn As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
    strAppendCase = "12"
    strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Registration"
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
    
    strAppendCase = "13"
    strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
    Set cnn.ActiveConnection = medixmasterconn
    cnn.CommandText = "Case_Registration"
    cnn.CommandType = adCmdStoredProc
    cnn.CommandTimeout = 300
    Set Prm1 = cnn.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cnn.Parameters.Append Prm1
    Set Prm2 = cnn.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cnn.Parameters.Append Prm2
    MBMExc.CursorLocation = adUseClient
    MBMExc.CursorType = adOpenForwardOnly
    MBMExc.CacheSize = 100
    Set MBMExc = cnn.Execute
    
    lstMember.ListItems.Clear
    If MBM.EOF = False Then
        Dim MBMListing As ListItem
            Set MBMListing = lstMember.ListItems.Add(, , "00")
            With MBM
                If Len(Trim(!MBMName)) <> 0 Then
                    MBMListing.SubItems(1) = ChkStr(!MBMName)
                End If
                If Len(Trim(!MBMIcBcPp)) <> 0 Then
                    MBMListing.SubItems(2) = !MBMIcBcPp
                End If
                If Len(Trim(!MBMMemberType)) Then
                    MBMListing.SubItems(3) = !MBMMemberType
                End If
                If Len(Trim(!MBMDOB)) Then
                    MBMListing.SubItems(4) = !MBMDOB
                End If
                If MBMExc.recordCount Then
                    If Len(MBMExc!MBMExclusion) Then
                        txtPrincipalExc = MBMExc!MBMExclusion
                    End If
                End If
                'If Len(MBMExc!MBMTopupJoinDate) Then
                 '   txtMBMTopUpJoinDate = Format((MBMExc!MBMTopupJoinDate), "dd/mm/yyyy")
                'End If
                
                If Trim(!MBMTopupInd) = "Y" Then
                    'lblTopupStatus.Visible = True
                    tmpMBMTopupInd = True
                Else
                    'lblTopupStatus.Visible = False
                    tmpMBMTopupInd = False
                End If
                
                If Trim(!PAYChkHosp) = "Y" Then
                    tmpHospChk = "Y"
                Else
                    tmpHospChk = ""
                End If
                
             
                
            End With
    End If
    MBM.Close
    Set MBM = Nothing
    MBMExc.Close
    Set MBMExc = Nothing
    
    
End Sub
' * End

' * Start

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
' * End

Private Sub SaveRepudiation()
    Dim REP As New ADODB.Recordset
    Dim CAS As New ADODB.Recordset
    Dim strsqlCAS As String
    Dim strsqlCASRep As String
    Dim TempSysDate As String
    Dim cmd As New ADODB.Command
    Dim cnn As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    ' INSERT INTO CASRepudiation TABLE
    
    TempSysDate = Format(GetServerTime, "yyyy/mm/dd")
    'StrSql = "Select CASNumber,CASClaimability  From CAS where CASNumber = '" & Trim(txtCaseNo) & "'"
    'CAS.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'strsqlCASRep = "insert into CASRepudiation (CasNumber, CASREPCode1, CASREPCode2,CASREPCode3,CASREPIssuedBy,CASREPDate) " & _
                "Values('" & ChkStr(txtCaseNo) & "', '" & Mid(cboREPCode1, 1, 7) & "', '" & Mid(cboREPCode2, 1, 7) & "', '" & Mid(cboREPCode3, 1, 7) & "', '" & Logon & "', '" & TempSysDate & "')"
    strsqlCASRep = "EXEC Case_RegInsertCASRep '" & ChkStr(txtCaseNo) & "', '" & IIf(Len(cboREPCode1), Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - IIf(InStr(1, cboREPCode1, "-") = 0, 0, 1))), "") & "', " & _
                        "'" & IIf(Len(cboREPCode2), Trim(Mid(cboREPCode2, 1, InStr(1, cboREPCode2, "-") - IIf(InStr(1, cboREPCode2, "-") = 0, 0, 1))), "") & "', '" & IIf(Len(cboREPCode3), Trim(Mid(cboREPCode3, 1, InStr(1, cboREPCode3, "-") - IIf(InStr(1, cboREPCode3, "-") = 0, 0, 1))), "") & "', " & _
                        "'" & Logon & "', '" & TempSysDate & "', '" & IIf(Len(Document), IIf(InStr(1, Document, "-") = 0, Trim(Document), Trim(Mid(Document, 1, InStr(1, Document, "-") - IIf(InStr(1, Document, "-") = 0, 0, 1)))), "") & "', " & _
                        "'" & IIf(Len(Personnel), IIf(InStr(1, Personnel, "-") = 0, Trim(Personnel), Trim(Mid(Personnel, 1, InStr(1, Personnel, "-") - IIf(InStr(1, Personnel, "-") = 0, 0, 1)))), "") & "', " & _
                        "'" & IIf(Len(Attention), IIf(InStr(1, Attention, "-") = 0, Trim(Attention), Trim(Mid(Attention, 1, InStr(1, Attention, "-") - IIf(InStr(1, Attention, "-") = 0, 0, 1)))), "") & "', " & _
                        "'" & CASUnderlying & "'"
                            
    medixmasterconn.Execute strsqlCASRep
    
    'strsqlCAS = "EXEC Case_RegUpdateCAS '" & Trim(txtCaseNo) & "', 'R'"
    'medixmasterconn.Execute strsqlCAS
    'With CAS
    '    !CASClaimability = "R"
    'End With
    'CAS.Update
    'CAS.Close
    'Set CAS = Nothing
End Sub

Private Sub SaveCASOthers(CaseNumber As String)
Dim CASOthers As New ADODB.Recordset
Dim RedAmt As String
Dim strsqlCASOthers As String
Dim FirstCall, FirstTreated As String
Dim PROCINVDone1, PROCINVDone2, PROCINVDone3
Dim CASCptCode1, CASCptCode2, CASCptCode3, CASCptCode4, CASCptCode5, CASCptCode6 As String
Dim PROCRequired1, PROCRequired2, PROCRequired3, PROCRequired4, PROCRequired5, PROCRequired6
Dim MEDSURDone1, MEDSURDone2, MEDSURDone3
Dim CASONonSurgPROCRequired1, CASONonSurgPROCRequired2, CASONonSurgPROCRequired3, CASONonSurgPROCRequired4, CASONonSurgPROCRequired5, CASONonSurgPROCRequired6 As String
Dim CASNonSurgCptCode1, CASNonSurgCptCode2, CASNonSurgCptCode3, CASNonSurgCptCode4, CASNonSurgCptCode5, CASNonSurgCptCode6 As String


    TotalProc = 0
    
    If txtFirstCall <> "__/__/____" Then
        FirstCall = ChkStr(txtFirstCall)
    Else
        FirstCall = Empty
    End If
    If txtFirstTreated <> "__/__/____" Then
        FirstTreated = ChkStr(txtFirstTreated)
    Else
        FirstTreated = Empty
    End If
    
    'If Mid(cboCondwillOcc, 1, 1) = "Y" Then
    '    !CASOCondWillRecur = "Y"
    'Else
    '    !CASOCondWillRecur = "N"
    'End If
    
    'If Mid(cboFLWTreatment, 1, 1) = "Y" Then
    '    !CASOFollowUpRequired = "Y"
    'Else
    '    !CASOFollowUpRequired = "N"
    'End If

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
    If txtProcedure6 <> "" Then
        TotalProc = CDbl(TotalProc) + 1
    End If
    
    
    RedAmt = CDbl(IIf(IsNumeric(txtREDAmt) = True, txtREDAmt, 0))
        
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
    
    
    CASONonSurgPROCRequired1 = ChkStr(Replace(txtNSProcedure1, "'", "''"))
    CASONonSurgPROCRequired2 = ChkStr(Replace(txtNSProcedure2, "'", "''"))
    CASONonSurgPROCRequired3 = ChkStr(Replace(txtNSProcedure3, "'", "''"))
    CASONonSurgPROCRequired4 = ChkStr(Replace(txtNSProcedure4, "'", "''"))
    CASONonSurgPROCRequired5 = ChkStr(Replace(txtNSProcedure5, "'", "''"))
    CASONonSurgPROCRequired6 = ChkStr(Replace(txtNSProcedure6, "'", "''"))
    
    
    CASNonSurgCptCode1 = IIf(Len(txtNSProcCode1), txtNSProcCode1, "")
    CASNonSurgCptCode2 = IIf(Len(txtNSProcCode2), txtNSProcCode2, "")
    CASNonSurgCptCode3 = IIf(Len(txtNSProcCode3), txtNSProcCode3, "")
    CASNonSurgCptCode4 = IIf(Len(txtNSProcCode4), txtNSProcCode4, "")
    CASNonSurgCptCode5 = IIf(Len(txtNSProcCode5), txtNSProcCode5, "")
    CASNonSurgCptCode6 = IIf(Len(txtNSProcCode6), txtNSProcCode6, "")
    
    MEDSURDone1 = ChkStr(Replace(txtMEDSURDone1, "'", "''"))
    MEDSURDone2 = ChkStr(Replace(txtMEDSURDone2, "'", "''"))
    MEDSURDone3 = ChkStr(Replace(txtMEDSURDone3, "'", "''"))
    
    strsqlCASOthers = " EXEC Case_RegInsertCASOthers '" & ChkStr(CaseNumber) & "', " & _
                    "'" & PROCINVDone1 & "', '" & PROCINVDone2 & "' , '" & PROCINVDone3 & "', " & _
                    "'" & MEDSURDone1 & "', '" & MEDSURDone2 & "', '" & MEDSURDone3 & "', " & _
                    "'" & Mid(cboCondwillOcc, 1, 1) & "', '" & Mid(cboFLWTreatment, 1, 1) & "', '" & FirstCall & "', '" & FirstTreated & "', " & _
                    "" & RedAmt & ", '" & ChkStr(txtINVCode1) & "', '" & ChkStr(txtINVCode2) & "', '" & ChkStr(txtINVCode3) & "', '" & Mid(cboFolUp, 1, 1) & "', '" & Mid(cboPayPending, 1, 1) & "', " & _
                    "'" & PROCRequired1 & "', '" & PROCRequired2 & "', '" & PROCRequired3 & "', '" & PROCRequired4 & "', '" & PROCRequired5 & "','" & PROCRequired6 & "',  " & _
                    "'" & CASCptCode1 & "', '" & CASCptCode2 & "', '" & CASCptCode3 & "', '" & CASCptCode4 & "', '" & CASCptCode5 & "',  '" & CASCptCode6 & "'," & _
                    "'" & CASONonSurgPROCRequired1 & "', '" & CASONonSurgPROCRequired2 & "', '" & CASONonSurgPROCRequired3 & "', '" & CASONonSurgPROCRequired4 & "', '" & CASONonSurgPROCRequired5 & "', '" & CASONonSurgPROCRequired6 & "'," & _
                    "'" & CASNonSurgCptCode1 & "', '" & CASNonSurgCptCode2 & "', '" & CASNonSurgCptCode3 & "', '" & CASNonSurgCptCode4 & "', '" & CASNonSurgCptCode5 & "', '" & CASNonSurgCptCode6 & "', '" & txtPatientAge & "', '" & Mid(cboYearDay, 1, 1) & "'"
                    
    medixmasterconn.Execute strsqlCASOthers
    
End Sub

Private Function SaveCASOthers2(CaseNumber As String) As Boolean
Dim CASOthersTwo As New ADODB.Recordset
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
Dim Reason As String
Dim RefBy As String
Dim strsqlCASOthers2 As String
Dim TmpFirstDocDate As String
Dim TmpSubSeqDocDate As String
Dim TmpIntExtTO As String
Dim TmpTHMed1 As String
Dim TmpTHMed2 As String
Dim TmpTHMed3 As String
Dim TmpTHMed4 As String
Dim TmpTHMed5 As String
Dim TmpTHMed6 As String
Dim TmpTHMed7 As String
Dim TmpTHInd1 As String
Dim TmpTHInd2 As String
Dim TmpTHInd3 As String
Dim TmpTHInd4 As String
Dim TmpTHInd5 As String
Dim TmpTHInd6 As String
Dim TmpTHInd7 As String
Dim TmpTHDOS1 As String
Dim TmpTHDOS2 As String
Dim TmpTHDOS3 As String
Dim TmpTHDOS4 As String
Dim TmpTHDOS5 As String
Dim TmpTHDOS6 As String
Dim TmpTHDOS7 As String
Dim TmpTHDur1 As String
Dim TmpTHDur2 As String
Dim TmpTHDur3 As String
Dim TmpTHDur4 As String
Dim TmpTHDur5 As String
Dim TmpTHDur6 As String
Dim TmpTHDur7 As String
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
Dim TmpTHCharges1 As Double
Dim TmpTHCharges2 As Double
Dim TmpTHCharges3 As Double
Dim TmpTHCharges4 As Double
Dim TmpTHCharges5 As Double
Dim TmpTHCharges6 As Double
Dim TmpTHCharges7 As Double
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
TmpCASComplaintIssue = ""
TmpCASComplainant = ""
TmpCASCompAction = txtCompAction
    If Trim(txtHOSName) <> "" Then
        tmpHospName = Replace(txtHOSName, "'", "''")
    Else
        tmpHospName = ""
    End If
    

    ' INSERT INTO CASOthersTwo TABLE
    'CASOthersTwo.Open "Select CASNumber, * from CASOthersTwo", medixmasterconn, adOpenKeyset, adLockOptimistic
            'CASOthersTwo.AddNew
            'With CASOthersTwo
                '!CasNumber = ChkStr(CaseNumber)
                
                If TempDRName <> "" Then
                '    !CASAttendDoctor1 = ChkStr(TempDRName)
                    TotalDoctor = CDbl(TotalDoctor) + 1
                End If
                If TempDRName2 <> "" Then
                '    !CASAttendDoctor2 = ChkStr(TempDRName2)
                    TotalDoctor = CDbl(TotalDoctor) + 1
                End If
                If TempDRName3 <> "" Then
                '    !CASAttendDoctor3 = ChkStr(TempDRName3)
                    TotalDoctor = CDbl(TotalDoctor) + 1
                End If
                If TempDRName4 <> "" Then
                '    !CASAttendDoctor4 = ChkStr(TempDRName4)
                    TotalDoctor = CDbl(TotalDoctor) + 1
                End If
                If TempDRName5 <> "" Then
                '    !CASAttendDoctor5 = ChkStr(TempDRName5)
                    TotalDoctor = CDbl(TotalDoctor) + 1
                End If
                If TempAnaesName <> "" Then
                '    !CASAnaesName1 = ChkStr(TempAnaesName)
                    TotalAnaes = CDbl(TotalAnaes) + 1
                End If
                If TempAnaesName2 <> "" Then
                '    !CASAnaesName2 = ChkStr(TempAnaesName2)
                    TotalAnaes = CDbl(TotalAnaes) + 1
                End If
                If TempAnaesName3 <> "" Then
                '    !CASAnaesName3 = ChkStr(TempAnaesName3)
                    TotalAnaes = CDbl(TotalAnaes) + 1
                End If
                
                If txtSubDocDate <> "__/__/____" Then
                    TmpSubSeqDocDate = Format(txtSubDocDate, "mm/dd/yyyy")
                Else
                    TmpSubSeqDocDate = Empty
                End If
                TmpIntExtTO = IIf(Len(cboIntExtTO), Mid(cboIntExtTO, 1, 1), "")
                                
                If txtFirstDocDate <> "__/__/____" Then
                    TmpFirstDocDate = Format(txtFirstDocDate, "mm/dd/yyyy")
                Else
                    TmpFirstDocDate = Empty
                End If
                
                If txtDOATime <> "__:__" Then
                    TmpCASDOATime = Format(txtDOATime, "HH:MM")
                Else
                    TmpCASDOATime = Empty
                End If
                
                'If txtDODTime <> "__:__" Then
                '    TmpCASDODTime = Format(txtDODTime, "HH:MM")
                'Else
                '    TmpCASDODTime = Empty
                'End If

                If txtDOALGTime <> "__:__" Then
                    TmpCASDOALGTime = Format(txtDOALGTime, "HH:MM")
                Else
                    TmpCASDOALGTime = Empty
                End If
                
                If txtDODLGTime <> "__:__" Then
                    TmpCASDODLGTime = Format(txtDODLGTime, "HH:MM")
                Else
                    TmpCASDODLGTime = Empty
                End If

                
                'SaveCASOthers2 = txtDOATime_Check
                'If SaveCASOthers2 = False Then Exit Function
                
                'SaveCASOthers2 = txtDODTime_Check
                'If SaveCASOthers2 = False Then Exit Function
                
                'SaveCASOthers2 = txtDOALGTime_Check
                'If SaveCASOthers2 = False Then Exit Function
                
                'SaveCASOthers2 = txtDODLGTime_Check
                'If SaveCASOthers2 = False Then Exit Function
                
                'If txtDOAComment <> "__/__/____" Then
                '    TmpCASDOAComment = Format(txtDOAComment, "mm/dd/yyyy")
                'Else
                '    TmpCASDOAComment = Empty
                'End If
                
                If txtDODComment <> "__/__/____" Then
                    TmpCASDODComment = Format(txtDODComment, "mm/dd/yyyy")
                Else
                    TmpCASDODComment = Empty
                End If
                
                'If txtComplaintDate <> "__/__/____" Then
                '    TmpCASComplaintDate = Format(txtComplaintDate, "mm/dd/yyyy")
                'Else
                    TmpCASComplaintDate = Empty
                'End If
               
                '!CASTotalDoctors = TotalDoctor
                '!CASTotalAnaes = TotalAnaes
                '!CASTotalFinalDiag = TotalFinalDiag
                '!CASTotalProc = TotalProc
            
            'End With
        'CASOthersTwo.Update
        'CASOthersTwo.Close
        'Set CASOthersTwo = Nothing
        
        strsqlCASOthers2 = " EXEC Case_RegInsertCASOthers_2015 '" & ChkStr(CaseNumber) & "', " & _
                        "'" & ChkStr(TempDRName) & "', '" & ChkStr(TempDRName2) & "', '" & ChkStr(TempDRName3) & "', " & _
                        "'" & ChkStr(TempDRName4) & "', '" & ChkStr(TempDRName5) & "', '" & ChkStr(TempAnaesName) & "', '" & ChkStr(TempAnaesName2) & "', " & _
                        "'" & ChkStr(TempAnaesName3) & "', '" & TotalDoctor & "', '" & TotalAnaes & "', '" & TotalFinalDiag & "', '" & TotalProc & "', '" & Reason & "', '" & RefBy & "', " & _
                        "'" & TmpFirstDocDate & "', '" & TmpSubSeqDocDate & "', '" & TmpIntExtTO & "', '" & tmpHospName & "','" & Replace(txtAuditReview, vbNewLine, "~") & "'"

                        '"'" & TmpCASDOATime & "','" & TmpCASDODTime & "', '" & TmpCASDOALGTime & "', '" & TmpCASDODLGTime & "'," & _
                        '"'" & TmpCASDOAComment & "','" & TmpCASDODComment & "', '" & TmpCASComplaintDate & "', '" & TmpCASComplaintIssue & "', " & _
                        '"'" & TmpCASComplainant & "', '" & TmpCASCompAction & "'"
        medixmasterconn.Execute strsqlCASOthers2
End Function




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
Sub crt(pth As String)
Dim f
f = FreeFile
Open App.path + "\temp.html" For Output As #f
    Print #f, "<html><body><center><img src=""" & pth & """></center></body></html>"
Close #f
DoEvents
'WebBrowser1.Navigate App.Path & "\temp.html"
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)

Screen.MousePointer = vbHourglass

Dim frmTermsStatus As Integer
Dim frmPanelHospStatus As Integer
Dim tmpHOSCode As String
Dim PolList As ListItem
Dim HospNameList As ListItem

    If SSTab1.Tab = 4 Then
        lstBenefitDetails.ListItems.Clear
        lstNewBenefitDetails.ListItems.Clear
        Call BNFListing
        If txtUpgPLNCode <> "" Then
            Call NewBNFListing
        End If
    End If
    
    If SSTab1.Tab = 5 Then
        If lstMBMHis.ListItems.Count = 0 Then
            Call MBMHISListing
        End If
    End If
    
    If SSTab1.Tab = 7 Then
        If lstMBMHis.ListItems.Count = 0 Then
            Call SearchPrincipalIll
        End If
    End If
    
    If SSTab1.Tab = 9 Then
        If txtPLNCode <> "" Then
                frmTermsStatus = 1
                Call PolicyTermsListing(txtPLNCode, frmTermsStatus, txtPAYCode, txtHLTCode)
        End If
    End If
        
    If SSTab1.Tab = 10 Then
        If lstSurgProc.ListItems.Count = 0 Then
            frmTermsStatus = 1
            Call ScheduleSurg(txtPAYCode, txtPLNCode, frmTermsStatus)
        End If
    End If
    
    If SSTab1.Tab = 11 Then
        If Trim(tmpPolicyWording) <> "" Then
            Call LoadDoc
        End If
    End If
Screen.MousePointer = Default
End Sub

Private Sub SSTab1_DblClick()
Screen.MousePointer = vbHourglass

    If SSTab1.Tab = 4 Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        lstBenefitDetails.ListItems.Clear
        lstNewBenefitDetails.ListItems.Clear
        Call BNFListing
        If txtUpgPLNCode <> "" Then
            Call NewBNFListing
        End If
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing

    End If
    
    If SSTab1.Tab = 5 Then
        If lstMBMHis.ListItems.Count = 0 Then
         '   medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
            Call MBMHISListing
          '  medixmasterconnMaint.Close
          '  Set medixmasterconnMaint = Nothing
        End If
    End If
    
    If SSTab1.Tab = 7 Then
        If lstMBMHis.ListItems.Count = 0 Then
           ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call SearchPrincipalIll
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        End If
    End If
    
Screen.MousePointer = Default

End Sub

Private Sub SSTab1_LostFocus()
    If SSTab1 = 1 Then
        cboPatientList.SetFocus
    End If
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
    txtTotalAnaesMMARate = f
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
    txtTotalAnaesMMARate = f
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
    txtTotalAnaesMMARate = f
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
    txtTotalAnaesMMARate = f
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
    txtTotalAnaesMMARate = f
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
    txtTotalAnaesMMARate = f
End Sub

Private Sub txtAnnualLimit_GotFocus()
  KeyPreview = True
End Sub

Private Sub txtAnnualLimit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SSTab1.Tab = 1
    '    tabIndexDone = True
    End If
End Sub

Private Sub txtAuditReview_Click()
  KeyPreview = False
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

Private Sub txtCASRemarkTwo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 126 Then
        MsgBox "You are not allowed to use this symbol as a Keyword", vbOKOnly + vbCritical, DialogBoxTitle
        KeyAscii = 0
    End If
End Sub

Private Sub txtCASReserveAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtCASLGAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtDateAdmitted_GotFocus()
    PrevDOA = txtDateAdmitted
End Sub

Private Sub txtDateAdmitted_LostFocus()
Dim DOAErrStatus As Boolean
Dim SelectAge As String

    StartAgeStatus = False
    EndAgeStatus = False
    DOAErrStatus = False
    
    If ChkDOA.Value = 1 Then
        If IsDate(txtDateAdmitted) Then
            If txtDateDischarge <> "__/__/____" Then
                If (CDate(txtDateDischarge) < CDate(txtDateAdmitted)) Then
                    MsgBox "Admission Date cannot greater than Discharge Date! ", vbCritical
                    txtDateAdmitted = PrevDOA
                    txtDateAdmitted.SetFocus
                    DOAErrStatus = True
                End If
            End If
            If txtDateAdmitted <> "__/__/____" Then
                If CDate(txtDateAdmitted) > CDate(GetServerTime) And ChkDOA.Value = 1 Then
                    MsgBox "Admission Date cannot greater than system date ! ", vbCritical
                    txtDateAdmitted.SetFocus
                    DOAErrStatus = True
                End If
            End If
            If (txtDateAdmitted <> "__/__/____") Then
                If (CDate(txtDateAdmitted) < CDate(txtMBMEffDate)) Then
                    MsgBox "Admission Date cannot be less than Effective Date! ", vbCritical
                    'txtDateAdmitted = PrevDOA
                    txtDateAdmitted.SetFocus
                    DOAErrStatus = True
                End If
            End If
            
            If (txtDateAdmitted <> "__/__/____") Then
                If Mid(CboPayType, 1, 1) = "C" Then
                    If (CDate(txtMBMExpiryDate) < CDate(txtDateAdmitted)) Then
                        MsgBox "Admission Date cannot be greater than Expiry Date! ", vbCritical
                        'txtDateAdmitted = PrevDOA
                        txtDateAdmitted.SetFocus
                        DOAErrStatus = True
                    End If
                End If
            End If
        Else
            If txtDateAdmitted <> "__/__/____" Then
                MsgBox "Invalid date format ! ", vbCritical
                txtDateAdmitted.SetFocus
                DOAErrStatus = True
            End If
        End If
    End If
    
    If DOAErrStatus = False Then
        If txtMBMNumber <> "" And cboPatientList <> "" Then
            
            If Mid(cboPatientList, 1, 2) = "00" Then
                If txtDateAdmitted <> "__/__/____" And tmpMBMDOB <> "__/__/____" Then
                    SelectAge = GetAgeDays(CDate(txtDateAdmitted), CDate(tmpMBMDOB))
                    Call GetAgeCat(SelectAge)
                    If IsNumeric(tmpStartAge) And Trim(AgeCat) <> "Y" Then
                        If tmpStartAge >= AgeOnly Then
                            StartAgeStatus = True
                        Else
                            StartAgeStatus = False
                        End If
                    Else
                        StartAgeStatus = False
                    End If
                    
                    If IsNumeric(EndAgeStatus) And Trim(AgeCat) <> "D" Then
                        If Trim(tmpEndAge) <> "" Then
                            If tmpEndAge < AgeOnly Then
                                EndAgeStatus = True
                            Else
                                EndAgeStatus = False
                            End If
                        Else
                            EndAgeStatus = False
                        End If
                    Else
                        EndAgeStatus = False
                    End If
                    
                    
                End If
            Else
                If txtDateAdmitted <> "__/__/____" And tmpMBMCDOB <> "__/__/____" Then
                    SelectAge = GetAgeDays(CDate(txtDateAdmitted), CDate(tmpMBMCDOB))
                    Call GetAgeCat(SelectAge)
                    If IsNumeric(tmpStartAge) And Trim(AgeCat) <> "Y" Then
                        If tmpStartAge >= AgeOnly Then
                            StartAgeStatus = True
                        Else
                            StartAgeStatus = False
                        End If
                    Else
                        StartAgeStatus = False
                    End If
                    
                    If IsNumeric(EndAgeStatus) And Trim(AgeCat) <> "D" Then
                        If tmpEndAge <> "" Then
                            If tmpEndAge < AgeOnly Then
                                EndAgeStatus = True
                            Else
                                EndAgeStatus = False
                            End If
                        Else
                            EndAgeStatus = False
                        End If
                    Else
                        EndAgeStatus = False
                    End If
                    
                End If
            End If
             
            If StartAgeStatus = True Then
                MsgBox "Please verify Patient's Age!", vbOKOnly + vbInformation
            End If
        
            If EndAgeStatus = True Then
                MsgBox "Please verify Patient's Age!", vbOKOnly + vbInformation
            End If
        End If
    End If
End Sub

Private Sub txtDateDischarge_LostFocus()
    If ChkDOD.Value = 1 Then
        If IsDate(txtDateDischarge) Then
            If txtDateAdmitted <> "__/__/____" Then
                If (CDate(txtDateDischarge) < CDate(txtDateAdmitted)) Then
                    MsgBox "Discharge Date must be greater then Admission date! ", vbCritical
                    txtDateDischarge.SetFocus
                End If
            End If
            If txtDateDischarge <> "__/__/____" Then
                If CDate(txtDateDischarge) > CDate(GetServerTime) And ChkDOD.Value = 1 Then
                    MsgBox "Discharge Date must be greater than system date ! ", vbCritical
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
            
            If (txtDateDischarge <> "__/__/____") Then
                If Mid(CboPayType, 1, 1) = "C" Then
                    If (CDate(txtMBMExpiryDate) < CDate(txtDateDischarge)) Then
                        MsgBox "Discharge Date cannot be greater than Expiry Date! ", vbCritical
                        'txtDateAdmitted = PrevDOA
                        txtDateDischarge.SetFocus
                    End If
                End If
            End If
            
        Else
            If txtDateDischarge <> "__/__/____" Then
                MsgBox "Invalid Date Format! ", vbCritical
                txtDateDischarge.SetFocus
            End If
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

Private Sub txtDOALGTime_LostFocus()
    Dim TempDate As Date
    If txtDOALGTime.Text <> "__:__" Then
        Dim strdate As String
        strdate = txtDOALGTime & ":00"
        
        On Error Resume Next
        TempDate = DateTime.TimeValue(strdate)
     
        If Err.Number <> 0 Then
            MsgBox strdate & " is Invalid Time Format"
            txtDOALGTime.SetFocus
        End If
    End If
End Sub

Private Sub txtDODLGTime_LostFocus()
Dim TempDate As Date
    If txtDODLGTime.Text <> "__:__" Then
        Dim strdate As String
        strdate = txtDODLGTime & ":00"
        
        On Error Resume Next
        TempDate = DateTime.TimeValue(strdate)
     
        If Err.Number <> 0 Then
            MsgBox strdate & " is Invalid Time Format"
            txtDODLGTime.SetFocus
        End If
    End If
End Sub

Private Sub txtDODTime_LostFocus()
Dim TempDate As Date
    If txtDODTime.Text <> "__:__" Then
        Dim strdate As String
        strdate = txtDODTime & ":00"
        
        On Error Resume Next
        TempDate = DateTime.TimeValue(strdate)
     
        If Err.Number <> 0 Then
            MsgBox strdate & " is Invalid Time Format"
            txtDODTime.SetFocus
        End If
    End If
End Sub


Private Sub txtDODComment_LostFocus()
If IsDate(txtDODComment) Then
    Else
        If txtDODComment <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDODComment.SetFocus
        End If
   End If
End Sub

'Private Sub txtComplaintDate_LostFocus()
'    If IsDate(txtComplaintDate) Then
'    Else
'        If txtComplaintDate <> "__/__/____" Then
'            MsgBox "Invalid Date Format! ", vbCritical
'            txtComplaintDate.SetFocus
'        End If
'    End If
'End Sub

'Private Sub txtDOAComment_LostFocus()
'     If IsDate(txtDOAComment) Then
'         Else
'            If txtDOAComment <> "__/__/____" Then
'                MsgBox "Invalid Date Format! ", vbCritical
'                txtDOAComment.SetFocus
'            End If
'    End If
'End Sub

'Private Sub txtDOB_GotFocus()
'    KeyPreview = True
'    If txtDOB <> "__/__/____" Then
'        txtMorW = Year(GetServerTime) - Year(txtDOB)
'        txtMorW = txtMorW & " Years Old"
'    End If
'End Sub

'Private Sub txtDOB_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        SSTab1.Tab = 1
'    End If
'End Sub

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

Private Sub txtHOSCode_Click()
Dim frmTermsStatus As Integer
Dim frmPanelHospStatus As Integer
Dim tmpHOSCode As String
Dim MBM As New ADODB.Recordset
Dim txtPLNCode As String
Dim sqlstr As String


   ' frmTermsStatus = 2
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
        If txtHOSCode <> "" Then
            tmpHOSCode = HOSArr.Item(txtHOSCode)
            'StrAppend = StrAppend + " AND HOS.HOSCode = '" & Trim(tmpHOSCode) & "'"
        End If
    
   ' Call PanelHospListing(tmpPlnCode, frmTermsStatus, frmPanelHospStatus, txtPAYCode, tmpHOSCode, txtHOSCode, txtHLTCode)
   
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    


    End Sub

Private Sub txtHOSCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    Clipboard.Clear
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtHOSCode, txtHOSCode.SelStart) + Chr(KeyAscii) + Mid(txtHOSCode, txtHOSCode.SelStart + txtHOSCode.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To txtHOSCode.ListCount - 1

        If NewText = LCase(Left(txtHOSCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = txtHOSCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               txtHOSCode = ValidValue
               txtHOSCode.SelStart = 0
               txtHOSCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub txtHOSCode_LostFocus()
Dim tmpchkHopsCode As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpAXAClerical As String
Dim tmpPrivate As Boolean
    
    tmpPrivate = False
    
    If Len(txtHOSCode) Then
    
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            If (txtHOSCode = "DENTAL CLINIC" Or txtHOSCode = "DIALYSIS CENTRE" Or txtHOSCode = "FOREIGN HOSPITAL" Or txtHOSCode = "GENERAL PRACTITIONER" Or txtHOSCode = "GOVERNMENT HOSPITAL" Or txtHOSCode = "NON-PANEL HOSPITAL" Or txtHOSCode = "SPECIALIST CLINIC") And Trim(txtHOSName) = "" Then
                MsgBox "Please Enter Hospital Name or Clinic Name!", vbOKOnly + vbInformation
                txtHOSName.Enabled = True
            ElseIf (txtHOSCode <> "DENTAL CLINIC" And txtHOSCode <> "DIALYSIS CENTRE" And txtHOSCode <> "FOREIGN HOSPITAL" And txtHOSCode <> "GENERAL PRACTITIONER" And txtHOSCode <> "GOVERNMENT HOSPITAL" And txtHOSCode <> "NON-PANEL HOSPITAL" And txtHOSCode <> "SPECIALIST CLINIC") And Trim(txtHOSName) <> "" Then
                txtHOSName.Enabled = False
                txtHOSName = ""
            End If
            
            
        If (txtHOSCode = "DENTAL CLINIC" Or txtHOSCode = "DIALYSIS CENTRE" Or txtHOSCode = "FOREIGN HOSPITAL" Or txtHOSCode = "GENERAL PRACTITIONER" Or txtHOSCode = "GOVERNMENT HOSPITAL" Or txtHOSCode = "NON-PANEL HOSPITAL" Or txtHOSCode = "SPECIALIST CLINIC") Then
        
        Else
        
            If Mid(CboPayType, 1, 1) = "C" Then
                tmpchkHopsCode = HOSArr.Item(txtHOSCode)
                strsqlHOS = ""
                strsqlHOS = "select * from HOS where HOSCode = '" & Trim(tmpchkHopsCode) & "'"
                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                   ' HOS.Close
                   ' Set HOS = Nothing
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
                            
                            
                            If txtPAYCode = "AX" And txtMBMPolNo = "03215666" Then
                                tmpAXAClerical = "AX - AXA AFFIN STAFF (NON-EXECUTIVE)"
                                strsqlHOS = ""
                                strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(txtPAYCode) & "' and HosPanelCode = '" & tmpAXAClerical & "'"
                                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                            ElseIf txtPAYCode = "AX" And (txtMBMPolNo = "03215814" Or txtMBMPolNo = "03215886") Then
                                tmpAXAClerical = "AX - SPNB"
                                strsqlHOS = ""
                                strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(txtPAYCode) & "' and HosPanelCode = '" & tmpAXAClerical & "'"
                                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                            ElseIf txtPAYCode = "AX" And txtMBMPolNo = "03312049" Then
                                tmpAXAClerical = "AX - EMBASSY OF LIBYA (STAFF&MT)"
                                strsqlHOS = ""
                                strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(txtPAYCode) & "' and HosPanelCode = '" & tmpAXAClerical & "'"
                                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                            
                            Else
                                tmpAXAClerical = "AX - AXA AFFIN"
                                strsqlHOS = "select * from PanelHospitals where HOSCode = '" & Trim(tmpchkHopsCode) & "' and PayCode = '" & Trim(txtPAYCode) & "'  and HosPanelCode = '" & tmpAXAClerical & "'"
                                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                            End If
                                                        
                                                        
                            
                            If HOS.EOF = True Then
                                MsgBox "The selected hospital is not on the panel!", vbOKOnly + vbInformation
                                cmdRegister.Enabled = False
                            Else
                                cmdRegister.Enabled = True
                            End If
                            HOS.Close
                            Set HOS = Nothing
                        End If
                    End If
                End If
            End If
        End If
    '    medixmasterconn.Close
     '   Set medixmasterconn = Nothing
    End If
End Sub

Private Sub txtHOSName_LostFocus()
    txtHOSName = ChkStr(txtHOSName)
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

Private Sub txtMBMNumber_GotFocus()
    cmdShow.Default = True
    cmdRegister.Enabled = False
End Sub

Private Sub txtMBMNumber_LostFocus()
txtMBMNumber = ChkStr(txtMBMNumber)
cmdShow.Default = False
End Sub

Private Sub txtMBMPolYr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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
    txtTotalNSMMARate = f

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
    txtTotalNSMMARate = f

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
    txtTotalNSMMARate = f

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
    txtTotalNSMMARate = f

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
    txtTotalNSMMARate = f

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
    txtTotalNSMMARate = f
End Sub

Private Sub txtNSProcCode1_Change()
    If Len(Trim(txtNSProcCode1)) Then
        Call SearchMMARate(txtNSProcCode1)
        txtNSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate1 = 0
    End If
End Sub

Private Sub txtNSProcCode2_Change()
    If Len(Trim(txtNSProcCode2)) Then
        Call SearchMMARate(txtNSProcCode2)
        txtNSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate2 = 0
    End If
End Sub

Private Sub txtNSProcCode3_Change()
    If Len(Trim(txtNSProcCode3)) Then
        Call SearchMMARate(txtNSProcCode3)
        txtNSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate3 = 0
    End If
End Sub

Private Sub txtNSProcCode4_Change()
    If Len(Trim(txtNSProcCode4)) Then
        Call SearchMMARate(txtNSProcCode4)
        txtNSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate4 = 0
    End If
End Sub

Private Sub txtNSProcCode5_Change()
    If Len(Trim(txtNSProcCode5)) Then
        Call SearchMMARate(txtNSProcCode5)
        txtNSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate5 = 0
    End If
End Sub

Private Sub txtNSProcCode6_Change()
    If Len(Trim(txtNSProcCode6)) Then
        Call SearchMMARate(txtNSProcCode6)
        txtNSMMARate6 = IIf(Len(MMASurgFees), MMASurgFees, 0)
    Else
        txtNSMMARate6 = 0
    End If
End Sub

Private Sub txtPatientAge_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPatientAge_LostFocus()
Dim tmpPatientAge As String

    StartPatientAge = False
    EndPatientAge = False
    If IsNumeric(txtPatientAge) And Mid(cboYearDay, 1, 1) <> "Y" And cboYearDay <> "" Then
        If Mid(cboYearDay, 1, 1) = "W" Then
            tmpPatientAge = txtPatientAge * 7
        Else
            tmpPatientAge = txtPatientAge * 30
        End If
        If tmpStartAge <> "" Then
            If CDbl(tmpStartAge) >= CDbl(tmpPatientAge) Then
                StartPatientAge = True
            Else
                StartPatientAge = False
            End If
        Else
            StartPatientAge = False
        End If
    Else
        StartPatientAge = False
    End If
    
    If IsNumeric(txtPatientAge) And Mid(cboYearDay, 1, 1) = "Y" And cboYearDay <> "" Then
        If tmpEndAge <> "" Then
            If CDbl(tmpEndAge) < CDbl(txtPatientAge) Then
                EndPatientAge = True
            Else
                EndPatientAge = False
            End If
        Else
            EndPatientAge = False
        End If
    Else
        EndPatientAge = False
    End If
    
    If StartPatientAge = True Then
        MsgBox "Please verify Patient's Age!", vbOKOnly + vbInformation
    End If

    If EndPatientAge = True Then
        MsgBox "Please verify Patient's Age!", vbOKOnly + vbInformation
    End If

    
End Sub

Private Sub txtPatientList_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyCharacter(KeyAscii)
End Sub

Private Sub txtPAYCode_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
End Sub

Private Sub txtPAYCode_LostFocus()
    If Len(txtPAYCode) > 2 Then
        MsgBox "Payor cannot excceed 2 character!", vbCritical
        txtPAYCode.SetFocus
    End If
End Sub

Private Sub txtPDDOS1_LostFocus()

    If IsDate(txtPDDOS1) Then
    Else
        If txtPDDOS1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPDDOS1.SetFocus
        End If
    End If
End Sub

Private Sub txtPDDOS2_LostFocus()

    If IsDate(txtPDDOS2) Then
    
    Else
    If txtPDDOS2 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtPDDOS2.SetFocus
    End If
    End If
End Sub

Private Sub txtPDDOS3_LostFocus()
        
    If IsDate(txtPDDOS3) Then
    Else
    If txtPDDOS3 <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtPDDOS3.SetFocus
    End If
    End If
End Sub



Private Sub txtProcCode1_Change()
    If Len(Trim(txtProcCode1)) Then
        Call SearchMMARate(txtProcCode1)
        txtSMMARate1 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate1 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate1 = 0
         txtAnaesMMARate1 = 0
    End If
End Sub

Private Sub txtProcCode2_Change()
    If Len(Trim(txtProcCode2)) Then
        Call SearchMMARate(txtProcCode2)
        txtSMMARate2 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate2 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate2 = 0
         txtAnaesMMARate2 = 0
    End If
End Sub

Private Sub txtProcCode3_Change()
    If Len(Trim(txtProcCode3)) Then
        Call SearchMMARate(txtProcCode3)
        txtSMMARate3 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate3 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate3 = 0
         txtAnaesMMARate3 = 0
    End If
End Sub

Private Sub txtProcCode4_Change()
    If Len(Trim(txtProcCode4)) Then
        Call SearchMMARate(txtProcCode4)
        txtSMMARate4 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate4 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate4 = 0
         txtAnaesMMARate4 = 0
    End If
End Sub

Private Sub txtProcCode5_Change()
    If Len(Trim(txtProcCode5)) Then
        Call SearchMMARate(txtProcCode5)
        txtSMMARate5 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate5 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate5 = 0
         txtAnaesMMARate5 = 0
    End If
End Sub

Private Sub txtProcCode6_Change()
    If Len(Trim(txtProcCode6)) Then
        Call SearchMMARate(txtProcCode6)
        txtSMMARate6 = IIf(Len(MMASurgFees), MMASurgFees, 0)
        txtAnaesMMARate6 = IIf(Len(MMAAnaeFees), MMAAnaeFees, 0)
    Else
         txtSMMARate6 = 0
         txtAnaesMMARate6 = 0
    End If

End Sub

Private Sub txtReason_Click()
    KeyPreview = False
End Sub

Private Sub txtReason_LostFocus()
    KeyPreview = True
End Sub


Private Sub txtRoomNBoardDays_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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
    txtSTotalMMARate1 = f
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
    txtSTotalMMARate1 = f
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
    txtSTotalMMARate1 = f
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
    txtSTotalMMARate1 = f
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
    txtSTotalMMARate1 = f
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
    txtSTotalMMARate1 = f

End Sub

Private Sub txtState_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SSTab1.Tab = 1
        'tabIndexDone = True
    End If
End Sub

Private Sub txtState_LostFocus()
    SSTab1.Tab = 1
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

Private Sub MBMHISListing()
    Dim MBMHIS As New ADODB.Recordset
    Dim HISList As ListItem
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String
    
    If Len(Trim(txtMBMIcBcPP)) <> 0 Then
        If lstMBMHis.ListItems.Count = 0 Then
            'sql = " select MBMNumber, MBMPolicyNo, " & _
                " MBMPayorEffDate, MBMPayorExpDate, " & _
                " INSCode, PLNCode from MBMCrossReference " & _
                " where MBMIcBcPp ='" & txtMBMIcBcPP & "'"
            'MBMHIS.Open sql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppendCase = "3"
            strAppend = " AND MBMIcBcPp ='" & txtMBMIcBcPP & "'"
            Set cmd.ActiveConnection = medixmasterconnMaint
            cmd.CommandText = "Case_Registration"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append Prm2
            MBMHIS.CursorLocation = adUseClient
            MBMHIS.CursorType = adOpenForwardOnly
            MBMHIS.CacheSize = 100
            Set MBMHIS = cmd.Execute
            
            If MBMHIS.EOF = False Then
                Do Until MBMHIS.EOF
                    Set HISList = lstMBMHis.ListItems.Add(, , MBMHIS!MBMNumber)
                    With MBMHIS
                        HISList.SubItems(1) = !MBMPolicyNo
                        HISList.SubItems(2) = !MBMPayorEffDate
                        HISList.SubItems(3) = !MBMPayorEXPDate
                        HISList.SubItems(4) = !INSCode
                        HISList.SubItems(5) = !PLNCode
                        MBMHIS.MoveNext
                    End With
                Loop
            End If
            MBMHIS.Close
            Set MBMHIS = Nothing
        End If
    End If
End Sub


Public Sub CallLimit()
Dim sqlstr As String
Dim AnnualLimit As String
Dim PLN As String
Dim INS As String
Dim HLT As String
Dim Effdate As Date
Dim MBM As New ADODB.Recordset
    
    If txtMBMNumber <> "" Then
        sqlstr = "Select PLNCode, INSCode, HLTCode, MBMPayorEffDate from MBMCrossReference Where MBMNumber = '" & RemStr(txtMBMNumber) & "'"
        MBM.Open sqlstr, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
        PLN = Trim(MBM!PLNCode)
        INS = Trim(MBM!INSCode)
        HLT = Trim(MBM!HLTCode)
        Effdate = CDate(MBM!MBMPayorEffDate)
        
        txtAnnualLimit = ""
        txtAnnualLimit = Format(GetAnnualLimit(INS, PLN, HLT, Effdate), "#,##0.00")
    End If
    MBM.Close
    Set MBM = Nothing
End Sub

Private Sub funcword()
'Dim MBM As New ADODB.Recordset
'Dim strsqlMBM As String
On Error GoTo openerror
    
    ' make sure only one word application launched
    'txtpathLoc
    'txtSVRPath = CrdPath & "Template\"
    
    'strsqlMBM = "Select PLNCode, LGInd from PLN where PLNCode = '" & Trim(txtPLNCode) & "'"
    'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'LocCardPath = crdPath
    If Letter = "L" And Trim(tmpPLNPerVisit) <> "Y" Then
    
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\New LG-2.DOT")
        End If
        
        If bookmark("New LG-2.DOT") = True Then
            Me.Show
        End If
    ElseIf Letter = "L" And Trim(tmpPLNPerVisit) = "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\LGVisit.DOT")
        End If
        
        If bookmark("LGVisit.DOT") = True Then
            Me.Show
        End If
    ElseIf Letter = "B" And Trim(tmpPLNPerVisit) <> "Y" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\BTBG.DOT")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
              vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              objdoc.Save
              objdoc.Close
            ElseIf vbCancel Then
              Exit Sub
            End If
        End If
        
        If bookmark("BTBG.DOT") = True Then
                Me.Show
        End If
        Exit Sub
    ElseIf Letter = "B" And Trim(tmpPLNPerVisit) = "Y" Then
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
                Me.Show
        End If
        Exit Sub
    
    ElseIf Letter = "B" And Trim(tmpPLNPerVisit) = "Y" Then
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
                Me.Show
        End If
        Exit Sub
    
    ElseIf Letter = "C" And Trim(tmpPLNPerVisit) <> "Y" Then
    
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\New LG-2.DOT")
        End If
        
        If bookmark("New LG-2.DOT") = True Then
            Me.Show
        End If
    
    ElseIf Letter = "R" Then
    Set objword = Nothing
    If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            'If Trim(Mid(cboREPCode1, 1, 5)) = "RC6" Then
            '    Set objdoc = objword.Documents.Add(LocCardPath & "Case" & "\IMB.DOT\")
            '    Set objdoc = Nothing
            '    Set objdoc = objword.Documents.Add(LocCardPath & "Case" & "\REPLetterCOR2.DOT\")
            'Else
                'Set objdoc = objword.Documents.Add(LocCardPath & "Case" & "\IMB.DOT\")
                Set objdoc = Nothing
                Set objdoc = objword.Documents.Add(LocCardPath & "Case" & "\REPLetter.DOT")
            'End If
    ElseIf Not objdoc Is Nothing Then
          If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
            vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
            objdoc.Save
            objdoc.Close
          ElseIf vbCancel Then
            Exit Sub
          End If
            'Set objdoc = Nothing
            'Set objdoc = objword.Documents.Add(LocCardPath & "Case" & "\IMB.DOT\")
            Set objdoc = Nothing
            Set objdoc = objword.Documents.Add(LocCardPath & "Case" & "REPLetter.DOT")
    End If
        bookmark ("REPLetter.DOT")
        Me.Show
        'If bookmark("REPLetter.DOT") = True Then
        '    objdoc.SaveAs (LocCardPath & "\Repudiation\" & Trim(txtCaseNo) & ".doc")
        '    Me.show
        '    Exit Sub
        'End If
    
    ElseIf Letter = "P" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "Case" & "\PriorLetter.DOT")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
              vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              objdoc.Save
              objdoc.Close
            ElseIf vbCancel Then
              Exit Sub
            End If
        End If
        
        If bookmark("PriorLetter.DOT") = True Then
            Me.Show
            Exit Sub
        End If

    ElseIf Letter = "E" Then
        Set objword = Nothing
        If objword Is Nothing Then
            Set objword = New Word.Application
            objword.Visible = True
            Set objdoc = objword.Documents.Add(LocCardPath & "Case" & "\PriorExpLetter.dot")
        ElseIf Not objdoc Is Nothing Then
            If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
              vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
              objdoc.Save
              objdoc.Close
            ElseIf vbCancel Then
              Exit Sub
            End If
        End If
        
        If bookmark("PriorExpLetter.dot") = True Then
            Me.Show
            Exit Sub
        End If

    'End If
    
    ElseIf Letter = "T" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\Take Home Drugs.doc\")
        End If
        
        
        If bookmark("Take Home Drugs.doc") = True Then
            Me.Show
        End If
        Exit Sub
        
    ElseIf Letter = "A" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\DocRequestAdm.dot")
        End If
        
        
        If bookmark("DocRequestAdm.dot") = True Then
            Me.Show
        End If
        Exit Sub
        
    ElseIf Letter = "D" Then
        Set objword = Nothing
        If objword Is Nothing Then
                Set objword = New Word.Application
                objword.Visible = True
                
                Set objdoc = objword.Documents.Add(LocCardPath & "\Case\" & "\DocRequestDis.dot")
        End If
        
        
        If bookmark("DocRequestDis.dot") = True Then
            Me.Show
        End If
        Exit Sub
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
    'End If
End Sub

Private Function bookmark(Optional strAppend As String) As Integer
Dim strsqlPay As String
Dim PAY As New ADODB.Recordset
Dim strsqlRep As String
Dim REP As New ADODB.Recordset
Dim REPCode As String
Dim HOS As New ADODB.Recordset
Dim strsqlHOS As String
Dim strsqlLG As String
Dim LG As New ADODB.Recordset
Dim REPCat As New ADODB.Recordset
Dim strsqlREPCat As String
Dim ContactPerson As String
Dim MBM As New ADODB.Recordset
Dim strsqlMBM As String
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
Dim cmd9 As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
    ContactPerson = ""
    Including = ""
    Terms = ""
    
    HOspitalName = Replace(txtHOSCode, "'", "''")
    'strsqlHOS = "select * from HOS where HOSName = '" & Trim(HOspitalName) & "'"
    'HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "16"
    StrAppend1 = " AND HOSName = '" & Trim(HOspitalName) & "'"
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "Case_Registration"
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
    
    If Letter = "A" Or Letter = "D" Then
        bookmark = False
        objdoc.Bookmarks("SystemDate").Range.Text = Format(GetServerTime, "DD/MM/YYYY")
        
        objdoc.Bookmarks("MedixStaffName").Range.Text = LogonName
        
        If Len(txtCaseNo) Then
            objdoc.Bookmarks("CaseNo").Range.Text = txtCaseNo
        End If
        
        If Len(txtHOSCode) Then
            objdoc.Bookmarks("HOSName").Range.Text = txtHOSCode
        End If
        
        If Len(HOS!HOSContact) Then
            objdoc.Bookmarks("HOSStaff").Range.Text = HOS!HOSContact
        End If
        
        If Len(HOS!HOSFaxNo) Then
            objdoc.Bookmarks("HospFax").Range.Text = HOS!HOSFaxNo
        End If

        If Trim(cboPatientList) <> "" Then
            objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
        Else
            objdoc.Bookmarks("patient").Range.Text = txtPatientList
        End If

        If txtMBMNumber <> "" Then
            objdoc.Bookmarks("mbmno").Range.Text = txtMBMNumber
        End If
        
        If txtPatientIC <> "" Then
            objdoc.Bookmarks("PatientIC").Range.Text = txtPatientIC
        End If
        
        If txtDateAdmitted <> "" Then
            objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
        End If
                
        If txtDateDischarge <> "" Then
            objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
        End If
                
        bookmark = True
    
    ElseIf Letter = "T" Then
        bookmark = False
        If Len(txtCaseNo) Then
            objdoc.Bookmarks("CaseNo").Range.Text = txtCaseNo
        End If

        If Logon <> "" Then
            objdoc.Bookmarks("ppby").Range.Text = Logon
        End If
        
        If Trim(cboPatientList) <> "" Then
            objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
        Else
            objdoc.Bookmarks("patient").Range.Text = txtPatientList
        End If

        If txtMBMNumber <> "" Then
            objdoc.Bookmarks("mbmno").Range.Text = txtMBMNumber
        End If
        
        If txtMBMPolNo <> "" Then
            objdoc.Bookmarks("PolicyNo").Range.Text = txtMBMPolNo
        End If
        
        If txtDateAdmitted <> "" Then
            objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
        End If
                
        If txtDateDischarge <> "" Then
            objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
        End If
                
        bookmark = True
    
    Else
        strAppendCase = "15"
        StrAppend1 = " AND PAYCode = '" & Trim(txtPAYCode) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Registration"
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
            
        
        If Letter = "L" Or Letter = "C" Then
            bookmark = False
            If Trim(cboPatientList) <> "" Then
                objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
            Else
                objdoc.Bookmarks("patient").Range.Text = txtPatientList
            End If
            If LGNumber <> "" Then
                objdoc.Bookmarks("LGno").Range.Text = LGNumber
            End If
            
            'If LGLogon <> "" Then
            objdoc.Bookmarks("LGID").Range.Text = Trim(Logon)
            'End If
            
            If txtMBMNumber <> "" Then
                objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
            End If
            If txtHOSCode <> "" Then
                objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
            End If
            
            If txtDateAdmitted <> "" Then
                objdoc.Bookmarks("dateofadmission").Range.Text = txtDateAdmitted
            End If
            If txtMBMPolNo <> "" Then
                objdoc.Bookmarks("policyno").Range.Text = txtMBMPolNo
            End If
            If txtCASLGAmt <> "" Then
                objdoc.Bookmarks("amount").Range.Text = Format(txtCASLGAmt, "#,##0.00")
            End If
            If Trim(cboPatientList) <> "" Then
                If tmpRBType <> "" Then
                    objdoc.Bookmarks("RB").Range.Text = tmpRBType
                ElseIf RB <> "" Then
                    objdoc.Bookmarks("RB").Range.Text = RB
                End If
            End If
            If Len(HOS!HOSContact) Then
                objdoc.Bookmarks("Attention").Range.Text = HOS!HOSContact
            End If
            If Len(HOS!HOSFaxNo) Then
                objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
            End If
            
            If TxtLGRemark <> "" Then
                objdoc.Bookmarks("Remark").Range.Text = TxtLGRemark
            End If
            If Letter = "C" Then
                If tmpGRPCompany <> "" Then
                    objdoc.Bookmarks("GRPCompany").Range.Text = " - " & tmpGRPCompany
                End If
            End If
            
            If txtHLTCode = "S" Then
                'strsqlMBM = "Select PLNCode, CoPaymentInd, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(txtPLNCode) & "'"
                strAppendCase = "17"
                StrAppend1 = " AND PLNCode = '" & Trim(txtPLNCode) & "'"
            Else
                'strsqlMBM = "Select PLNCode, CoPaymentInd, PLNMeal, PLNNursing, PLNTax, PLNMRI from PLN where PLNCode = '" & Trim(txtPLNCode) & "' AND PAYCode = '" & Trim(txtPAYCode) & "'"
                strAppendCase = "17"
                StrAppend1 = " AND PLNCode = '" & Trim(txtPLNCode) & "' AND PAYCode = '" & Trim(txtPAYCode) & "'"
            End If
            'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
            Set cmd2.ActiveConnection = medixmasterconn
            cmd2.CommandText = "Case_Registration"
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
            Else
                COPay = False
                MRI = False
            End If
            MBM.Close
            Set MBM = Nothing
            
            If Mid(txtPLNCode, 1, 1) = "4" Then
                Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                         vbNewLine & "       physician." & _
                         vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                         vbNewLine & "4)    The following items are not covered: -" & _
                         vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                         vbNewLine & "           (b) Non-Medical items." & _
                         vbNewLine & "           (c) Medical Report Charges." & _
                         vbNewLine & "           (d) Registration Fees, Admission Fees, Billing Fees and Admission Kit." & _
                         vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Lenses, Pacemaker, Braces, Corset " & _
                         vbNewLine & "           (f) Lodger" & _
                         vbNewLine & "           (g) Supplies & Services not related to illness injury." & _
                         vbNewLine & "5)    Co-Payment - In the event the Room and Board limit is exceeded, the Insurance will pay only 80% of " & _
                         vbNewLine & "       other eligible costs."
                         
                
            ElseIf COPay = True And MRI = True And Trim(Letter) <> "P" Then
    
                Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                         vbNewLine & "       physician." & _
                         vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                         vbNewLine & "4)    The following items are not covered: -" & _
                         vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                         vbNewLine & "           (b) Non-Medical items." & _
                         vbNewLine & "           (c) Medical Report Charges." & _
                         vbNewLine & "           (d) Registration Fees, Admission Fees, Billing Fees and Admission Kit." & _
                         vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Lenses, Pacemaker, Braces, Corset, Arm sling and Cervical" & _
                         vbNewLine & "                Collar." & _
                         vbNewLine & "           (f) Lodger" & _
                         vbNewLine & "           (g) Supplies & Services not related to illness injury." & _
                         vbNewLine & "5)    Co-Payment - In the event the Room and Board limit is exceeded, the Insurance will pay only 80% of " & _
                         vbNewLine & "       other eligible costs." & _
                         vbNewLine & "6)    Insurance will pay only 80% of MRI, CT Scan and Heart Scan charges."
                                
            ElseIf COPay = False And MRI = True And Trim(Letter) <> "P" Then
                
                Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                         vbNewLine & "       physician." & _
                         vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                         vbNewLine & "4)    The following items are not covered: -" & _
                         vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                         vbNewLine & "           (b) Non-Medical items." & _
                         vbNewLine & "           (c) Medical Report Charges." & _
                         vbNewLine & "           (d) Registration Fees, Admission Fees, Billing Fees, Dispensing Fees and Admission Kit." & _
                         vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Lenses, Pacemaker, Braces, Corset, Arm sling and Cervical" & _
                         vbNewLine & "                Collar." & _
                         vbNewLine & "           (f) Lodger" & _
                         vbNewLine & "           (g) Supplies & Services not related to illness injury." & _
                         vbNewLine & "5)    Insurance will pay only 80% of MRI, CT Scan and Heart Scan charges."
                                     
            ElseIf COPay = True And MRI = False And Trim(Letter) <> "P" Then
                     
                Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                         vbNewLine & "       physician." & _
                         vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                         vbNewLine & "4)    The following items are not covered: -" & _
                         vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                         vbNewLine & "           (b) Non-Medical items." & _
                         vbNewLine & "           (c) Medical Report Charges." & _
                         vbNewLine & "           (d) Registration Fees, Admission Fees, Billing Fees, Dispensing Fees and Admission Kit." & _
                         vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Lenses, Pacemaker, Braces, Corset, Arm sling and Cervical" & _
                         vbNewLine & "                Collar." & _
                         vbNewLine & "           (f) Lodger" & _
                         vbNewLine & "           (g) Supplies & Services not related to illness injury." & _
                         vbNewLine & "5)    Co-Payment - In the event the Room and Board limit is exceeded, the Insurance will pay only 80% of " & _
                         vbNewLine & "       other eligible costs."
                
            ElseIf COPay = False And MRI = False And Trim(Letter) <> "P" Then
                
                Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                         vbNewLine & "       physician." & _
                         vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                         vbNewLine & "4)    The following items are not covered: -" & _
                         vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                         vbNewLine & "           (b) Non-Medical items." & _
                         vbNewLine & "           (c) Medical Report Charges." & _
                         vbNewLine & "           (d) Registration Fees, Admission Fees, Billing Fees, Dispensing Fees and Admission Kit." & _
                         vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Lenses, Pacemaker, Braces, Corset, Arm sling and Cervical" & _
                         vbNewLine & "                Collar." & _
                         vbNewLine & "           (f) Lodger" & _
                         vbNewLine & "           (g) Supplies & Services not related to illness injury."
                                       
                                                   
            ElseIf Trim(Letter) = "C" Then
                
                Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                         vbNewLine & "       physician." & _
                         vbNewLine & "3)    All Doctors and Anaesthetists charges shall NOT exceed the fees as per the MMA schedule of Fees." & _
                         vbNewLine & "4)    The following items are not covered: -" & _
                         vbNewLine & "           (a) Items of personal nature such as telephone calls, newspaper, laundry, television etc." & _
                         vbNewLine & "           (b) Non-Medical items." & _
                         vbNewLine & "           (c) Admission Kit." & _
                         vbNewLine & "           (e) Appliances, Prosthesis - Crutches, Pacemaker, Braces, Corset" & _
                         vbNewLine & "                Collar." & _
                         vbNewLine & "           (f) Lodger" & _
                         vbNewLine & "           (g) Supplies & Services not related to illness injury."
            End If
            
            objdoc.Bookmarks("including").Range.Text = Including
            objdoc.Bookmarks("string").Range.Text = Terms
            bookmark = True
            
        ElseIf Letter = "B" Then
            bookmark = False
            
            If Trim(cboPatientList) <> "" Then
                objdoc.Bookmarks("patient").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
            Else
                objdoc.Bookmarks("patient").Range.Text = " Please key in patient name here!"
            End If
            
            If LGNumber <> "" Then
                objdoc.Bookmarks("BTBNo").Range.Text = LGNumber
            End If
            
            objdoc.Bookmarks("BTBID").Range.Text = Trim(Logon)
            
            If txtMBMNumber <> "" Then
                objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
            End If
            If txtHOSCode <> "" Then
                objdoc.Bookmarks("hospital").Range.Text = Trim(txtHOSCode)
            End If
            If txtDateAdmitted <> "" Then
                objdoc.Bookmarks("dateofadmission").Range.Text = txtDateAdmitted
            End If
            If txtMBMPolNo <> "" Then
                objdoc.Bookmarks("policyno").Range.Text = txtMBMPolNo
            End If
            
            If RB <> "" Then
                objdoc.Bookmarks("RB").Range.Text = RB
            End If
            
            If Len(HOS!HOSContact) Then
                objdoc.Bookmarks("Attention").Range.Text = HOS!HOSContact
            End If
        
            If Len(HOS!HOSFaxNo) Then
                objdoc.Bookmarks("Fax").Range.Text = HOS!HOSFaxNo
            End If
                    
            If txtHLTCode = "S" Then
                strAppendCase = "17"
                StrAppend1 = " AND PLNCode = '" & Trim(txtPLNCode) & "'"
            Else
                strAppendCase = "17"
                StrAppend1 = " AND PLNCode = '" & Trim(txtPLNCode) & "' AND PAYCode = '" & Trim(txtPAYCode) & "'"
            End If
            
            Set cmd2.ActiveConnection = medixmasterconn
            cmd2.CommandText = "Case_Registration"
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
    
            Terms = "2)    Medical Information relating to the patient and treatment must be truthfully disclosed by the attending " & _
                         vbNewLine & "        physician." & _
                          vbNewLine & "3)    Items of personal nature such as telephone calls, newspaper, laundry, television etc."
            
            objdoc.Bookmarks("including").Range.Text = Including
            objdoc.Bookmarks("string").Range.Text = Terms
            bookmark = True
        
        ElseIf Letter = "R" Then
            bookmark = False
            objdoc.Bookmarks("PrincipalName").Range.Text = txtMBMName
            objdoc.Bookmarks("Add1").Range.Text = txtAdd1
            objdoc.Bookmarks("Add2").Range.Text = txtAdd2 & txtAdd3
            objdoc.Bookmarks("PostCode").Range.Text = txtPostCode & "" & txtCity
            objdoc.Bookmarks("State").Range.Text = txtState
                
            If txtCaseNo <> "" Then
                objdoc.Bookmarks("CASENo").Range.Text = txtCaseNo
            End If
        
            If Trim(cboPatientList) <> "" Then
              objdoc.Bookmarks("patientname").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
            Else
              objdoc.Bookmarks("patientname").Range.Text = " Please key in patient name here!"
            End If
            If txtMBMNumber <> "" Then
                objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
            End If
            
            If txtFinalDiag1 <> "" Then
                Dim strsql As String
                strsql = strsql + txtFinalDiag1
                If txtFinalDiag2 <> "" Then
                    strsql = strsql + ", " + txtFinalDiag2
                    If txtFinalDiag3 <> "" Then
                        strsql = strsql + ", " + txtFinalDiag3
                    End If
                End If
                objdoc.Bookmarks("Reasons").Range.Text = strsql
            End If
            
            If Underlying = True Then
                If txtUD1 <> "" Then
                    Dim strsqlUD As String
                    strsqlUD = " With underlying illness as "
                    strsqlUD = strsqlUD + txtUD1
                    If txtUD2 <> "" Then
                        strsqlUD = strsqlUD + ", " + txtUD2
                        If txtUD3 <> "" Then
                            strsqlUD = strsqlUD + ", " + txtUD3
                        End If
                    End If
                    objdoc.Bookmarks("Underlying").Range.Text = strsqlUD + "."
                End If
            End If
            
            If txtHOSCode <> "" Then
                objdoc.Bookmarks("hospital").Range.Text = txtHOSCode
            End If
            If txtHOSCode <> "" Then
                objdoc.Bookmarks("hospital2").Range.Text = txtHOSCode
            End If
            If txtMBMExpiryDate <> "" Then
                objdoc.Bookmarks("ExpDate").Range.Text = txtMBMExpiryDate
            End If
            If txtDateAdmitted <> "" Then
                objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
            End If
            If txtMBMPolNo <> "" Then
                objdoc.Bookmarks("policynumber").Range.Text = txtMBMPolNo
            End If
            If txtMBMName <> "" Then
                objdoc.Bookmarks("InsuredName").Range.Text = txtMBMName
            End If
            If Document <> "" Then
                objdoc.Bookmarks("REPDocument").Range.Text = Trim(Mid(Document, InStr(1, Document, "-") + 1, Len(Document)))
            End If
            If Personnel <> "" Then
                objdoc.Bookmarks("REPPersonnel").Range.Text = Trim(Mid(Personnel, InStr(1, Personnel, "-") + 1, Len(Personnel)))
            End If
            If Attention <> "" Then
                objdoc.Bookmarks("REPAttention").Range.Text = Trim(Mid(Attention, InStr(1, Attention, "-") + 1, Len(Attention)))
            End If
            
                If Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - 1)) <> "RC6" Then
                    If Len(Trim(cboREPCode1)) Then
                        'strsqlREP = "Select * from REPCode where RepCode = '" & Trim(Mid(cboREPCode1, 1, 7)) & "'"
                        'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
                        strAppendCase = "18"
                        StrAppend1 = " AND RepCode = '" & Trim(Mid(cboREPCode1, 1, InStr(1, cboREPCode1, "-") - 1)) & "'"
                        Set cmd3.ActiveConnection = medixmasterconn
                        cmd3.CommandText = "Case_Registration"
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
                        
                        strAppendCase = "19"
                        StrAppend1 = " AND REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                        Set cmd4.ActiveConnection = medixmasterconn
                        cmd4.CommandText = "Case_Registration"
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
                    strAppendCase = "18"
                    StrAppend1 = " AND RepCode = '" & Trim(Mid(cboREPCode2, 1, InStr(1, cboREPCode2, "-") - 1)) & "'"
                    Set cmd5.ActiveConnection = medixmasterconn
                    cmd5.CommandText = "Case_Registration"
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
                    
                    strAppendCase = "19"
                    StrAppend1 = " AND REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                    Set cmd6.ActiveConnection = medixmasterconn
                    cmd6.CommandText = "Case_Registration"
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
                    
                    If REPCat.EOF = False Then
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
                    strAppendCase = "18"
                    StrAppend1 = " AND RepCode = '" & Trim(Mid(cboREPCode3, 1, InStr(1, cboREPCode3, "-") - 1)) & "'"
                    Set cmd7.ActiveConnection = medixmasterconn
                    cmd7.CommandText = "Case_Registration"
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
                    
                    strAppendCase = "19"
                    StrAppend1 = " AND REPCatCode = '" & Trim(REP!REPCatCode) & "'"
                    Set cmd8.ActiveConnection = medixmasterconn
                    cmd8.CommandText = "Case_Registration"
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
                    
                    If REP.recordCount Then
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
            If txtPAYCode <> "" Then
                objdoc.Bookmarks("Payor").Range.Text = PAY!PAYCompanyName
            End If
            
            If txtPAYCode <> "" Then
                objdoc.Bookmarks("Payor2").Range.Text = PAY!PAYCompanyName
            End If
            
            If Len(PAY!PAYAdmissionContact) Then
                    objdoc.Bookmarks("Pic").Range.Text = Trim(PAY!PAYAdmissionContact)
            End If
            
            'strsqlLG = "Select CASENo from CASLG where CASENo = '" & Trim(txtCaseNo) & "'"
            'LG.Open strsqlLG, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            strAppendCase = "4"
            StrAppend1 = " AND CASENo = '" & Trim(txtCaseNo) & "'"
            Set cmd9.ActiveConnection = medixmasterconnMaint
            cmd9.CommandText = "Case_Registration"
            cmd9.CommandType = adCmdStoredProc
            cmd9.CommandTimeout = 300
            Set Prm1 = cmd9.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
            cmd9.Parameters.Append Prm1
            Set Prm2 = cmd9.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd9.Parameters.Append Prm2
            LG.CursorLocation = adUseClient
            LG.CursorType = adOpenForwardOnly
            LG.CacheSize = 100
            Set LG = cmd9.Execute
            
            If LG.recordCount > 0 Then
                objdoc.Bookmarks("LDDesc").Range.Text = "MediExpress has facilitated your admission and settled your hospitalization " & _
                                                        "charges amounting to RM2,500.00. We enclose a copy of the bill for your records. " & _
                                                        "We would be pleased if you could let us have your payment to MediExpress(M) " & _
                                                        "Sdn. Bhd. within 14 days herewith."
            End If
            LG.Close
            Set LG = Nothing
            bookmark = True
            
        ElseIf Letter = "P" Then
            bookmark = False
            If PAY.recordCount Then
                objdoc.Bookmarks("Payor").Range.Text = PAY!PAYCompanyName
                objdoc.Bookmarks("ADD1").Range.Text = IIf(Len(PAY!PAYAddress1), PAY!PAYAddress1, "")
                objdoc.Bookmarks("ADD2").Range.Text = IIf(Len(PAY!PAYAddress2), PAY!PAYAddress2, "")
                objdoc.Bookmarks("PostCode").Range.Text = IIf(Len(PAY!PAYAddress3), PAY!PAYAddress3, "")
                ContactPerson = IIf(Len(PAY!PAYMAINContact), PAY!PAYMAINContact, "")
                
            End If
            objdoc.Bookmarks("REGDate").Range.Text = GetServerTime
            objdoc.Bookmarks("Month").Range.Text = Format(txtMBMExpiryDate, "MMMM YYYY")
                
            If txtCaseNo <> "" Then
                objdoc.Bookmarks("CASENo").Range.Text = txtCaseNo
            End If
        
            If Trim(cboPatientList) <> "" Then
              objdoc.Bookmarks("patientname").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
            Else
              objdoc.Bookmarks("patientname").Range.Text = " Please key in patient name here!"
            End If
            If txtMBMPolNo <> "" Then
                objdoc.Bookmarks("policynumber").Range.Text = txtMBMPolNo
            End If
            If txtMBMExpiryDate <> "__/__/____" Then
                objdoc.Bookmarks("EXPDate").Range.Text = txtMBMExpiryDate
            End If
            If txtMBMNumber <> "" Then
                objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
            End If
            If txtDateAdmitted <> "" Then
                objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
            End If
            If txtDateDischarge <> "__/__/____" Then
                objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
            End If
            If txtHOSCode <> "" Then
                objdoc.Bookmarks("hospital").Range.Text = txtHOSCode
            End If
            If txtReason <> "" Then
                objdoc.Bookmarks("Diag").Range.Text = txtReason
            End If
            bookmark = True
            
        ElseIf Letter = "E" Then
            bookmark = False
            If PAY.recordCount Then
                objdoc.Bookmarks("Payor").Range.Text = PAY!PAYCompanyName
                objdoc.Bookmarks("ADD1").Range.Text = IIf(Len(PAY!PAYAddress1), PAY!PAYAddress1, "")
                objdoc.Bookmarks("ADD2").Range.Text = IIf(Len(PAY!PAYAddress2), PAY!PAYAddress2, "")
                objdoc.Bookmarks("PostCode").Range.Text = IIf(Len(PAY!PAYAddress3), PAY!PAYAddress3, "")
                ContactPerson = IIf(Len(PAY!PAYMAINContact), PAY!PAYMAINContact, "")
                
            End If
            objdoc.Bookmarks("REGDate").Range.Text = GetServerTime
            objdoc.Bookmarks("Month").Range.Text = Format(txtMBMExpiryDate, "MMMM YYYY")
                
            If txtCaseNo <> "" Then
                objdoc.Bookmarks("CASENo").Range.Text = txtCaseNo
            End If
        
            If Trim(cboPatientList) <> "" Then
              objdoc.Bookmarks("patientname").Range.Text = Trim(Mid(cboPatientList, InStr(1, cboPatientList, "-") + 1, Len(cboPatientList) - InStr(1, cboPatientList, "-")))
            Else
              objdoc.Bookmarks("patientname").Range.Text = " Please key in patient name here!"
            End If
            If txtMBMPolNo <> "" Then
                objdoc.Bookmarks("policynumber").Range.Text = txtMBMPolNo
            End If
            If txtMBMExpiryDate <> "__/__/____" Then
                objdoc.Bookmarks("EXPDate").Range.Text = txtMBMExpiryDate
            End If
            If txtMBMNumber <> "" Then
                objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMNumber
            End If
            If txtDateAdmitted <> "" Then
                objdoc.Bookmarks("DOA").Range.Text = txtDateAdmitted
            End If
            If txtDateDischarge <> "__/__/____" Then
                objdoc.Bookmarks("DOD").Range.Text = txtDateDischarge
            End If
            If txtHOSCode <> "" Then
                objdoc.Bookmarks("hospital").Range.Text = txtHOSCode
            End If
            If txtReason <> "" Then
                objdoc.Bookmarks("Diag").Range.Text = txtReason
            End If
            
            bookmark = True
            
        End If
        HOS.Close
        Set HOS = Nothing
        'End If
    End If
End Function

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

Private Sub GetAccountSummary()
    Dim Acc As New ADODB.Recordset
    Dim strsql As String
    Dim Paid2MBM, PaidByMem, PayableMem As Double
    
    strsql = " select casnumber, totalApprove, totalPayable2Hos, " & _
            " totalPaid2Hos, totalPayable2Mbm, totalPaidByMbm , CLMFinalApprove, " & _
            " totalPaid2MBM , totalExcess , totalReimburse , CLMFinalApprove from view_account_summary where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    Acc.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                
                
    If Not Acc.EOF Then
        With Acc
            PayableMem = Format(!totalPayable2Mbm, "#,#.00;(#,#.00)")
            If IsNumeric(PayableMem) = False Then
                PayableMem = 0
            End If
            If Trim(!totalPaid2MBM) <> "" Then
                Paid2MBM = Format(!totalPaid2MBM, "#,#.00")
            End If
            If IsNumeric(Paid2MBM) = False Then
                Paid2MBM = 0
            End If
            If Trim(!totalPaidByMbm) <> "" Then
                PaidByMem = Format(!totalPaidByMbm, "#,#.00")
            End If
            If IsNumeric(PaidByMem) = False Then
                PaidByMem = 0
            End If
            'txtRecoveryOS = Format(CDbl(PayableMem) - CDbl(Paid2MBM) + CDbl(PaidByMem), "#,#.00;(#,#.00)")
        End With
    End If
    Acc.Close
    Set Acc = Nothing

End Sub

Private Sub NewBNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim bnf As ListItem
    Dim sql As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim strAppendCase As String
    Dim StrAppend1 As String
    
    If lstNewBenefitDetails.ListItems.Count = 0 And Len(Trim(txtHLTCode)) <> 0 _
     And Len(Trim(txtINSCode1)) <> 0 And Len(Trim(txtUpgPLNCode)) <> 0 Then
        'sql = " select BNFCode, BNFNoOFDays , " & _
          " BNFclaimableAmount, BNFdescription, BNFRMPerDis " & _
          " from View_BNF where PLNCOde ='" & Trim(txtUpgPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(txtHLTCode) & "'" & _
          " AND INSCODe ='" & Trim(txtINSCode1) & "'"
        'rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "20"
        StrAppend1 = " AND PLNCOde ='" & Trim(txtUpgPLNCode) & "' AND HLTCode ='" & Trim(txtHLTCode) & "' AND INSCODe ='" & Trim(txtINSCode1) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_Registration"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
            Set bnf = lstNewBenefitDetails.ListItems.Add(, , rst2!BNFCode)
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
                End If
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
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
    ' Don't allow user to affect workbook
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

Private Sub SearchPrincipalIll()
Dim rst As New ADODB.Recordset
Dim strsql As String
Dim cmdIll As New ADODB.Command
Dim PrmIll As ADODB.Parameter
Dim strAppend As String
Dim CASIllListing As ListItem

    PreIllCAT = ""
    PreIllCATOthers = ""
    UndIllCAT = ""
    UndIllCATOthers = ""
    
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
        PreIllCAT = IIf(Len(rst!CASPreIllCAT), rst!CASPreIllCAT, "")
        PreIllCATOthers = IIf(Len(rst!CASPreIllOthers), rst!CASPreIllOthers, "")
        UndIllCAT = IIf(Len(rst!CASUndIllCAT), rst!CASUndIllCAT, "")
        UndIllCATOthers = IIf(Len(rst!CASUndIllOthers), rst!CASUndIllOthers, "")

        If SSTab1.Tab = 7 And lstPreUndIllness.ListItems.Count = 0 Then
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

Private Sub SearchSuppIll()
Dim rst As New ADODB.Recordset
Dim strsql As String


    strsql = "Select MBMCPreIllCAT,MBMCUndIllCAT FROM MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 2) & "'"
    rst.Open strsql, medixmasterconnPat, adOpenKeyset, adLockOptimistic
   
    If rst.recordCount Then
        With rst
            PreIllCAT = IIf(Len(!MBMCPreIllCAT), !MBMCPreIllCAT, "")
            UndIllCAT = IIf(Len(!MBMCUndIllCAT), !MBMCUndIllCAT, "")
        End With
    End If
    rst.Close
    Set rst = Nothing

End Sub

Private Function SearchMMARate(OPCSCode As String)
Dim strsqlMMA As String
Dim MAA As New ADODB.Recordset
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    strsqlMMA = "Select SurgeonsFees, AnaesthetistsFees  from VIEW_MMA_Schedule where  OPCSCode = '" & Trim(OPCSCode) & "'"
    MAA.Open strsqlMMA, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        Do While Not MAA.EOF
            MMASurgFees = IIf(Len(MAA!SurgeonsFees), MAA!SurgeonsFees, 0)
            MMAAnaeFees = IIf(Len(MAA!AnaesthetistsFees), MAA!AnaesthetistsFees, 0)
        MAA.MoveNext
        Loop
    MAA.Close
    Set MAA = Nothing
  '  medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End Function

Private Sub GetAgeCat(AgeValue As String)
    AgeLenght = Len(AgeValue)
    AgeCat = Mid(AgeValue, AgeLenght, 1)
    AgeOnly = Mid(AgeValue, 1, InStr(1, AgeValue, AgeCat) - 1)
End Sub

Private Sub CheckLimit()
    Dim strAppend, StrSqlAnnLmt As String
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
    
    LifetimeAnnBal = 0
    tmpTotalBalLmt = 0
    

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
                    
                    'End If
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
                    
                    StrSqlAnnLmt = "Select DisabilityLmt, SuppLimitStatus,AnnLimit,SuppLimit,LifeTimeLimit, PLNCode from AnnualLimit where PLNCode = '" & Trim(txtPLNCode) & "' And INSCode = '" & Trim(tmpINSType) & "' And AnnualEffDate <= '" & Format(CAS!CASDateAdmitted, "mm/dd/yyyy") & "' order by AnnualEffDate  desc"
                    AnnLmt.Open StrSqlAnnLmt, medixmasterconn, adOpenKeyset, adLockOptimistic
                    
                    tmpMBMAnnLmt = IIf(IsNumeric(AnnLmt!AnnLimit), AnnLmt!AnnLimit, 0)
                    tmpMBMSuppAnnLmt = IIf(IsNumeric(AnnLmt!SuppLimit), AnnLmt!SuppLimit, 0)
                    tmpMBMLTAnnLmt = IIf(IsNumeric(AnnLmt!LifeTimeLimit), AnnLmt!LifeTimeLimit, 0)
                                            
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
                            
                            'ElseIf Topup = True Then
                            '    If CASTopup = False Then
                            '        tmpChkAnnBalLmt = Format(CDbl(tmpMBMSuppAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                            '    Else
                            '        If IsNumeric(TopNewBalLmt) = False Then
                            '            TopNewBalLmt = 0
                            '        End If
                            '            OldBalLmt = Format(CDbl(tmpMBMSuppAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                            '        If CDate(AdmissionDate) > CDate(TopEffDate) And CDate(AdmissionDate) <= CDate(PolExpDate) Then
                             '               tmpChkAnnBalLmt = Format(CDbl(OldBalLmt) + CDbl(TopNewBalLmt), "#,#.00")
                            '        ElseIf CDate(AdmissionDate) > CDate(PolExpDate) Then
                            '                tmpChkAnnBalLmt = Format(TopNewBalLmt, "#,#.00")
                            '        ElseIf CDate(AdmissionDate) < CDate(TopEffDate) Then
                            '                tmpChkAnnBalLmt = Format(OldBalLmt, "#,#.00")
                            '        End If
                            '    End If
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
                            'ElseIf Topup = True Then
                            '    If CASTopup = False Then
                            '            tmpChkAnnBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                            '    Else
                            '        If IsNumeric(TopNewBalLmt) = False Then
                            '            TopNewBalLmt = 0
                            '        End If
                            '        OldBalLmt = Format(CDbl(tmpMBMSuppAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                            '        If CDate(AdmissionDate) > CDate(TopEffDate) And CDate(AdmissionDate) <= CDate(PolExpDate) Then
                            '                tmpChkAnnBalLmt = Format(CDbl(OldBalLmt) + CDbl(TopNewBalLmt), "#,#.00")
                            '        ElseIf CDate(AdmissionDate) > CDate(PolExpDate) Then
                            '                tmpChkAnnBalLmt = Format(TopNewBalLmt, "#,#.00")
                            '        ElseIf CDate(AdmissionDate) < CDate(TopEffDate) Then
                            '                tmpChkAnnBalLmt = Format(OldBalLmt, "#,#.00")
                            '        End If
                            '    End If
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
                                'If Len(CAS!MBMAvailableLimit) Then
                                    tmpChkAnnBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                                'End If
                            Else
                            '    If IsNumeric(TopNewBalLmt) = False Then
                            '        TopNewBalLmt = 0
                            '    End If
                            '    OldBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                            '    If CDate(AdmissionDate) > CDate(TopEffDate) And CDate(AdmissionDate) <= CDate(PolExpDate) Then
                            '            tmpChkAnnBalLmt = Format(CDbl(OldBalLmt) + CDbl(TopNewBalLmt), "#,#.00")
                            '    ElseIf CDate(AdmissionDate) > CDate(PolExpDate) Then
                            '            tmpChkAnnBalLmt = Format(TopNewBalLmt, "#,#.00")
                            '    ElseIf CDate(AdmissionDate) < CDate(TopEffDate) Then
                            '            tmpChkAnnBalLmt = Format(CDbl(tmpMBMAnnLmt) - CDbl(tmpTotalBordx), "#,#.00")
                            '    End If
                            End If
                        End If
                    End If
                    
                    OldBalLmt = tmpChkAnnBalLmt
                    AnnLmt.Close
                    Set AnnLmt = Nothing
                    
                    
                If IsNumeric(tmpChkAnnBalLmt) = False Then
                    tmpChkAnnBalLmt = 0
                End If
                
                If CDbl(txtCASLGAmt) > CDbl(tmpChkAnnBalLmt) Then
                    LGStatus = False
                Else
                    LGStatus = True
                End If
            End If
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

Sub addTheFiles()
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


Private Sub ShowExclusion()
    Dim rst As New ADODB.Recordset
    Dim strsql As String
    
        strsql = " SELECT * from View_Get_MemberExclusion " & _
                " where PAYCode = '" & Trim(txtPAYCode) & "' " & _
                " And MBMDOB = '" & Format(txtDOB, "yyyy/mm/dd") & "' " & _
                " And INSCode = '" & Trim(txtINSCode1) & "'" & _
                " And MBMName = '" & Trim(Replace(txtMBMName, "'", "''")) & "'"
        rst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            Do While Not rst.EOF
                If Trim(txtPrincipalExc) = "" Then
                    txtPrincipalExc = IIf(Len(Trim(rst!MBMExclusion)), Trim(rst!MBMExclusion), "")
                Else
                    txtPrincipalExc = txtPrincipalExc & vbNewLine & IIf(Len(Trim(rst!MBMExclusion)), Trim(rst!MBMExclusion), "")
                End If
            rst.MoveNext
            Loop
        rst.Close
        Set rst = Nothing
    'End If
End Sub





