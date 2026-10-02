VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCLMCrpd 
   Caption         =   "Claim Correspondence"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   60
      Top             =   1320
   End
   Begin TabDlg.SSTab SSTabCLMCrpd 
      Height          =   7935
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   13996
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Payor"
      TabPicture(0)   =   "frmCLMCrpd.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMembershipNo(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPrincipleName(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDateOfAdmission(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPolicyNo(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPatientName(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblMembershipNo(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "mskDateOfDischargePayor"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "mskDateOfAdmissionPayor"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtMembershipNoPayor"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtPrincipleNamePayor"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtPolicyNoPayor"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPatientNamePayor"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "btnSavePayor"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "btnResetPayor"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "btnPrintPayor"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdExitPayor"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Member"
      TabPicture(1)   =   "frmCLMCrpd.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblMembershipNo(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblPatientName(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblPolicyNo(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblDateOfAdmission(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblPrincipleName(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblMembershipNo(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "mskDateOfDischargeMember"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "mskDateOfAdmissionMember"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtPatientNameMember"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtPolicyNoMember"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtPrincipleNameMember"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtMembershipNoMember"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lstReminder"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "btnAddMember"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "btnEditMember"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "btnSaveMember"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "btnPrintMember"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cmdExitMember"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Hospital"
      TabPicture(2)   =   "frmCLMCrpd.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblMembershipNo(4)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblPrincipleName(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblDateOfAdmission(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblPolicyNo(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblPatientName(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblMembershipNo(5)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "mskDateOfDischargeHospital"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "mskDateOfAdmissionHospital"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lstMedDocReq"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "btnPrintHospital"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "btnSaveHospital"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "btnEditHospital"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "btnAddHospital"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtMembershipNoHospital"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtPrincipleNameHospital"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtPolicyNoHospital"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtPatientNameHospital"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Frame3"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "cmdExitHos"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).ControlCount=   19
      Begin VB.CommandButton cmdExitPayor 
         Appearance      =   0  'Flat
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10800
         TabIndex        =   101
         Top             =   4560
         Width           =   690
      End
      Begin VB.CommandButton cmdExitMember 
         Appearance      =   0  'Flat
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -64020
         TabIndex        =   100
         Top             =   6120
         Width           =   570
      End
      Begin VB.CommandButton cmdExitHos 
         Appearance      =   0  'Flat
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -64080
         TabIndex        =   99
         Top             =   5760
         Width           =   570
      End
      Begin VB.CommandButton btnPrintPayor 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10080
         TabIndex        =   24
         Top             =   4560
         Width           =   690
      End
      Begin VB.CommandButton btnResetPayor 
         Appearance      =   0  'Flat
         Caption         =   "&Reset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9360
         TabIndex        =   23
         Top             =   4560
         Width           =   690
      End
      Begin VB.CommandButton btnSavePayor 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8640
         TabIndex        =   22
         Top             =   4560
         Width           =   690
      End
      Begin VB.Frame Frame3 
         Caption         =   "Actions"
         Height          =   2295
         Left            =   -74880
         TabIndex        =   76
         Top             =   1800
         Width           =   11415
         Begin VB.CheckBox chkOtherHos 
            Caption         =   "Other"
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
            Left            =   840
            TabIndex        =   87
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox chkBreakdown 
            Caption         =   "Breakdown"
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
            Left            =   7200
            TabIndex        =   86
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox chkXrayCtMRIReport 
            Caption         =   "X-ray/Ct/MRI Report"
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
            Left            =   7200
            TabIndex        =   85
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkHisthophatologyReport 
            Caption         =   "Histhophatology Report"
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
            Left            =   7200
            TabIndex        =   84
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox chkItemisedBreakdown 
            Caption         =   "Itemised Breakdown"
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
            Left            =   4020
            TabIndex        =   83
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox chkLaboratoryReports 
            Caption         =   "Laboratory Reports"
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
            Left            =   4020
            TabIndex        =   82
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkMedicalReimbPart2 
            Caption         =   "Medical/Reimb. Part 2"
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
            Left            =   4020
            TabIndex        =   81
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox chkIVUReport 
            Caption         =   "IVU Report"
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
            Left            =   840
            TabIndex        =   80
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox chkReferralLetterHos 
            Caption         =   "Referral Letter"
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
            Left            =   840
            TabIndex        =   79
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkPresentingSymptoms 
            Caption         =   "Presenting Symptoms"
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
            Left            =   840
            TabIndex        =   78
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtOtherHospital 
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
            TabIndex        =   77
            Top             =   1485
            Width           =   6975
         End
         Begin MSMask.MaskEdBox mskReminderDateHospital 
            Height          =   315
            Left            =   2520
            TabIndex        =   88
            Top             =   1800
            Width           =   2175
            _ExtentX        =   3836
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
            Caption         =   "Reminder Date"
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
            Left            =   840
            TabIndex        =   89
            Top             =   1830
            Width           =   1095
         End
      End
      Begin VB.TextBox txtPatientNameHospital 
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
         Left            =   -67560
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtPolicyNoHospital 
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
         Left            =   -67560
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtPrincipleNameHospital 
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
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtMembershipNoHospital 
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
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton btnAddHospital 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66480
         TabIndex        =   70
         Top             =   5760
         Width           =   570
      End
      Begin VB.CommandButton btnEditHospital 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -65880
         TabIndex        =   69
         Top             =   5760
         Width           =   570
      End
      Begin VB.CommandButton btnSaveHospital 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -65280
         TabIndex        =   68
         Top             =   5760
         Width           =   570
      End
      Begin VB.CommandButton btnPrintHospital 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -64680
         TabIndex        =   67
         Top             =   5760
         Width           =   570
      End
      Begin VB.CommandButton btnPrintMember 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -64620
         TabIndex        =   50
         Top             =   6120
         Width           =   570
      End
      Begin VB.CommandButton btnSaveMember 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -65220
         TabIndex        =   49
         Top             =   6120
         Width           =   570
      End
      Begin VB.CommandButton btnEditMember 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -65820
         TabIndex        =   48
         Top             =   6120
         Width           =   570
      End
      Begin VB.CommandButton btnAddMember 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66420
         TabIndex        =   47
         Top             =   6120
         Width           =   570
      End
      Begin MSComctlLib.ListView lstReminder 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   46
         Top             =   4440
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "Actions"
         Height          =   2535
         Left            =   120
         TabIndex        =   66
         Top             =   1800
         Width           =   11415
         Begin MSMask.MaskEdBox mskWorksheet 
            Height          =   315
            Left            =   2520
            TabIndex        =   98
            Top             =   1080
            Width           =   2175
            _ExtentX        =   3836
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
         Begin VB.TextBox txtOther 
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
            MaxLength       =   200
            TabIndex        =   19
            Top             =   1485
            Width           =   6975
         End
         Begin MSMask.MaskEdBox mskAppealLetter 
            Height          =   315
            Left            =   2520
            TabIndex        =   12
            Top             =   390
            Width           =   2175
            _ExtentX        =   3836
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
         Begin VB.CheckBox chkOther 
            Caption         =   "Other"
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
            Left            =   840
            TabIndex        =   18
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CheckBox chkRelevantDocumentsNote 
            Caption         =   "Relevant Documents/Note"
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
            Left            =   5700
            TabIndex        =   17
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox chkOfficialBillsAndReceipt 
            Caption         =   "Official Bills and Receipt"
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
            Left            =   5700
            TabIndex        =   16
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkReports 
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
            Height          =   375
            Left            =   5700
            TabIndex        =   15
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox chkWorksheet 
            Caption         =   "Worksheet"
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
            Left            =   840
            TabIndex        =   14
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox chkRepudiationLetter 
            Caption         =   "Repudiation Letter"
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
            Left            =   840
            TabIndex        =   13
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkAppealLetter 
            Caption         =   "Appeal Letter"
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
            Left            =   840
            TabIndex        =   11
            Top             =   360
            Width           =   2535
         End
         Begin MSMask.MaskEdBox mskReminderDatePayor 
            Height          =   315
            Left            =   2520
            TabIndex        =   21
            Top             =   1920
            Width           =   2175
            _ExtentX        =   3836
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
         Begin VB.Label lblReminderDate 
            Caption         =   "Reminder Date"
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
            Left            =   840
            TabIndex        =   20
            Top             =   1950
            Width           =   1095
         End
      End
      Begin VB.TextBox txtMembershipNoMember 
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
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtPrincipleNameMember 
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
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtPolicyNoMember 
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
         Left            =   -67560
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtPatientNameMember 
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
         Left            =   -67560
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Actions"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   58
         Top             =   1800
         Width           =   11415
         Begin VB.TextBox txtOther2 
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
            TabIndex        =   43
            Top             =   1822
            Width           =   6975
         End
         Begin VB.TextBox txtOther1 
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
            TabIndex        =   41
            Top             =   1485
            Width           =   6975
         End
         Begin VB.CheckBox chkMemberReimbursementForm 
            Caption         =   "Member Reimbursement Form"
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
            Left            =   840
            TabIndex        =   31
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox chkUndertakingLetter 
            Caption         =   "Undertaking Letter"
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
            Left            =   840
            TabIndex        =   32
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkOfficalReceipt 
            Caption         =   "Official Receipt"
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
            Left            =   840
            TabIndex        =   33
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox chkOriginalInvoice 
            Caption         =   "Original Invoice"
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
            Left            =   4020
            TabIndex        =   34
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox chkReferralLetter 
            Caption         =   "Referral Letter"
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
            Left            =   4020
            TabIndex        =   35
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkStudentID 
            Caption         =   "Student ID"
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
            Left            =   4020
            TabIndex        =   36
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox chkCompleteMedicalForm 
            Caption         =   "Complete Medical Form"
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
            Left            =   7200
            TabIndex        =   37
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox chkInternationalPassport 
            Caption         =   "International Passport"
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
            Left            =   7200
            TabIndex        =   38
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkPoliceReport 
            Caption         =   "Police Report"
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
            Left            =   7200
            TabIndex        =   39
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CheckBox chkOther1 
            Caption         =   "Other 1"
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
            Left            =   840
            TabIndex        =   59
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox chkOther2 
            Caption         =   "Other 2"
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
            Left            =   840
            TabIndex        =   42
            Top             =   1800
            Width           =   1095
         End
         Begin MSMask.MaskEdBox mskReminderDateMember 
            Height          =   315
            Left            =   2520
            TabIndex        =   45
            Top             =   2160
            Width           =   2175
            _ExtentX        =   3836
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
            Caption         =   "Reminder Date"
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
            Left            =   840
            TabIndex        =   44
            Top             =   2190
            Width           =   1095
         End
      End
      Begin VB.TextBox txtPatientNamePayor 
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
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1065
         Width           =   2175
      End
      Begin VB.TextBox txtPolicyNoPayor 
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
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   705
         Width           =   2175
      End
      Begin VB.TextBox txtPrincipleNamePayor 
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
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1065
         Width           =   2175
      End
      Begin VB.TextBox txtMembershipNoPayor 
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
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   705
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mskDateOfAdmissionPayor 
         Height          =   315
         Left            =   2640
         TabIndex        =   9
         Top             =   1410
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox mskDateOfDischargePayor 
         Height          =   315
         Left            =   7440
         TabIndex        =   10
         Top             =   1410
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox mskDateOfAdmissionMember 
         Height          =   315
         Left            =   -72360
         TabIndex        =   29
         Top             =   1425
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox mskDateOfDischargeMember 
         Height          =   315
         Left            =   -67560
         TabIndex        =   30
         Top             =   1425
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
      Begin MSComctlLib.ListView lstMedDocReq 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   71
         Top             =   4200
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSMask.MaskEdBox mskDateOfAdmissionHospital 
         Height          =   315
         Left            =   -72360
         TabIndex        =   90
         Top             =   1425
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox mskDateOfDischargeHospital 
         Height          =   315
         Left            =   -67560
         TabIndex        =   91
         Top             =   1425
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
      Begin VB.Label lblMembershipNo 
         Caption         =   "Date of Discharge"
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
         Left            =   -69240
         TabIndex        =   97
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label lblPatientName 
         Caption         =   "Patient Name"
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
         Left            =   -69240
         TabIndex        =   96
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label lblPolicyNo 
         Caption         =   "Policy No."
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
         Left            =   -69240
         TabIndex        =   95
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label lblDateOfAdmission 
         Caption         =   "Date of Admission"
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
         Left            =   -74040
         TabIndex        =   94
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label lblPrincipleName 
         Caption         =   "Principle Name"
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
         Left            =   -74040
         TabIndex        =   93
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label lblMembershipNo 
         Caption         =   "Membership No."
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
         Left            =   -74040
         TabIndex        =   92
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label lblMembershipNo 
         Caption         =   "Membership No."
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
         Left            =   -74040
         TabIndex        =   65
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label lblPrincipleName 
         Caption         =   "Principle Name"
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
         Left            =   -74040
         TabIndex        =   64
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label lblDateOfAdmission 
         Caption         =   "Date of Admission"
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
         Left            =   -74040
         TabIndex        =   63
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label lblPolicyNo 
         Caption         =   "Policy No."
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
         Left            =   -69240
         TabIndex        =   62
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label lblPatientName 
         Caption         =   "Patient Name"
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
         Left            =   -69240
         TabIndex        =   61
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label lblMembershipNo 
         Caption         =   "Date of Discharge"
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
         Left            =   -69240
         TabIndex        =   60
         Top             =   1455
         Width           =   1335
      End
      Begin VB.Label lblMembershipNo 
         Caption         =   "Date of Discharge"
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
         Left            =   5760
         TabIndex        =   57
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblPatientName 
         Caption         =   "Patient Name"
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
         Left            =   5760
         TabIndex        =   56
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblPolicyNo 
         Caption         =   "Policy No."
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
         Left            =   5760
         TabIndex        =   55
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblDateOfAdmission 
         Caption         =   "Date of Admission"
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
         Left            =   960
         TabIndex        =   54
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblPrincipleName 
         Caption         =   "Principle Name"
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
         Left            =   960
         TabIndex        =   53
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblMembershipNo 
         Caption         =   "Membership No."
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
         Left            =   960
         TabIndex        =   52
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame CASNumberFrame 
      Height          =   495
      Left            =   120
      TabIndex        =   40
      Top             =   360
      Width           =   11535
      Begin VB.ComboBox cboCasNumber 
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
         ItemData        =   "frmCLMCrpd.frx":0054
         Left            =   2760
         List            =   "frmCLMCrpd.frx":0056
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Default         =   -1  'True
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
         Left            =   3480
         TabIndex        =   3
         Top             =   135
         Width           =   570
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
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   1
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label65 
         Caption         =   "Cas Number"
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
         Left            =   120
         TabIndex        =   51
         Top             =   165
         Width           =   1095
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claim Correspondence"
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
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmCLMCrpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const APPEAL_LETTER = "Copy of Appeal Letter from Insured / Agent dated "
Private Const REPUDIATION_LETTER = "Copy of Repudiation Letter"
Private Const WORKSHEET = "Copy of Worksheet dated "
Private Const REPORTS = "Copy of Reports (X-ray / MRI / CT Scan / HPE / Lab Investigation Reports)"
Private Const OFFICIAL_BILLS_AND_RECEIPT = "Copy of the Official Bills and Receipt"
Private Const RELEVANT_DOCUMENTS_NOTE = "Copy of the relevant claims documents & case management note"
Private Const Others = "Others: "
Private Const PRESENTING_SYMPTOMS = "Presenting Symptoms: "
Private Const PS_DURATION = "Duration: "
Private Const MEDICAL_REIMB_PART2 = "Please complete Medical/Reimbursement Form Part 2 - Treatement Detials"
Private Const HISTHOPHATOLOGY_REPORT = "Copy of Histhophatology Report"
Private Const REFERRAL_LETTER = "Copy of Referral Letter"
Private Const LAB_REPORT = "Copy of all Laboratory Investigation Report"
Private Const X_RAY_REPORT = "Copy of X-ray Report / Ct Scan / MRI report"
Private Const IVU_REPORT = "Copy of IVU Report"
Private Const HOS_OTHER = "Other: "
Private Const ITEMISED_BREAKDOWN = "Itemised Breakdown"
Private Const BREAKDOWN = "Breakdown"
Private Const CRPD_Tab1 = 1
Private Const CRPD_Tab2 = 2
Private Const CRPD_Tab3 = 3

Private blnHasSearched As Boolean
Private strCasNumber As String
Private strMembershipNo As String
Private strPrincipleName As String
Private strMemberAddress As String
Private strPrincipleIC As String
Private strPatientIC As String
Private strPolicyNo As String
Private strPatientName As String
Private strHosCode As String
Private strHosName As String
Private strHosContact As String
Private strHosFaxNo As String
Private strHosAdd1 As String
Private strHosAdd2 As String
Private strHosAdd3 As String
Private dteDOA As Date
Private dteDOD As Date
Private strSelectedVersion As String
Private blnCanSaveMember As Boolean
Private blnCanAddMember As Boolean
Private blnCanEditMember As Boolean
Private blnCanAddHospital As Boolean
Private blnCanSaveHospital As Boolean
Private blnCanEditHospital As Boolean
Private blnHasPayorActionData As Boolean
Private strSelectedRemNo As String
Private strSelectedReqNo As String
Private strAttendDoctor As String
Private strReason As String
Private strPayCode As String
Private strClientCompany As String
Private strClientCompanyAddress As String

Private Sub btnAddMember_Click()
    Dim intRemCount As Integer
    intRemCount = lstReminder.listitems.Count
    
    If intRemCount > 2 Then
        MsgBox "Only three reminders are allowed"
        Exit Sub
    End If

    SetAddEditSavePrintButtonsForMember False, False, True, False
    blnCanSaveMember = True
    blnCanAddMember = True
    SetActionsForMember True
    ResetActionsForMember
    CheckIfAddEditSavePrintCanSetForMember
End Sub

Private Sub btnAddHospital_Click()
    Dim intRemCount As Integer
    intRemCount = lstMedDocReq.listitems.Count
    
    If intRemCount > 2 Then
        MsgBox "Only three reminders are allowed"
        Exit Sub
    End If

    SetAddEditSavePrintButtonsForHospital False, False, True, False
    blnCanSaveHospital = True
    blnCanAddHospital = True
    SetActionsForHospital True
    ResetActionsForHospital
    CheckIfAddEditSavePrintCanSetForHospital
End Sub

Private Sub btnEditHospital_Click()
    SetAddEditSavePrintButtonsForHospital False, False, True, False
    blnCanSaveHospital = True
    blnCanEditHospital = True
    SetActionsForHospital True
End Sub

Private Sub btnEditMember_Click()
    SetAddEditSavePrintButtonsForMember False, False, True, False
    blnCanSaveMember = True
    blnCanEditMember = True
    SetActionsForMember True
End Sub

Private Sub btnPrintHospital_Click()
    btnPrintHospital.Enabled = False
    GenerateWordDocHospital
End Sub

Private Sub btnPrintMember_Click()
    btnPrintMember.Enabled = False
    GenerateWordDocMember
End Sub

Private Sub btnPrintPayor_Click()
    GenerateWordDocPayor
End Sub

Private Sub btnSaveHospital_Click()
    Dim intReqCount As Integer
    Dim dbCommand As ADODB.Command
    Dim prmReqNo As ADODB.Parameter
    Dim prmCaseNo As ADODB.Parameter
    Dim prmReminderDateHospital As ADODB.Parameter
    Dim prmVersionNo As ADODB.Parameter
    Dim prmPresentingSymptoms As ADODB.Parameter
    Dim prmMedicalReimbPart2 As ADODB.Parameter
    Dim prmHisthophatologyReport As ADODB.Parameter
    Dim prmLaboratoryReports As ADODB.Parameter
    Dim prmReferralLetterHos As ADODB.Parameter
    Dim prmXrayCtMRIReport As ADODB.Parameter
    Dim prmIVUReport As ADODB.Parameter
    Dim prmItemisedBreakdown As ADODB.Parameter
    Dim prmBreakdown As ADODB.Parameter
    Dim prmOtherHos As ADODB.Parameter
    Dim prmOtherHospital As ADODB.Parameter
    Dim strReqNo As String

    SetAddEditSavePrintButtonsForHospital True, False, False, False
    blnCanSaveHospital = False
    SetActionsForHospital False

    If blnCanAddHospital Then
        intReqCount = lstMedDocReq.listitems.Count
        
        If intReqCount < 3 Then
            strReqNo = txtCaseNo & "-" & CStr(intReqCount + 1) & "R"
            'MsgBox strRemNo
            
            Screen.MousePointer = vbHourglass
            'open connection to HISDB
            On Error GoTo ErrorCloseConnectionForAdd
            medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Set dbCommand = New ADODB.Command
            
            Set dbCommand.ActiveConnection = medixmasterconn
            dbCommand.CommandType = adCmdStoredProc
            dbCommand.CommandText = "CC_SaveCLMCrpdHospital"
            
            Set prmReqNo = dbCommand.CreateParameter("ReminderNo", adVarChar, adParamInput, 18, UCase(strReqNo))
            Set prmCaseNo = dbCommand.CreateParameter("CaseNo", adVarChar, adParamInput, 15, UCase(txtCaseNo))
            Set prmVersionNo = dbCommand.CreateParameter("VersionNo", adInteger, adParamInput, 4, cboCasNumber.Text)
            Set prmReminderDateHospital = dbCommand.CreateParameter("ReminderDateHospital", adDate, adParamInput, 8, Format(mskReminderDateHospital, "dd/MM/yyyy"))
            Set prmPresentingSymptoms = dbCommand.CreateParameter("PresentingSymptoms", adTinyInt, adParamInput, 1, chkPresentingSymptoms)
            Set prmMedicalReimbPart2 = dbCommand.CreateParameter("MedicalReimbPart2", adTinyInt, adParamInput, 1, chkMedicalReimbPart2)
            Set prmHisthophatologyReport = dbCommand.CreateParameter("HisthophatologyReport", adTinyInt, adParamInput, 1, chkHisthophatologyReport)
            Set prmLaboratoryReports = dbCommand.CreateParameter("LaboratoryReports", adTinyInt, adParamInput, 1, chkLaboratoryReports)
            Set prmReferralLetterHos = dbCommand.CreateParameter("ReferralLetterHos", adTinyInt, adParamInput, 1, chkReferralLetterHos)
            Set prmXrayCtMRIReport = dbCommand.CreateParameter("XrayCtMRIReport", adTinyInt, adParamInput, 1, chkXrayCtMRIReport)
            Set prmIVUReport = dbCommand.CreateParameter("IVUReport", adTinyInt, adParamInput, 1, chkIVUReport)
            Set prmItemisedBreakdown = dbCommand.CreateParameter("ItemisedBreakdown", adTinyInt, adParamInput, 1, chkItemisedBreakdown)
            Set prmBreakdown = dbCommand.CreateParameter("Breakdown", adTinyInt, adParamInput, 1, chkBreakdown)
            Set prmOtherHos = dbCommand.CreateParameter("OtherHos", adTinyInt, adParamInput, 1, chkOtherHos)
            Set prmOtherHospital = dbCommand.CreateParameter("OtherHospital", adVarChar, adParamInput, 200, txtOtherHospital)
            
            dbCommand.Parameters.Append prmReqNo
            dbCommand.Parameters.Append prmCaseNo
            dbCommand.Parameters.Append prmVersionNo
            dbCommand.Parameters.Append prmReminderDateHospital
            dbCommand.Parameters.Append prmPresentingSymptoms
            dbCommand.Parameters.Append prmMedicalReimbPart2
            dbCommand.Parameters.Append prmHisthophatologyReport
            dbCommand.Parameters.Append prmLaboratoryReports
            dbCommand.Parameters.Append prmReferralLetterHos
            dbCommand.Parameters.Append prmXrayCtMRIReport
            dbCommand.Parameters.Append prmIVUReport
            dbCommand.Parameters.Append prmItemisedBreakdown
            dbCommand.Parameters.Append prmBreakdown
            dbCommand.Parameters.Append prmOtherHos
            dbCommand.Parameters.Append prmOtherHospital

            dbCommand.Execute
            Screen.MousePointer = vbDefault
            
            'start to generate word document
            'GenerateWordDocPayor
            If MsgBox("Print Medical Document Request", vbQuestion Or vbYesNo, "KPI Helpdesk") = vbYes Then
                GenerateWordDocHospital
            End If
ErrorCloseConnectionForAdd:
            medixmasterconn.Close

        End If
    ElseIf blnCanEditHospital Then
        'CC_UpdateCLMCrpdMember
        'Debug.Print strSelectedRemNo
        
        Screen.MousePointer = vbHourglass
        'open connection to HISDB
        On Error GoTo ErrorCloseConnectionForEdit
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Set dbCommand = New ADODB.Command
        
        Set dbCommand.ActiveConnection = medixmasterconn
        dbCommand.CommandType = adCmdStoredProc
        dbCommand.CommandText = "CC_UpdateCLMCrpdHospital"
        
        Set prmReqNo = dbCommand.CreateParameter("ReminderNo", adVarChar, adParamInput, 18, UCase(strSelectedReqNo))
        Set prmCaseNo = dbCommand.CreateParameter("CaseNo", adVarChar, adParamInput, 15, UCase(txtCaseNo))
        Set prmVersionNo = dbCommand.CreateParameter("VersionNo", adInteger, adParamInput, 4, cboCasNumber.Text)
        Set prmReminderDateHospital = dbCommand.CreateParameter("ReminderDateHospital", adDate, adParamInput, 8, Format(mskReminderDateHospital, "dd/MM/yyyy"))
        Set prmPresentingSymptoms = dbCommand.CreateParameter("PresentingSymptoms", adTinyInt, adParamInput, 1, chkPresentingSymptoms)
        Set prmMedicalReimbPart2 = dbCommand.CreateParameter("MedicalReimbPart2", adTinyInt, adParamInput, 1, chkMedicalReimbPart2)
        Set prmHisthophatologyReport = dbCommand.CreateParameter("HisthophatologyReport", adTinyInt, adParamInput, 1, chkHisthophatologyReport)
        Set prmLaboratoryReports = dbCommand.CreateParameter("LaboratoryReports", adTinyInt, adParamInput, 1, chkLaboratoryReports)
        Set prmReferralLetterHos = dbCommand.CreateParameter("ReferralLetterHos", adTinyInt, adParamInput, 1, chkReferralLetterHos)
        Set prmXrayCtMRIReport = dbCommand.CreateParameter("XrayCtMRIReport", adTinyInt, adParamInput, 1, chkXrayCtMRIReport)
        Set prmIVUReport = dbCommand.CreateParameter("IVUReport", adTinyInt, adParamInput, 1, chkIVUReport)
        Set prmItemisedBreakdown = dbCommand.CreateParameter("ItemisedBreakdown", adTinyInt, adParamInput, 1, chkItemisedBreakdown)
        Set prmBreakdown = dbCommand.CreateParameter("Breakdown", adTinyInt, adParamInput, 1, chkBreakdown)
        Set prmOtherHos = dbCommand.CreateParameter("OtherHos", adTinyInt, adParamInput, 1, chkOtherHos)
        Set prmOtherHospital = dbCommand.CreateParameter("OtherHospital", adVarChar, adParamInput, 200, IIf(IsEmpty(txtOtherHospital) Or chkOtherHos = 0, "", txtOtherHospital))
        
        dbCommand.Parameters.Append prmReqNo
        dbCommand.Parameters.Append prmCaseNo
        dbCommand.Parameters.Append prmVersionNo
        dbCommand.Parameters.Append prmReminderDateHospital
        dbCommand.Parameters.Append prmPresentingSymptoms
        dbCommand.Parameters.Append prmMedicalReimbPart2
        dbCommand.Parameters.Append prmHisthophatologyReport
        dbCommand.Parameters.Append prmLaboratoryReports
        dbCommand.Parameters.Append prmReferralLetterHos
        dbCommand.Parameters.Append prmXrayCtMRIReport
        dbCommand.Parameters.Append prmIVUReport
        dbCommand.Parameters.Append prmItemisedBreakdown
        dbCommand.Parameters.Append prmBreakdown
        dbCommand.Parameters.Append prmOtherHos
        dbCommand.Parameters.Append prmOtherHospital
        
        dbCommand.Execute
        Screen.MousePointer = vbDefault
        
        'start to generate word document
        'GenerateWordDocPayor
ErrorCloseConnectionForEdit:
        medixmasterconn.Close
        
    End If
    
    blnCanAddHospital = False
    blnCanEditHospital = False
    cmdShow_Click

End Sub

Private Sub btnSaveMember_Click()
    Dim intRemCount As Integer
    Dim dbCommand As ADODB.Command
    Dim prmReminderNo As ADODB.Parameter
    Dim prmCaseNo As ADODB.Parameter
    Dim prmVersionNo As ADODB.Parameter
    Dim prmReminderDate As ADODB.Parameter
    Dim prmMRF As ADODB.Parameter
    Dim prmOrigInv As ADODB.Parameter
    Dim prmMedForm As ADODB.Parameter
    Dim prmUTLetter As ADODB.Parameter
    Dim prmRefLetter As ADODB.Parameter
    Dim prmInterPassport As ADODB.Parameter
    Dim prmOrigReceipt As ADODB.Parameter
    Dim prmStudentID As ADODB.Parameter
    Dim prmPoliceReport As ADODB.Parameter
    Dim prmOther1 As ADODB.Parameter
    Dim prmOther2 As ADODB.Parameter
    Dim prmOther1Text As ADODB.Parameter
    Dim prmOther2Text As ADODB.Parameter
    Dim strRemNo As String
        
    SetAddEditSavePrintButtonsForMember True, False, False, False
    blnCanSaveMember = False
    SetActionsForMember False
    
    If blnCanAddMember Then
        intRemCount = lstReminder.listitems.Count
        
        If intRemCount < 3 Then
            strRemNo = txtCaseNo & "-" & CStr(intRemCount + 1)
            'MsgBox strRemNo
            
            Screen.MousePointer = vbHourglass
            'open connection to HISDB
            On Error GoTo ErrorCloseConnectionForAdd
            medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Set dbCommand = New ADODB.Command
            
            Set dbCommand.ActiveConnection = medixmasterconn
            dbCommand.CommandType = adCmdStoredProc
            dbCommand.CommandText = "CC_SaveCLMCrpdMember"
            
            Set prmReminderNo = dbCommand.CreateParameter("ReminderNo", adVarChar, adParamInput, 18, UCase(strRemNo))
            Set prmCaseNo = dbCommand.CreateParameter("CaseNo", adVarChar, adParamInput, 15, UCase(txtCaseNo))
            Set prmVersionNo = dbCommand.CreateParameter("VersionNo", adInteger, adParamInput, 4, cboCasNumber.Text)
            Set prmReminderDate = dbCommand.CreateParameter("ReminderDate", adDate, adParamInput, 8, Format(mskReminderDateMember, "dd/MM/yyyy"))
            Set prmMRF = dbCommand.CreateParameter("MRF", adTinyInt, adParamInput, 1, chkMemberReimbursementForm)
            Set prmOrigInv = dbCommand.CreateParameter("OrigInv", adTinyInt, adParamInput, 1, chkOriginalInvoice)
            Set prmMedForm = dbCommand.CreateParameter("MedForm", adTinyInt, adParamInput, 1, chkCompleteMedicalForm)
            Set prmUTLetter = dbCommand.CreateParameter("UTLetter", adTinyInt, adParamInput, 1, chkUndertakingLetter)
            Set prmRefLetter = dbCommand.CreateParameter("RefLetter", adTinyInt, adParamInput, 200, chkReferralLetter)
            Set prmInterPassport = dbCommand.CreateParameter("InterPassport", adTinyInt, adParamInput, 1, chkInternationalPassport)
            Set prmOrigReceipt = dbCommand.CreateParameter("OrigReceipt", adTinyInt, adParamInput, 1, chkOfficalReceipt)
            Set prmStudentID = dbCommand.CreateParameter("StudentID", adTinyInt, adParamInput, 1, chkStudentID)
            Set prmPoliceReport = dbCommand.CreateParameter("PoliceReport", adTinyInt, adParamInput, 1, chkPoliceReport)
            Set prmOther1 = dbCommand.CreateParameter("Other1", adTinyInt, adParamInput, 1, chkOther1)
            Set prmOther2 = dbCommand.CreateParameter("Other2", adTinyInt, adParamInput, 1, chkOther2)
            Set prmOther1Text = dbCommand.CreateParameter("Other1Text", adVarChar, adParamInput, 200, txtOther1)
            Set prmOther2Text = dbCommand.CreateParameter("Other2Text", adVarChar, adParamInput, 200, txtOther2)
            
            dbCommand.Parameters.Append prmReminderNo
            dbCommand.Parameters.Append prmCaseNo
            dbCommand.Parameters.Append prmVersionNo
            dbCommand.Parameters.Append prmReminderDate
            dbCommand.Parameters.Append prmMRF
            dbCommand.Parameters.Append prmOrigInv
            dbCommand.Parameters.Append prmMedForm
            dbCommand.Parameters.Append prmUTLetter
            dbCommand.Parameters.Append prmRefLetter
            dbCommand.Parameters.Append prmInterPassport
            dbCommand.Parameters.Append prmOrigReceipt
            dbCommand.Parameters.Append prmStudentID
            dbCommand.Parameters.Append prmPoliceReport
            dbCommand.Parameters.Append prmOther1
            dbCommand.Parameters.Append prmOther2
            dbCommand.Parameters.Append prmOther1Text
            dbCommand.Parameters.Append prmOther2Text
                    
            dbCommand.Execute
            Screen.MousePointer = vbDefault
            
            'start to generate word document
            If MsgBox("Print Document Request?", vbQuestion Or vbYesNo, "KPI Helpdesk") = vbYes Then
                GenerateWordDocMember
            End If
ErrorCloseConnectionForAdd:
            medixmasterconn.Close

        End If
    ElseIf blnCanEditMember Then
        'CC_UpdateCLMCrpdMember
        'Debug.Print strSelectedRemNo
        
        Screen.MousePointer = vbHourglass
        'open connection to HISDB
        On Error GoTo ErrorCloseConnectionForEdit
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Set dbCommand = New ADODB.Command
        
        Set dbCommand.ActiveConnection = medixmasterconn
        dbCommand.CommandType = adCmdStoredProc
        dbCommand.CommandText = "CC_UpdateCLMCrpdMember"
        
        Set prmReminderNo = dbCommand.CreateParameter("ReminderNo", adVarChar, adParamInput, 18, UCase(strSelectedRemNo))
        Set prmCaseNo = dbCommand.CreateParameter("CaseNo", adVarChar, adParamInput, 15, UCase(txtCaseNo))
        Set prmVersionNo = dbCommand.CreateParameter("VersionNo", adInteger, adParamInput, 4, cboCasNumber.Text)
        Set prmReminderDate = dbCommand.CreateParameter("ReminderDate", adDate, adParamInput, 8, Format(mskReminderDateMember, "MM/dd/yyyy"))
        Set prmMRF = dbCommand.CreateParameter("MRF", adTinyInt, adParamInput, 1, chkMemberReimbursementForm)
        Set prmOrigInv = dbCommand.CreateParameter("OrigInv", adTinyInt, adParamInput, 1, chkOriginalInvoice)
        Set prmMedForm = dbCommand.CreateParameter("MedForm", adTinyInt, adParamInput, 1, chkCompleteMedicalForm)
        Set prmUTLetter = dbCommand.CreateParameter("UTLetter", adTinyInt, adParamInput, 1, chkUndertakingLetter)
        Set prmRefLetter = dbCommand.CreateParameter("RefLetter", adTinyInt, adParamInput, 1, chkReferralLetter)
        Set prmInterPassport = dbCommand.CreateParameter("InterPassport", adTinyInt, adParamInput, 1, chkInternationalPassport)
        Set prmOrigReceipt = dbCommand.CreateParameter("OrigReceipt", adTinyInt, adParamInput, 1, chkOfficalReceipt)
        Set prmStudentID = dbCommand.CreateParameter("StudentID", adTinyInt, adParamInput, 1, chkStudentID)
        Set prmPoliceReport = dbCommand.CreateParameter("PoliceReport", adTinyInt, adParamInput, 1, chkPoliceReport)
        Set prmOther1 = dbCommand.CreateParameter("Other1", adTinyInt, adParamInput, 1, chkOther1)
        Set prmOther2 = dbCommand.CreateParameter("Other2", adTinyInt, adParamInput, 1, chkOther2)
        Set prmOther1Text = dbCommand.CreateParameter("Other1Text", adVarChar, adParamInput, 200, IIf(IsEmpty(txtOther1) Or chkOther1 = 0, "", txtOther1))
        Set prmOther2Text = dbCommand.CreateParameter("Other2Text", adVarChar, adParamInput, 200, IIf(IsEmpty(txtOther2) Or chkOther2 = 0, "", txtOther2))
        
        dbCommand.Parameters.Append prmReminderNo
        dbCommand.Parameters.Append prmCaseNo
        dbCommand.Parameters.Append prmVersionNo
        dbCommand.Parameters.Append prmReminderDate
        dbCommand.Parameters.Append prmMRF
        dbCommand.Parameters.Append prmOrigInv
        dbCommand.Parameters.Append prmMedForm
        dbCommand.Parameters.Append prmUTLetter
        dbCommand.Parameters.Append prmRefLetter
        dbCommand.Parameters.Append prmInterPassport
        dbCommand.Parameters.Append prmOrigReceipt
        dbCommand.Parameters.Append prmStudentID
        dbCommand.Parameters.Append prmPoliceReport
        dbCommand.Parameters.Append prmOther1
        dbCommand.Parameters.Append prmOther2
        dbCommand.Parameters.Append prmOther1Text
        dbCommand.Parameters.Append prmOther2Text
        
        dbCommand.Execute
        Screen.MousePointer = vbDefault
        
        'start to generate word document
        'GenerateWordDocPayor
ErrorCloseConnectionForEdit:
        medixmasterconn.Close
        
    End If
    
    blnCanAddMember = False
    blnCanEditMember = False
    cmdShow_Click
End Sub

Private Sub btnSavePayor_Click()
    Dim dbCommand As ADODB.Command
    Dim prmCaseNo As ADODB.Parameter
    Dim prmVersionNo As ADODB.Parameter
    Dim prmAppealLetter As ADODB.Parameter
    Dim prmRepudiationLetter As ADODB.Parameter
    Dim prmWorksheet As ADODB.Parameter
    Dim prmReports As ADODB.Parameter
    Dim prmOfficialBillsAndReceipt As ADODB.Parameter
    Dim prmRelevantDocumentsNote As ADODB.Parameter
    Dim prmOther As ADODB.Parameter
    Dim prmOtherText As ADODB.Parameter
    Dim prmReminderDate As ADODB.Parameter
    Dim prmAppealLetterDate As ADODB.Parameter
    Dim prmWorksheetDate As ADODB.Parameter
    
    If chkOther And txtOther = "" Then
        MsgBox "Please enter description in 'Others'!", vbCritical + vbOKOnly, "Correspondence"
        txtOther.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    'open connection to HISDB
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Set dbCommand = New ADODB.Command

    Set dbCommand.ActiveConnection = medixmasterconn
    dbCommand.CommandType = adCmdStoredProc
    
    If blnHasPayorActionData Then
        dbCommand.CommandText = "CC_SaveCLMCrpdPayor"
    Else
        dbCommand.CommandText = "CC_SaveNewCLMCrpdPayor"
    End If
    
    Set prmCaseNo = dbCommand.CreateParameter("CASNumber", adVarChar, adParamInput, 15, txtCaseNo.Text)
    Set prmVersionNo = dbCommand.CreateParameter("versionNo", adVarChar, adParamInput, 2, cboCasNumber.Text)
    Set prmAppealLetter = dbCommand.CreateParameter("appealLetter", adTinyInt, adParamInput, 1, chkAppealLetter)
    Set prmRepudiationLetter = dbCommand.CreateParameter("repudiationLetter", adTinyInt, adParamInput, 1, chkRepudiationLetter)
    Set prmWorksheet = dbCommand.CreateParameter("worksheet", adTinyInt, adParamInput, 1, chkWorksheet)
    Set prmReports = dbCommand.CreateParameter("reports", adTinyInt, adParamInput, 1, chkReports)
    Set prmOfficialBillsAndReceipt = dbCommand.CreateParameter("officialBillsAndReceipt", adTinyInt, adParamInput, 1, chkOfficialBillsAndReceipt)
    Set prmRelevantDocumentsNote = dbCommand.CreateParameter("relevantDocumentsNote", adTinyInt, adParamInput, 1, chkRelevantDocumentsNote)
    Set prmOther = dbCommand.CreateParameter("other", adTinyInt, adParamInput, 1, chkOther)
    Set prmOtherText = dbCommand.CreateParameter("otherText", adVarChar, adParamInput, 200, txtOther)
    Set prmReminderDate = dbCommand.CreateParameter("reminderDate", adDate, adParamInput, 8, Format(mskReminderDatePayor, "yyyy-mm-dd"))
    Set prmAppealLetterDate = dbCommand.CreateParameter("appealLetterDate", adDate, adParamInput, 8, Format(mskAppealLetter, "yyyy-mm-dd"))
    Set prmWorksheetDate = dbCommand.CreateParameter("worksheetDate", adDate, adParamInput, 8, Format(mskWorksheet, "yyyy-mm-dd"))

    dbCommand.Parameters.Append prmCaseNo
    dbCommand.Parameters.Append prmVersionNo
    dbCommand.Parameters.Append prmAppealLetter
    dbCommand.Parameters.Append prmRepudiationLetter
    dbCommand.Parameters.Append prmWorksheet
    dbCommand.Parameters.Append prmReports
    dbCommand.Parameters.Append prmOfficialBillsAndReceipt
    dbCommand.Parameters.Append prmRelevantDocumentsNote
    dbCommand.Parameters.Append prmOther
    dbCommand.Parameters.Append prmOtherText
    dbCommand.Parameters.Append prmReminderDate
    dbCommand.Parameters.Append prmAppealLetterDate
    dbCommand.Parameters.Append prmWorksheetDate
    dbCommand.Execute
    Screen.MousePointer = vbDefault
    
    'start to generate word document
    If MsgBox("Print Appeal Acknowledgement?", vbQuestion Or vbYesNo, "KPI Helpdesk") = vbYes Then
        GenerateWordDocPayor
    End If
    medixmasterconn.Close
    cmdShow_Click
End Sub

Private Sub btnResetPayor_Click()
    chkAppealLetter = False
    chkRepudiationLetter = False
    chkWorksheet = False
    chkReports = False
    chkOfficialBillsAndReceipt = False
    chkRelevantDocumentsNote = False
    chkOther = False
    txtOther = ""
    mskReminderDatePayor = Format(DateTime.Date, "dd/MM/yyyy")
    mskAppealLetter = Format(DateTime.Date, "dd/MM/yyyy")
End Sub

Private Sub cboCasNumber_Click()
    strSelectedVersion = cboCasNumber.Text
    tmrRefresh.Enabled = True
End Sub

Private Sub chkAppealLetter_Click()
    mskAppealLetter.Enabled = chkAppealLetter
    CheckIfSaveResetPrintCanSetForPayor
End Sub

Private Sub chkBreakdown_Click()
    CheckIfAddEditSavePrintCanSetForHospital
End Sub

Private Sub chkHisthophatologyReport_Click()
    CheckIfAddEditSavePrintCanSetForHospital
End Sub

Private Sub chkItemisedBreakdown_Click()
    CheckIfAddEditSavePrintCanSetForHospital
End Sub

Private Sub chkIVUReport_Click()
    CheckIfAddEditSavePrintCanSetForHospital
End Sub

Private Sub chkLaboratoryReports_Click()
    CheckIfAddEditSavePrintCanSetForHospital
End Sub

Private Sub chkMedicalReimbPart2_Click()
    CheckIfAddEditSavePrintCanSetForHospital
End Sub

Private Sub chkOfficialBillsAndReceipt_Click()
    CheckIfSaveResetPrintCanSetForPayor
End Sub

Private Sub chkOther_Click()
    txtOther.Enabled = chkOther
    CheckIfSaveResetPrintCanSetForPayor
    If txtOther.Enabled Then txtOther.SetFocus Else txtOther.Text = ""
End Sub

Private Sub chkOtherHos_Click()
    CheckIfAddEditSavePrintCanSetForHospital
    If chkOtherHos.Enabled Then
        txtOtherHospital.Enabled = chkOtherHos
        If chkOtherHos Then txtOtherHospital.SetFocus
    Else
        txtOtherHospital.Enabled = False
    End If
End Sub

Private Sub chkPresentingSymptoms_Click()
    CheckIfAddEditSavePrintCanSetForHospital
End Sub

Private Sub chkReferralLetterHos_Click()
    CheckIfAddEditSavePrintCanSetForHospital
End Sub

Private Sub chkRelevantDocumentsNote_Click()
    CheckIfSaveResetPrintCanSetForPayor
End Sub

Private Sub chkReports_Click()
    CheckIfSaveResetPrintCanSetForPayor
End Sub

Private Sub chkRepudiationLetter_Click()
    CheckIfSaveResetPrintCanSetForPayor
End Sub

Private Sub chkWorksheet_Click()
    mskWorksheet.Enabled = chkWorksheet
    CheckIfSaveResetPrintCanSetForPayor
End Sub

Private Sub chkCompleteMedicalForm_Click()
    CheckIfAddEditSavePrintCanSetForMember
End Sub

Private Sub chkInternationalPassport_Click()
    CheckIfAddEditSavePrintCanSetForMember
End Sub

Private Sub chkMemberReimbursementForm_Click()
    CheckIfAddEditSavePrintCanSetForMember
End Sub

Private Sub chkOfficalReceipt_Click()
    CheckIfAddEditSavePrintCanSetForMember
End Sub

Private Sub chkOriginalInvoice_Click()
    CheckIfAddEditSavePrintCanSetForMember
End Sub

Private Sub chkOther1_Click()
    CheckIfAddEditSavePrintCanSetForMember
    If chkOther1.Enabled Then
        txtOther1.Enabled = chkOther1
        If chkOther1 Then txtOther1.SetFocus
    Else
        txtOther1.Enabled = False
    End If
End Sub

Private Sub chkOther2_Click()
    CheckIfAddEditSavePrintCanSetForMember
    If chkOther2.Enabled Then
        txtOther2.Enabled = chkOther2
        If chkOther2 Then txtOther2.SetFocus
    Else
        txtOther2.Enabled = False
    End If
End Sub

Private Sub chkPoliceReport_Click()
    CheckIfAddEditSavePrintCanSetForMember
End Sub

Private Sub chkReferralLetter_Click()
    CheckIfAddEditSavePrintCanSetForMember
End Sub

Private Sub chkStudentID_Click()
    CheckIfAddEditSavePrintCanSetForMember
End Sub

Private Sub chkUndertakingLetter_Click()
    CheckIfAddEditSavePrintCanSetForMember
End Sub

Private Sub chkXrayCtMRIReport_Click()
    CheckIfAddEditSavePrintCanSetForHospital
End Sub

Private Sub cmdExitHos_Click()
    Unload Me
End Sub

Private Sub cmdExitPayor_Click()
    Unload Me
End Sub

Private Sub cmdShow_Click()
    Dim dbCommand As ADODB.Command
    Dim dbRecordset As ADODB.Recordset
    Dim prmCaseNo As ADODB.Parameter
    Dim prmClmVer As ADODB.Parameter
    Dim prmHosCode As ADODB.Parameter

    If txtCaseNo <> "" Then
        Screen.MousePointer = vbHourglass
        'open connection to HISDB
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        Set dbCommand = New ADODB.Command
        Set dbRecordset = New ADODB.Recordset
        dbRecordset.CacheSize = 100
        dbRecordset.CursorLocation = adUseClient
        dbRecordset.CursorType = adOpenForwardOnly

        Set dbCommand.ActiveConnection = medixmasterconn
        dbCommand.CommandType = adCmdStoredProc
        dbCommand.CommandText = "CC_SearchCLMCrpd"
        'prmClmVer (claim version) must be set to -1 to disable claim version as a criteria in
        'the stored procedure so that all the version in a particular case are retrieved
        Set prmCaseNo = dbCommand.CreateParameter("CASNumber", adVarChar, adParamInput, 15, txtCaseNo.Text)
        Set prmClmVer = dbCommand.CreateParameter("ClaimVersion", adInteger, adParamInput, 4, -1)
        dbCommand.Parameters.Append prmCaseNo
        dbCommand.Parameters.Append prmClmVer
    
        Set dbRecordset = dbCommand.Execute
    
        'if the record is found then populate the claim version combo box
        If Not dbRecordset.EOF Then
            blnHasSearched = True
            cboCasNumber.Clear
        
            While Not dbRecordset.EOF
                cboCasNumber.AddItem dbRecordset!claimversion
                dbRecordset.MoveNext
            Wend
            
            If strSelectedVersion = "" Or strCasNumber <> txtCaseNo.Text Then strSelectedVersion = cboCasNumber.List(cboCasNumber.TopIndex)
        
            strCasNumber = txtCaseNo.Text
        
            'remove the claim version parameter then create a new one with value set to value not -1
            dbCommand.Parameters.Delete 1
            Set prmClmVer = dbCommand.CreateParameter("ClaimVersion", adInteger, adParamInput, 4, strSelectedVersion)
    
            dbCommand.Parameters.Append prmClmVer
            Set dbRecordset = dbCommand.Execute
    
            If Not dbRecordset.EOF Then
                With dbRecordset
                    strMembershipNo = Trim(!MBMNumber)
                    strPrincipleName = Trim(!MBMName)
                    strPolicyNo = Trim(!MBMPolicyNo)
                    strPatientName = IIf(!MBMCName <> Null, Trim(!MBMCName), Trim(!MBMName))
                    strPatientIC = IIf(!MBMCName <> Null, Trim(!MBMCICBCPPNo), Trim(!MBMIcBcPp))
                    dteDOA = Trim(!CASDateAdmitted)
                    dteDOD = Trim(!CASDischargeDate)
                    strHosCode = Trim(!CASHOSCode)
                    strPrincipleIC = Trim(!MBMIcBcPp)
                    strPayCode = Trim(!PAYCode)
                    strMemberAddress = _
                        IIf(IsNull(!MBMAddress1) Or !MBMAddress1 = "", "", !MBMAddress1 & vbCrLf) & _
                        IIf(IsNull(!MBMAddress2) Or !MBMAddress2 = "", "", !MBMAddress2 & vbCrLf) & _
                        IIf(IsNull(!MBMAddress3) Or !MBMAddress3 = "", "", !MBMAddress3 & vbCrLf) & _
                        IIf(IsNull(!MBMPostCode) Or !MBMPostCode = "", "", !MBMPostCode)
                        
                End With
                dbRecordset.MoveNext
            End If

            Dim idx As Integer, limit As Integer
            
            limit = dbCommand.Parameters.Count
            
            limit = limit - 1
            For idx = 0 To limit
                dbCommand.Parameters.Delete 0
            Next idx
            
            'search pay company's company name and address
            
            dbCommand.CommandText = "CC_SearchClientCompany"
            dbCommand.Parameters.Append dbCommand.CreateParameter("CASNumber", adVarChar, adParamInput, 2, strPayCode)
            dbRecordset.Close
            Set dbRecordset = dbCommand.Execute
            With dbRecordset
                strClientCompany = !PAYCompanyName
                strClientCompanyAddress = _
                    IIf(IsNull(!PAYAddress1) Or !PAYAddress1 = "", "", !PAYAddress1 & vbCrLf) & _
                    IIf(IsNull(!PAYAddress2) Or !PAYAddress2 = "", "", !PAYAddress2 & vbCrLf) & _
                    IIf(IsNull(!PAYAddress3) Or !PAYAddress3 = "", "", !PAYAddress3 & vbCrLf)
            End With
            
            'then search for the hospital name with HOSCode
            dbCommand.CommandText = "CC_SearchHosName"
            Set prmHosCode = dbCommand.CreateParameter("HOSCode", adVarChar, adParamInput, 5, strHosCode)
            
            limit = dbCommand.Parameters.Count
            
            limit = limit - 1
            For idx = 0 To limit
                dbCommand.Parameters.Delete 0
            Next idx
            
            dbCommand.Parameters.Append prmHosCode
            dbRecordset.Close
            Set dbRecordset = dbCommand.Execute
            
            If Not dbRecordset.EOF Then
                strHosName = IIf(IsNull(dbRecordset!HOSName), "N/A", dbRecordset!HOSName)
                strHosFaxNo = IIf(IsNull(dbRecordset!HOSFaxNo), "N/A", dbRecordset!HOSFaxNo)
                strHosContact = IIf(IsNull(dbRecordset!HOSContact), "N/A", dbRecordset!HOSContact)
                strHosAdd1 = IIf(IsNull(dbRecordset!HOSAdd1), "N/A", dbRecordset!HOSAdd1)
                strHosAdd2 = IIf(IsNull(dbRecordset!HOSAdd2), "", dbRecordset!HOSAdd2)
                strHosAdd3 = IIf(IsNull(dbRecordset!HOSAdd3), "", dbRecordset!HOSAdd3)
                strHosFaxNo = IIf(IsNull(dbRecordset!HOSFaxNo), "N/A", dbRecordset!HOSFaxNo)
            Else
                strHosName = strHosCode
                strHosFaxNo = "N/A"
                strHosAdd1 = "N/A"
                strHosAdd2 = ""
                strHosAdd3 = ""
            End If
            
            'populate the checkboxes in payor tab
            dbCommand.CommandText = "CC_SearchActionsPayor"
            Set prmCaseNo = dbCommand.CreateParameter("CASNumber", adVarChar, adParamInput, 15, txtCaseNo.Text)
            dbCommand.Parameters.Delete 0
            dbCommand.Parameters.Append prmCaseNo
            dbRecordset.Close
            Set dbRecordset = dbCommand.Execute
            
            If Not dbRecordset.EOF Then
                With dbRecordset
                    chkAppealLetter = IIf(IsNull(!CASAppealLetter), False, !CASAppealLetter)
                    chkRepudiationLetter = IIf(IsNull(!CASRepudiationLetter), False, !CASRepudiationLetter)
                    chkWorksheet = IIf(IsNull(!CASWorksheet), False, !CASWorksheet)
                    chkReports = IIf(IsNull(!CASReports), False, !CASReports)
                    chkOfficialBillsAndReceipt = IIf(IsNull(!CASOfficialBillsAndReceipt), False, !CASOfficialBillsAndReceipt)
                    chkRelevantDocumentsNote = IIf(IsNull(!CASRelevantDocumentsNote), False, !CASRelevantDocumentsNote)
                    chkOther = IIf(IsNull(!CASOther), False, !CASOther)
                    txtOther = IIf(IsNull(!CASOtherText), "", !CASOtherText)
                    mskAppealLetter.Text = Format(IIf(IsNull(!CASAppealLetterDate), DateTime.Now, !CASAppealLetterDate), "dd/MM/yyyy")
                    mskWorksheet.Text = Format(IIf(IsNull(!CASWorksheetDate), DateTime.Now, !CASWorksheetDate), "dd/MM/yyyy")
                    mskReminderDatePayor.Text = Format(IIf(IsNull(!CASReminderDate), DateTime.Now, !CASReminderDate), "dd/MM/yyyy")
                End With
                
                dbCommand.CommandText = "CC_SearchActionsPayor2"
                dbRecordset.Close
                Set dbRecordset = dbCommand.Execute

                If Not dbRecordset.EOF Then
                    With dbRecordset
                        strAttendDoctor = IIf(IsNull(!CASAttendDoctor1), "N/A", !CASAttendDoctor1)
                        strReason = IIf(IsNull(!CASAdmissionReason), "N/A", !CASAdmissionReason)
                    End With
                End If
                
                blnHasPayorActionData = True
            Else
                'provides the default value for each of the action controls when data cannot be found
                chkAppealLetter = False
                chkRepudiationLetter = False
                chkWorksheet = False
                chkReports = False
                chkOfficialBillsAndReceipt = False
                chkRelevantDocumentsNote = False
                chkOther = False
                txtOther = ""
                mskAppealLetter.Text = Format(DateTime.Now, "dd/mm/yyyy")
                mskWorksheet.Text = Format(DateTime.Now, "dd/mm/yyyy")
                mskReminderDatePayor.Text = Format(DateTime.Now, "dd/mm/yyyy")
                blnHasPayorActionData = False
            End If
            
            dbCommand.CommandText = "CC_SearchActionsMember"
            Set prmCaseNo = dbCommand.CreateParameter("CaseNo", adVarChar, adParamInput, 15, txtCaseNo.Text)
            Set prmClmVer = dbCommand.CreateParameter("VersionNo", adInteger, adParamInput, 4, strSelectedVersion)
            dbCommand.Parameters.Delete 0
            dbCommand.Parameters.Append prmCaseNo
            dbCommand.Parameters.Append prmClmVer
            dbRecordset.Close
            Set dbRecordset = dbCommand.Execute
            
            lstReminder.listitems.Clear
            lstReminder.FullRowSelect = True
            lstReminder.Gridlines = True
            
            If Not dbRecordset.EOF Then
                With dbRecordset
                    While Not dbRecordset.EOF
                        Dim itemReminder As ListItem
                        Set itemReminder = lstReminder.listitems.Add(, !ReminderNo, !ReminderNo)
                        itemReminder.SubItems(1) = !CaseNo
                        itemReminder.SubItems(2) = Format(IIf(IsNull(!ReminderDate), DateTime.Now, !ReminderDate), "dd/MM/yyyy")
                        itemReminder.SubItems(3) = IIf(IsNull(!MRF) Or !MRF = 0, "N", "Y")
                        itemReminder.SubItems(4) = IIf(IsNull(!OrigInv) Or !OrigInv = 0, "N", "Y")
                        itemReminder.SubItems(5) = IIf(IsNull(!MedForm) Or !MedForm = 0, "N", "Y")
                        itemReminder.SubItems(6) = IIf(IsNull(!UTLetter) Or !UTLetter = 0, "N", "Y")
                        itemReminder.SubItems(7) = IIf(IsNull(!RefLetter) Or !RefLetter = 0, "N", "Y")
                        itemReminder.SubItems(8) = IIf(IsNull(!InterPassport) Or !InterPassport = 0, "N", "Y")
                        itemReminder.SubItems(9) = IIf(IsNull(!OrigReceipt) Or !OrigReceipt = 0, "N", "Y")
                        itemReminder.SubItems(10) = IIf(IsNull(!StudentID) Or !StudentID = 0, "N", "Y")
                        itemReminder.SubItems(11) = IIf(IsNull(!PoliceReport) Or !PoliceReport = 0, "N", "Y")
                        itemReminder.SubItems(12) = IIf(IsNull(!OTHER1) Or !OTHER1 = 0, "N", "Y")
                        itemReminder.SubItems(13) = IIf(IsNull(!OTHER2) Or !OTHER2 = 0, "N", "Y")
                        itemReminder.SubItems(14) = IIf(IsNull(!Other1Text), "", !Other1Text)
                        itemReminder.SubItems(15) = IIf(IsNull(!Other2Text), "", !Other2Text)
                        lstReminder.Refresh
                        .MoveNext
                    Wend
                End With
            Else
                With dbRecordset
                    chkMemberReimbursementForm = False
                    chkUndertakingLetter = False
                    chkOfficalReceipt = False
                    chkOriginalInvoice = False
                    chkReferralLetter = False
                    chkStudentID = False
                    chkCompleteMedicalForm = False
                    chkInternationalPassport = False
                    chkPoliceReport = False
                    chkOther1 = False
                    chkOther2 = False
                    txtOther1 = ""
                    txtOther2 = ""
                    mskReminderDateMember = Format(DateTime.Now, "dd/MM/yyyy")
                    lstReminder.listitems.Clear
                    lstReminder.FullRowSelect = True
                    lstReminder.Gridlines = True
                End With
            End If
            
            dbCommand.CommandText = "CC_SearchActionsHospital"
            Set prmCaseNo = dbCommand.CreateParameter("CaseNo", adVarChar, adParamInput, 15, txtCaseNo.Text)
            Set prmClmVer = dbCommand.CreateParameter("VersionNo", adInteger, adParamInput, 4, strSelectedVersion)
            dbCommand.Parameters.Delete 0
            dbCommand.Parameters.Delete 0
            dbCommand.Parameters.Append prmCaseNo
            dbCommand.Parameters.Append prmClmVer
            dbRecordset.Close
            Set dbRecordset = dbCommand.Execute
            
            lstMedDocReq.listitems.Clear
            lstMedDocReq.FullRowSelect = True
            lstMedDocReq.Gridlines = True
            
            If Not dbRecordset.EOF Then
                With dbRecordset
                    While Not dbRecordset.EOF
                        Dim itemRequest As ListItem
                        Set itemRequest = lstMedDocReq.listitems.Add(, !ReqNo, !ReqNo)
                        itemRequest.SubItems(1) = !CaseNo
                        itemRequest.SubItems(2) = Format(IIf(IsNull(!ReminderDateHospital), DateTime.Now, !ReminderDateHospital), "dd/MM/yyyy")
                        itemRequest.SubItems(3) = IIf(IsNull(!PresentingSymptoms) Or !PresentingSymptoms = 0, "N", "Y")
                        itemRequest.SubItems(4) = IIf(IsNull(!MedicalReimbPart2) Or !MedicalReimbPart2 = 0, "N", "Y")
                        itemRequest.SubItems(5) = IIf(IsNull(!HisthophatologyReport) Or !HisthophatologyReport = 0, "N", "Y")
                        itemRequest.SubItems(6) = IIf(IsNull(!ReferralLetterHos) Or !ReferralLetterHos = 0, "N", "Y")
                        itemRequest.SubItems(7) = IIf(IsNull(!LaboratoryReports) Or !LaboratoryReports = 0, "N", "Y")
                        itemRequest.SubItems(8) = IIf(IsNull(!XrayCtMRIReport) Or !XrayCtMRIReport = 0, "N", "Y")
                        itemRequest.SubItems(9) = IIf(IsNull(!IVUReport) Or !IVUReport = 0, "N", "Y")
                        itemRequest.SubItems(10) = IIf(IsNull(!ItemisedBreakdown) Or !ItemisedBreakdown = 0, "N", "Y")
                        itemRequest.SubItems(11) = IIf(IsNull(!BREAKDOWN) Or !BREAKDOWN = 0, "N", "Y")
                        itemRequest.SubItems(12) = IIf(IsNull(!OtherHos) Or !OtherHos = 0, "N", "Y")
                        itemRequest.SubItems(13) = IIf(IsNull(!OtherHosText) Or !OtherHos = 0, "", !OtherHosText)
                        lstMedDocReq.Refresh
                        .MoveNext
                    Wend
                End With
            Else
                ResetActionsForHospital
            End If

            InitFormOutputFields CRPD_Tab1
            InitFormOutputFields CRPD_Tab2
            InitFormOutputFields CRPD_Tab3
            cboCasNumber.Text = strSelectedVersion
            'SetSaveResetPrintButtonsForPayor True, True, True
            SetAddEditSavePrintButtonsForMember True, False, False, False
            SetAddEditSavePrintButtonsForHospital True, False, False, False
            SetActionsForPayor True
            SetActionsForMember False
            SetActionsForHospital False
        Else
            MsgBox "CAS Number not found!", vbCritical + vbOKOnly, "Correspondence"
        End If

        dbRecordset.Close
        Set dbRecordset = Nothing
    
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        Screen.MousePointer = vbDefault
    Else
        MsgBox "Please enter CAS Number!", vbCritical + vbOKOnly, "Correspondence"
    End If
End Sub

Private Sub cmdExitMember_Click()
    Unload Me
End Sub

Private Sub lstMedDocReq_ItemClick(ByVal Item As MSComctlLib.ListItem)
    chkPresentingSymptoms = IIf(UCase(Item.SubItems(3)) = "N", 0, 1)
    chkMedicalReimbPart2 = IIf(UCase(Item.SubItems(4)) = "N", 0, 1)
    chkHisthophatologyReport = IIf(UCase(Item.SubItems(5)) = "N", 0, 1)
    chkLaboratoryReports = IIf(UCase(Item.SubItems(7)) = "N", 0, 1)
    chkReferralLetterHos = IIf(UCase(Item.SubItems(6)) = "N", 0, 1)
    chkXrayCtMRIReport = IIf(UCase(Item.SubItems(8)) = "N", 0, 1)
    chkIVUReport = IIf(UCase(Item.SubItems(9)) = "N", 0, 1)
    chkItemisedBreakdown = IIf(UCase(Item.SubItems(10)) = "N", 0, 1)
    chkBreakdown = IIf(UCase(Item.SubItems(11)) = "N", 0, 1)
    chkOtherHos = IIf(UCase(Item.SubItems(12)) = "N", 0, 1)
    txtOtherHospital = IIf(IsNull(Item.SubItems(13)), "", Item.SubItems(13))
    mskReminderDateHospital.Text = Item.SubItems(2)
    strSelectedReqNo = lstMedDocReq.SelectedItem
    
    SetAddEditSavePrintButtonsForHospital True, True, False, True
    SetActionsForHospital False
    blnCanAddHospital = False
    blnCanEditHospital = False

End Sub

Private Sub lstReminder_ItemClick(ByVal Item As MSComctlLib.ListItem)
    chkMemberReimbursementForm = IIf(UCase(Item.SubItems(3)) = "N", 0, 1)
    chkUndertakingLetter = IIf(UCase(Item.SubItems(6)) = "N", 0, 1)
    chkOfficalReceipt = IIf(UCase(Item.SubItems(9)) = "N", 0, 1)
    chkOriginalInvoice = IIf(UCase(Item.SubItems(4)) = "N", 0, 1)
    chkReferralLetter = IIf(UCase(Item.SubItems(7)) = "N", 0, 1)
    chkStudentID = IIf(UCase(Item.SubItems(10)) = "N", 0, 1)
    chkCompleteMedicalForm = IIf(UCase(Item.SubItems(5)) = "N", 0, 1)
    chkInternationalPassport = IIf(UCase(Item.SubItems(8)) = "N", 0, 1)
    chkPoliceReport = IIf(UCase(Item.SubItems(11)) = "N", 0, 1)
    chkOther1 = IIf(UCase(Item.SubItems(12)) = "N", 0, 1)
    chkOther2 = IIf(UCase(Item.SubItems(13)) = "N", 0, 1)
    txtOther1 = IIf(IsNull(Item.SubItems(14)), "", Item.SubItems(14))
    txtOther2 = IIf(IsNull(Item.SubItems(15)), "", Item.SubItems(15))
    mskReminderDateMember = Item.SubItems(2)
    strSelectedRemNo = lstReminder.SelectedItem
    
    SetAddEditSavePrintButtonsForMember True, True, False, True
    SetActionsForMember False
    blnCanAddMember = False
    blnCanEditMember = False
   
End Sub

Private Sub SSTabCLMCrpd_Click(PreviousTab As Integer)
    Static blnTabChangeWithCode As Boolean
    
    If blnHasSearched Then
        If Not blnTabChangeWithCode Then
            If (PreviousTab <> SSTabCLMCrpd.Tab) And (blnCanSaveMember Or blnCanSaveHospital) Then
                If MsgBox("Do you want to cancel this operation?", vbYesNo Or vbExclamation, "KPI Helpdesk") = vbYes Then
                    Call MsgBox("Save operation cancelled", vbOKOnly, "KPI Helpdesk")
                    blnCanSaveMember = False
                    blnCanSaveHospital = False
                    SetAddEditSavePrintButtonsForMember True, False, False, False
                    SetActionsForMember False
                    SetAddEditSavePrintButtonsForHospital True, False, False, False
                    SetActionsForHospital False
                Else
                    blnTabChangeWithCode = True
                    SSTabCLMCrpd.Tab = PreviousTab
                    Exit Sub
                End If
            End If
        Else
            blnTabChangeWithCode = False
        End If
        
        Select Case SSTabCLMCrpd.Tab
        Case 0:
            InitFormOutputFields CRPD_Tab1
            'Debug.Print "ST: " & SSTabCLMCrpd.Tab & ", CRPD_Tab: " & CRPD_Tab1 & vbCrLf
        Case 1:
            InitFormOutputFields CRPD_Tab2
            'Debug.Print "ST: " & SSTabCLMCrpd.Tab & ", CRPD_Tab: " & CRPD_Tab2 & vbCrLf
        Case 2:
            InitFormOutputFields CRPD_Tab3
            'Debug.Print "ST: " & SSTabCLMCrpd.Tab & ", CRPD_Tab: " & CRPD_Tab3 & vbCrLf
        End Select
    End If
End Sub

Private Sub SSTabCLMCrpd_DblClick()
    If blnHasSearched Then
        Select Case SSTabCLMCrpd.Tab
        Case 0:
            InitFormOutputFields CRPD_Tab1
            'Debug.Print "ST: " & SSTabCLMCrpd.Tab & ", CRPD_Tab: " & CRPD_Tab1 & vbCrLf
        Case 1:
            InitFormOutputFields CRPD_Tab2
            'Debug.Print "ST: " & SSTabCLMCrpd.Tab & ", CRPD_Tab: " & CRPD_Tab2 & vbCrLf
        Case 2:
            InitFormOutputFields CRPD_Tab3
            'Debug.Print "ST: " & SSTabCLMCrpd.Tab & ", CRPD_Tab: " & CRPD_Tab3 & vbCrLf
        End Select
    End If
End Sub

Private Sub Form_Load()
    blnHasSearched = False
    SetupReminderListView
    SetupMedDocReqListView
    'payor tab panel initial setting
    SetSaveResetPrintButtonsForPayor False, False, False
    SetActionsForPayor False
    mskAppealLetter.Enabled = chkAppealLetter
    mskWorksheet.Enabled = chkWorksheet
    txtOther.Enabled = chkOther
    mskReminderDatePayor.Text = Format(DateTime.Date, "dd/MM/yyyy")
    mskAppealLetter.Text = Format(DateTime.Date, "dd/MM/yyyy")
    mskWorksheet.Text = Format(DateTime.Date, "dd/MM/yyyy")
   
    'member tab panel intial setting
    SetAddEditSavePrintButtonsForMember False, False, False, False
    SetActionsForMember False
    txtOther1.Enabled = chkOther1
    txtOther2.Enabled = chkOther2
    mskReminderDateMember.Text = Format(DateTime.Date, "dd/MM/yyyy")
    'lstReminder.Enabled = False
    
    'hospital tab panel intial setting
    SetAddEditSavePrintButtonsForHospital False, False, False, False
    SetActionsForHospital False
    mskReminderDateHospital.Text = Format(DateTime.Date, "dd/MM/yyyy")
End Sub

Private Sub InitFormOutputFields(ByVal intTabNo As Integer)
    Select Case intTabNo
    Case 1:
        txtMembershipNoPayor.Text = strMembershipNo
        txtPrincipleNamePayor.Text = strPrincipleName
        txtPolicyNoPayor.Text = strPolicyNo
        txtPatientNamePayor.Text = strPatientName
        mskDateOfAdmissionPayor.Text = Format(dteDOA, "dd/MM/yyyy")
        mskDateOfDischargePayor.Text = Format(dteDOD, "dd/MM/yyyy")
    Case 2:
        txtMembershipNoMember.Text = strMembershipNo
        txtPrincipleNameMember.Text = strPrincipleName
        txtPolicyNoMember.Text = strPolicyNo
        txtPatientNameMember.Text = strPatientName
        mskDateOfAdmissionMember.Text = Format(dteDOA, "dd/MM/yyyy")
        mskDateOfDischargeMember.Text = Format(dteDOD, "dd/MM/yyyy")
    Case 3:
        txtMembershipNoHospital.Text = strMembershipNo
        txtPrincipleNameHospital.Text = strPrincipleName
        txtPolicyNoHospital.Text = strPolicyNo
        txtPatientNameHospital.Text = strPatientName
        mskDateOfAdmissionHospital.Text = Format(dteDOA, "dd/MM/yyyy")
        mskDateOfDischargeHospital.Text = Format(dteDOD, "dd/MM/yyyy")
    End Select
End Sub

Private Sub SetActionsForPayor(ByVal blnEnable As Boolean)
    chkAppealLetter.Enabled = blnEnable
    chkRepudiationLetter.Enabled = blnEnable
    chkWorksheet.Enabled = blnEnable
    chkReports.Enabled = blnEnable
    chkOfficialBillsAndReceipt.Enabled = blnEnable
    chkRelevantDocumentsNote.Enabled = blnEnable
    chkOther.Enabled = blnEnable
    mskReminderDatePayor.Enabled = blnEnable
End Sub

Private Sub SetActionsForMember(ByVal blnEnable As Boolean)
    chkMemberReimbursementForm.Enabled = blnEnable
    chkUndertakingLetter.Enabled = blnEnable
    chkOfficalReceipt.Enabled = blnEnable
    chkOriginalInvoice.Enabled = blnEnable
    chkReferralLetter.Enabled = blnEnable
    chkStudentID.Enabled = blnEnable
    chkCompleteMedicalForm.Enabled = blnEnable
    chkInternationalPassport.Enabled = blnEnable
    chkPoliceReport.Enabled = blnEnable
    chkOther1.Enabled = blnEnable
    chkOther2.Enabled = blnEnable
    If chkOther1 And chkOther1.Enabled Then txtOther1.Enabled = True Else txtOther1.Enabled = False
    If chkOther2 And chkOther2.Enabled Then txtOther2.Enabled = True Else txtOther2.Enabled = False
    mskReminderDateMember.Enabled = blnEnable
End Sub

Private Sub SetActionsForHospital(ByVal blnEnable As Boolean)
    chkPresentingSymptoms.Enabled = blnEnable
    chkMedicalReimbPart2.Enabled = blnEnable
    chkHisthophatologyReport.Enabled = blnEnable
    chkLaboratoryReports.Enabled = blnEnable
    chkReferralLetterHos.Enabled = blnEnable
    chkXrayCtMRIReport.Enabled = blnEnable
    chkIVUReport.Enabled = blnEnable
    chkItemisedBreakdown.Enabled = blnEnable
    chkBreakdown.Enabled = blnEnable
    chkOtherHos.Enabled = blnEnable
    If chkOtherHos And chkOtherHos.Enabled Then txtOtherHospital.Enabled = True Else txtOtherHospital.Enabled = False
    mskReminderDateHospital.Enabled = blnEnable
End Sub

Private Sub SetSaveResetPrintButtonsForPayor(ByVal blnSave As Boolean, ByVal blnReset As Boolean, ByVal blnPrint As Boolean)
    btnSavePayor.Enabled = blnSave
    btnResetPayor.Enabled = blnReset
    btnPrintPayor.Enabled = blnPrint
End Sub

Private Sub SetAddEditSavePrintButtonsForMember(ByVal blnAdd As Boolean, ByVal blnEdit As Boolean, ByVal blnSave As Boolean, ByVal blnPrint As Boolean)
    btnEditMember.Enabled = blnEdit
    btnAddMember.Enabled = blnAdd
    btnSaveMember.Enabled = blnSave
    btnPrintMember.Enabled = blnPrint
End Sub

Private Sub GetAddEditSavePrintButtonsForMember(ByRef blnAdd As Boolean, ByRef blnEdit As Boolean, ByRef blnSave As Boolean, ByRef blnPrint As Boolean)
    blnEdit = btnEditMember.Enabled
    blnAdd = btnAddMember.Enabled
    blnSave = btnSaveMember.Enabled
    blnPrint = btnPrintMember.Enabled
End Sub

Private Sub SetAddEditSavePrintButtonsForHospital(ByRef blnAdd As Boolean, ByRef blnEdit As Boolean, ByRef blnSave As Boolean, ByRef blnPrint As Boolean)
    btnEditHospital.Enabled = blnEdit
    btnAddHospital.Enabled = blnAdd
    btnSaveHospital.Enabled = blnSave
    btnPrintHospital.Enabled = blnPrint
End Sub

Private Sub CheckIfSaveResetPrintCanSetForPayor()
    If blnHasSearched Then
        If chkAppealLetter Or chkRepudiationLetter Or _
            chkWorksheet Or chkReports Or _
            chkOfficialBillsAndReceipt Or chkRelevantDocumentsNote _
            Or chkOther Then
            SetSaveResetPrintButtonsForPayor True, True, True
        Else
            SetSaveResetPrintButtonsForPayor False, False, False
        End If
    End If
End Sub

Private Sub CheckIfAddEditSavePrintCanSetForMember()
    If blnHasSearched Then
        If blnCanSaveMember Then
            If chkMemberReimbursementForm Or chkUndertakingLetter Or _
                chkOfficalReceipt Or chkOriginalInvoice Or _
                chkReferralLetter Or chkStudentID _
                Or chkCompleteMedicalForm Or chkInternationalPassport Or _
                chkPoliceReport Or chkOther1 Or chkOther2 Then
                SetAddEditSavePrintButtonsForMember False, False, True, False
            Else
                SetAddEditSavePrintButtonsForMember False, False, False, False
            End If
        End If
    End If
End Sub

Private Sub CheckIfAddEditSavePrintCanSetForHospital()
    If blnHasSearched Then
        If blnCanSaveHospital Then
            If chkPresentingSymptoms Or chkMedicalReimbPart2 Or _
                chkHisthophatologyReport Or chkLaboratoryReports Or _
                chkReferralLetterHos Or _
                chkXrayCtMRIReport Or chkIVUReport _
                Or chkItemisedBreakdown Or chkBreakdown Or _
                chkOtherHos Then
                SetAddEditSavePrintButtonsForHospital False, False, True, False
            Else
                SetAddEditSavePrintButtonsForHospital False, False, False, False
            End If
        End If
    End If
End Sub

Private Sub ResetActionsForMember()
    chkMemberReimbursementForm = 0
    chkUndertakingLetter = 0
    chkOfficalReceipt = 0
    chkOriginalInvoice = 0
    chkReferralLetter = 0
    chkStudentID = 0
    chkCompleteMedicalForm = 0
    chkInternationalPassport = 0
    chkPoliceReport = 0
    chkOther1 = 0
    chkOther2 = 0
    txtOther1 = ""
    txtOther2 = ""
    mskReminderDateMember.Text = Format(DateTime.Now, "dd/MM/yyyy")
End Sub

Private Sub ResetActionsForHospital()
    chkPresentingSymptoms = 0
    chkMedicalReimbPart2 = 0
    chkHisthophatologyReport = 0
    chkLaboratoryReports = 0
    chkReferralLetterHos = 0
    chkXrayCtMRIReport = 0
    chkIVUReport = 0
    chkItemisedBreakdown = 0
    chkBreakdown = 0
    chkOtherHos = 0
    txtOtherHospital = ""
    mskReminderDateHospital.Text = Format(DateTime.Now, "dd/MM/yyyy")
    lstMedDocReq.FullRowSelect = True
    lstMedDocReq.Gridlines = True
End Sub

Private Sub SetupReminderListView()
    Dim col As ColumnHeader
    
    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 1
    col.Text = "Rem No."
    
    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 2
    col.Text = "Case No."

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 3
    col.Text = "Rem. Date"

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 4
    col.Text = "MRF"

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 5
    col.Text = "Orig. Inv."

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 6
    col.Text = "Med. Form"

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 7
    col.Text = "U/taking"

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 8
    col.Text = "Ref. Letter"
    
    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 9
    col.Text = "Inter. Passport"

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 10
    col.Text = "Orig. Recpt."
    
    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 11
    col.Text = "Stud. ID"

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 12
    col.Text = "Police Rpt."

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 13
    col.Text = "Other1"

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 14
    col.Text = "Other2"
    
    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 15
    col.Text = "Other1Text"

    Set col = lstReminder.ColumnHeaders.Add
    col.Position = 16
    col.Text = "Other2Text"

End Sub

Private Sub SetupMedDocReqListView()
    Dim col As ColumnHeader
    
    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 1
    col.Text = "Req No."
    
    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 2
    col.Text = "Case No."
    
    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 3
    col.Text = "Rem. Date"
    
    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 4
    col.Text = "Present Symp."
    
    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 5
    col.Text = "MRF Part 2"

    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 6
    col.Text = "Histho. Rpt."

    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 7
    col.Text = "Ref. Letter"

    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 8
    col.Text = "Lab Rpt."

    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 9
    col.Text = "X-ray/Ct/MRI"

    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 10
    col.Text = "IVU Rpt."
    
    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 11
    col.Text = "Itemised Brkdown"
    
    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 12
    col.Text = "Brkdown"

    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 13
    col.Text = "Other"

    Set col = lstMedDocReq.ColumnHeaders.Add
    col.Position = 14
    col.Text = "OtherText"

End Sub

Private Sub GenerateWordDocHospital()

    Dim wordApp As New Word.Application
    Dim wordDoc As Word.Document
    Dim caseRange As Word.Range
    Dim intIdx As Integer
    Set wordDoc = wordApp.Documents.Add(LocCardPath & "Correspondence\HosDocRequest.dot", , , True)
    wordDoc.Activate
    
    'initializes all the bookmarks with values from the correspondence form
    Set caseRange = wordDoc.Bookmarks("BK_CaseNo").Range
    caseRange.Text = UCase(strCasNumber)
    caseRange.Select
    caseRange.Bold = True
    caseRange.End = caseRange.Start
    caseRange.Select
    
    wordDoc.Bookmarks("BK_ReminderDate").Range.Text = Format(mskReminderDateHospital, "d MMM yyyy")
    'wordDoc.Bookmarks("BK_InsuredAddress").Range.Text = strMemberAddress
    'wordDoc.Bookmarks("BK_InsuredSalute").Range.Text = strMembershipNo
    wordDoc.Bookmarks("BK_PatientName").Range.Text = strPatientName
    wordDoc.Bookmarks("BK_PatientIC").Range.Text = IIf(strPatientIC = "", "N/A", strPatientIC)
    wordDoc.Bookmarks("BK_HospitalName").Range.Text = strHosName
    wordDoc.Bookmarks("BK_HospitalAddress1").Range.Text = strHosAdd1
    wordDoc.Bookmarks("BK_HospitalAddress2").Range.Text = strHosAdd2
    wordDoc.Bookmarks("BK_HospitalAddress3").Range.Text = strHosAdd3
    wordDoc.Bookmarks("BK_MailFax").Range.Text = strHosFaxNo
    wordDoc.Bookmarks("BK_DOA").Range.Text = Format(dteDOA, "dd/MM/yyyy")
    'wordDoc.Bookmarks("BK_DrName").Range.Text = strAttendDoctor
    wordDoc.Bookmarks("BK_DrName").Range.Text = strHosContact

    intIdx = 1
    
'Private Const PS_DURATION = "Duration: "

    If chkPresentingSymptoms Then
        wordDoc.Bookmarks("BK_Item1").Range.Text = CStr(intIdx) & ". " & PRESENTING_SYMPTOMS & strReason & _
            vbTab & "Duration: _____________ (dd/mm/yy)"
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Item1"
        DeleteBookmark wordDoc, "BK_AfterItem1"
    End If
    
    If chkMedicalReimbPart2 Then
        wordDoc.Bookmarks("BK_Item2").Range.Text = CStr(intIdx) & ". " & MEDICAL_REIMB_PART2
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Item2"
        DeleteBookmark wordDoc, "BK_AfterItem2"
    End If
    
     If chkHisthophatologyReport Then
        wordDoc.Bookmarks("BK_Item3").Range.Text = CStr(intIdx) & ". " & HISTHOPHATOLOGY_REPORT
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Item3"
        DeleteBookmark wordDoc, "BK_AfterItem3"
    End If
   
    If chkReferralLetterHos Then
        wordDoc.Bookmarks("BK_Item4").Range.Text = CStr(intIdx) & ". " & REFERRAL_LETTER
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Item4"
        DeleteBookmark wordDoc, "BK_AfterItem4"
    End If
   
    If chkLaboratoryReports Then
        wordDoc.Bookmarks("BK_Item5").Range.Text = CStr(intIdx) & ". " & LAB_REPORT
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Item5"
        DeleteBookmark wordDoc, "BK_AfterItem5"
    End If
   
    If chkXrayCtMRIReport Then
        wordDoc.Bookmarks("BK_Item6").Range.Text = CStr(intIdx) & ". " & X_RAY_REPORT
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Item6"
        DeleteBookmark wordDoc, "BK_AfterItem6"
    End If

    If chkIVUReport Then
        wordDoc.Bookmarks("BK_Item7").Range.Text = CStr(intIdx) & ". " & IVU_REPORT
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Item7"
        DeleteBookmark wordDoc, "BK_AfterItem7"
    End If

    If chkOtherHos Then
        wordDoc.Bookmarks("BK_Item8").Range.Text = CStr(intIdx) & ". " & HOS_OTHER & txtOtherHospital
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Item8"
        DeleteBookmark wordDoc, "BK_AfterItem8"
    End If

    If chkItemisedBreakdown Then
        Dim BK_Item9Table_Range As Word.Range
        
        wordDoc.Bookmarks("BK_Item9").Range.Text = CStr(intIdx) & ". " & ITEMISED_BREAKDOWN
        
        Set BK_Item9Table_Range = wordDoc.Bookmarks("BK_Item9Table").Range
        wordDoc.Bookmarks("BK_Item9Table").Range.Tables.Add BK_Item9Table_Range, 6, 3
        BK_Item9Table_Range.Tables(1).Rows.SetLeftIndent 55, wdAdjustNone
        BK_Item9Table_Range.Tables(1).Columns(1).SetWidth 32.4, wdAdjustNone
        BK_Item9Table_Range.Tables(1).Columns(2).SetWidth 234, wdAdjustNone
        BK_Item9Table_Range.Tables(1).Columns(3).SetWidth 86.4, wdAdjustNone
        BK_Item9Table_Range.Tables(1).Cell(1, 1).Range.Text = "No."
        BK_Item9Table_Range.Tables(1).Cell(1, 1).Range.Select
        BK_Item9Table_Range.Tables(1).Cell(1, 1).Range.Bold = True
        BK_Item9Table_Range.Tables(1).Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                
        BK_Item9Table_Range.Tables(1).Cell(1, 2).Range.Text = "Description"
        BK_Item9Table_Range.Tables(1).Cell(1, 2).Range.Select
        BK_Item9Table_Range.Tables(1).Cell(1, 2).Range.Bold = True
        BK_Item9Table_Range.Tables(1).Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        BK_Item9Table_Range.Tables(1).Cell(1, 3).Range.Text = "Amount"
        BK_Item9Table_Range.Tables(1).Cell(1, 3).Range.Select
        BK_Item9Table_Range.Tables(1).Cell(1, 3).Range.Bold = True
        BK_Item9Table_Range.Tables(1).Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Item9"
        DeleteBookmark wordDoc, "BK_AfterItem9"
        DeleteBookmark wordDoc, "BK_Item9Table"
        DeleteBookmark wordDoc, "BK_AfterItem9Table"
    End If

    If chkBreakdown Then
        Dim BK_Item10Table_Range As Word.Range
        
        wordDoc.Bookmarks("BK_Item10").Range.Text = CStr(intIdx) & ". " & BREAKDOWN
        
        Set BK_Item10Table_Range = wordDoc.Bookmarks("BK_Item10Table").Range
        wordDoc.Bookmarks("BK_Item10Table").Range.Tables.Add BK_Item10Table_Range, 6, 5
        BK_Item10Table_Range.Tables(1).Rows.SetLeftIndent 55, wdAdjustNone
        BK_Item10Table_Range.Tables(1).Columns(1).SetWidth 40, wdAdjustNone
        BK_Item10Table_Range.Tables(1).Columns(2).SetWidth 160, wdAdjustNone
        BK_Item10Table_Range.Tables(1).Columns(3).SetWidth 60, wdAdjustNone
        BK_Item10Table_Range.Tables(1).Columns(4).SetWidth 65, wdAdjustNone
        BK_Item10Table_Range.Tables(1).Columns(5).SetWidth 80, wdAdjustNone
        
        BK_Item10Table_Range.Tables(1).Cell(1, 1).Range.Text = "Date"
        BK_Item10Table_Range.Tables(1).Cell(1, 1).Range.Select
        BK_Item10Table_Range.Tables(1).Cell(1, 1).Range.Bold = True
        BK_Item10Table_Range.Tables(1).Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                
        BK_Item10Table_Range.Tables(1).Cell(1, 2).Range.Text = "Description/Medication"
        BK_Item10Table_Range.Tables(1).Cell(1, 2).Range.Select
        BK_Item10Table_Range.Tables(1).Cell(1, 2).Range.Bold = True
        BK_Item10Table_Range.Tables(1).Cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        BK_Item10Table_Range.Tables(1).Cell(1, 3).Range.Text = "Qty"
        BK_Item10Table_Range.Tables(1).Cell(1, 3).Range.Select
        BK_Item10Table_Range.Tables(1).Cell(1, 3).Range.Bold = True
        BK_Item10Table_Range.Tables(1).Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        BK_Item10Table_Range.Tables(1).Cell(1, 4).Range.Text = "Duration"
        BK_Item10Table_Range.Tables(1).Cell(1, 4).Range.Select
        BK_Item10Table_Range.Tables(1).Cell(1, 4).Range.Bold = True
        BK_Item10Table_Range.Tables(1).Cell(1, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        BK_Item10Table_Range.Tables(1).Cell(1, 5).Range.Text = "Charges(RM)"
        BK_Item10Table_Range.Tables(1).Cell(1, 5).Range.Select
        BK_Item10Table_Range.Tables(1).Cell(1, 5).Range.Bold = True
        BK_Item10Table_Range.Tables(1).Cell(1, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Item10"
        DeleteBookmark wordDoc, "BK_AfterItem10"
        DeleteBookmark wordDoc, "BK_Item10Table"
    End If

    wordDoc.ActiveWindow.Visible = True
End Sub

Private Sub GenerateWordDocMember()
    Dim wordApp As New Word.Application
    Dim wordDoc As Word.Document
    Dim caseRange As Word.Range
    Dim intIdx As Integer
    
    Set wordDoc = wordApp.Documents.Add(LocCardPath & "Correspondence\DocRequest.dot", , , True)
    wordDoc.Activate
    
    'initializes all the bookmarks with values from the correspondence form
    Set caseRange = wordDoc.Bookmarks("BK_CaseNo").Range
    caseRange.Text = UCase(strCasNumber)
    caseRange.Select
    caseRange.Bold = True
    caseRange.End = caseRange.Start
    caseRange.Select
    
    wordDoc.Bookmarks("BK_ReminderDate").Range.Text = Format(mskReminderDateMember, "d MMM yyyy")
    wordDoc.Bookmarks("BK_InsuredName").Range.Text = strPrincipleName
    wordDoc.Bookmarks("BK_InsuredAddress").Range.Text = strMemberAddress
    Debug.Print strMemberAddress
    wordDoc.Bookmarks("BK_MembershipNo").Range.Text = strMembershipNo
    wordDoc.Bookmarks("BK_PolicyNo").Range.Text = strPolicyNo
    wordDoc.Bookmarks("BK_InsuredName2").Range.Text = strPrincipleName
    wordDoc.Bookmarks("BK_PatientName").Range.Text = strPatientName
    wordDoc.Bookmarks("BK_HospitalName").Range.Text = strHosName
    wordDoc.Bookmarks("BK_DOA").Range.Text = Format(dteDOA, "dd/MM/yyyy")
    wordDoc.Bookmarks("BK_DOD").Range.Text = Format(dteDOD, "dd/MM/yyyy")
    
    intIdx = 1
    
    If chkMemberReimbursementForm Then
        wordDoc.Bookmarks("BK_MemberReimbursementForm").Range.Text = CStr(intIdx) & ". " & chkMemberReimbursementForm.Caption
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_MemberReimbursementForm"
    End If
    
    If chkOriginalInvoice Then
        wordDoc.Bookmarks("BK_OriginalInvoice").Range.Text = CStr(intIdx) & ". " & chkOriginalInvoice.Caption
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_OriginalInvoice"
    End If
    
     If chkCompleteMedicalForm Then
        wordDoc.Bookmarks("BK_CompleteMedicalForm").Range.Text = CStr(intIdx) & ". " & chkCompleteMedicalForm.Caption
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_CompleteMedicalForm"
    End If
   
    If chkUndertakingLetter Then
        wordDoc.Bookmarks("BK_UndertakingLetter").Range.Text = CStr(intIdx) & ". " & chkUndertakingLetter.Caption
        intIdx = intIdx + 1
        
        Dim wordApp1 As New Word.Application
        Dim wordDoc1 As Word.Document
        Dim caseRange1 As Word.Range
            
        Set wordDoc1 = wordApp1.Documents.Add(LocCardPath & "Correspondence\UndertakingLetter.dot", , , True)
        wordDoc1.Activate
        
        wordDoc1.Bookmarks("BK_PrincipleName").Range.Text = UCase(Trim(strPrincipleName))
        wordDoc1.Bookmarks("BK_PrincipleIC").Range.Text = UCase(Trim(strPrincipleIC))
        wordDoc1.Bookmarks("BK_PolicyNo").Range.Text = UCase(Trim(strPolicyNo))
        wordDoc1.Bookmarks("BK_MemberNumber").Range.Text = UCase(Trim(strMembershipNo))
        wordDoc1.Bookmarks("BK_CaseNo").Range.Text = UCase(Trim(strCasNumber))
        wordDoc1.Bookmarks("BK_PatientName").Range.Text = UCase(Trim(strPatientName))
        wordDoc1.Bookmarks("BK_DOA").Range.Text = Format(IIf(IsNull(dteDOA), "N/A", dteDOA), "dd/MM/yyyy")
        wordDoc1.Bookmarks("BK_DOD").Range.Text = Format(IIf(IsNull(dteDOD), "N/A", dteDOA), "dd/MM/yyyy")
        wordDoc1.Bookmarks("BK_AdmissionReason").Range.Text = UCase(Trim(strReason))
        wordDoc1.Bookmarks("BK_HospitalName").Range.Text = UCase(Trim(strHosName))
        wordDoc1.Bookmarks("BK_PrincipleName1").Range.Text = UCase(Trim(strPrincipleName))
        wordDoc1.Bookmarks("BK_PrincipleIC1").Range.Text = UCase(Trim(strPrincipleIC))
        
        wordDoc1.ActiveWindow.Visible = True
    Else
        DeleteBookmark wordDoc, "BK_UndertakingLetter"
    End If
   
    If chkReferralLetter Then
        wordDoc.Bookmarks("BK_ReferralLetter").Range.Text = CStr(intIdx) & ". " & chkReferralLetter.Caption
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_ReferralLetter"
    End If
   
    If chkInternationalPassport Then
        wordDoc.Bookmarks("BK_InternationalPassport").Range.Text = CStr(intIdx) & ". " & chkInternationalPassport.Caption
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_InternationalPassport"
    End If

    If chkOfficalReceipt Then
        wordDoc.Bookmarks("BK_OfficalReceipt").Range.Text = CStr(intIdx) & ". " & chkOfficalReceipt.Caption
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_OfficalReceipt"
    End If

    If chkStudentID Then
        wordDoc.Bookmarks("BK_StudentID").Range.Text = CStr(intIdx) & ". " & chkStudentID.Caption
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_StudentID"
    End If

    If chkPoliceReport Then
        wordDoc.Bookmarks("BK_PoliceReport").Range.Text = CStr(intIdx) & ". " & chkPoliceReport.Caption
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_PoliceReport"
    End If

    If chkOther1 Then
        wordDoc.Bookmarks("BK_Other1").Range.Text = CStr(intIdx) & ". " & txtOther1
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Other1"
    End If

    If chkOther2 Then
        wordDoc.Bookmarks("BK_Other2").Range.Text = CStr(intIdx) & ". " & txtOther2
        intIdx = intIdx + 1
    Else
        DeleteBookmark wordDoc, "BK_Other2"
    End If

    wordDoc.ActiveWindow.Visible = True
End Sub

Private Sub GenerateWordDocPayor()
    Dim wordApp As New Word.Application
    Dim wordDoc As Word.Document
    Dim caseRange As Word.Range
    
    Set wordDoc = wordApp.Documents.Add(LocCardPath & "Correspondence\appealack.dot", , , True)
    wordDoc.Activate
    
    'initializes all the bookmarks with values from the correspondence form
    Set caseRange = wordDoc.Bookmarks("casNumber").Range
    caseRange.Text = UCase(strCasNumber)
    caseRange.Select
    caseRange.Bold = True
    caseRange.End = caseRange.Start
    caseRange.Select
    
    wordDoc.Bookmarks("reminderDate").Range.Text = Format(mskAppealLetter, "d MMM yyyy")
    wordDoc.Bookmarks("payCompanyName").Range.Text = strClientCompany
    wordDoc.Bookmarks("payCompanyAddress").Range.Text = strClientCompanyAddress
    wordDoc.Bookmarks("mbmNo").Range.Text = strMembershipNo
    wordDoc.Bookmarks("policyNo").Range.Text = strPolicyNo
    wordDoc.Bookmarks("insuredName").Range.Text = strPrincipleName
    wordDoc.Bookmarks("patientName").Range.Text = strPatientName
    wordDoc.Bookmarks("hospitalName").Range.Text = strHosName
    wordDoc.Bookmarks("DOA").Range.Text = Format(dteDOA, "dd/MM/yyyy")
    wordDoc.Bookmarks("insuredName2").Range.Text = strPrincipleName
    
    If chkAppealLetter Then
        wordDoc.Bookmarks("appealLetter").Range.Text = ChrW(8730) & " " & APPEAL_LETTER & Format(mskAppealLetter, "d MMM yyyy")
    Else
        DeleteBookmark wordDoc, "appealLetter"
        DeleteBookmark wordDoc, "afterAppealLetter"
    End If
    
    If chkRepudiationLetter Then
        wordDoc.Bookmarks("repudiationLetter").Range.Text = ChrW(8730) & " " & REPUDIATION_LETTER
    Else
        DeleteBookmark wordDoc, "repudiationLetter"
        DeleteBookmark wordDoc, "afterRepudiationLetter"
    End If
    
     If chkWorksheet Then
        wordDoc.Bookmarks("worksheet").Range.Text = ChrW(8730) & " " & WORKSHEET & Format(mskWorksheet, "d MMM yyyy")
    Else
        DeleteBookmark wordDoc, "worksheet"
        DeleteBookmark wordDoc, "afterWorksheet"
    End If
   
    If chkReports Then
        wordDoc.Bookmarks("reports").Range.Text = ChrW(8730) & " " & REPORTS
    Else
        DeleteBookmark wordDoc, "reports"
        DeleteBookmark wordDoc, "afterReports"
    End If
   
    If chkOfficialBillsAndReceipt Then
        wordDoc.Bookmarks("officialBillsAndReceipt").Range.Text = ChrW(8730) & " " & OFFICIAL_BILLS_AND_RECEIPT
    Else
        DeleteBookmark wordDoc, "officialBillsAndReceipt"
        DeleteBookmark wordDoc, "afterOfficialBillsAndReceipt"
    End If
   
    If chkRelevantDocumentsNote Then
        wordDoc.Bookmarks("relevantDocumentsNote").Range.Text = ChrW(8730) & " " & RELEVANT_DOCUMENTS_NOTE
    Else
        DeleteBookmark wordDoc, "relevantDocumentsNote"
        DeleteBookmark wordDoc, "afterRelevantDocumentsNote"
    End If

    If chkOther Then
        wordDoc.Bookmarks("others").Range.Text = ChrW(8730) & " " & Others & txtOther
    Else
        DeleteBookmark wordDoc, "others"
        DeleteBookmark wordDoc, "afterOthers"
    End If

    wordDoc.ActiveWindow.Visible = True
End Sub

Private Sub DeleteBookmark(ByRef doc As Word.Document, ByRef strRange As String)
    Dim theRange As Word.Range

    Set theRange = doc.Bookmarks(strRange).Range
    theRange.Select
    theRange.Delete
    doc.Bookmarks(strRange).Delete
End Sub

Private Sub tmrRefresh_Timer()
'    MsgBox "timer"
    Call cmdShow_Click
    tmrRefresh.Enabled = False
End Sub
