VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNewMBMEnquiry 
   Caption         =   "Member Enquiry"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   11850
   Begin TabDlg.SSTab SSTab1 
      Height          =   5760
      Left            =   120
      TabIndex        =   13
      Top             =   1965
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10160
      _Version        =   393216
      Tabs            =   17
      TabsPerRow      =   17
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Prin &Detail"
      TabPicture(0)   =   "frmNewMBMEnquiry.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Supp"
      TabPicture(1)   =   "frmNewMBMEnquiry.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "lstCovAdd"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Adjust History"
      TabPicture(2)   =   "frmNewMBMEnquiry.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstAdjustmentHis"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "HIS &Benefits"
      TabPicture(3)   =   "frmNewMBMEnquiry.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstBenefitDetails"
      Tab(3).Control(1)=   "lstNewBenefitDetails"
      Tab(3).Control(2)=   "lstSuppBenefitDetails"
      Tab(3).Control(3)=   "Label105(1)"
      Tab(3).Control(4)=   "lblCovName"
      Tab(3).Control(5)=   "Label37"
      Tab(3).Control(6)=   "Label105(0)"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "&Current Case"
      TabPicture(4)   =   "frmNewMBMEnquiry.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label7"
      Tab(4).Control(1)=   "lstCaseList"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Case History"
      TabPicture(5)   =   "frmNewMBMEnquiry.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstCASHistory"
      Tab(5).Control(1)=   "lstCASHistoryTwo"
      Tab(5).Control(2)=   "lblTotalCasesDOB"
      Tab(5).Control(3)=   "lblTotalCasesIC"
      Tab(5).Control(4)=   "Label1(26)"
      Tab(5).Control(5)=   "Label1(25)"
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "Mem &History"
      TabPicture(6)   =   "frmNewMBMEnquiry.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lstMBMHis"
      Tab(6).Control(1)=   "lstMBMHisTwo"
      Tab(6).Control(2)=   "lblTotalMBMDOB"
      Tab(6).Control(3)=   "lblTotalMBMIC"
      Tab(6).Control(4)=   "Label1(28)"
      Tab(6).Control(5)=   "Label1(27)"
      Tab(6).ControlCount=   6
      TabCaption(7)   =   "Account"
      TabPicture(7)   =   "frmNewMBMEnquiry.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cmdTotal"
      Tab(7).Control(1)=   "Frame12"
      Tab(7).Control(2)=   "Frame13"
      Tab(7).Control(3)=   "Frame9"
      Tab(7).Control(4)=   "lstMBMAcc"
      Tab(7).Control(5)=   "Frame11"
      Tab(7).ControlCount=   6
      TabCaption(8)   =   "Acc History"
      TabPicture(8)   =   "frmNewMBMEnquiry.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lstACCHistory"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Fees"
      TabPicture(9)   =   "frmNewMBMEnquiry.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame2"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Non Disclosure"
      TabPicture(10)  =   "frmNewMBMEnquiry.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lstPreUndIllness"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "Policy Terms"
      TabPicture(11)  =   "frmNewMBMEnquiry.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Fgrid"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "LT Limit"
      TabPicture(12)  =   "frmNewMBMEnquiry.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "lstLTLmt"
      Tab(12).ControlCount=   1
      TabCaption(13)  =   "Policy Wording"
      TabPicture(13)  =   "frmNewMBMEnquiry.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "lstPolicyWording"
      Tab(13).Control(1)=   "CmnDlg"
      Tab(13).ControlCount=   2
      TabCaption(14)  =   "T/O Claims"
      TabPicture(14)  =   "frmNewMBMEnquiry.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "lstTOClaims"
      Tab(14).ControlCount=   1
      TabCaption(15)  =   "Notes"
      TabPicture(15)  =   "frmNewMBMEnquiry.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "txtMgmtNotes"
      Tab(15).ControlCount=   1
      TabCaption(16)  =   "WS"
      TabPicture(16)  =   "frmNewMBMEnquiry.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "lstWorksheet"
      Tab(16).ControlCount=   1
      Begin VB.TextBox txtMgmtNotes 
         Height          =   4905
         Left            =   -74880
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   286
         Top             =   480
         Width           =   11460
      End
      Begin VB.ListBox lstPolicyWording 
         Height          =   4935
         ItemData        =   "frmNewMBMEnquiry.frx":01DC
         Left            =   -74880
         List            =   "frmNewMBMEnquiry.frx":01DE
         TabIndex        =   263
         ToolTipText     =   "right click to randomize contents"
         Top             =   400
         Width           =   4455
      End
      Begin VB.Frame Frame2 
         Height          =   1275
         Left            =   -74880
         TabIndex        =   193
         Top             =   480
         Width           =   2505
         Begin VB.Line Line1 
            X1              =   120
            X2              =   2400
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   2400
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label txtMBMProviderRedCharges 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   7560
            TabIndex        =   211
            Top             =   555
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label txtMBMProviderEffCharges 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   7560
            TabIndex        =   210
            Top             =   915
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label77 
            Caption         =   "Effective Access Fee"
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
            Left            =   5880
            TabIndex        =   209
            Top             =   915
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label Label79 
            Caption         =   "Acces Fee Refund"
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
            Left            =   5880
            TabIndex        =   208
            Top             =   555
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Label61 
            Caption         =   "MCO Refund"
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
            Left            =   2955
            TabIndex        =   207
            Top             =   555
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label txtMBMMCORed 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   4440
            TabIndex        =   206
            Top             =   555
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label txtMBMPremiumRed 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   1320
            TabIndex        =   205
            Top             =   560
            Width           =   1035
         End
         Begin VB.Label Label58 
            Caption         =   "Prem Refund"
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
            TabIndex        =   204
            Top             =   560
            Width           =   1020
         End
         Begin VB.Label txtMBMPremium 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   1320
            TabIndex        =   203
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label lblPremium 
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
            Height          =   255
            Left            =   120
            TabIndex        =   202
            Top             =   195
            Width           =   960
         End
         Begin VB.Label txtMBMProviderCharges 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   7560
            TabIndex        =   201
            Top             =   180
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label txtMCOFeesEff 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   4440
            TabIndex        =   200
            Top             =   915
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label txtMBMPremiumEff 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   1320
            TabIndex        =   199
            Top             =   915
            Width           =   1035
         End
         Begin VB.Label txtMCOFees 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   4440
            TabIndex        =   198
            Top             =   180
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label83 
            Caption         =   "Effective Prem "
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
            TabIndex        =   197
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label Label85 
            Caption         =   "Effective MCO Fee "
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
            Left            =   2955
            TabIndex        =   196
            Top             =   915
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblMCOFee 
            Caption         =   "MCO Fees"
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
            Left            =   2955
            TabIndex        =   195
            Top             =   195
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.Label Label11 
            Caption         =   "Access Fee"
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
            Left            =   5880
            TabIndex        =   194
            Top             =   180
            Visible         =   0   'False
            Width           =   1140
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3435
         Left            =   -74880
         TabIndex        =   108
         Top             =   360
         Width           =   11460
         Begin VB.TextBox txtCovExclusion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   6000
            MaxLength       =   2500
            MultiLine       =   -1  'True
            TabIndex        =   126
            Top             =   1920
            Width           =   3375
         End
         Begin VB.ComboBox cboCovSalutation 
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
            Left            =   1080
            TabIndex        =   125
            Text            =   "MS"
            Top             =   195
            Width           =   1020
         End
         Begin VB.TextBox txtCovName 
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
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   124
            Top             =   180
            Width           =   3660
         End
         Begin VB.TextBox txtCovIcBcPp 
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
            MaxLength       =   30
            TabIndex        =   123
            Top             =   180
            Width           =   3015
         End
         Begin VB.ComboBox cboCovRelationship 
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
            Left            =   1080
            TabIndex        =   122
            Text            =   "W -Wife"
            Top             =   480
            Width           =   1020
         End
         Begin VB.ComboBox cboCovAdultChild 
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
            Left            =   1080
            TabIndex        =   121
            Text            =   "A"
            Top             =   720
            Width           =   1020
         End
         Begin VB.ComboBox cboCovSex 
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
            Left            =   1080
            TabIndex        =   116
            Text            =   "M"
            Top             =   980
            Width           =   1020
         End
         Begin VB.TextBox txtPrevPLNCode 
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
            Left            =   1080
            TabIndex        =   120
            Top             =   1260
            Width           =   1020
         End
         Begin VB.TextBox txtPrevINSCom 
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
            MaxLength       =   60
            TabIndex        =   119
            Top             =   1520
            Width           =   1020
         End
         Begin VB.ComboBox cboSuppStaffDisc 
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
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   115
            Top             =   1760
            Width           =   1020
         End
         Begin VB.TextBox txtSuppPLNCode 
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
            Left            =   6000
            MaxLength       =   8
            TabIndex        =   214
            Top             =   480
            Width           =   1140
         End
         Begin VB.TextBox txtRBAmt 
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
            MaxLength       =   8
            TabIndex        =   118
            Top             =   2040
            Width           =   1020
         End
         Begin VB.TextBox txtLoadingPrem 
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
            MaxLength       =   8
            TabIndex        =   187
            Top             =   2300
            Width           =   1020
         End
         Begin VB.TextBox txtUndExcess 
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
            MaxLength       =   8
            TabIndex        =   222
            Top             =   2550
            Width           =   1020
         End
         Begin MSMask.MaskEdBox txtSuppEffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   8280
            TabIndex        =   243
            Top             =   435
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSuppExpDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   10200
            TabIndex        =   244
            Top             =   435
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCovBirthdate 
            Height          =   285
            Left            =   3480
            TabIndex        =   128
            Top             =   480
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtSuppPayPremGross 
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
            Left            =   3480
            TabIndex        =   114
            Top             =   720
            Width           =   1140
         End
         Begin VB.TextBox txtSuppPayMCOGross 
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
            Left            =   3480
            TabIndex        =   113
            Top             =   960
            Width           =   1140
         End
         Begin VB.TextBox txtPolicyDisc 
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
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   109
            Top             =   1200
            Width           =   1140
         End
         Begin VB.TextBox txtPolicyDiscAmt 
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
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   186
            Top             =   1440
            Width           =   1140
         End
         Begin VB.TextBox txtSuppPayPremNet 
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
            Left            =   6000
            TabIndex        =   111
            Top             =   720
            Width           =   1140
         End
         Begin VB.TextBox txtSuppPayMCONet 
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
            Left            =   6000
            TabIndex        =   110
            Top             =   975
            Width           =   1140
         End
         Begin VB.TextBox txtRenewalDisc 
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
            Left            =   6000
            MaxLength       =   8
            TabIndex        =   112
            Top             =   1200
            Width           =   1140
         End
         Begin VB.TextBox txtRenewalDiscAmt 
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
            Left            =   6000
            MaxLength       =   8
            TabIndex        =   188
            Top             =   1440
            Width           =   1140
         End
         Begin VB.TextBox txtCOVPayRemarks 
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
            Left            =   8280
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   117
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtPaySeqID 
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
            MaxLength       =   2
            TabIndex        =   249
            Top             =   1000
            Width           =   1095
         End
         Begin VB.TextBox txtClientID 
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
            MaxLength       =   2
            TabIndex        =   250
            Top             =   1260
            Width           =   1095
         End
         Begin VB.TextBox txtCovAllergic 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   2280
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   127
            Top             =   1920
            Width           =   2415
         End
         Begin VB.TextBox txtPayMBMNo 
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
            MaxLength       =   8
            TabIndex        =   247
            Top             =   1000
            Width           =   1095
         End
         Begin VB.TextBox txtClientNo 
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
            MaxLength       =   8
            TabIndex        =   248
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label45 
            Caption         =   "Client ID"
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
            Left            =   9480
            TabIndex        =   254
            Top             =   1300
            Width           =   615
         End
         Begin VB.Label Label45 
            Caption         =   "Client No"
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
            Left            =   7320
            TabIndex        =   253
            Top             =   1290
            Width           =   975
         End
         Begin VB.Label Label45 
            Caption         =   "Pay Seq"
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
            Left            =   9480
            TabIndex        =   252
            Top             =   1050
            Width           =   735
         End
         Begin VB.Label Label45 
            Caption         =   "Pay MBM No"
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
            Left            =   7320
            TabIndex        =   251
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label Label62 
            Caption         =   "ExpDate"
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
            Left            =   9480
            TabIndex        =   246
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label62 
            Caption         =   "EffDate"
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
            Left            =   7320
            TabIndex        =   245
            Top             =   480
            Width           =   645
         End
         Begin VB.Label lblCovHomeNurseBalLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10340
            TabIndex        =   237
            Top             =   3120
            Width           =   1000
         End
         Begin VB.Label lblCovCancerBalLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7880
            TabIndex        =   236
            Top             =   3120
            Width           =   1000
         End
         Begin VB.Label lblCovKidneyBalLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5450
            TabIndex        =   235
            Top             =   3120
            Width           =   1000
         End
         Begin VB.Label Label33 
            Caption         =   "Kidney Dialysis Bal"
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
            Index           =   1
            Left            =   4000
            TabIndex        =   234
            Top             =   3120
            Width           =   1380
         End
         Begin VB.Label Label33 
            Caption         =   "Home Nursing Bal"
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
            Index           =   5
            Left            =   8950
            TabIndex        =   233
            Top             =   3120
            Width           =   1425
         End
         Begin VB.Label Label33 
            Caption         =   "Cancer Treat. Bal"
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
            Index           =   3
            Left            =   6530
            TabIndex        =   232
            Top             =   3120
            Width           =   1305
         End
         Begin VB.Label lblCovHomeNurseAnnLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10340
            TabIndex        =   231
            Top             =   2760
            Width           =   1000
         End
         Begin VB.Label lblCovCancerAnnLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7880
            TabIndex        =   230
            Top             =   2760
            Width           =   1000
         End
         Begin VB.Label lblCovKidneyAnnLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5450
            TabIndex        =   229
            Top             =   2760
            Width           =   1000
         End
         Begin VB.Label Label33 
            Caption         =   "Kidney Dialysis Ann"
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
            Left            =   4000
            TabIndex        =   228
            Top             =   2760
            Width           =   1380
         End
         Begin VB.Label Label33 
            Caption         =   "Home Nursing Ann"
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
            Index           =   4
            Left            =   8950
            TabIndex        =   227
            Top             =   2760
            Width           =   1395
         End
         Begin VB.Label Label33 
            Caption         =   "Cancer Treat. Ann"
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
            Index           =   2
            Left            =   6525
            TabIndex        =   226
            Top             =   2760
            Width           =   1410
         End
         Begin VB.Line Line3 
            X1              =   2160
            X2              =   11360
            Y1              =   3045
            Y2              =   3045
         End
         Begin VB.Label lblCovAnnLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2950
            TabIndex        =   225
            Top             =   2760
            Width           =   960
         End
         Begin VB.Label lblCovAvLmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2950
            TabIndex        =   224
            Top             =   3120
            Width           =   960
         End
         Begin VB.Line Line2 
            X1              =   2160
            X2              =   11360
            Y1              =   2685
            Y2              =   2685
         End
         Begin VB.Label Label6 
            Caption         =   "Pln Code"
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
            Left            =   4800
            TabIndex        =   215
            Top             =   525
            Width           =   885
         End
         Begin VB.Label Label6 
            Caption         =   "Policy Disc."
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
            Left            =   2280
            TabIndex        =   150
            Top             =   1200
            Width           =   1125
         End
         Begin VB.Label Label6 
            Caption         =   "Renewal Disc."
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
            Left            =   4800
            TabIndex        =   149
            Top             =   1230
            Width           =   1035
         End
         Begin VB.Label Label6 
            Caption         =   "Pay MCO Net"
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
            Index           =   17
            Left            =   4800
            TabIndex        =   148
            Top             =   1000
            Width           =   1125
         End
         Begin VB.Label Label6 
            Caption         =   "Pay MCO Gross"
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
            Left            =   2280
            TabIndex        =   147
            Top             =   960
            Width           =   1125
         End
         Begin VB.Label Label6 
            Caption         =   "Pay Prem Net"
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
            Left            =   4800
            TabIndex        =   146
            Top             =   750
            Width           =   1125
         End
         Begin VB.Label Label6 
            Caption         =   "Pay Prem Gross"
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
            Left            =   2280
            TabIndex        =   145
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label49 
            Caption         =   "Pay Remarks"
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
            Left            =   7320
            TabIndex        =   143
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label91 
            Caption         =   "T/O Ins "
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
            TabIndex        =   142
            Top             =   1500
            Width           =   660
         End
         Begin VB.Label Label6 
            Caption         =   "R&&B Amt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   141
            Top             =   2060
            Width           =   900
         End
         Begin VB.Label Label51 
            Caption         =   "Prev. Pln"
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
            TabIndex        =   140
            Top             =   1250
            Width           =   645
         End
         Begin VB.Label Label6 
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
            Index           =   11
            Left            =   2280
            TabIndex        =   139
            Top             =   200
            Width           =   1140
         End
         Begin VB.Label Label6 
            Caption         =   "Adult/Child"
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
            Index           =   3
            Left            =   120
            TabIndex        =   138
            Top             =   750
            Width           =   825
         End
         Begin VB.Label Label6 
            Caption         =   "Sex "
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
            TabIndex        =   137
            Top             =   1000
            Width           =   825
         End
         Begin VB.Label Label66 
            Caption         =   "Allergic"
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
            Left            =   2280
            TabIndex        =   136
            Top             =   1680
            Width           =   600
         End
         Begin VB.Label Label6 
            Caption         =   "Salutation"
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
            TabIndex        =   135
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label94 
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
            Left            =   7320
            TabIndex        =   134
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label6 
            Caption         =   "DOB"
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
            Left            =   2280
            TabIndex        =   133
            Top             =   460
            Width           =   900
         End
         Begin VB.Label Label6 
            Caption         =   "Supp Lmt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   14
            Left            =   2200
            TabIndex        =   132
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "R'Ship"
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
            Left            =   120
            TabIndex        =   131
            Top             =   500
            Width           =   1005
         End
         Begin VB.Label Label98 
            Caption         =   "Supplementary Warranty "
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
            Left            =   4800
            TabIndex        =   130
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label Label6 
            Caption         =   "Bal. Limit"
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
            Index           =   4
            Left            =   2200
            TabIndex        =   129
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Load'g Prem"
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
            TabIndex        =   191
            Top             =   2300
            Width           =   1005
         End
         Begin VB.Label Label6 
            Caption         =   "Pol. Disc. Amt"
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
            Left            =   2280
            TabIndex        =   190
            Top             =   1440
            Width           =   1365
         End
         Begin VB.Label Label6 
            Caption         =   "Ren. Disc. Amt"
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
            Index           =   18
            Left            =   4800
            TabIndex        =   189
            Top             =   1480
            Width           =   1125
         End
         Begin VB.Label Label6 
            Caption         =   "Disc Req"
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
            Left            =   120
            TabIndex        =   144
            Top             =   1750
            Width           =   825
         End
         Begin VB.Label Label6 
            Caption         =   "Und Excess"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   20
            Left            =   120
            TabIndex        =   223
            Top             =   2560
            Width           =   900
         End
      End
      Begin VB.CommandButton cmdTotal 
         Caption         =   "Total"
         Height          =   375
         Left            =   -64500
         TabIndex        =   107
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame12 
         Caption         =   "Payable to Member"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -69360
         TabIndex        =   91
         Top             =   420
         Width           =   4815
         Begin VB.TextBox txtPayableMem 
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
            Left            =   2280
            TabIndex        =   95
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtPaid2MBM 
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
            Left            =   2280
            TabIndex        =   94
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtPaidByMem 
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
            Left            =   2280
            TabIndex        =   93
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtMemOS 
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
            Left            =   2280
            TabIndex        =   92
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label81 
            Caption         =   "Paid To Mem"
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
            TabIndex        =   99
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label82 
            Caption         =   "O/S"
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
            TabIndex        =   98
            Top             =   960
            Width           =   1425
         End
         Begin VB.Label Label84 
            Caption         =   "Received From Mem"
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
            TabIndex        =   97
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label txtMBMPayable 
            Caption         =   "Payable to/(by) Mem"
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
            TabIndex        =   96
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Worksheet Summary"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   82
         Top             =   420
         Width           =   5055
         Begin VB.TextBox txtAnnualLmt 
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
            Left            =   1320
            TabIndex        =   86
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtAvailableAnnualLmt 
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
            Left            =   1320
            TabIndex        =   85
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtApprovedAmt 
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
            Left            =   1320
            TabIndex        =   84
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtBordxAmt 
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
            Left            =   1320
            TabIndex        =   83
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label89 
            Caption         =   "Bordx Amt"
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
            TabIndex        =   90
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label86 
            Caption         =   "Annual Limit"
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
            TabIndex        =   89
            Top             =   255
            Width           =   1065
         End
         Begin VB.Label Label87 
            Caption         =   "Approved Amt"
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
            TabIndex        =   88
            Top             =   735
            Width           =   1065
         End
         Begin VB.Label Label88 
            Caption         =   "Balance Limit"
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
            TabIndex        =   87
            Top             =   495
            Width           =   945
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Payable To Hospital"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   75
         Top             =   1740
         Width           =   5055
         Begin VB.TextBox txtPayable2Hos 
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
            Left            =   1440
            TabIndex        =   78
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtPaid2Hos 
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
            Left            =   1440
            TabIndex        =   77
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtHosOS 
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
            Left            =   1440
            TabIndex        =   76
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label59 
            Caption         =   "Bill Amt"
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
            TabIndex        =   81
            Top             =   255
            Width           =   1065
         End
         Begin VB.Label Label67 
            Caption         =   "O/S "
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
            TabIndex        =   80
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label68 
            Caption         =   "Paid Amount"
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
            TabIndex        =   79
            Top             =   495
            Width           =   945
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5250
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   11445
         Begin VB.TextBox txtMBMExclusion 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1395
            Left            =   120
            MaxLength       =   40000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   306
            Top             =   3720
            Width           =   3540
         End
         Begin VB.TextBox txtRemarks 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1395
            Left            =   3720
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   305
            Top             =   3720
            Width           =   3660
         End
         Begin VB.TextBox txtSpecialNote 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1395
            Left            =   7440
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   304
            Top             =   3720
            Width           =   3900
         End
         Begin VB.TextBox txtMBMAdd1 
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
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   38
            Top             =   195
            Width           =   3705
         End
         Begin VB.TextBox cboINSCode 
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
            Left            =   6120
            TabIndex        =   37
            Top             =   210
            Width           =   1035
         End
         Begin VB.TextBox cboPLNCode 
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
            Left            =   8280
            TabIndex        =   36
            Top             =   210
            Width           =   1080
         End
         Begin VB.TextBox txtMBMAdd2 
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
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   35
            Top             =   450
            Width           =   3705
         End
         Begin VB.TextBox txtMBMAdd3 
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
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   34
            Top             =   710
            Width           =   3705
         End
         Begin VB.TextBox txtMBMPostCode 
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
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   33
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cboCTYCode 
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
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   32
            Top             =   960
            Width           =   1305
         End
         Begin VB.ComboBox cboSTACode 
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
            Left            =   1200
            TabIndex        =   31
            Top             =   1220
            Width           =   3705
         End
         Begin VB.ComboBox cboMBMSex 
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
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   1505
            Width           =   1215
         End
         Begin VB.ComboBox cboMBMTakeOver 
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
            Left            =   3600
            TabIndex        =   29
            Top             =   1505
            Width           =   1305
         End
         Begin VB.ComboBox cboMBMMaritalStatus 
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
            Left            =   3600
            TabIndex        =   28
            Top             =   1770
            Width           =   1305
         End
         Begin VB.ComboBox cboNATCode 
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
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   2030
            Width           =   1305
         End
         Begin MSMask.MaskEdBox txtMBMBirthDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   42
            Top             =   1770
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboRACCode 
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
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   2030
            Width           =   1215
         End
         Begin VB.TextBox txtAGECode 
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
            Left            =   3600
            TabIndex        =   17
            Top             =   2280
            Width           =   1305
         End
         Begin VB.ComboBox CboInsMode 
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
            Left            =   1200
            TabIndex        =   16
            Top             =   2280
            Width           =   1200
         End
         Begin MSMask.MaskEdBox txtNewPolEffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   9480
            TabIndex        =   43
            Top             =   3840
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtPLNCode 
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
            TabIndex        =   15
            Top             =   4080
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.ComboBox cboMBMRenewal 
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
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   24
            Top             =   2570
            Width           =   1200
         End
         Begin VB.ComboBox CboGuaRenewal 
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
            ItemData        =   "frmNewMBMEnquiry.frx":01E0
            Left            =   3600
            List            =   "frmNewMBMEnquiry.frx":01E2
            TabIndex        =   220
            Top             =   2570
            Width           =   1320
         End
         Begin MSMask.MaskEdBox txtPolEffDate 
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
            Left            =   6120
            TabIndex        =   288
            Top             =   480
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
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
            Height          =   315
            Left            =   6120
            TabIndex        =   39
            Top             =   735
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPolNxDue 
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
            Left            =   8280
            TabIndex        =   294
            Top             =   480
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMPolicyYear 
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
            Left            =   10200
            MaxLength       =   25
            TabIndex        =   21
            Top             =   240
            Width           =   1155
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
            Left            =   8280
            TabIndex        =   41
            Top             =   720
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMBordxDate 
            Height          =   300
            Left            =   6120
            TabIndex        =   40
            Top             =   980
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPolIssDate 
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
            Left            =   6120
            TabIndex        =   296
            Top             =   1200
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMDateJoin 
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
            Left            =   10200
            MaxLength       =   25
            TabIndex        =   22
            Top             =   480
            Width           =   1155
         End
         Begin VB.TextBox txtMBMEmployeeNo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Left            =   10200
            MaxLength       =   15
            TabIndex        =   20
            Top             =   770
            Width           =   1155
         End
         Begin MSMask.MaskEdBox txtEndtDate 
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
            Left            =   8280
            TabIndex        =   298
            Top             =   1200
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMPymtDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   295
            Left            =   10200
            TabIndex        =   300
            Top             =   990
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDOR 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   295
            Left            =   10200
            TabIndex        =   192
            Top             =   1260
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMGroupCompany 
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
            Left            =   6120
            MaxLength       =   150
            TabIndex        =   18
            Top             =   1460
            Width           =   3240
         End
         Begin VB.TextBox txtGrpCat 
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
            Left            =   10200
            TabIndex        =   277
            Top             =   1530
            Width           =   1155
         End
         Begin VB.TextBox txtMBMBranch 
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
            Left            =   10200
            TabIndex        =   25
            Top             =   1750
            Width           =   1155
         End
         Begin VB.TextBox txtDeptName 
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
            Left            =   6120
            MaxLength       =   150
            TabIndex        =   19
            Top             =   1740
            Width           =   3240
         End
         Begin VB.TextBox txtMBMAgentCode 
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
            Left            =   10200
            MaxLength       =   15
            TabIndex        =   23
            Top             =   2040
            Width           =   1155
         End
         Begin VB.TextBox txtCostCenter 
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
            Left            =   8280
            TabIndex        =   275
            Top             =   2030
            Width           =   1080
         End
         Begin VB.TextBox txtMBMPosition 
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
            Left            =   6120
            TabIndex        =   279
            Top             =   2030
            Width           =   1035
         End
         Begin VB.TextBox txtSubsidiary 
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
            Left            =   6120
            TabIndex        =   292
            Top             =   2300
            Width           =   1035
         End
         Begin VB.TextBox txtSubDivision 
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
            Left            =   8280
            TabIndex        =   281
            Top             =   2300
            Width           =   1080
         End
         Begin VB.TextBox txtDivision 
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
            Left            =   10200
            TabIndex        =   283
            Top             =   2300
            Width           =   1155
         End
         Begin VB.TextBox txtMBMAppName 
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
            Left            =   6120
            MaxLength       =   150
            TabIndex        =   302
            Top             =   2580
            Width           =   3240
         End
         Begin MSMask.MaskEdBox txtAlterDate 
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
            Left            =   3600
            TabIndex        =   307
            Top             =   2850
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtAlterType 
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
            Left            =   1200
            MaxLength       =   150
            TabIndex        =   308
            Top             =   2830
            Width           =   1200
         End
         Begin VB.TextBox txtRBType 
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
            Left            =   10200
            TabIndex        =   315
            Top             =   2580
            Width           =   1155
         End
         Begin MSMask.MaskEdBox txtFinalExpDate 
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
            Left            =   10200
            TabIndex        =   311
            Top             =   2865
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtProductName 
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
            Left            =   6120
            TabIndex        =   290
            Top             =   2840
            Width           =   3240
         End
         Begin VB.TextBox txtMBMPAYRemarks 
            Height          =   300
            Left            =   6120
            MaxLength       =   60
            MultiLine       =   -1  'True
            TabIndex        =   318
            Top             =   3120
            Width           =   3240
         End
         Begin MSMask.MaskEdBox txtrtDate 
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
            Left            =   3600
            TabIndex        =   320
            Top             =   3120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Retire Date"
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
            Index           =   44
            Left            =   2520
            TabIndex        =   321
            Top             =   3120
            Width           =   1155
         End
         Begin VB.Label Label8 
            Caption         =   "Pay Remarks"
            Height          =   375
            Left            =   5040
            TabIndex        =   319
            Top             =   3150
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Room"
            Height          =   255
            Index           =   0
            Left            =   9480
            TabIndex        =   316
            Top             =   2640
            Width           =   915
         End
         Begin VB.Label Label16 
            Caption         =   "Final Exp "
            Height          =   255
            Index           =   1
            Left            =   9480
            TabIndex        =   312
            Top             =   2925
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Alter Type"
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
            Left            =   120
            TabIndex        =   310
            Top             =   2880
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Alter Date"
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
            Left            =   2520
            TabIndex        =   309
            Top             =   2880
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Applicant"
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
            Left            =   5040
            TabIndex        =   303
            Top             =   2620
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "Pymt Date"
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
            Left            =   9450
            TabIndex        =   301
            Top             =   1050
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Endt Date"
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
            Left            =   7200
            TabIndex        =   299
            Top             =   1250
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Pol Issue Date"
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
            Left            =   5040
            TabIndex        =   297
            Top             =   1250
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Next Due Date"
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
            Left            =   7200
            TabIndex        =   295
            Top             =   525
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Subsidiary"
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
            Left            =   5040
            TabIndex        =   293
            Top             =   2350
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Product"
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
            Left            =   5040
            TabIndex        =   291
            Top             =   2880
            Width           =   1080
         End
         Begin VB.Label Label1 
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
            Index           =   34
            Left            =   5040
            TabIndex        =   289
            Top             =   495
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Division"
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
            Left            =   9480
            TabIndex        =   284
            Top             =   2350
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Sub Division"
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
            Left            =   7200
            TabIndex        =   282
            Top             =   2350
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Position"
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
            Left            =   5040
            TabIndex        =   280
            Top             =   2080
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Grp Cat."
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
            Left            =   9450
            TabIndex        =   278
            Top             =   1520
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Cost Center"
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
            Left            =   7200
            TabIndex        =   276
            Top             =   2080
            Width           =   960
         End
         Begin VB.Label lblCancerAnn 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7155
            TabIndex        =   241
            Top             =   5220
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblHomeNurseAnn 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10215
            TabIndex        =   240
            Top             =   5220
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblCancerBal 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7155
            TabIndex        =   239
            Top             =   5550
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblHomeNurseBal 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10215
            TabIndex        =   238
            Top             =   5550
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Gua. Renewal"
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
            Left            =   2520
            TabIndex        =   221
            Top             =   2640
            Width           =   1140
         End
         Begin VB.Label Label108 
            Caption         =   "Special Note"
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
            Left            =   7440
            TabIndex        =   216
            Top             =   3480
            Width           =   1590
         End
         Begin VB.Label Label1 
            Caption         =   "DOR"
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
            Left            =   9450
            TabIndex        =   212
            Top             =   1280
            Width           =   1005
         End
         Begin VB.Label Label23 
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
            Left            =   3720
            TabIndex        =   74
            Top             =   3480
            Width           =   1590
         End
         Begin VB.Label Label40 
            Caption         =   "Principal Warranty "
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
            TabIndex        =   73
            Top             =   3480
            Width           =   1590
         End
         Begin VB.Label Label1 
            Caption         =   "New PLN Code"
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
            Left            =   8400
            TabIndex        =   72
            Top             =   3960
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "New Eff Date"
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
            Left            =   9240
            TabIndex        =   71
            Top             =   3720
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label27 
            Caption         =   "Marital Status"
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
            Left            =   2520
            TabIndex        =   70
            Top             =   1815
            Width           =   1215
         End
         Begin VB.Label Label71 
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
            Height          =   270
            Index           =   0
            Left            =   2520
            TabIndex        =   69
            Top             =   2050
            Width           =   960
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   68
            Top             =   2140
            Width           =   420
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   7
            Left            =   9480
            TabIndex        =   67
            Top             =   1800
            Width           =   720
         End
         Begin VB.Label Label63 
            Caption         =   "TakeOver"
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
            Left            =   2520
            TabIndex        =   66
            Top             =   1580
            Width           =   825
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
            Left            =   5040
            TabIndex        =   65
            Top             =   750
            Width           =   1275
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
            Left            =   7200
            TabIndex        =   64
            Top             =   750
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Pol Year"
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
            Left            =   9450
            TabIndex        =   63
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "Sex "
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
            TabIndex        =   62
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Address"
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
            TabIndex        =   61
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label57 
            Caption         =   "City/Town"
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
            Left            =   2520
            TabIndex        =   60
            Top             =   1020
            Width           =   1020
         End
         Begin VB.Label Label1 
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
            TabIndex        =   59
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "Birthdate"
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
            TabIndex        =   58
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "PostCode"
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
            Left            =   120
            TabIndex        =   57
            Top             =   1015
            Width           =   1140
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   15
            Left            =   2520
            TabIndex        =   56
            Top             =   2340
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Plan Code"
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
            Left            =   7200
            TabIndex        =   55
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Ins Type"
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
            Left            =   5040
            TabIndex        =   54
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label14 
            Caption         =   "Agent"
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
            Left            =   9480
            TabIndex        =   53
            Top             =   2080
            Width           =   660
         End
         Begin VB.Label Label1 
            Caption         =   "Bordx. Date"
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
            Left            =   5040
            TabIndex        =   52
            Top             =   990
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Batch Num"
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
            Left            =   7200
            TabIndex        =   51
            Top             =   1000
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Date Join"
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
            Left            =   9450
            TabIndex        =   50
            Top             =   550
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Emp Num"
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
            Left            =   9450
            TabIndex        =   49
            Top             =   800
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "Grp Company"
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
            Left            =   5040
            TabIndex        =   48
            Top             =   1500
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "Dept. Name"
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
            Left            =   5040
            TabIndex        =   47
            Top             =   1800
            Width           =   1545
         End
         Begin VB.Label Label1 
            Caption         =   "Renewal "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   46
            Top             =   2640
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Ins mode"
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
            TabIndex        =   44
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label txtMBMBatchNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8280
            TabIndex        =   45
            Top             =   975
            Width           =   1080
         End
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   1605
         Left            =   -74880
         TabIndex        =   151
         Top             =   3900
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   2831
         SortKey         =   8
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
         NumItems        =   47
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Salutation"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "IC No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Birthdate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sex"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Adult/Child"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Relationship"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Status "
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Allergic"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Exclusion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Status"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "AnnualLimit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Available Limit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Plan"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "EffDate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "ExpDate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   " "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Premium"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Prev Plan"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "RB Amt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Payor Remarks"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Prev Insurance Com"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Clm Non Pay Fr"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Clm Non Pay To"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   25
            Text            =   "Payor Prem"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Payor Prem Net"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "Payor MCO Gross"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Payor MCO Net"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "RenewalDisc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Pol Disc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "PolDiscAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "RenDiscAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "LoadingPrem"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   " "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "KidneyAnn"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "KidneyBal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "CancerAnn"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "CancelBal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "HomeNurseAnn"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "HomeNurseBal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "PAYMBMNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "PayMBMID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "PayClientNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   44
            Text            =   "PayClientID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   45
            Text            =   "BordxDate"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   46
            Text            =   "BatchNo"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstAdjustmentHis 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   152
         Top             =   360
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   8916
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Batch No"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Adjust. Date"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Endt. Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Endt. Eff Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstBenefitDetails 
         Height          =   4155
         Left            =   -74940
         TabIndex        =   153
         Top             =   840
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   7329
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
            Text            =   "Code"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "BNF Desc"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Days"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM PDay"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Dis"
            Object.Width           =   1587
         EndProperty
      End
      Begin MSComctlLib.ListView lstNewBenefitDetails 
         Height          =   675
         Left            =   -70800
         TabIndex        =   154
         Top             =   3480
         Visible         =   0   'False
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   1191
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
            Text            =   "Code"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "BNF Desc"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Days"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM PDay"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Dis"
            Object.Width           =   1587
         EndProperty
      End
      Begin MSComctlLib.ListView lstCaseList 
         Height          =   4755
         Left            =   -74880
         TabIndex        =   155
         Top             =   600
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   8387
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mem No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOA"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DOD"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Claimability"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "FinalDiag"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Case Status"
            Object.Width           =   1905
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Hosp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Payor"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstACCHistory 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   156
         Top             =   360
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   8864
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mem No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DOA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DOR"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Claim No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DN Note"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "DN Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "DN Amt"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Receipt No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Receipt Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Receipt Amt"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "O/S Amt"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBMAcc 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   157
         Top             =   3060
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   4048
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MBMNumber"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cov ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Approved Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pay2HOS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Paid2HOS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PAY2MBM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Paid2MBM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PaidByMBM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Reimburse"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Excess"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Bordx Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Excess Paid"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBMHis 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   158
         Top             =   720
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8281
         SortKey         =   6
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mem No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CovID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Topup"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "TopUp MBM Number"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pol No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Eff Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Exp Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Data Status"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Pol Status"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Endt Bordx Date"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "MBM Reg Date"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Pricipal Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Principal DOB"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstCASHistory 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   159
         Top             =   720
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   8070
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mem No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOA"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DOD"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Claimability"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "FinalDiag"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Case Status"
            Object.Width           =   1905
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Hosp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Payor"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lstPreUndIllness 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   257
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8705
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin MSFlexGridLib.MSFlexGrid Fgrid 
         Height          =   2835
         Left            =   -74880
         TabIndex        =   258
         Top             =   540
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   5001
         _Version        =   393216
         Rows            =   11
         ScrollBars      =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lstLTLmt 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   259
         Top             =   480
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   8864
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
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "MBM Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Bal Ann Lmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Bal LTL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Current Claims"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total Claims"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Current Claim Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Total Claim Amt"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstSuppBenefitDetails 
         Height          =   4155
         Left            =   -69150
         TabIndex        =   260
         Top             =   840
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   7329
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
            Text            =   "Code"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "BNF Desc"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Days"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM PDay"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Dis"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Frame Frame11 
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -69360
         TabIndex        =   100
         Top             =   1740
         Width           =   4815
         Begin VB.TextBox txtReimbursement 
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
            Left            =   2280
            TabIndex        =   103
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtExcess 
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
            Left            =   2280
            TabIndex        =   102
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtExcessPaid 
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
            Left            =   2280
            TabIndex        =   101
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label92 
            Caption         =   "Excess Paid"
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
            TabIndex        =   106
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label78 
            Caption         =   "Total Excess"
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
            TabIndex        =   105
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label80 
            Caption         =   "Total Reimburse"
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
            TabIndex        =   104
            Top             =   240
            Width           =   1455
         End
      End
      Begin MSComDlg.CommonDialog CmnDlg 
         Left            =   -68520
         Top             =   4560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lstCASHistoryTwo 
         Height          =   915
         Left            =   -74880
         TabIndex        =   264
         Top             =   3240
         Visible         =   0   'False
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   1614
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mem No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOA"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DOD"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Claimability"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "FinalDiag"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Case Status"
            Object.Width           =   1905
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Hosp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Payor"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBMHisTwo 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   268
         Top             =   3120
         Visible         =   0   'False
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   4048
         SortKey         =   6
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mem No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CovID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Topup"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "TopUp MBM Number"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pol No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Eff Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Exp Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Data Status"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Pol Status"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Endt Bordx Date"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "MBM Reg Date"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Principal Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Principal DOB"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstTOClaims 
         Height          =   4995
         Left            =   -74880
         TabIndex        =   285
         Top             =   360
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   8811
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Eff Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Exp Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient Name"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Patient IC"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOA"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DOD"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Hosp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Diagnosis"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Dr. Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Bill Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Approved Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Excess Amt"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstWorksheet 
         Height          =   4875
         Left            =   -74920
         TabIndex        =   287
         Top             =   480
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   8599
         SortKey         =   13
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
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Version"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Mem No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Patient"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Rel"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DOA"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "DOD"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Claimability"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "FinalDiag"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Case Status"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Hosp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Payor"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "WS Reg Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Act Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "App Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Per Disability Lmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Major Med Fee"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Bordx No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Bordx Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblTotalCasesDOB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -73320
         TabIndex        =   273
         Top             =   2950
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblTotalCasesIC 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -73560
         TabIndex        =   272
         Top             =   400
         Width           =   975
      End
      Begin VB.Label lblTotalMBMDOB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -73320
         TabIndex        =   271
         Top             =   2820
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblTotalMBMIC 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -73440
         TabIndex        =   270
         Top             =   405
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "By Name and DOB"
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
         Left            =   -74850
         TabIndex        =   269
         Top             =   2850
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "By IC Number"
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
         Left            =   -74850
         TabIndex        =   267
         Top             =   400
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "By Name and DOB"
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
         Left            =   -74850
         TabIndex        =   266
         Top             =   2980
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "By IC Number"
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
         Left            =   -74850
         TabIndex        =   265
         Top             =   420
         Width           =   1020
      End
      Begin VB.Label Label105 
         Caption         =   "* Please click Supplementary to get benefit listing"
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
         Height          =   240
         Index           =   1
         Left            =   -69120
         TabIndex        =   262
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblCovName 
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
         Left            =   -67080
         TabIndex        =   261
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label7 
         Caption         =   "Double click to open the case"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   162
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label37 
         Caption         =   "Principal Bnf Listing"
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
         Left            =   -74880
         TabIndex        =   161
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label105 
         Caption         =   "Supplementary Bnf Listing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   -69120
         TabIndex        =   160
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame7 
      Height          =   525
      Left            =   120
      TabIndex        =   163
      Top             =   7800
      Width           =   11655
      Begin VB.CommandButton cmdplan 
         Caption         =   "&PLAN"
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
         Left            =   6480
         TabIndex        =   322
         Top             =   120
         Width           =   840
      End
      Begin VB.CommandButton cmdPrintHistory 
         Caption         =   "Print History"
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
         Left            =   7320
         TabIndex        =   274
         Top             =   120
         Width           =   1080
      End
      Begin VB.CommandButton CmdPrint2 
         Caption         =   "P&rint Report"
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
         Left            =   9600
         TabIndex        =   165
         Top             =   120
         Width           =   1080
      End
      Begin VB.CommandButton CmdMBMOthers 
         Caption         =   "&Others Details"
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
         Left            =   8400
         TabIndex        =   168
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton CmdClinical 
         Caption         =   "&Clinical"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8520
         TabIndex        =   169
         Top             =   150
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtcount 
         Height          =   285
         Left            =   360
         TabIndex        =   166
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print MBM"
         Height          =   315
         Left            =   8760
         TabIndex        =   164
         Top             =   150
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
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
         Height          =   315
         Left            =   10680
         TabIndex        =   167
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   11655
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
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
         Left            =   3240
         TabIndex        =   5
         Top             =   225
         Width           =   690
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
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdSearchMBM 
         Caption         =   "&Search"
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
         Left            =   5280
         TabIndex        =   4
         Top             =   225
         Width           =   675
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Pre"
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
         Left            =   3920
         TabIndex        =   3
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
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
         Left            =   4610
         TabIndex        =   2
         Top             =   225
         Width           =   690
      End
      Begin MSMask.MaskEdBox txtSuspendDate 
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
         TabIndex        =   313
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label25 
         Caption         =   "Suspend Date"
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
         Index           =   1
         Left            =   7440
         TabIndex        =   314
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label75 
         Caption         =   "Hlt Type"
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
         Left            =   10560
         TabIndex        =   218
         Top             =   600
         Width           =   675
      End
      Begin VB.Label cboHLTCode 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   11280
         TabIndex        =   217
         Top             =   600
         Width           =   375
      End
      Begin VB.Label txtMBMStatus 
         Alignment       =   2  'Center
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
         Left            =   6960
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblEXPStatus 
         Alignment       =   2  'Center
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
         Left            =   11040
         TabIndex        =   11
         Top             =   240
         Width           =   375
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
         Height          =   330
         Left            =   10080
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
      Begin VB.Label txtDataStatus 
         Alignment       =   2  'Center
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
         Left            =   10560
         TabIndex        =   9
         Top             =   600
         Width           =   375
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
         Height          =   330
         Left            =   9600
         TabIndex        =   8
         Top             =   600
         Width           =   915
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
         TabIndex        =   7
         Top             =   210
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
         Height          =   330
         Index           =   0
         Left            =   6120
         TabIndex        =   6
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.PictureBox crptSummary 
      Height          =   480
      Left            =   480
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   317
      Top             =   360
      Width           =   1200
   End
   Begin VB.Frame Frame10 
      Height          =   1155
      Left            =   120
      TabIndex        =   170
      Top             =   790
      Width           =   11655
      Begin VB.TextBox txtMBMName 
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
         Left            =   2680
         MaxLength       =   50
         TabIndex        =   175
         Top             =   187
         Width           =   3075
      End
      Begin VB.ComboBox cboMBMSalutation 
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
         Left            =   880
         TabIndex        =   174
         Top             =   195
         Width           =   795
      End
      Begin VB.TextBox txtMBMPolNo 
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
         Left            =   2680
         MaxLength       =   25
         TabIndex        =   172
         Top             =   480
         Width           =   3075
      End
      Begin VB.ComboBox cboPAYCode 
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
         Left            =   880
         TabIndex        =   171
         Top             =   480
         Width           =   795
      End
      Begin VB.TextBox txtMBMIcBcPP 
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
         Left            =   2680
         MaxLength       =   30
         TabIndex        =   173
         Top             =   780
         Width           =   3075
      End
      Begin VB.Label Label4 
         Caption         =   "Per Disability"
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
         Left            =   5880
         TabIndex        =   256
         Top             =   555
         Width           =   965
      End
      Begin VB.Label lblPerDisability 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   6840
         TabIndex        =   255
         Top             =   525
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Ann Bal Lmt"
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
         TabIndex        =   242
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblLimitType 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   8880
         TabIndex        =   219
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label LblHnS 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   8325
         TabIndex        =   184
         Top             =   525
         Width           =   480
      End
      Begin VB.Label txtAvailableAnnualLimit 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   6840
         TabIndex        =   183
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label txtAnnualLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   182
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label5 
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
         Height          =   210
         Left            =   7080
         TabIndex        =   181
         Top             =   300
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label74 
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
         Left            =   120
         TabIndex        =   180
         Top             =   525
         Width           =   915
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
         Left            =   1800
         TabIndex        =   179
         Top             =   525
         Width           =   900
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
         Left            =   1800
         TabIndex        =   178
         Top             =   240
         Width           =   1020
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
         Left            =   1800
         TabIndex        =   177
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "Salutation"
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
         TabIndex        =   176
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "H&&C"
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
         Left            =   7920
         TabIndex        =   185
         Top             =   555
         Width           =   375
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Member Enquiry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   213
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmNewMBMEnquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bExit As Boolean
Public bDataStatus As Boolean
Public ReportCriteria As String
Dim ScanCode As String
Dim txtSRVPath As String
Dim objword As Word.Application
Dim objdoc As Word.Document

Dim ADOParameterMBM As ADODB.Parameter
Dim ADOCommandMBM As New ADODB.Command
Dim ADOrstMBM As New ADODB.Recordset

Dim ADOParameterMBMCov As ADODB.Parameter
Dim ADOCommandMBMCov As New ADODB.Command
Dim ADOrstMBMCov As New ADODB.Recordset
Dim tmpCISPln As String
Dim tmpCISIns As String
Dim tmpCisEffDate As String
Dim tmpCisExpDate As String
Dim HisCisEffDate As String
Dim HISCISExpDate As String
Dim HISCisPLN As String
Dim HISCiSIns As String
Dim SptLmtInd As String
Dim tmpPolicyWording As String
Dim cntCasesHisIC As Integer
Dim cntCasesHisDOB As Integer
Dim cntMBMHisIC As Integer
Dim cntMBMHisDOB As Integer
Dim tmpMBMFullCover As String
Dim tmpAnnLmtIND As String
Dim tmpVIP As Boolean
Dim tmpMBMGLIssuance As String
Dim txtMBMAdultChild As String

Dim Ciamamt As String
Dim AnnlLmt As String
Dim Bal As String

Dim PictureFileName As String
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

Private m_blnDirection(1 To 9) As Boolean


Private Sub CmdClinical_Click()
    frmMBMClinical.Source = 3
    frmMBMClinical.Show
    frmMBMClinical.Frame1.Enabled = False
End Sub


Private Sub CmdMBMOthers_Click()
    frmNewMBMOthers.Source = 2
    frmNewMBMOthers.Show
End Sub

Private Sub cmdplan_Click()
Load frmplandetails
Call showMBMPLAN
frmplandetails.Show vbModal
End Sub


Private Function showMBMPLAN()
Dim sql1 As String
Dim ADOrstPLN As New ADODB.Recordset
Dim Statuscat As String
sql1 = "Exec SearchPlandetails '" & Trim(txtMBMNumber) & "'"
    Set ADOrstPLN = New ADODB.Recordset
    ADOrstPLN.Open sql1, medixmasterconn, adOpenKeyset, adLockOptimistic
        If Not ADOrstPLN.EOF Then
            frmplandetails.txtgrpcompany = IIf(Len(Trim(ADOrstPLN!MBMGroupCompany)), Trim(ADOrstPLN!MBMGroupCompany), "")
            frmplandetails.txtAnnualLimit = Format(IIf(Len(Trim(ADOrstPLN!AnnLimit)), Trim(ADOrstPLN!AnnLimit), ""), "#,##0.00")
            frmplandetails.txtmajormedlimit = Format(IIf(Len(Trim(ADOrstPLN!MajorMedicalLmt)), Trim(ADOrstPLN!MajorMedicalLmt), "0"), "#,##0.00")
            frmplandetails.txtperdsilimit = Format(IIf(Len(Trim(ADOrstPLN!DisabilityLmt)), Trim(ADOrstPLN!DisabilityLmt), "0"), "#,##0.00")
            frmplandetails.txtcopayind = IIf(Len(Trim(ADOrstPLN!PLNCoPayIND)), Trim(ADOrstPLN!PLNCoPayIND), "N")
            frmplandetails.txtcopayper = IIf(Len(Trim(ADOrstPLN!PLNCoPayPerc)), Trim(ADOrstPLN!PLNCoPayPerc), "0")
            frmplandetails.txtEffDate = IIf(Len(Trim(ADOrstPLN!MBMPayorEffDate)), Trim(ADOrstPLN!MBMPayorEffDate), "")
            frmplandetails.txtExpDate = IIf(Len(Trim(ADOrstPLN!MBMPayorEXPDate)), Trim(ADOrstPLN!MBMPayorEXPDate), "")
            frmplandetails.txtvipind = IIf(Len(Trim(ADOrstPLN!VIPInd)), Trim(ADOrstPLN!VIPInd), "N")
            'frmplandetails.txtlimittype = IIf(Len(Trim(ADOrstPLN!AnnLmtInd)), Trim(ADOrstPLN!AnnLmtInd), "N")
            Statuscat = IIf(Len(Trim(ADOrstPLN!SuppLimitStatus)), Trim(ADOrstPLN!SuppLimitStatus), "")
            'frmplandetails.txtlimitcat
            If Statuscat = "A" Then
                frmplandetails.txtlimitcat = "A - Shared Limit for family"
            ElseIf Statuscat = "B" Then
                frmplandetails.txtlimitcat = "B - Shared Limit for dependent"
            ElseIf Statuscat = "C" Then
                frmplandetails.txtlimitcat = "C - Separate Limit for family"
            ElseIf Statuscat = "D" Then
                frmplandetails.txtlimitcat = "D - Shared Limit for family"
            End If
            frmplandetails.txtLGNo = IIf(Len(Trim(ADOrstPLN!PLNLGPrivate)), Trim(ADOrstPLN!PLNLGPrivate), "")
            frmplandetails.LblMBMNo = IIf(Len(Trim(ADOrstPLN!MBMNumber)), Trim(ADOrstPLN!MBMNumber), "")
            
        ADOrstPLN.Close
 Set ADOrstPLN = Nothing
 End If
End Function


Private Sub cmdPrintHistory_Click()
    Screen.MousePointer = vbHourglass
        
    If txtMBMNumber <> "" Then
                
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    '    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
            Call funcword2
        
     '   medixmasterconn.Close
     '   Set medixmasterconn = Nothing
     '   medixmasterconnMaint.Close
      '  Set medixmasterconnMaint = Nothing
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
              
    Screen.MousePointer = Default
    

End Sub

Private Sub funcword2()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim i As Integer
Dim ii As Integer
Dim RecordSelected As Integer
Dim CaseRec As New ADODB.Recordset
Dim strsqlCaseRec As String
Dim WSHistoryRec As New ADODB.Recordset
Dim strsqlWSHistoryRec As String

Dim TmpREc1 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpName As String
Dim strAppend As String
Dim tmpDOb As String

Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim TmpREc2 As New ADODB.Recordset

Dim cmd3 As New ADODB.Command
Dim Prm5 As New ADODB.Parameter
Dim prm6 As New ADODB.Parameter
Dim TmpREc3 As New ADODB.Recordset

Dim cmd4 As New ADODB.Command
Dim prm7 As New ADODB.Parameter
Dim prm8 As New ADODB.Parameter
Dim TmpREc4 As New ADODB.Recordset

Dim strsql As String
Dim CAS As New ADODB.Recordset

Dim strsql2 As String
Dim CAS2 As New ADODB.Recordset

Dim strsql3 As String
Dim CAS3 As New ADODB.Recordset

Dim strsql4 As String
Dim CAS4 As New ADODB.Recordset


    If txtMBMNumber <> "" Then
        i = 0
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
        End If
        
        On Error GoTo HANDLE_ERROR
        objExcel.Interactive = False
        
        
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
        
            objExcel.WindowState = xlMaximized
    
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Case\PrintHistory1.xlt")
        
            Set objWorksheet = objWorkbook.Sheets(1)
            'Turn off the alerts, otherwise user will have to confirm my actions.
            objExcel.DisplayAlerts = False
        
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
                            
                
                    'start-----Principal by IC'
                    If Len(Trim(txtMBMIcBcPP)) > 5 Then
                        tmpName = Trim(Replace(txtMBMName, "'", "''"))
                        tmpDOb = Format(txtMBMBirthDate, "yyyy/mm/dd")
                        
                        strAppend = " And HISMaintenance.dbo.MBMCrossReference.MBMIcBcPp = '" & Trim(txtMBMIcBcPP) & "'"
                        Set cmd.ActiveConnection = medixmasterconn
                        cmd.CommandText = "SearchMBMHistory06"
                        cmd.CommandType = adCmdStoredProc
                        cmd.CommandTimeout = 300
                        Set Prm1 = cmd.CreateParameter("MBMType", adVarChar, adParamInput, 10, "Princ")
                        cmd.Parameters.Append Prm1
                        Set Prm2 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                        cmd.Parameters.Append Prm2
                    
                        TmpREc1.CursorLocation = adUseClient
                        TmpREc1.CursorType = adOpenForwardOnly
                        TmpREc1.CacheSize = 100
                        Set TmpREc1 = cmd.Execute
                        
                        Do While Not TmpREc1.EOF
                            i = CDbl(i) + 1
                            objWorksheet.Cells(6 + CDbl(i), "A") = TmpREc1!MBMNumber
                            objWorksheet.Cells(6 + CDbl(i), "B") = IIf(Len(TmpREc1!MBMCOVID), TmpREc1!MBMCOVID, "")
                            objWorksheet.Cells(6 + CDbl(i), "C") = IIf(Len(TmpREc1!MBMName), TmpREc1!MBMName, "")
                            objWorksheet.Cells(6 + CDbl(i), "D") = IIf(Len(TmpREc1!MBMDOB), TmpREc1!MBMDOB, "")
                            objWorksheet.Cells(6 + CDbl(i), "E") = IIf(Len(TmpREc1!MBMPolicyNo), TmpREc1!MBMPolicyNo, "")
                            objWorksheet.Cells(6 + CDbl(i), "F") = IIf(TmpREc1!MBMPayorEffDate <> "", TmpREc1!MBMPayorEffDate, "")
                            objWorksheet.Cells(6 + CDbl(i), "G") = IIf(TmpREc1!MBMPayorEXPDate <> "", TmpREc1!MBMPayorEXPDate, "")
                            objWorksheet.Cells(6 + CDbl(i), "H") = IIf(Len(TmpREc1!MBMStatus), TmpREc1!MBMStatus, "")
                            objWorksheet.Cells(6 + CDbl(i), "I") = IIf(Len(TmpREc1!MBMBordxDate), TmpREc1!MBMBordxDate, "")
                            objWorksheet.Cells(6 + CDbl(i), "J") = IIf(Len(TmpREc1!ProductName), TmpREc1!ProductName, "")
                            TmpREc1.MoveNext
                        Loop
                            TmpREc1.Close
                            Set TmpREc1 = Nothing
                        'End-----Principal by IC'
                        
                            'Start----Supp by IC
                            strAppend = ""
                            strAppend = " And dbo.MBMCoveredPersons.MBMCICBCPPNo = '" & Trim(txtMBMIcBcPP) & "'"
                    
                            Set cmd2.ActiveConnection = medixmasterconn
                            cmd2.CommandText = "SearchMBMHistory06"
                            cmd2.CommandType = adCmdStoredProc
                            cmd2.CommandTimeout = 300
                            Set Prm3 = cmd2.CreateParameter("MBMType", adVarChar, adParamInput, 10, "Supp")
                            cmd2.Parameters.Append Prm3
                            Set prm4 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                            cmd2.Parameters.Append prm4
                        
                            TmpREc2.CursorLocation = adUseClient
                            TmpREc2.CursorType = adOpenForwardOnly
                            TmpREc2.CacheSize = 100
                            Set TmpREc2 = cmd2.Execute
                            
                            Do While Not TmpREc2.EOF
                                i = CDbl(i) + 1
                                With TmpREc2
                                    objWorksheet.Cells(6 + CDbl(i), "A") = !MBMNumber
                                    objWorksheet.Cells(6 + CDbl(i), "B") = IIf(Len(!MBMCOVID), !MBMCOVID, "")
                                    objWorksheet.Cells(6 + CDbl(i), "C") = IIf(Len(!MBMName), !MBMName, "")
                                    objWorksheet.Cells(6 + CDbl(i), "D") = IIf(Len(!MBMDOB), !MBMDOB, "")
                                    objWorksheet.Cells(6 + CDbl(i), "E") = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                                    objWorksheet.Cells(6 + CDbl(i), "F") = IIf(!MBMPayorEffDate <> "", !MBMPayorEffDate, "")
                                    objWorksheet.Cells(6 + CDbl(i), "G") = IIf(!MBMPayorEXPDate <> "", !MBMPayorEXPDate, "")
                                    objWorksheet.Cells(6 + CDbl(i), "H") = IIf(Len(!MBMStatus), !MBMStatus, "")
                                    objWorksheet.Cells(6 + CDbl(i), "I") = IIf(Len(!MBMBordxDate), !MBMBordxDate, "")
                                End With
                                TmpREc2.MoveNext
                            Loop
                            TmpREc2.Close
                            Set TmpREc2 = Nothing
                            'End----Supp by IC
                        End If
                        
                        'Start----Principal by Name & DOB
                        If Len(Trim(txtMBMName)) <> "" And txtMBMBirthDate <> "__/__/____" Then
                            i = 0
                            strAppend = ""
                            tmpName = Trim(Replace(txtMBMName, "'", "''"))
                            tmpDOb = Format(txtMBMBirthDate, "yyyy/mm/dd")
                            strAppend = " And HISMaintenance.dbo.MBMCrossReference.MBMName = '" & tmpName
                            strAppend = strAppend & "' And HISMaintenance.dbo.MBMCrossReference.MBMDOB = '" & Format(tmpDOb, "yyyy/mm/dd") & "'"
                            
                            Set cmd3.ActiveConnection = medixmasterconn
                            cmd3.CommandText = "SearchMBMHistory06"
                            cmd3.CommandType = adCmdStoredProc
                            cmd3.CommandTimeout = 300
                            Set Prm5 = cmd3.CreateParameter("MBMType", adVarChar, adParamInput, 10, "Princ")
                            cmd3.Parameters.Append Prm5
                            Set prm6 = cmd3.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                            cmd3.Parameters.Append prm6
                        
                            TmpREc3.CursorLocation = adUseClient
                            TmpREc3.CursorType = adOpenForwardOnly
                            TmpREc3.CacheSize = 100
                            Set TmpREc3 = cmd3.Execute
                                Do While Not TmpREc3.EOF
                                    i = CDbl(i) + 1
                                    With TmpREc3
                                        objWorksheet.Cells(28 + CDbl(i), "A") = !MBMNumber
                                        objWorksheet.Cells(28 + CDbl(i), "B") = IIf(Len(!MBMCOVID), !MBMCOVID, "")
                                        objWorksheet.Cells(28 + CDbl(i), "C") = IIf(Len(!MBMName), !MBMName, "")
                                        objWorksheet.Cells(28 + CDbl(i), "D") = IIf(Len(!MBMDOB), !MBMDOB, "")
                                        objWorksheet.Cells(28 + CDbl(i), "E") = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                                        objWorksheet.Cells(28 + CDbl(i), "F") = IIf(!MBMPayorEffDate <> "", !MBMPayorEffDate, "")
                                        objWorksheet.Cells(28 + CDbl(i), "G") = IIf(!MBMPayorEXPDate <> "", !MBMPayorEXPDate, "")
                                        objWorksheet.Cells(28 + CDbl(i), "H") = IIf(Len(!MBMStatus), !MBMStatus, "")
                                        objWorksheet.Cells(28 + CDbl(i), "I") = IIf(Len(!MBMBordxDate), !MBMBordxDate, "")
                                    End With
                                    TmpREc3.MoveNext
                                Loop
                                TmpREc3.Close
                                Set TmpREc3 = Nothing
                            'End----Principal by Name & DOB
                            
                            'Start----Supp by Name & DOB
                            strAppend = ""
                            strAppend = " And dbo.MBMCoveredPersons.MBMCName = '" & tmpName
                            strAppend = strAppend & "' And dbo.MBMCoveredPersons.MBMCDOB= '" & Format(tmpDOb, "yyyy/mm/dd") & "'"
                            Set cmd4.ActiveConnection = medixmasterconn
                            cmd4.CommandText = "SearchMBMHistory06"
                            cmd4.CommandType = adCmdStoredProc
                            cmd4.CommandTimeout = 300
                            Set prm7 = cmd4.CreateParameter("MBMType", adVarChar, adParamInput, 10, "Supp")
                            cmd4.Parameters.Append prm7
                            Set prm8 = cmd4.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                            cmd4.Parameters.Append prm8
                        
                            TmpREc4.CursorLocation = adUseClient
                            TmpREc4.CursorType = adOpenForwardOnly
                            TmpREc4.CacheSize = 100
                            Set TmpREc4 = cmd4.Execute
                                Do While Not TmpREc4.EOF
                                    i = CDbl(i) + 1
                                    With TmpREc4
                                        objWorksheet.Cells(28 + CDbl(i), "A") = !MBMNumber
                                        objWorksheet.Cells(28 + CDbl(i), "B") = IIf(Len(!MBMCOVID), !MBMCOVID, "")
                                        objWorksheet.Cells(28 + CDbl(i), "C") = IIf(Len(!MBMName), !MBMName, "")
                                        objWorksheet.Cells(28 + CDbl(i), "D") = IIf(Len(!MBMDOB), !MBMDOB, "")
                                        objWorksheet.Cells(28 + CDbl(i), "E") = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                                        objWorksheet.Cells(28 + CDbl(i), "F") = IIf(!MBMPayorEffDate <> "", !MBMPayorEffDate, "")
                                        objWorksheet.Cells(28 + CDbl(i), "G") = IIf(!MBMPayorEXPDate <> "", !MBMPayorEXPDate, "")
                                        objWorksheet.Cells(28 + CDbl(i), "H") = IIf(Len(!MBMStatus), !MBMStatus, "")
                                        objWorksheet.Cells(28 + CDbl(i), "I") = IIf(Len(!MBMBordxDate), !MBMBordxDate, "")
                                    End With
                                    TmpREc4.MoveNext
                                Loop
                                TmpREc4.Close
                                Set TmpREc4 = Nothing
                        'End----Supp by Name & DOB
                    End If
                    
                    Set objWorksheet = objWorkbook.Sheets(2)
                    If Len(Trim(txtMBMIcBcPP)) > 5 Then
                        'Start---Principal Cases by IC
                        i = 0
                        strsql = "Exec SearchCASHistory '" & Trim(txtMBMIcBcPP) & "'"
                        Set CAS = New ADODB.Recordset
                        CAS.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                    
                        Screen.MousePointer = vbHourglass
                            
                        Do While Not CAS.EOF
                            i = CDbl(i) + 1
                            With CAS
                                objWorksheet.Cells(5 + CDbl(i), "A") = !CasNumber
                                objWorksheet.Cells(5 + CDbl(i), "B") = !MBMNumber
                                If CAS!MBMPatientCovID = "00" Then
                                    objWorksheet.Cells(5 + CDbl(i), "C") = IIf(Len(Trim(!MBMName)), Trim(!MBMName), "")
                                    objWorksheet.Cells(5 + CDbl(i), "D") = "P"
                                Else
                                    objWorksheet.Cells(5 + CDbl(i), "C") = IIf(Len(Trim(!MBMCName)), Trim(!MBMCName), "")
                                    objWorksheet.Cells(5 + CDbl(i), "D") = IIf(Len(!MBMCRELCode), !MBMCRELCode, "")
                                End If
                                
                                If Len(Trim(!CASDateAdmitted)) Then
                                    objWorksheet.Cells(5 + CDbl(i), "E") = !CASDateAdmitted
                                End If
                                If Len(Trim(!CASDischargeDate)) Then
                                    objWorksheet.Cells(5 + CDbl(i), "F") = !CASDischargeDate
                                End If
                                objWorksheet.Cells(5 + CDbl(i), "G") = IIf(Len(CAS!DIAFinalDiag1), CAS!DIAFinalDiag1, "")
                                objWorksheet.Cells(5 + CDbl(i), "H") = IIf(Len(CAS!CASClaimability), CAS!CASClaimability, "")
                                objWorksheet.Cells(5 + CDbl(i), "I") = IIf(Len(CAS!HosName), CAS!HosName, "")
                            End With
                            CAS.MoveNext
                            
                        Loop
                        
                        CAS.Close
                        Set CAS = Nothing
                        
                        'End---Cases by IC
    
    
                        'Start---Supp Cases by IC
                        
                        strsql2 = "Exec SearchCasHistorySupp '" & Trim(txtMBMIcBcPP) & "'"
                        Set CAS2 = New ADODB.Recordset
                        CAS2.Open strsql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                                            
                        Do While Not CAS2.EOF
                            i = CDbl(i) + 1
                            With CAS2
                                objWorksheet.Cells(5 + CDbl(i), "A") = !CasNumber
                                objWorksheet.Cells(5 + CDbl(i), "B") = !MBMNumber
                                If CAS!MBMPatientCovID = "00" Then
                                    objWorksheet.Cells(5 + CDbl(i), "C") = IIf(Len(Trim(!MBMName)), Trim(!MBMName), "")
                                    objWorksheet.Cells(5 + CDbl(i), "D") = "P"
                                Else
                                    objWorksheet.Cells(5 + CDbl(i), "C") = IIf(Len(Trim(!MBMCName)), Trim(!MBMCName), "")
                                    objWorksheet.Cells(5 + CDbl(i), "D") = IIf(Len(!MBMCRELCode), !MBMCRELCode, "")
                                End If
                                
                                If Len(Trim(!CASDateAdmitted)) Then
                                    objWorksheet.Cells(5 + CDbl(i), "E") = !CASDateAdmitted
                                End If
                                If Len(Trim(!CASDischargeDate)) Then
                                    objWorksheet.Cells(5 + CDbl(i), "F") = !CASDischargeDate
                                End If
                                objWorksheet.Cells(5 + CDbl(i), "G") = IIf(Len(CAS!DIAFinalDiag1), CAS!DIAFinalDiag1, "")
                                objWorksheet.Cells(5 + CDbl(i), "H") = IIf(Len(CAS!CASClaimability), CAS!CASClaimability, "")
                                objWorksheet.Cells(5 + CDbl(i), "I") = IIf(Len(CAS!HosName), CAS!HosName, "")
                            End With
                            CAS2.MoveNext
                            
                        Loop
                        
                        CAS2.Close
                        Set CAS2 = Nothing
                        'End---Supp Cases by IC
                    End If
                    
                    
                    If Len(Trim(txtMBMName)) <> "" And txtMBMBirthDate <> "__/__/____" Then
                        'Start---Principal Cases by Name and DOB
                        tmpName = Trim(Replace(txtMBMName, "'", "''"))
                        tmpDOb = Format(txtMBMBirthDate, "yyyy/mm/dd")
                        
                        i = 0
                        strsql3 = "Exec SearchCASHistoryTwo '" & Trim(tmpName) & "', '" & Format(tmpDOb, "yyyy/mm/dd") & "'"
                        Set CAS3 = New ADODB.Recordset
                        CAS3.Open strsql3, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                    
                            
                        Do While Not CAS3.EOF
                            i = CDbl(i) + 1
                            With CAS3
                                objWorksheet.Cells(28 + CDbl(i), "A") = !CasNumber
                                objWorksheet.Cells(28 + CDbl(i), "B") = !MBMNumber
                                If !MBMPatientCovID = "00" Then
                                    objWorksheet.Cells(28 + CDbl(i), "C") = IIf(Len(Trim(!MBMName)), Trim(!MBMName), "")
                                    objWorksheet.Cells(28 + CDbl(i), "D") = "P"
                                Else
                                    objWorksheet.Cells(28 + CDbl(i), "C") = IIf(Len(Trim(!MBMCName)), Trim(!MBMCName), "")
                                    objWorksheet.Cells(28 + CDbl(i), "D") = IIf(Len(!MBMCRELCode), !MBMCRELCode, "")
                                End If
                                
                                If Len(Trim(!CASDateAdmitted)) Then
                                    objWorksheet.Cells(28 + CDbl(i), "E") = !CASDateAdmitted
                                End If
                                If Len(Trim(!CASDischargeDate)) Then
                                    objWorksheet.Cells(28 + CDbl(i), "F") = !CASDischargeDate
                                End If
                                objWorksheet.Cells(28 + CDbl(i), "G") = IIf(Len(!DIAFinalDiag1), !DIAFinalDiag1, "")
                                objWorksheet.Cells(28 + CDbl(i), "H") = IIf(Len(!CASClaimability), !CASClaimability, "")
                                objWorksheet.Cells(28 + CDbl(i), "I") = IIf(Len(!HosName), !HosName, "")
                            End With
                            CAS3.MoveNext
                            
                        Loop
                        CAS3.Close
                        Set CAS3 = Nothing
    
                        'End---Principal Cases by Name and DOB
                
                
                        'Start---Supp Cases by Name and DOB
                        
                        strsql4 = "Exec SearchCasHistorySuppTwo '" & Trim(tmpName) & "', '" & Format(tmpDOb, "yyyy/mm/dd") & "'"
                        Set CAS4 = New ADODB.Recordset
                        CAS4.Open strsql4, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                                            
                            Do While Not CAS4.EOF
                                With CAS4
                                    i = CDbl(i) + 1
                                    objWorksheet.Cells(28 + CDbl(i), "A") = !CasNumber
                                    objWorksheet.Cells(28 + CDbl(i), "B") = !MBMNumber
                                    If !MBMPatientCovID = "00" Then
                                        objWorksheet.Cells(28 + CDbl(i), "C") = IIf(Len(Trim(!MBMName)), Trim(!MBMName), "")
                                        objWorksheet.Cells(28 + CDbl(i), "D") = "P"
                                    Else
                                        objWorksheet.Cells(28 + CDbl(i), "C") = IIf(Len(Trim(!MBMCName)), Trim(!MBMCName), "")
                                        objWorksheet.Cells(28 + CDbl(i), "D") = IIf(Len(!MBMCRELCode), !MBMCRELCode, "")
                                    End If
                                    
                                    If Len(Trim(!CASDateAdmitted)) Then
                                        objWorksheet.Cells(28 + CDbl(i), "E") = !CASDateAdmitted
                                    End If
                                    If Len(Trim(!CASDischargeDate)) Then
                                        objWorksheet.Cells(28 + CDbl(i), "F") = !CASDischargeDate
                                    End If
                                    objWorksheet.Cells(28 + CDbl(i), "G") = IIf(Len(!DIAFinalDiag1), !DIAFinalDiag1, "")
                                    objWorksheet.Cells(28 + CDbl(i), "H") = IIf(Len(!CASClaimability), !CASClaimability, "")
                                    objWorksheet.Cells(28 + CDbl(i), "I") = IIf(Len(!HosName), !HosName, "")
                                End With
                            CAS4.MoveNext
                            
                        Loop
                        CAS4.Close
                        Set CAS4 = Nothing
        
                    'End---Supp Cases by Name and DOB
                    End If
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            Exit Sub
    End If
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

'Private Sub cmdPrint_Click()
'
'    Screen.MousePointer = vbHourglass
'
'    If txtMBMNumber <> "" Then
'
'        crptSummary.ReportFileName = crdPath & "\Report\RptMbmEnquiry1.rpt"
'        crptSummary.Destination = crptToWindow
'        crptSummary.WindowState = crptMaximized
'        crptSummary.SelectionFormula = "{VIEW_MBM_ENQUIRY.MBMNumber} = '" & Trim(txtMBMNumber) & "'"
'        crptSummary.ParameterFields(0) = "MemberNo;" & Trim(txtMBMNumber) & ";true"
'        crptSummary.Action = 1
'        crptSummary.Reset
'    Else
 '       MsgBox "Report cannot be printed without any record.", vbInformation
'    End If
 '
'    Screen.MousePointer = Default
'End Sub

Private Sub cmdPrint2_Click()
    
    Screen.MousePointer = vbHourglass
        
    If txtMBMNumber <> "" Then
        
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
     '   medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
     '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Call funcword
     '   medixmasterconn.Close
     '   Set medixmasterconn = Nothing
     '   medixmasterconnMaint.Close
     '   Set medixmasterconnMaint = Nothing
     '   medixmasterconnWS.Close
     ''   Set medixmasterconnWS = Nothing
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
              
    Screen.MousePointer = Default
End Sub

Private Sub CmdTotal_Click()
Dim Acc As New ADODB.Recordset
Dim strsql As String
Dim masterlist As ListItem
Dim strsqlMBM As String
Dim MBM As New ADODB.Recordset
Dim strsqlMBMArc As String
Dim MBMArc As New ADODB.Recordset
Dim totalApprove As Variant
Dim totalPayable2Hos As Variant
Dim totalPaid2Hos As Variant
Dim totalPayable2Mbm As Variant
Dim totalPaidByMbm As Variant
Dim totalPaid2MBM As Variant
Dim TotalExcess As Variant
Dim TotalReimburse As Variant
Dim MBMTotalAnnualLmt As Variant
Dim MBMTotalAvailLmt As Variant

    If lstMBMAcc.ListItems.Count <> 0 Then
        strsqlMBM = " Select MBMNumber, MBMMemberType, MBMCAnnualLimit, " & _
                    " MBMCAvailableLimit " & _
                    " From View_MBM_MemberType where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    If MBM.recordCount Then
        If MBM!MBMMemberType = "P" Then
            txtAnnualLmt = txtAnnualLimit
            txtAvailableAnnualLmt = txtAvailableAnnualLimit
        Else
            Do Until MBM.EOF
                MBMTotalAnnualLmt = Format(MBMTotalAnnualLmt + MBM!MBMCAnnualLimit, "#,##0.00")
                MBMTotalAvailLmt = Format(MBMTotalAvailLmt + MBM!MBMCAvailableLimit, "#,##0.00")
            MBM.MoveNext
            Loop
                txtAnnualLmt = Format(MBMTotalAnnualLmt + txtAnnualLimit, "#,##0.00")
                txtAvailableAnnualLmt = Format(MBMTotalAvailLmt + txtAvailableAnnualLimit, "#,##0.00")
        End If
        MBM.Close
        Set MBM = Nothing
    Else
    
        strsqlMBMArc = " Select MBMNumber, MBMMemberType, MBMCAnnualLimit, " & _
                       " MBMCAvailableLimit, From View_MBM_MemTypeArc where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBMArc.Open strsqlMBMArc, medixmasterconn, adOpenKeyset, adLockOptimistic

        If MBMArc.recordCount Then
            If MBMArc!MBMMemberType = "P" Then
                txtAnnualLmt = Format(txtAnnualLimit, "#,##0.00")
                txtAvailableAnnualLmt = Format(txtAvailableAnnualLimit, "#,##0.00")
            Else
                Do Until MBM.EOF
                    MBMTotalAnnualLmt = Format(MBMTotalAnnualLmt + MBMArc!MBMCAnnualLimit, "#,##0.00")
                    MBMTotalAvailLmt = Format(MBMTotalAvailLmt + MBMArc!MBMCAvailableLimit, "#,##0.00")
                MBMArc.MoveNext
                Loop
                    txtAnnualLmt = Format(MBMTotalAnnualLmt + txtAnnualLimit, "#,##0.00")
                    txtAvailableAnnualLmt = Format(MBMTotalAvailLmt + txtAvailableAnnualLimit, "#,##0.00")
            End If
        End If
        MBMArc.Close
        Set MBMArc = Nothing
    End If

        strsql = " select totalApprove, totalPayable2Hos, " & _
                " totalPaid2Hos, totalPayable2Mbm, totalPaidByMbm, MBMPatientCOVID, " & _
                " totalPaid2MBM , totalExcess , totalReimburse from view_MBMTotal_summary where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Acc.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If Not Acc.EOF Then
            Do Until Acc.EOF
            With Acc
                totalApprove = Format(totalApprove + !totalApprove, "#,##0.00")
                totalPayable2Hos = Format(totalPayable2Hos + !totalPayable2Hos, "#,##0.00")
                If Trim(!totalPaid2Hos) <> "" Then
                    totalPaid2Hos = Format(totalPaid2Hos + !totalPaid2Hos, "#,##0.00")
                End If
                totalPayable2Mbm = Format(totalPayable2Mbm + !totalPayable2Mbm, "#,##0.00")
                If Trim(!totalPaid2MBM) <> "" Then
                    totalPaid2MBM = Format(totalPaid2MBM + totalPaid2MBM, "#,##0.00")
                End If
                If Trim(!totalPaidByMbm) <> "" Then
                    totalPaidByMbm = Format(totalPaidByMbm + totalPaidByMbm, "#,##0.00")
                End If
                'txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
                TotalReimburse = Format(TotalReimburse + TotalReimburse, "#,##0.00")
                TotalExcess = Format(TotalExcess + !TotalExcess, "#,##0.00")
            End With
            Acc.MoveNext
            Loop
            
            txtApprovedAmt = Format(totalApprove, "#,##0.00")
            
            txtPayable2Hos = totalPayable2Hos
            If IsNumeric(txtPayable2Hos) = False Then
                txtPayable2Hos = "0.00"
            End If
            
            txtPaid2Hos = totalPaid2Hos
            If IsNumeric(txtPaid2Hos) = False Then
                txtPaid2Hos = "0.00"
            End If
            ' Out standing hos
            txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,##0.00;(#,##0.00)")
            
            txtPayableMem = Format(totalPayable2Mbm, "#,##0.00;(#,##0.00)")
            
            txtPaid2MBM = Format(totalPaid2MBM, "#,##0.00")
            If IsNumeric(txtPaid2MBM) = False Then
                txtPaid2MBM = "0.00"
            End If
                
            txtPaidByMem = Format(totalPaidByMbm, "#,##0.00")
            If IsNumeric(txtPaidByMem) = False Then
                txtPaidByMem = "0.00"
            End If
            ' Out standing MBM
            txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
            txtReimbursement = Format(TotalReimburse, "#,##0.00;(#,##0.00)")
            txtExcess = Format(TotalExcess, "#,##0.00;(#,##0.00)")
        End If
        Acc.Close
        Set Acc = Nothing
    End If
    
End Sub

Private Sub Form_Load()
    Call DoInitialSettings
    Call ComboListing
    bDataStatus = True
    cmdShow.Enabled = True
    txtMBMNumber.Enabled = True
    cmdSearchMBM.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdExit.Enabled = True
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

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub ClearEntries()
On Error Resume Next
Dim ctl As Control
Dim s As Integer

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtMBMNumber" Then
               ctl.Enabled = True
               ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is ComboBox Then
                ctl.Enabled = True
                ctl.Text = ""
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Enabled = True
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        End If
    Next
    'lblKidneyAnn = ""
    'lblKidneyBal = ""
    lblCancerAnn = ""
    lblCancerBal = ""
    lblHomeNurseAnn = ""
    lblHomeNurseBal = ""
    lblCovAnnLmt = ""
    lblCovAvLmt = ""
    lblCovKidneyAnnLmt = ""
    lblCovKidneyBalLmt = ""
    lblCovCancerAnnLmt = ""
    lblCovCancerBalLmt = ""
    lblCovHomeNurseAnnLmt = ""
    lblCovHomeNurseBalLmt = ""
    lblPerDisability = ""
    lstPreUndIllness.ListItems.Clear
    lstLTLmt.ListItems.Clear
    lstSuppBenefitDetails.ListItems.Clear
    Fgrid.col = 1
    For s = 1 To 10
        Fgrid.row = (s)
        Fgrid = ""
    Next
    lstPolicyWording.Clear
    lblTotalCasesDOB = 0
    lblTotalCasesIC = 0
    lblTotalMBMIC = 0
    lblTotalMBMDOB = 0
    
    'Fgrid.Clear
End Sub


Private Sub ComboListing()
    cboMBMTakeOver.AddItem "N - None"
    cboMBMTakeOver.AddItem "Y - Take Over"
    cboMBMTakeOver.AddItem "C - Conversion"
    
    CboInsMode.AddItem "A - Annually"
    CboInsMode.AddItem "Q - Quarterly"
    CboInsMode.AddItem "S - Semi Annual"
    CboInsMode.AddItem "M - Monthly"
    
    CboGuaRenewal.Clear
    CboGuaRenewal.AddItem "  "
    CboGuaRenewal.AddItem "1 - No Life Time Limit"
    CboGuaRenewal.AddItem "3 - With Life Time Limit"

End Sub

Private Sub BNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim bnf As ListItem
    Dim sql, cnt As String
    Dim CurrEffDate, PrevEffDate As String
    
    If lstBenefitDetails.ListItems.Count = 0 And Len(Trim(cboHLTCode)) <> 0 _
     And Len(Trim(cboINSCode)) <> 0 And Len(Trim(cboPLNCode)) <> 0 And IsDate(txtMBMPayorEffDate) = True Then
        sql = " select BNFCode, BNFNoOFDays , BNFEffDate, " & _
          " BNFclaimableAmount, BNFdescription, BNFRMPerDIs " & _
          " from View_BNF where PLNCOde ='" & Trim(cboPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(cboHLTCode) & "'" & _
          " AND INSCODe ='" & Trim(cboINSCode) & "'" & _
          " and BNFEffDate <='" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' ORDER BY BNFEffDate DESC"
        rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        cnt = 0
        Do Until rst2.EOF
            cnt = cnt + 1
            If cnt = 1 Then
                PrevEffDate = rst2!BNFEffDate
            End If
            CurrEffDate = rst2!BNFEffDate
            If PrevEffDate <> CurrEffDate Then
                Exit Do
            End If
            Set bnf = lstBenefitDetails.ListItems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFdescription) <> "" Then
                    bnf.SubItems(1) = rst2!BNFdescription
                End If
                
                bnf.SubItems(2) = IIf(IsNull(rst2!BNFNoOfDays), 0, rst2!BNFNoOfDays)
                If IsNumeric(rst2!BNFClaimAbleAmount) Then
                    bnf.SubItems(3) = Format(Round(rst2!BNFClaimAbleAmount, 2), "###,###,###.00")
                Else
                    bnf.SubItems(3) = IIf(Trim(rst2!BNFClaimAbleAmount) <> "", rst2!BNFClaimAbleAmount, "")
                End If
                If Trim(rst2!BNFRMperDis) <> "" Then
                    bnf.SubItems(4) = rst2!BNFRMperDis
                End If
                
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub BNFListingSupp()
    Dim rst2 As New ADODB.Recordset
    Dim bnf As ListItem
    Dim sql, cnt As String
    Dim CurrEffDate, PrevEffDate As String
    lstSuppBenefitDetails.ListItems.Clear
    If lstSuppBenefitDetails.ListItems.Count = 0 And Len(Trim(cboHLTCode)) <> 0 _
     And Len(Trim(cboINSCode)) <> 0 And Len(Trim(txtSuppPLNCode)) <> 0 And IsDate(txtMBMPayorEffDate) = True Then
        sql = " select BNFCode, BNFNoOFDays , BNFEffDate, " & _
          " BNFclaimableAmount, BNFdescription, BNFRMPerDIs " & _
          " from View_BNF where PLNCOde ='" & Trim(txtSuppPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(cboHLTCode) & "'" & _
          " AND INSCODe ='" & Trim(cboINSCode) & "'" & _
          " and BNFEffDate <='" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' ORDER BY BNFEffDate DESC"
        rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        cnt = 0
        Do Until rst2.EOF
            cnt = cnt + 1
            If cnt = 1 Then
                PrevEffDate = rst2!BNFEffDate
            End If
            CurrEffDate = rst2!BNFEffDate
            If PrevEffDate <> CurrEffDate Then
                Exit Do
            End If
            Set bnf = lstSuppBenefitDetails.ListItems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFdescription) <> "" Then
                    bnf.SubItems(1) = rst2!BNFdescription
                End If
                
                bnf.SubItems(2) = IIf(IsNull(rst2!BNFNoOfDays), 0, rst2!BNFNoOfDays)
                If IsNumeric(rst2!BNFClaimAbleAmount) Then
                    bnf.SubItems(3) = Format(Round(rst2!BNFClaimAbleAmount, 2), "###,###,###.00")
                Else
                    bnf.SubItems(3) = IIf(Trim(rst2!BNFClaimAbleAmount) <> "", rst2!BNFClaimAbleAmount, "")
                End If
                If Trim(rst2!BNFRMperDis) <> "" Then
                    bnf.SubItems(4) = rst2!BNFRMperDis
                End If
                
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub NewBNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim bnf As ListItem
    Dim sql, cnt As String
    Dim CurrEffDate, PrevEffDate As String
    If lstNewBenefitDetails.ListItems.Count = 0 And Len(Trim(cboHLTCode)) <> 0 _
     And Len(Trim(cboINSCode)) <> 0 And Len(Trim(txtPLNCode)) <> 0 And IsDate(txtMBMPayorEffDate) = True Then
        sql = " select BNFCode, BNFNoOFDays , " & _
          " BNFclaimableAmount, BNFdescription, BNFRMPerDIs,BNFEffDate " & _
          " from View_BNF where PLNCOde ='" & Trim(txtPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(cboHLTCode) & "'" & _
          " AND INSCODe ='" & Trim(cboINSCode) & "'" & _
          " and BNFEffDate <='" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' ORDER BY BNFEffDate DESC"
            rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            cnt = 0
            Do Until rst2.EOF
                cnt = cnt + 1
                If cnt = 1 Then
                    PrevEffDate = rst2!BNFEffDate
                End If
                CurrEffDate = rst2!BNFEffDate
                If PrevEffDate <> CurrEffDate Then
                    Exit Do
                End If
                Set bnf = lstNewBenefitDetails.ListItems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFdescription) <> "" Then
                    bnf.SubItems(1) = rst2!BNFdescription
                End If
                
                bnf.SubItems(2) = IIf(IsNull(rst2!BNFNoOfDays), 0, rst2!BNFNoOfDays)
                If Trim(rst2!BNFClaimAbleAmount) <> "" Then
                    bnf.SubItems(3) = Format(Round(rst2!BNFClaimAbleAmount, 2), "###,###,###.00")
                End If
                If Trim(rst2!BNFRMperDis) <> "" Then
                    bnf.SubItems(4) = rst2!BNFRMperDis
                End If
                
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub MBMHISListing()
    If lstMBMHis.ListItems.Count = 0 Then
        If Len(Trim(txtMBMIcBcPP)) > 4 Then
            Call MBMHistory("Princ")
            Call MBMHistory("Supp")
        'Else
        '    Call MBMHistory("Princ", False)
        '    Call MBMHistory("Supp", False)
        End If
    End If
End Sub

Private Sub MBMHISListingTwo()
    If lstMBMHisTwo.ListItems.Count = 0 Then
        If Trim(txtMBMName) <> "" And txtMBMBirthDate <> "__/__/____" Then
            Call MBMHistoryTwo("Princ")
            Call MBMHistoryTwo("Supp")
        End If
    End If
End Sub
Private Sub ADJListing()
    Dim ADJ As New ADODB.Recordset
    Dim ADJList As ListItem
    Dim sql As String
    
    If Len(Trim(txtMBMNumber)) Then
         sql = " select MBMAmendID, MBMNumber, " & _
           " MBMAmendEdtType, MBMAmendDate, MBMAENDtEffDate, " & _
           " MBMAEndtBordxDate, MBMAEndtBatchNo " & _
           " From MBMAmendHistory where MBMnumber = '" & Trim(txtMBMNumber) & "'"
         ADJ.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
         Do Until ADJ.EOF
         
             Set ADJList = lstAdjustmentHis.ListItems.Add(, , IIf(Len(Trim(ADJ!MBMAENDtBordxDate)) <> 0, ADJ!MBMAENDtBordxDate, "-"))
                 If Len(Trim(ADJ!MBMAEndtBatchNo)) Then
                    ADJList.SubItems(1) = ADJ!MBMAEndtBatchNo
                 End If
                 If Len(Trim(ADJ!MBMAmendDate)) Then
                    ADJList.SubItems(2) = ADJ!MBMAmendDate
                 End If

                If Len(ADJ!MBMAmendEdtType) Then
                    ADJList.SubItems(3) = ADJ!MBMAmendEdtType
                End If
                
                If Len(ADJ!MBMAEndtEffDate) Then
                    ADJList.SubItems(4) = ADJ!MBMAEndtEffDate
                End If

               
                 ADJ.MoveNext
         Loop
         ADJ.Close
         Set ADJ = Nothing
    End If
End Sub


' ######## List out all the cases which have been discharged
Private Sub ListALLCases()
Dim ForLoop As Integer
Dim TempIcNo As String
Dim StrSqlMBMsupp As String
Dim MBMSupp As New ADODB.Recordset
Dim MBMSuppRecStatus As Boolean
    cntCasesHisIC = 0
    ForLoop = 0
    TempIcNo = ""
    lstCASHistory.ListItems.Clear
    MBMSuppRecStatus = False
    If Len(Trim(txtMBMIcBcPP)) > 5 Then
        Call ListCases1(True, "", txtMBMIcBcPP, False)
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
    lblTotalCasesIC = cntCasesHisIC
    

End Sub

Private Sub ListALLCasesTwo()
Dim ForLoop As Integer
Dim TempIcNo As String
Dim StrSqlMBMsupp As String
Dim MBMSupp As New ADODB.Recordset
Dim MBMSuppRecStatus As Boolean
Dim tmpMBMCName As String
Dim tmpMBMCDOB As String
Dim tmpMBMName As String

    ForLoop = 0
    TempIcNo = ""
    tmpMBMCName = ""
    tmpMBMCDOB = ""
    lstCASHistoryTwo.ListItems.Clear
    MBMSuppRecStatus = False
    If Trim(txtMBMName) <> "" And txtMBMBirthDate <> "__/__/____" Then
        tmpMBMName = Trim(Replace(txtMBMName, "'", "''"))
        Call ListCasesTwo(True, "", tmpMBMName, txtMBMBirthDate, False)
        StrSqlMBMsupp = ""
        StrSqlMBMsupp = "Select MBMCName,MBMCDOB from MBMCoveredPersons where MBMCName = '" & Trim(tmpMBMName) & "' And MBMCDOB = '" & Format(txtMBMBirthDate, "yyyy/mm/dd") & "'"
        MBMSupp.Open StrSqlMBMsupp, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        MBMSuppRecStatus = False
        Do While Not MBMSupp.EOF
           MBMSuppRecStatus = True
           MBMSupp.MoveNext
           Exit Do
        Loop
        MBMSupp.Close
        Set MBMSupp = Nothing
        If MBMSuppRecStatus = True Then
            Call ListCasesTwo(True, "", tmpMBMName, txtMBMBirthDate, True)
        End If
    End If
    If Mid(cboINSCode, 1, 1) = "F" Or Mid(cboINSCode, 1, 1) = "H" Then
        For ForLoop = 1 To lstCovAdd.ListItems.Count
            MBMSuppRecStatus = False
            'TempIcNo = lstCovAdd.ListItems(ForLoop).SubItems(2)
            tmpMBMCName = Trim(Replace(lstCovAdd.ListItems(ForLoop).Text, "'", "''"))
            tmpMBMCDOB = lstCovAdd.ListItems(ForLoop).SubItems(3)
            If Trim(tmpMBMCName) <> "" And tmpMBMCDOB <> "" Then
                StrSqlMBMsupp = ""
                tmpMBMName = Trim(Replace(txtMBMName, "'", "''"))
                StrSqlMBMsupp = "Select MBMName,MBMDOB from MBMCrossReference where MBMName = '" & Trim(tmpMBMName) & "' And MBMDOB = '" & Format(txtMBMBirthDate, "yyyy/mm/dd") & "'"
                MBMSupp.Open StrSqlMBMsupp, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
                Do While Not MBMSupp.EOF
                    MBMSuppRecStatus = True
                    MBMSupp.MoveNext
                    Exit Do
                Loop
                MBMSupp.Close
                Set MBMSupp = Nothing
            End If
            If MBMSuppRecStatus = True Then
                Call ListCasesTwo(True, "", tmpMBMCName, tmpMBMCDOB, False)
            End If
        Next ForLoop
    End If
    lblTotalCasesDOB = cntCasesHisDOB

End Sub
' ####################### Calculation's Function ###############
Private Sub MemberCallAge()
Dim SelectAge As String
    If Trim(txtMBMPayorEffDate) <> "__/__/____" _
        And Len(Trim(txtMBMBirthDate)) <> 0 _
        And Trim(txtMBMBirthDate) <> "__/__/____" _
        And Len(Trim(cboHLTCode)) <> 0 Then
        If IsDate(txtMBMBirthDate) Then
            SelectAge = GetAge(CDate(HisCisEffDate), CDate(txtMBMBirthDate))
            txtAGECode = GetAgeBand(Trim(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1)), SelectAge, HISCisPLN, HISCiSIns)
            If Len(Trim(txtAGECode)) = 0 Then
                MsgBox "Please Enter A Valid Date OR Age Out of Acceptable Range", vbOKOnly + vbCritical, DialogBoxTitle
                txtMBMPayorEffDate.SetFocus
            Else
                If Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)) = "01" Then
                    txtMBMAdultChild = "C"
                Else
                    txtMBMAdultChild = "A"
                End If
            End If
        End If
    End If
End Sub

Public Sub CallLimit()
    txtAnnualLimit = ""
    If Len(Trim(cboPLNCode)) Then
        txtAnnualLimit = Format(GetAnnualLimit(cboINSCode, cboPLNCode, cboHLTCode, txtMBMPayorEffDate), "#,##0.00")
    End If
End Sub


Private Sub cmdPrev_Click()
    Dim MBM As New ADODB.Recordset
    Dim sql As String
    Dim MBMNumber As String
    Dim status As Boolean
        
        'lblTopupStatus.Visible = False
        
        status = False
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        sql = " select MBMNumber from MBM where MBMnumber < '" & RemStr(txtMBMNumber) & "' " & _
            " Order by MBMnumber DESC "
        MBM.MaxRecords = 1
        MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.recordCount Then
            MBMNumber = MBM!MBMNumber
            status = True
        End If
        MBM.Close
        Set MBM = Nothing
       ' medixmasterconn.Close
      '  Set medixmasterconn = Nothing
        
        If status = True Then
            txtMBMNumber = MBMNumber
            cmdShow_Click
        Else
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
        End If
End Sub

Private Sub CmdNext_Click()
    Dim MBM As New ADODB.Recordset
    Dim sql As String
    Dim MBMNumber As String
    Dim status As Boolean
        
        'lblTopupStatus.Visible = False
        
        status = False
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        sql = " select MBMNumber from MBM where MBMnumber > '" & RemStr(txtMBMNumber) & "' " & _
            " Order by MBMnumber "
        MBM.MaxRecords = 1
        MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.recordCount Then
            MBMNumber = MBM!MBMNumber
            status = True
        End If
        MBM.Close
        Set MBM = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
        If status = True Then
            txtMBMNumber = MBMNumber
            cmdShow_Click
        Else
            MsgBox "Last Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
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
    
    If Right(B, 4) = ".pdf" Then
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
        
        If Right(B, 4) = ".PDF" Then
            lstPolicyWording.AddItem B
        End If

    Else
        Exit Do
    End If
Loop
'List1.Height = Me.Height - 2520 - 350
End Sub
Private Sub lstACCHistory_DblClick()
    If lstACCHistory.ListItems.Count Then
        frmCaseInquiry2016.txtCaseNo = Trim(Mid(lstACCHistory.SelectedItem.SubItems(4), 1, IIf(InStr(1, lstACCHistory.SelectedItem.SubItems(4), "-") = 0, 10, InStr(1, lstACCHistory.SelectedItem.SubItems(4), "-") - 1)))
        Call frmCaseInquiry2016.cmdShow_Click
        frmCaseInquiry2016.Show
    End If
End Sub

Private Sub lstAdjustmentHis_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstAdjustmentHis.SortKey = ColumnHeader.Index - 1
    lstAdjustmentHis.Sorted = True
    If lstAdjustmentHis.SortOrder = lvwAscending Then
        lstAdjustmentHis.SortOrder = lvwDescending
    Else
        lstAdjustmentHis.SortOrder = lvwAscending
    End If

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

Private Sub lstCASHistory_DblClick()
    If lstCaseList.ListItems.Count Then
        frmCaseInquiry2016.txtCaseNo = ChkStr(lstCASHistory.SelectedItem.Text)
        Call frmCaseInquiry2016.cmdShow_Click
        frmCaseInquiry2016.Show
    End If
End Sub

Private Sub lstNewBenefitDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstNewBenefitDetails.SortKey = ColumnHeader.Index - 1
    lstNewBenefitDetails.Sorted = True
    If lstNewBenefitDetails.SortOrder = lvwAscending Then
        lstNewBenefitDetails.SortOrder = lvwDescending
    Else
        lstNewBenefitDetails.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCaseList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCovAdd.SortKey = ColumnHeader.Index - 1
    lstCovAdd.Sorted = True
    If lstCovAdd.SortOrder = lvwAscending Then
        lstCovAdd.SortOrder = lvwDescending
    Else
        lstCovAdd.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCovAdd_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCovAdd.SortKey = ColumnHeader.Index - 1
    lstCovAdd.Sorted = True
    If lstCovAdd.SortOrder = lvwAscending Then
        lstCovAdd.SortOrder = lvwDescending
    Else
        lstCovAdd.SortOrder = lvwAscending
    End If

End Sub

Private Sub lstCovAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstCovAdd.ListItems.Count Then
        txtCovName = lstCovAdd.SelectedItem.Text
        lblCovName = lstCovAdd.SelectedItem.Text
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtSuppPLNCode = IIf(Len(lstCovAdd.SelectedItem.SubItems(14)), lstCovAdd.SelectedItem.SubItems(14), cboPLNCode)
        txtSuppEffDate = IIf(Len(lstCovAdd.SelectedItem.SubItems(15)), lstCovAdd.SelectedItem.SubItems(15), txtMBMPayorEffDate)
        txtSuppExpDate = IIf(Len(lstCovAdd.SelectedItem.SubItems(16)), lstCovAdd.SelectedItem.SubItems(16), txtMBMPayorExpDate)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        lblCovAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        lblCovAvLmt = Format(lstCovAdd.SelectedItem.SubItems(13), "###,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(19)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(20)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(21)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(22)
        'If Len(lstCovAdd.SelectedItem.SubItems(23)) Then
        '    TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(23)
        'End If
        'If Len(lstCovAdd.SelectedItem.SubItems(24)) Then
        '    TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(24)
        'End If
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(25)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(26)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(27)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(28)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(29)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(30)
        txtPolicyDiscAmt = lstCovAdd.SelectedItem.SubItems(31)
        txtRenewalDiscAmt = lstCovAdd.SelectedItem.SubItems(32)
        txtLoadingPrem = lstCovAdd.SelectedItem.SubItems(33)
        lblCovKidneyAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(35), "#,##0.00")
        lblCovKidneyBalLmt = Format(lstCovAdd.SelectedItem.SubItems(36), "#,##0.00")
        lblCovCancerAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(37), "#,##0.00")
        lblCovCancerBalLmt = Format(lstCovAdd.SelectedItem.SubItems(38), "#,##0.00")
        lblCovHomeNurseAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(39), "#,##0.00")
        lblCovHomeNurseBalLmt = Format(lstCovAdd.SelectedItem.SubItems(40), "#,##0.00")
        txtPayMBMNo = IIf(Len(lstCovAdd.SelectedItem.SubItems(41)), Trim(lstCovAdd.SelectedItem.SubItems(41)), "")
        txtPaySeqID = IIf(Len(lstCovAdd.SelectedItem.SubItems(42)), Trim(lstCovAdd.SelectedItem.SubItems(42)), "")
        txtClientNo = IIf(Len(lstCovAdd.SelectedItem.SubItems(43)), Trim(lstCovAdd.SelectedItem.SubItems(43)), "")
        txtClientID = IIf(Len(lstCovAdd.SelectedItem.SubItems(44)), Trim(lstCovAdd.SelectedItem.SubItems(44)), "")
    End If
End Sub

Private Sub lstCovAdd_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstCovAdd.ListItems.Count Then
        txtCovName = lstCovAdd.SelectedItem.Text
        lblCovName = lstCovAdd.SelectedItem.Text
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        lblCovAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        lblCovAvLmt = Format(lstCovAdd.SelectedItem.SubItems(13), "###,##0.00")
        txtSuppPLNCode = IIf(Len(lstCovAdd.SelectedItem.SubItems(14)), lstCovAdd.SelectedItem.SubItems(14), cboPLNCode)
        txtSuppEffDate = IIf(Len(lstCovAdd.SelectedItem.SubItems(15)), lstCovAdd.SelectedItem.SubItems(15), txtMBMPayorEffDate)
        txtSuppExpDate = IIf(Len(lstCovAdd.SelectedItem.SubItems(16)), lstCovAdd.SelectedItem.SubItems(16), txtMBMPayorExpDate)
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(19)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(20)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(21)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(22)
        'If Len(lstCovAdd.SelectedItem.SubItems(23)) Then
        '    TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(23)
        'End If
        'If Len(lstCovAdd.SelectedItem.SubItems(24)) Then
        '    TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(24)
        'End If
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(25)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(26)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(27)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(28)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(29)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(30)
        txtPolicyDiscAmt = lstCovAdd.SelectedItem.SubItems(31)
        txtRenewalDiscAmt = lstCovAdd.SelectedItem.SubItems(32)
        txtLoadingPrem = lstCovAdd.SelectedItem.SubItems(33)
        lblCovKidneyAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(35), "#,##0.00")
        lblCovKidneyBalLmt = Format(lstCovAdd.SelectedItem.SubItems(36), "#,##0.00")
        lblCovCancerAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(37), "#,##0.00")
        lblCovCancerBalLmt = Format(lstCovAdd.SelectedItem.SubItems(38), "#,##0.00")
        lblCovHomeNurseAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(39), "#,##0.00")
        lblCovHomeNurseBalLmt = Format(lstCovAdd.SelectedItem.SubItems(40), "#,##0.00")
        txtPayMBMNo = IIf(Len(lstCovAdd.SelectedItem.SubItems(41)), Trim(lstCovAdd.SelectedItem.SubItems(41)), "")
        txtPaySeqID = IIf(Len(lstCovAdd.SelectedItem.SubItems(42)), Trim(lstCovAdd.SelectedItem.SubItems(42)), "")
        txtClientNo = IIf(Len(lstCovAdd.SelectedItem.SubItems(43)), Trim(lstCovAdd.SelectedItem.SubItems(43)), "")
        txtClientID = IIf(Len(lstCovAdd.SelectedItem.SubItems(44)), Trim(lstCovAdd.SelectedItem.SubItems(44)), "")
    End If
End Sub

Private Sub lstMBMAcc_Click()
Dim strsql As String
Dim MBM As New ADODB.Recordset
    
    If lstMBMAcc.ListItems.Count <> 0 Then
    
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        strsql = " Select MBMNumber, MBMMemberType, MBMCAnnualLimit, " & _
                 " MBMCAvailableLimit From View_MBM_MemberType where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.recordCount Then
        If MBM!MBMMemberType = "P" Then
            txtAnnualLmt = txtAnnualLimit
            txtAvailableAnnualLmt = txtAvailableAnnualLimit
        ElseIf MBM!MBMMemberType = "Q" And lstMBMAcc.SelectedItem.SubItems(1) <> 0 Then
            txtAnnualLmt = MBM!MBMCAnnualLimit
            txtAvailableAnnualLmt = MBM!MBMCAvailableLimit
        ElseIf MBM!MBMMemberType = "Q" And lstMBMAcc.SelectedItem.SubItems(1) = 0 Then
            txtAnnualLmt = txtAnnualLimit
            txtAvailableAnnualLmt = txtAvailableAnnualLimit
        End If
        
        txtApprovedAmt = Format(lstMBMAcc.SelectedItem.SubItems(2), "#,##0.00")
        txtPayable2Hos = Format(lstMBMAcc.SelectedItem.SubItems(3), "#,##0.00")
        If IsNumeric(txtPayable2Hos) = False Then
            txtPayable2Hos = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(4)) <> "" Then
            txtPaid2Hos = Format(lstMBMAcc.SelectedItem.SubItems(4), "#,##0.00")
        End If
        If IsNumeric(txtPaid2Hos) = False Then
            txtPaid2Hos = 0
        End If
        
        ' Out standing hos
        txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,##0.00;(#,##0.00)")
        
        txtPayableMem = Format(lstMBMAcc.SelectedItem.SubItems(5), "#,##0.00;(#,##0.00)")
        If IsNumeric(txtPayableMem) = False Then
            txtPayableMem = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(6)) <> "" Then
            txtPaid2MBM = Format(lstMBMAcc.SelectedItem.SubItems(6), "#,##0.00")
        End If
        If IsNumeric(txtPaid2MBM) = False Then
            txtPaid2MBM = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(7)) <> "" Then
            txtPaidByMem = Format(lstMBMAcc.SelectedItem.SubItems(7), "#,##0.00")
        End If
        If IsNumeric(txtPaidByMem) = False Then
            txtPaidByMem = 0
        End If
        txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,##0.00;(#,##0.00)")
        txtReimbursement = Format(lstMBMAcc.SelectedItem.SubItems(8), "#,##0.00")
        txtExcess = Format(lstMBMAcc.SelectedItem.SubItems(9), "#,##0.00;(#,##0.00)")
        txtBordxAmt = Format(lstMBMAcc.SelectedItem.SubItems(10), "#,##0.00;(#,##0.00)")
        If lstMBMAcc.SelectedItem.SubItems(9) = "" Then
            lstMBMAcc.SelectedItem.SubItems(9) = "0"
            txtExcessPaid = Format(lstMBMAcc.SelectedItem.SubItems(9) - lstMBMAcc.SelectedItem.SubItems(5), "#,#.00;(#,#.00)")
        End If

        End If
        
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
    End If
End Sub

Private Sub lstMBMHis_DblClick()
    If lstMBMHis.ListItems.Count Then
        txtMBMNumber = lstMBMHis.SelectedItem.Text
        cmdShow_Click
        SSTab1.Tab = 0
    End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Screen.MousePointer = vbHourglass
'Dim PolList As ListItem
Dim PolList As MSFlexGrid
Dim frmPanelHospStatus As Integer
Dim frmTermsStatus As Integer
Dim txtHOSCode As String
Dim txtPayCode As String
Dim txtHLTCode As String
Dim tmpHOSCode As String
Dim tmpHltCode As String

    If SSTab1.Tab = 1 Then
        If Mid(cboINSCode, 1, 1) = "I" Or Mid(cboINSCode, 1, 1) = "G" Then
            MsgBox "This Plan does not have supplementary", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 0
        Else
            SSTab1.Tab = 1
        End If
        
    End If
    
    If SSTab1.Tab = 2 Then
        lstAdjustmentHis.ListItems.Clear
        If lstAdjustmentHis.ListItems.Count = 0 Then
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call ADJListing
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        End If
    End If
    
    If SSTab1.Tab = 3 Then
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            If lstBenefitDetails.ListItems.Count = 0 Then
                Call BNFListing
            End If
            If cboINSCode = "F" Or cboINSCode = "H" Then
                'If lstSuppBenefitDetails.ListItems.Count = 0 Then
                Call BNFListingSupp
                'End If
            End If
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        'If lstNewBenefitDetails.ListItems.Count = 0 Then
        '    If txtPLNCode <> "" Or txtNewPolEffDate <> "__/__/____" Then
        '        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        '        Call NewBNFListing
        '        medixmasterconn.Close
        '        Set medixmasterconn = Nothing
        '    End If
        'End If
    End If
    
    If SSTab1.Tab = 4 Then
        If Trim(txtMBMNumber) <> "" And lstCaseList.ListItems.Count = 0 Then
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call ListCases1(False, txtMBMNumber, "", False)
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        End If
    End If
    
    If SSTab1.Tab = 5 Then
        If Trim(txtMBMNumber) <> "" And lstCASHistory.ListItems.Count = 0 Then
          '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
         '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            lblTotalCasesIC = 0
            lblTotalCasesDOB = 0
            cntCasesHisIC = 0
            cntCasesHisDOB = 0
            
            Call ListALLCases
            'Call ListALLCasesTwo
         '   medixmasterconn.Close
         '   Set medixmasterconn = Nothing
           ' medixmasterconnMaint.Close
           ' Set medixmasterconnMaint = Nothing
            
            'If CDbl(cntCasesHisIC) <> CDbl(cntCasesHisDOB) Then
            '    MsgBox "Please Check Case History", vbOKOnly + vbCritical, DialogBoxTitle
            'End If
        End If
    End If
    
    If SSTab1.Tab = 6 Then
        If txtMBMNumber = "" Then
            MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMNumber.SetFocus
        ElseIf lstMBMHis.ListItems.Count = 0 Then
           ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
          '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                lblTotalMBMIC = 0
                lblTotalMBMDOB = 0
                cntMBMHisIC = 0
                cntMBMHisDOB = 0
                
                Call MBMHISListing
                'Call MBMHISListingTwo
                lblTotalMBMIC = cntMBMHisIC
                lblTotalMBMDOB = cntMBMHisDOB
                'If CDbl(cntMBMHisIC) <> CDbl(cntMBMHisDOB) Then
                '    MsgBox "Please Check Member History", vbOKOnly + vbCritical, DialogBoxTitle
                'End If
                
            'medixmasterconn.Close
          '  Set medixmasterconn = Nothing
           ' medixmasterconnMaint.Close
           ' Set medixmasterconnMaint = Nothing
        End If
    End If
    
    If SSTab1.Tab = 7 Then
        If lstMBMAcc.ListItems.Count = 0 Then
          '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call GetAccountSummary
           ' medixmasterconn.Close
            'Set medixmasterconn = Nothing
        End If
    End If
    
    If SSTab1.Tab = 8 Then
        If lstACCHistory.ListItems.Count = 0 Then
            'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                Call GetAccHisListing
           ' medixmasterconnWS.Close
            'Set medixmasterconnWS = Nothing
        End If
    End If
    
    If SSTab1.Tab = 10 Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call SearchPrincipalIll
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
    End If
    
    If SSTab1.Tab = 11 Then
        If cboPLNCode <> "" Then
            frmTermsStatus = 4
           ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call PolicyTermsListing(cboPLNCode, frmTermsStatus, cboPAYCode, cboHLTCode)
         '   medixmasterconn.Close
         '   Set medixmasterconn = Nothing
        End If
    End If
    
    If SSTab1.Tab = 12 Then
        If cboPLNCode <> "" Then
            If lstLTLmt.ListItems.Count = 0 Then
               ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                Call LTLmtListing
                'medixmasterconn.Close
                'Set medixmasterconn = Nothing
            End If
        End If
    End If
    
    If SSTab1.Tab = 13 Then
        If Trim(tmpPolicyWording) <> "" Then
            Call LoadDoc
        End If
    End If

    If SSTab1.Tab = 14 Then
        lstTOClaims.ListItems.Clear
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call TakeoverClaims
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
    End If

   'If SSTab1.Tab = 12 Then
   '     If cboPLNCode <> "" Then
   '         If lstPanelHosp.listitems.Count = 0 Then
   '             frmTermsStatus = 1
   '             frmPanelHospStatus = 3
   '             medixmasterconn.Close
   '             medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   '           Call PanelHospListing(cboPLNCode, frmTermsStatus, frmPanelHospStatus, cboPAYCode, tmpHOSCode, txtHOSCode, tmpHLTCode)
                
   '         If medixmasterconn Is Nothing Then
   '             medixmasterconn.Close
   '         End If
            
   '         Set medixmasterconn = Nothing
   '         End If
   '     End If
   ' End If
   
       If SSTab1.Tab = 16 Then
        If Trim(txtMBMNumber) <> "" And lstCaseList.ListItems.Count = 0 Then
           ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call ListWorksheet(False, txtMBMNumber, "", False)
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        End If
    End If

Screen.MousePointer = Default
End Sub

Private Sub TakeoverClaims()
Dim CASList As ListItem
Dim strsql As String
Dim CAS As New ADODB.Recordset
    

    Screen.MousePointer = vbHourglass
        
    strsql = "Select * from CLMMAATakaful where CLMPolNo = '" & Trim(txtMBMPolNo) & "'"
    CAS.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        
    Do While Not CAS.EOF
        Set CASList = lstTOClaims.ListItems.Add(, , CAS!CLMEffDate)
        CASList.SubItems(1) = CAS!CLMExpDate
        CASList.SubItems(2) = IIf(Len(Trim(CAS!CLMPatientname)), Trim(CAS!CLMPatientname), "")
        If Len(Trim(CAS!CLMPatientIC)) Then
            CASList.SubItems(3) = CAS!CLMPatientIC
        End If
        If Len(Trim(CAS!CLMDOA)) Then
            CASList.SubItems(4) = CAS!CLMDOA
        End If
        
        If Len(Trim(CAS!CLMDOD)) Then
            CASList.SubItems(5) = CAS!CLMDOD
        End If
        
        CASList.SubItems(6) = IIf(Len(CAS!CLMHospital), CAS!CLMHospital, "")
        CASList.SubItems(7) = IIf(Len(CAS!CLMDiagnosis1), CAS!CLMDiagnosis1, "")
        If Len(CAS!CLMDiagnosis2) Then
            CASList.SubItems(7) = CASList.SubItems(7) & "/" & CAS!CLMDiagnosis2
        End If
        If Len(CAS!CLMDiagnosis3) Then
            CASList.SubItems(7) = CASList.SubItems(7) & "/" & CAS!CLMDiagnosis3
        End If
        CASList.SubItems(8) = IIf(Len(CAS!CLMDrName1), CAS!CLMDrName1, "")
        If Len(CAS!CLMDrName2) Then
            CASList.SubItems(8) = CASList.SubItems(8) & "/" & CAS!CLMDrName2
        End If
        
        CASList.SubItems(9) = IIf(Len(CAS!CLMBillAmt), CAS!CLMBillAmt, "")
        CASList.SubItems(10) = IIf(Len(CAS!CLMAppAmt), CAS!CLMAppAmt, "")
        CASList.SubItems(11) = IIf(Len(CAS!CLMExcessAmt), CAS!CLMExcessAmt, "")
        cntCasesHisIC = CDbl(cntCasesHisIC) + 1
        CAS.MoveNext
        
    Loop
    
    CAS.Close
    Set CAS = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub lstCaseList_DblClick()
    If lstCaseList.ListItems.Count Then
        frmCaseInquiry2016.txtCaseNo = ChkStr(lstCaseList.SelectedItem.Text)
        Call frmCaseInquiry2016.cmdShow_Click
        frmCaseInquiry2016.Show
    End If
End Sub
Public Sub CheckBalance(MBMNumber As String)
    Dim txtgrpcompany As String
    Dim Balcmd As New ADODB.Command
    Dim BalPrm1 As New ADODB.Parameter
    Dim BalPrm2 As New ADODB.Parameter
    Dim BalRes As New ADODB.Recordset
    Dim CovID As String
    Dim MBMNo As String
    
    MBMNo = MBMNumber
    
    If InStr(1, MBMNo, "-") Then
       Dim sArr() As String
        sArr = Split(MBMNo, "-")
        MBMNo = sArr(0)
        If (Len(sArr(1))) Then
            CovID = sArr(1)
        Else
            CovID = "00"
        End If
    Else
        CovID = "00"
    End If
    
        Set Balcmd.ActiveConnection = medixmasterconn
        Balcmd.CommandText = "sp_HISUtilization_android"
        Balcmd.CommandType = adCmdStoredProc
        Balcmd.CommandTimeout = 300
        Set BalPrm1 = Balcmd.CreateParameter("MBMNo", adVarChar, adParamInput, 2000, Trim(MBMNo))
        Balcmd.Parameters.Append BalPrm1
        Set BalPrm2 = Balcmd.CreateParameter("CoverID", adVarChar, adParamInput, 2000, Trim(CovID))
        Balcmd.Parameters.Append BalPrm2
        BalRes.CursorLocation = adUseClient
        BalRes.CursorType = adOpenForwardOnly
        BalRes.CacheSize = 100
        Set BalRes = Balcmd.Execute
        Do While Not BalRes.EOF
            Ciamamt = IIf(Len(BalRes!ClaimedAmt), BalRes!ClaimedAmt, 0)
            AnnlLmt = IIf(Len(BalRes!AnnualLimit), BalRes!AnnualLimit, 0)
            Bal = IIf(Len(BalRes!balance), BalRes!balance, 0)
            BalRes.MoveNext
        Loop
        BalRes.Close
        Set BalRes = Nothing
End Sub

Public Sub cmdShow_Click()

Dim MBMNumber As String
Dim tmpREcCnt As Integer
    SSTab1.Tab = 0
    MBMNumber = txtMBMNumber
    Call ClearEntries
    txtMBMNumber = MBMNumber
    'Call CheckBalance(MBMNumber)
    tmpVIP = False
    tmpMBMGLIssuance = ""
    Screen.MousePointer = vbHourglass
    If txtMBMNumber = "" Then
        MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
    Else
        ScanCode = Trim(txtMBMNumber)
        If Mid(ScanCode, 1, 2) = "LO" Then
            txtMBMNumber = ""
            tmpREcCnt = 2
            Call SearchForm(ScanCode)
        ElseIf Len(txtMBMNumber) = 8 Or Len(txtMBMNumber) = 9 Or (Len(Trim(txtMBMNumber)) = 11 And Mid(Trim(txtMBMNumber), 9, 1) <> "*") Then
            tmpREcCnt = 1
            frmSearchMBM05.Source = 3
            frmSearchMBM05.Show
            tmpREcCnt = frmSearchMBM05.SearchRecCnt
            If frmSearchMBM05.SearchRecCnt < 1 Then
                Unload frmSearchMBM05
            ElseIf frmSearchMBM05.SearchRecCnt = 1 Then
                txtMBMNumber = frmSearchMBM05.tmpMBMNo
                Unload frmSearchMBM05
            End If

        End If
        If tmpREcCnt < 2 Then
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                frmNewMBMOthers.Source = 2
                Call FindMBM
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
            
            If tmpMBMGLIssuance = "N" Then
                MsgBox "GL is NOT allowed for this member", vbOKOnly + vbInformation, DialogBoxTitle
            End If
           
            If Trim(txtMBMStatus) = "S" Then
                MsgBox "Membership has been suspended since " & txtSuspendDate & ".", vbCritical
            End If

        End If
    End If
    
    Screen.MousePointer = Default
    
End Sub

Private Function FindMBM() As Integer
Dim strCount As String
Dim sqlstr2 As String
Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim rs As New ADODB.Recordset
Dim MBMClnRec As New ADODB.Recordset
Dim MBMClnSql As String
    
    MBMClnSql = "Select PLNType, INSType, MBMPolEffDate, MBMPolExpDate from MBMCLNOthers where MBMNumber='" & Trim(txtMBMNumber) & "'"
    MBMClnRec.Open MBMClnSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMClnRec.recordCount Then
        tmpCISIns = Trim(MBMClnRec!InsType)
        tmpCISPln = Trim(MBMClnRec!PLNType)
        tmpCisEffDate = MBMClnRec!MBMPolEffDate
        tmpCisExpDate = MBMClnRec!MBMPolExpDate
        If Trim(cboPLNCode.Text) = "" Then
            HisCisEffDate = MBMClnRec!MBMPolEffDate
            HISCISExpDate = MBMClnRec!MBMPolExpDate
            HISCisPLN = Trim(MBMClnRec!PLNType)
            HISCiSIns = Trim(MBMClnRec!InsType)
        End If
    End If
    MBMClnRec.Close
    Set MBMClnRec = Nothing

    sql = "Exec SearchMBMCovPersonsProc '" & Trim(txtMBMNumber) & "'"
    Set ADOrstMBM = New ADODB.Recordset
    ADOrstMBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If Not ADOrstMBM.EOF Then
            Call showMBM
            lstAdjustmentHis.ListItems.Clear
            lstBenefitDetails.ListItems.Clear
            lstNewBenefitDetails.ListItems.Clear
            lstCaseList.ListItems.Clear
            lstMBMHis.ListItems.Clear
            lstMBMHisTwo.ListItems.Clear
            lstCASHistory.ListItems.Clear
            lstCASHistoryTwo.ListItems.Clear
            lstCovAdd.ListItems.Clear
            lstACCHistory.ListItems.Clear
            lstWorksheet.ListItems.Clear
            
            If Trim(ADOrstMBM!INSCode) = "F" Or Trim(ADOrstMBM!INSCode) = "H" _
            Or Trim(tmpCISIns) = "F" Or Trim(tmpCISIns) = "H" Then
                If Len(Trim(ADOrstMBM!MBMCCoverID)) And Trim(ADOrstMBM!MBMCCoverID) <> "00" Then
                    Call showMBMCov
                End If
            If Trim(tmpMBMFullCover) = "Y" Then
                MsgBox "Please give priority to this member!", vbCritical
            End If
            
            If tmpVIP = True Then
                MsgBox "Please give priority. This member is a VIP!", vbCritical
            End If
           
           
        End If
    
    Else
        MsgBox "No record found!", vbCritical
    End If
    ADOrstMBM.Close
    Set ADOrstMBM = Nothing
    
End Function

Private Function showMBM() As Integer
Dim sqlstr, GetProString As String
Dim MBM As New ADODB.Recordset
Dim sqlstqMBMRef As String
Dim MBMRef As New ADODB.Recordset
Dim MBMNumber As String
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim AnnLmt As New ADODB.Recordset
Dim StrSqlAnnLmt As String
Dim strsql As String
Dim AnnLmtType As New ADODB.Recordset
Dim tmpKidneyAnn As Double
Dim tmpKidneyBal As Double
Dim tmpCancerAnn As Double
Dim tmpCancerBal As Double
Dim tmpHomeNurseAnn As Double
Dim tmpHomeNurseBal As Double
Dim tmpPlnCode As String
Dim StrSqlMsgProm As String
Dim MsgProm As New ADODB.Recordset
Dim tmpMsgprom As String

    txtMBMPremiumRed = ""
    txtMBMMCORed = ""
    txtMBMProviderRedCharges = ""
    txtMBMProviderEffCharges = ""
    txtMCOFeesEff = ""
    txtMBMPremiumEff = ""
    lblPerDisability = ""

    With ADOrstMBM
        txtMBMName = Trim(!MBMName)
        LblHnS = IIf(Len(!HSCLNInd), !HSCLNInd, "H")
        txtMBMIcBcPP = IIf(Len(Trim(!MBMIcBcPp)), Trim(!MBMIcBcPp), "")
        txtMBMPolNo = IIf(Len(Trim(!MBMPolicyNo)), Trim(!MBMPolicyNo), "")
        txtDataStatus = IIf(Len(Trim(!MBMDataStatus)), Trim(!MBMDataStatus), "")
        
        txtMBMStatus = IIf(Len(Trim(!MBMStatus)), Trim(!MBMStatus), "")
        
        If !MBMSuspendStatus = "Y" Then
            txtMBMStatus = "S"
        End If
        
        If !MBMSuspendDate <> "" Then
            txtSuspendDate = Format(!MBMSuspendDate, "dd/mm/yyyy")
        End If
         If !MBMRetireDate <> "" Then
            txtrtDate = Format(!MBMRetireDate, "dd/mm/yyyy")
        End If
        lblEXPStatus = IIf(Len(Trim(!MBMExpiredStatus)), Trim(!MBMExpiredStatus), "")
        
        
        If Len(!INSCode) Then
            cboINSCode = Trim(!INSCode)
            HISCiSIns = Trim(!INSCode)
        End If
        If Len(!PLNCode) Then
            cboPLNCode = Trim(!PLNCode)
            HISCisPLN = Trim(!PLNCode)
        End If
        tmpPlnCode = cboPLNCode
        If Trim(cboHLTCode) <> "S" Then
            If !AnnLmtInd = "N" Then
                tmpAnnLmtIND = "N"
            Else
                tmpAnnLmtIND = "Y"
            End If
        Else
            tmpAnnLmtIND = "Y"
        End If
        
        lblLimitType = ""
        SptLmtInd = "A"
        StrSqlAnnLmt = "Select DisabilityLmt, SuppLimitStatus,MajorMedicalLmt from AnnualLimit where PLNCode = '" & Trim(!PLNCode) & _
                       "' AND INSCode='" & Trim(!INSCode) & _
                       "' And HLTCode = '" & Trim(!HLTCode) & _
                       "' AND AnnualEffDate <= '" & Format(!MBMPayorEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate  desc"
        AnnLmt.Open StrSqlAnnLmt, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        If AnnLmt.EOF = False Then
            SptLmtInd = IIf(Len(Trim(AnnLmt!SuppLimitStatus)), Trim(AnnLmt!SuppLimitStatus), "A")
            If !PLNMajorMedInd = "Y" Then
                lblPerDisability = IIf(IsNumeric(AnnLmt!MajorMedicalLmt), Format(AnnLmt!MajorMedicalLmt, "#,##0.00"), 0)
            Else
                lblPerDisability = IIf(IsNumeric(AnnLmt!DisabilityLmt), Format(AnnLmt!DisabilityLmt, "#,##0.00"), 0)
            End If
            
            If SptLmtInd = "A" Then
                lblLimitType = "A - Shared Limit for family"
            ElseIf SptLmtInd = "B" Then
                lblLimitType = "B - Shared Limit for dependent"
            ElseIf SptLmtInd = "C" Then
                lblLimitType = "C - Separate Limit for family"
            ElseIf SptLmtInd = "D" Then
                lblLimitType = "D - Shared Limit for family"
            End If
        End If
        AnnLmt.Close
        Set AnnLmt = Nothing
        'lblLimitType = ""
        'If Trim(cboINSCode) <> "" Then
        '    If Mid(cboINSCode, 1, 1) = "F" Or Mid(cboINSCode, 1, 1) = "H" Then
        '        lblLimitType = SptLmtInd
        '        strsql = "Select LimitDesc from AnnualLimitType where LimitCode='" & Trim(SptLmtInd) & " '"
        '        AnnLmtType.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        '        If AnnLmtType.EOF = False Then
        '            lblLimitType = SptLmtInd & " - " & AnnLmtType!LimitDesc
        '        End If
        '    End If
        'End If
        txtMBMAdd1 = IIf(Len(Trim(!MBMAddress1)), Trim(!MBMAddress1), "")
        txtMBMAdd2 = IIf(Len(Trim(!MBMAddress2)), Trim(!MBMAddress2), "")
        txtMBMAdd3 = IIf(Len(Trim(!MBMAddress3)), Trim(!MBMAddress3), "")
        txtMBMBatchNo = IIf(Len(Trim(!MBMBatchNo)), Trim(!MBMBatchNo), "")
        txtMBMPostCode = IIf(Len(Trim(!MBMPostCode)), Trim(!MBMPostCode), "")
        cboCTYCode = IIf(Len(Trim(!CTYCode)), Trim(!CTYCode), "")
        cboMBMRenewal = IIf(Len(Trim(!MBMRenewFlag)), Trim(!MBMRenewFlag), "")
        cboMBMSex = IIf(Len(Trim(!MBMSex)), Trim(!MBMSex), "")
        txtMBMBranch = IIf(Len(Trim(!MBMBranch)), Trim(!MBMBranch), "")
        cboRACCode = IIf(Len(Trim(!RACCode)), Trim(!RACCode), "")
        cboSTACode = IIf(Len(Trim(!STACode)), Trim(!STACode), "")
        cboPAYCode = IIf(Len(Trim(!PAYCode)), Trim(!PAYCode), "")
        txtAGECode = IIf(Len(Trim(!AGECode)), Trim(!AGECode), "")
        cboHLTCode = IIf(Len(Trim(!HLTCode)), Trim(!HLTCode), "")
        txtMBMAdultChild = IIf(Len(Trim(!MBMAdultChild)), Trim(!MBMAdultChild), "")
        txtSpecialNote = IIf(Len(Trim(!PLNSpecialNote)), Trim(!PLNSpecialNote), "")
        
        txtMBMBirthDate = Format(!MBMDOB, "dd/mm/yyyy")
        If Trim(!MBMTakeover) = "C" Then
            cboMBMTakeOver.ListIndex = 2
        ElseIf Trim(!MBMTakeover) = "Y" Then
            cboMBMTakeOver.ListIndex = 1
        Else
            cboMBMTakeOver.ListIndex = 0
        End If
        If Trim(!MBMBordxDate) <> "" Then
            txtMBMBordxDate = Format(!MBMBordxDate, "dd/mm/yyyy")
        End If
        If Trim(!MBMPayorEffDate) <> "" Then
            txtMBMPayorEffDate = Format(!MBMPayorEffDate, "dd/mm/yyyy")
            HisCisEffDate = txtMBMPayorEffDate
        End If
        If Trim(!MBMPayorEXPDate) <> "" Then
            txtMBMPayorExpDate = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
            HISCISExpDate = txtMBMPayorExpDate
        End If
        If Trim(!MBMValueDate) <> "" Then
            txtDOR = Format(!MBMValueDate, "dd/mm/yyyy")
        End If
        If Len(!MBMPolicyYear) Then
            txtMBMPolicyYear = Trim(!MBMPolicyYear)
        End If
        If Len(!MBMAvailableLimit) Then
            txtAvailableAnnualLimit = Format(Trim(!MBMAvailableLimit), "#,##0.00")
          ' txtAvailableAnnualLimit = Format(Trim(Bal), "#,##0.00")
        End If
        If Len(!MBMFinalExpiryDate) And !MBMFinalExpiryDate <> "01/01/1900" Then
            txtFinalExpDate = Format(!MBMFinalExpiryDate, "dd/mm/yyyy")
        End If
        'If Trim(!PAYGRPCode) = "ETS" Then
        '    Label1(34).Caption = "Next Due Date"
            If Len(!MBMNextDueDate) Then
                txtPolNxDue = Format(!MBMNextDueDate, "dd/mm/yyyy")
            End If
        'Else
        '    Label1(34).Caption = "Policy Eff Date"
            If Len(!MBMPolEffDate) Then
                txtPolEffDate = Format(!MBMPolEffDate, "dd/mm/yyyy")
            End If
        'End If
        
        tmpMBMGLIssuance = IIf(Len(!MBMGLIssuance), !MBMGLIssuance, "")
    End With
    
    If ADOrstMBM!MBMStatus = "C" Then
        MBMNumber = Trim(ADOrstMBM!MBMNumber)
        sqlstqMBMRef = "Select * from MBMRefund where MBMNumber =  '" & MBMNumber & "'"
        MBMRef.Open sqlstqMBMRef, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMRef.recordCount Then
            With MBMRef
                If MBMRef!MBMPremReduction <> "" Then
                    txtMBMPremiumRed = Format(Round(!MBMPremReduction, 0), "#,##0.00")
                End If
                If Len(!MBMMCORefund) Then
                    txtMBMMCORed = Format(Round(!MBMMCORefund, 0), "#,##0.00")
                End If
                If Len(!MBMProviderRef) Then
                    txtMBMProviderRedCharges = Format(Round(!MBMProviderRef, 1), "#,##0.00")
                End If
                If Len(!MBMProviderEff) Then
                    txtMBMProviderEffCharges = Format(Round(!MBMProviderEff, 1), "#,##0.00")
                End If
                If Len(!MBMMCOEffective) Then
                    txtMCOFeesEff = Format(Round(!MBMMCOEffective, 0), "#,##0.00")
                End If
                If Len(!MBMPREMEffective) Then
                    txtMBMPremiumEff = Format(Round(!MBMPREMEffective, 0), "#,##0.00")
                End If
            End With
        End If
        MBMRef.Close
        Set MBMRef = Nothing
    End If
    Call showMBMTwo
    Call showMBMOthers
    Call showMBMOthersTwo
    Call ShowExclusion
    Call ShowRemarks
    If tmpAnnLmtIND = "Y" Then
        Call CallLimit
    Else
        txtAvailableAnnualLimit = "FULL"
    End If

    
    Call MBMBenefitLimit(Trim(ADOrstMBM!MBMNumber), "00", tmpKidneyAnn, tmpKidneyBal, tmpCancerAnn, tmpCancerBal, tmpHomeNurseAnn, tmpHomeNurseBal)
   ' lblKidneyAnn = Format(tmpKidneyAnn, "#,##0.00")
    'lblKidneyBal = Format(tmpKidneyBal, "#,##0.00")
    lblCancerAnn = Format(tmpCancerAnn, "#,##0.00")
    lblCancerBal = Format(tmpCancerBal, "#,##0.00")
    lblHomeNurseAnn = Format(tmpHomeNurseAnn, "#,##0.00")
    lblHomeNurseBal = Format(tmpHomeNurseBal, "#,##0.00")
    
    
    If Len(txtSpecialNote) Then
        txtSpecialNote.BackColor = vbYellow
        MsgBox txtSpecialNote, vbOKOnly + vbInformation
    End If
    
    '==================Message Prom==============================================
    
    StrSqlMsgProm = "Select MessagePromt,PLNPolicyWording,productname,PLNRBType,PLNDescription  from PLN where PLNCode = '" & Trim(tmpPlnCode) & "'"
    MsgProm.Open StrSqlMsgProm, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    If MsgProm.EOF = False Then
        If MsgProm!MessagePromt <> "" Then
            tmpMsgprom = MsgProm!MessagePromt
        End If
        tmpPolicyWording = IIf(Trim(Len(MsgProm!PLNPolicyWording)), MsgProm!PLNPolicyWording, "")
        'txtProductName = IIf(Trim(Len(MsgProm!ProductName)), MsgProm!ProductName, "")
        txtProductName = txtProductName & "(" & IIf(Trim(Len(MsgProm!PLNDescription)), MsgProm!PLNDescription, "") & ")"
        txtRBType = IIf(Trim(Len(MsgProm!PLNRBType)), MsgProm!PLNRBType, "")
    Else
        tmpMsgprom = ""
        tmpPolicyWording = ""
        txtProductName = ""
        txtRBType = ""
    End If
    MsgProm.Close
    Set MsgProm = Nothing
    
    If tmpMsgprom <> "" Then
        MsgBox tmpMsgprom, vbOKOnly + vbInformation
    End If
End Function

Private Sub showMBMCov()
Dim MemberConveredPerson As ListItem
Dim tmpKidneyAnn As Double
Dim tmpKidneyBal As Double
Dim tmpCancerAnn As Double
Dim tmpCancerBal As Double
Dim tmpHomeNurseAnn As Double
Dim tmpHomeNurseBal As Double
            
    Do While Not ADOrstMBM.EOF
        With ADOrstMBM
            Set MemberConveredPerson = lstCovAdd.ListItems.Add(, , IIf(Trim(!MBMCName) <> "", !MBMCName, ""))
            'MemberConveredPerson.SubItems(44) = IIf(Len(!MBMCPAYSeqNo), Trim(!MBMCPAYSeqNo), "")
            'MemberConveredPerson.SubItems(45) = IIf(Len(!MBMCClientNO), Trim(!MBMCClientNO), "")
            'MemberConveredPerson.SubItems(46) = IIf(Len(!MBMCClientID), Trim(!MBMCClientID), "")
            'MemberConveredPerson.SubItems(47) = IIf(Len(!MBMCIcBcPp2nd), Trim(!MBMCIcBcPp2nd), "")
            MemberConveredPerson.SubItems(1) = IIf(Len(Trim(!MBMCSalutation)), Trim(!MBMCSalutation), "")
            MemberConveredPerson.SubItems(2) = IIf(Len(Trim(!MBMCICBCPPNo)), Trim(!MBMCICBCPPNo), "")
            If Len(Trim(!MBMCDOB)) Then
                MemberConveredPerson.SubItems(3) = !MBMCDOB
            End If
            MemberConveredPerson.SubItems(4) = IIf(Len(Trim(!MBMCSex)), Trim(!MBMCSex), "")
            MemberConveredPerson.SubItems(5) = IIf(Len(Trim(!MBMCAdultChild)), Trim(!MBMCAdultChild), "")
            MemberConveredPerson.SubItems(6) = IIf(Len(Trim(!MBMCRELCode)), Trim(!MBMCRELCode), "")
            MemberConveredPerson.SubItems(7) = IIf(Len(Trim(!MBMCStatus)), Trim(!MBMCStatus), "")
            MemberConveredPerson.SubItems(8) = IIf(Len(Trim(!MBMCAllergic)), Trim(!MBMCAllergic), "")
            MemberConveredPerson.SubItems(9) = IIf(Len(Trim(!MBMCExclusion)), Trim(!MBMCExclusion), "")
            MemberConveredPerson.SubItems(10) = IIf(Len(Trim(!MBMCCoverID)), Trim(!MBMCCoverID), "")
            If Len(Trim(!MBMCAnnualLimit)) Then
                MemberConveredPerson.SubItems(12) = !MBMCAnnualLimit
            End If
            If Len(Trim(!MBMCAvailableLimit)) Then
                MemberConveredPerson.SubItems(13) = !MBMCAvailableLimit
            End If
            MemberConveredPerson.SubItems(14) = IIf(Len(!MBMCPLNCode), Trim(!MBMCPLNCode), cboPLNCode)
            MemberConveredPerson.SubItems(15) = IIf(Len(!MBMCEffDate), !MBMCEffDate, txtMBMPayorEffDate)
            MemberConveredPerson.SubItems(16) = IIf(Len(!MBMCExpDate), !MBMCExpDate, txtMBMPayorExpDate)
            MemberConveredPerson.SubItems(17) = IIf(IsNumeric(!MBMCUndExcess), !MBMCUndExcess, 0)
            If Len(Trim(!MBMCPremium)) Then
                MemberConveredPerson.SubItems(18) = !MBMCPremium
            End If
            MemberConveredPerson.SubItems(19) = IIf(Len(Trim(!MBMCPrevPLNCode)), Trim(!MBMCPrevPLNCode), "")
            If Len(Trim(!MBMCRBAmt)) Then
                MemberConveredPerson.SubItems(20) = !MBMCRBAmt
            End If
            MemberConveredPerson.SubItems(21) = IIf(Len(Trim(!MBMCPAYRemarks)), Trim(!MBMCPAYRemarks), "")
            MemberConveredPerson.SubItems(22) = IIf(Len(Trim(!MBMCPrevInsCom)), Trim(!MBMCPrevInsCom), "")
            MemberConveredPerson.SubItems(23) = IIf(Len(Trim(!MBMCCLMNonPayFr)), Trim(!MBMCCLMNonPayFr), "")
            MemberConveredPerson.SubItems(24) = IIf(Len(Trim(!MBMCCLMNonPayTo)), Trim(!MBMCCLMNonPayTo), "")
            If Len(!MBMCPayorPremGross) Then
                MemberConveredPerson.SubItems(25) = !MBMCPayorPremGross
            End If
            If Len(!MBMCPayorPremNet) Then
                MemberConveredPerson.SubItems(26) = !MBMCPayorPremNet
            End If
            If Len(!MBMCPayorMCOGross) Then
                MemberConveredPerson.SubItems(27) = !MBMCPayorMCOGross
            End If
            If Len(!MBMCPayorMCONet) Then
                MemberConveredPerson.SubItems(28) = !MBMCPayorMCONet
            End If
            If Len(!MBMCRenewalDisc) Then
                MemberConveredPerson.SubItems(29) = !MBMCRenewalDisc
            End If
            If Len(!MBMCPolicyDisc) Then
                MemberConveredPerson.SubItems(30) = !MBMCPolicyDisc
            End If
            If Len(!MBMCFamDiscAmt) Then
                MemberConveredPerson.SubItems(31) = !MBMCFamDiscAmt
            End If
            If Len(!MBMCRenewDiscAmt) Then
                MemberConveredPerson.SubItems(32) = !MBMCRenewDiscAmt
            End If
            If Len(!MBMCLoadingPrem) Then
                MemberConveredPerson.SubItems(33) = !MBMCLoadingPrem
            End If
            Call MBMBenefitLimit(Trim(ADOrstMBM!MBMNumber), !MBMCCoverID, tmpKidneyAnn, tmpKidneyBal, tmpCancerAnn, tmpCancerBal, tmpHomeNurseAnn, tmpHomeNurseBal)
            MemberConveredPerson.SubItems(35) = tmpKidneyAnn
            MemberConveredPerson.SubItems(36) = tmpKidneyBal
            MemberConveredPerson.SubItems(37) = tmpCancerAnn
            MemberConveredPerson.SubItems(38) = tmpCancerBal
            MemberConveredPerson.SubItems(39) = tmpHomeNurseAnn
            MemberConveredPerson.SubItems(40) = tmpHomeNurseBal
            MemberConveredPerson.SubItems(41) = IIf(Len(!MBMCPayNo), Trim(!MBMCPayNo), "")
            MemberConveredPerson.SubItems(42) = IIf(Len(!MBMCPaySeqNo), Trim(!MBMCPaySeqNo), "")
            MemberConveredPerson.SubItems(43) = IIf(Len(!MBMCClientNo), Trim(!MBMCClientNo), "")
            MemberConveredPerson.SubItems(44) = IIf(Len(!MBMCClientID), Trim(!MBMCClientID), "")
            If Len(!MBMCBordxDate) Then
                MemberConveredPerson.SubItems(45) = Format(!MBMCBordxDate, "dd/mm/yyyy")
            End If
            If Len(!MBMCBatchNo) Then
                MemberConveredPerson.SubItems(46) = !MBMCBatchNo
            End If
        End With
        ADOrstMBM.MoveNext
    Loop
End Sub

Private Sub CmdExit_Click()
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    
    Call ClearForm(frmNewMBMOthers)
    Unload Me
End Sub

Private Sub cmdSearchMBM_Click()
    'lblTopupStatus.Visible = False
    bDataStatus = True
    frmSearchMBM.Source = 3
    frmSearchMBM.Show vbModal
End Sub


Private Sub lstCovAdd_Click()
    If lstCovAdd.ListItems.Count Then
        txtCovName = lstCovAdd.SelectedItem.Text
        lblCovName = "(" & lstCovAdd.SelectedItem.Text & ")"
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        lblCovAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        lblCovAvLmt = Format(lstCovAdd.SelectedItem.SubItems(13), "###,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(19)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(20)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(21)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(22)
        txtUndExcess = lstCovAdd.SelectedItem.SubItems(17)
        'TxtSuppClmNonPayFr.mask = ""
        'TxtSuppClmNonPayFr.Text = ""
        'TxtSuppClmNonPayFr.mask = "##/##/####"
        'TxtSuppClmNonPayTo.mask = ""
        'TxtSuppClmNonPayTo.Text = ""
        'TxtSuppClmNonPayTo.mask = "##/##/####"
        'If Len(lstCovAdd.SelectedItem.SubItems(23)) Then
        '    TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(23)
        'End If
        'If Len(lstCovAdd.SelectedItem.SubItems(24)) Then
        '    TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(24)
        'End If
        txtSuppPLNCode = IIf(Len(lstCovAdd.SelectedItem.SubItems(14)), lstCovAdd.SelectedItem.SubItems(14), cboPLNCode)
        txtSuppEffDate = IIf(Len(lstCovAdd.SelectedItem.SubItems(15)), lstCovAdd.SelectedItem.SubItems(15), txtMBMPayorEffDate)
        txtSuppExpDate = IIf(Len(lstCovAdd.SelectedItem.SubItems(16)), lstCovAdd.SelectedItem.SubItems(16), txtMBMPayorExpDate)
        txtPayMBMNo = IIf(Len(lstCovAdd.SelectedItem.SubItems(41)), Trim(lstCovAdd.SelectedItem.SubItems(41)), "")
        txtPaySeqID = IIf(Len(lstCovAdd.SelectedItem.SubItems(42)), Trim(lstCovAdd.SelectedItem.SubItems(42)), "")
        txtClientNo = IIf(Len(lstCovAdd.SelectedItem.SubItems(43)), Trim(lstCovAdd.SelectedItem.SubItems(43)), "")
        txtClientID = IIf(Len(lstCovAdd.SelectedItem.SubItems(44)), Trim(lstCovAdd.SelectedItem.SubItems(44)), "")

        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(25)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(26)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(27)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(28)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(29)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(30)
        txtPolicyDiscAmt = lstCovAdd.SelectedItem.SubItems(31)
        txtRenewalDiscAmt = lstCovAdd.SelectedItem.SubItems(32)
        txtLoadingPrem = lstCovAdd.SelectedItem.SubItems(33)
        lblCovKidneyAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(35), "#,##0.00")
        lblCovKidneyBalLmt = Format(lstCovAdd.SelectedItem.SubItems(36), "#,##0.00")
        lblCovCancerAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(37), "#,##0.00")
        lblCovCancerBalLmt = Format(lstCovAdd.SelectedItem.SubItems(38), "#,##0.00")
        lblCovHomeNurseAnnLmt = Format(lstCovAdd.SelectedItem.SubItems(39), "#,##0.00")
        lblCovHomeNurseBalLmt = Format(lstCovAdd.SelectedItem.SubItems(40), "#,##0.00")
    End If
End Sub

'  ################# END Suppplementary ######################

' ################################ Enable Tab / button   ######
Private Sub EnableEntries(status As Boolean)
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

Private Sub InsertPrem()
Dim MBMInsertPrem As Double
Dim MBMPrem As Double
Dim PAYExpDate As String

MBMPrem = txtMCOFees
PAYExpDate = txtMBMPayorExpDate

    MBMInsertPrem = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMBMPremium, "#,##0.00"))
    txtMBMPremium = ""
    txtMBMPremium = Format(MBMInsertPrem, "#,##0.00")
End Sub

Private Sub GetAccountSummary()
    Dim Acc As New ADODB.Recordset
    Dim strsql As String
    Dim masterlist As ListItem
        
    If lstMBMAcc.ListItems.Count = 0 Then
        lstMBMAcc.ListItems.Clear
        strsql = " select totalApprove, totalPayable2Hos, " & _
                " totalPaid2Hos, totalPayable2Mbm, totalPaidByMbm, MBMPatientCOVID, MBMPaidAmt," & _
                " totalPaid2MBM , totalExcess , totalReimburse from view_MBMTotal_summary where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Acc.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If Not Acc.EOF Then
            Do Until Acc.EOF
            With Acc
            Set masterlist = lstMBMAcc.ListItems.Add(, , txtMBMNumber)
                masterlist.SubItems(1) = Format(!MBMPatientCovID)
                masterlist.SubItems(2) = Format(!totalApprove, "#,##0.00")
                masterlist.SubItems(3) = Format(!totalPayable2Hos, "#,##0.00")
                If Trim(!totalPaid2Hos) <> "" Then
                    masterlist.SubItems(4) = Format(!totalPaid2Hos, "#,##0.00")
                End If
                If IsNumeric(masterlist.SubItems(4)) = False Then
                    masterlist.SubItems(4) = 0
                End If
                
                ' Out standing hos
                'txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,#.00;(#,#.00)")
                
                masterlist.SubItems(5) = Format(!totalPayable2Mbm, "#,##0.00;(#,##0.00)")
                If IsNumeric(masterlist.SubItems(5)) = False Then
                    masterlist.SubItems(5) = 0
                End If
                If Trim(!MBMPaidAmt) <> "" Then
                    masterlist.SubItems(6) = Format(!MBMPaidAmt, "#,##0.00")
                End If
                If IsNumeric(masterlist.SubItems(6)) = False Then
                    masterlist.SubItems(6) = 0
                End If
                If Trim(!totalPaidByMbm) <> "" Then
                    masterlist.SubItems(7) = Format(!totalPaidByMbm, "#,##0.00")
                End If
                If IsNumeric(masterlist.SubItems(7)) = False Then
                    masterlist.SubItems(7) = 0
                End If
                'txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
                masterlist.SubItems(8) = Format(!TotalReimburse, "#,##0.00")
                masterlist.SubItems(9) = Format(!TotalExcess, "#,##0.00;(#,##0.00)")
                masterlist.SubItems(10) = Format(!totalApprove, "#,##0.00;(#,##0.00)")
                'masterlist.SubItems(11) = Format(!CLMExcess, "#,##0.00;(#,##0.00)")
    
            End With
            Acc.MoveNext
            Loop
        End If
        Acc.Close
        Set Acc = Nothing
    End If
End Sub

Private Sub GetAccHisListing()
    Dim Acc As New ADODB.Recordset
    Dim strsql As String
    Dim masterlist As ListItem
    Dim RecAmt As Double
    Dim MBMAmt As Double
    
        lstACCHistory.ListItems.Clear
        strsql = " select CASNumber, ClaimVersion,DebitCreditDate,DebitCreditRefNo, " & _
                " CLMPayableByMedix2M, ReceiptNo, ReceiptDate, ReceiptAmt," & _
                " MBMIcBcPp,MBMNumber,CASDateAdmitted,CaseRegDate,MBMPayorEffDate from VIEW_Account_History where MBMIcBcPp = '" & Trim(txtMBMIcBcPP) & "'"
        Acc.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        If Not Acc.EOF Then
            Do Until Acc.EOF
            With Acc
            Set masterlist = lstACCHistory.ListItems.Add(, , !MBMNumber)
                masterlist.SubItems(1) = !MBMPayorEffDate
                masterlist.SubItems(2) = !CASDateAdmitted
                masterlist.SubItems(3) = !CaseRegDate
                masterlist.SubItems(4) = !CasNumber & "-" & !claimversion
                
                If (!CLMPayableByMedix2M) <> 0 Then
                    masterlist.SubItems(5) = !DebitCreditRefNo
                    If Len(!DebitCreditDate) Then
                        masterlist.SubItems(6) = !DebitCreditDate
                    End If
                End If
                
                'f Len(!DebitCreditDate) Then
                '    masterlist.SubItems(6) = !DebitCreditDate
                'End If
                
                masterlist.SubItems(7) = Format(!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                
                MBMAmt = IIf(IsNumeric(!CLMPayableByMedix2M) = True, !CLMPayableByMedix2M, 0)
                If Len(!RECEIPTNO) Then
                    masterlist.SubItems(8) = !RECEIPTNO
                End If
                If Len(!ReceiptDate) Then
                    masterlist.SubItems(9) = !ReceiptDate
                End If
                masterlist.SubItems(10) = Format(IIf(IsNumeric(!ReceiptAmt) = True, !ReceiptAmt, 0), "#,##0.00;(#,##0.00)")
                RecAmt = IIf(IsNumeric(!ReceiptAmt) = True, !ReceiptAmt, 0)
                masterlist.SubItems(11) = Format(CDbl(RecAmt) + CDbl(MBMAmt), "#,##0.00;(#,##0.00)")
    
            End With
            Acc.MoveNext
            Loop
        End If
        Acc.Close
        Set Acc = Nothing
    'End If
End Sub

Private Sub ShowExclusion()
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
    Dim rst As New ADODB.Recordset
    Dim strsql As String
    Dim tmpMBMName As String
        strAppendCase = 1
        tmpMBMName = Trim(Replace(txtMBMName, "'", "''"))
        
        strAppend = strAppend & "AND PAYCode = '" & Trim(cboPAYCode) & "'" & _
                " And MBMDOB = '" & Format(txtMBMBirthDate, "yyyy/mm/dd") & "'" & _
                " And MBMName = '" & Trim(tmpMBMName) & "'"
        
        'strAppend = strAppend & "AND MBMPolicyNo = '" & Trim(txtMBMPolNo) & "' and SUBSTRING (MBMNumber ,1,2)= '" + cboPAYCode + "'"

         
        'strsql = ""
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "MBM_GetExcRemarks"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 255, strAppendCase)
        cmd.Parameters.Append Prm2
        rst.CursorLocation = adUseClient
        rst.CursorType = adOpenForwardOnly
        rst.CacheSize = 100
        Set rst = cmd.Execute
            Do While Not rst.EOF
                If Trim(txtMBMExclusion) = "" Then
                    txtMBMExclusion = IIf(Len(Trim(rst!MBMExclusion)), rst!MBMNumber & " - " & Trim(rst!MBMExclusion), "")
                Else
                    txtMBMExclusion = txtMBMExclusion & vbNewLine & vbNewLine & IIf(Len(Trim(rst!MBMExclusion)), rst!MBMNumber & " - " & Trim(rst!MBMExclusion), "")
                End If
            rst.MoveNext
            Loop
        rst.Close
        Set rst = Nothing
    'End If
    
        'strsql = " SELECT * from View_Get_MemberExclusion " & _
        '        " where PAYCode = '" & Trim(cboPAYCode) & "' " & _
        '        " And MBMDOB = '" & Format(txtMBMBirthDate, "yyyy/mm/dd") & "' " & _
        '        " And INSCode = '" & Trim(cboINSCode) & "'" & _
        '        " And MBMName = '" & Trim(tmpMBMName) & "'"
        'rst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
End Sub
Private Sub ShowRemarks()
    Dim cmd As New ADODB.Command
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim strAppendCase As String
    Dim strAppend As String

    Dim rst As New ADODB.Recordset
    Dim strsql As String
    Dim tmpMBMName As String
        tmpMBMName = Trim(Replace(txtMBMName, "'", "''"))
        strsql = ""

        'strAppend = strAppend & "AND PAYCode = '" & Trim(cboPAYCode) & "'" & _
        '        " And MBMDOB = '" & Format(txtMBMBirthDate, "yyyy/mm/dd") & "'" & _
        '        " And INSCode = '" & Trim(cboINSCode) & "'" & _
        '        " And MBMName = '" & Trim(tmpMBMName) & "'"
        'strAppend = strAppend & "AND MBMIcBcPp = '" & Trim(txtMBMIcBcPP) & "' and SUBSTRING (MBMNumber ,1,2)= '" + cboPAYCode + "'"
        strAppend = strAppend & "AND MBMNumber = '" & Trim(txtMBMNumber) & "' "
        strAppendCase = 2
        
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "MBM_GetExcRemarks"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 255, strAppendCase)
        cmd.Parameters.Append Prm2
        rst.CursorLocation = adUseClient
        rst.CursorType = adOpenForwardOnly
        rst.CacheSize = 100
        Set rst = cmd.Execute
        
            Do While Not rst.EOF
                If Trim(txtRemarks) = "" Then
                    txtRemarks = IIf(Len(Trim(rst!MBMRemarks)), Trim(rst!MBMRemarks) & " - " & rst!MBMNumber, "")
                Else
                    txtRemarks = txtRemarks & vbNewLine & IIf(Len(Trim(rst!MBMRemarks)), Trim(rst!MBMRemarks) & " - " & rst!MBMNumber, "")
                End If
            rst.MoveNext
            Loop
        rst.Close
        Set rst = Nothing
    'End If
End Sub

Private Sub txtMBMNumber_GotFocus()
cmdShow.Default = True
End Sub

Private Sub txtMBMNumber_LostFocus()
txtMBMNumber = ChkStr(txtMBMNumber)
cmdShow.Default = False
End Sub

Private Function showMBMTwo() As Integer
Dim sqlstrTwo As String
Dim MBMTwo As New ADODB.Recordset
    
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    Call ClearForm(frmNewMBMOthers)
    
        With ADOrstMBM
            If Trim(!MBMSalutation) <> "" Then
                cboMBMSalutation = Trim(!MBMSalutation)
            End If
            'If Trim(!SRVCodeList) <> "" Then
            '    cboSRVCode = Trim(!SRVCodeList)
            'End If
            If Trim(!NATCode) <> "" Then
                cboNATCode = Trim(!NATCode)
            End If
            If Trim(!MBMTelnoH) <> "" Then
                frmNewMBMOthers.txtMBMTelNoH = Trim(!MBMTelnoH)
            End If
            
            If Trim(!MBMTelNoM) <> "" Then
                frmNewMBMOthers.txtMBMTelNoM = Trim(!MBMTelNoM)
            End If
            
            If Trim(!MBMTelNoO) <> "" Then
                frmNewMBMOthers.txtMBMTelNoO = Trim(!MBMTelNoO)
            End If
            
            If Len(!MBMEmail) Then
                frmNewMBMOthers.txtMBMEmail = !MBMEmail
            End If
            
            If Len(!LGNCode) Then
                frmNewMBMOthers.cboLGNCode = !LGNCode
            End If
            
            frmNewMBMOthers.txtagentEmail = IIf(Len(!MBMAgentEmail), !MBMAgentEmail, "")
            frmNewMBMOthers.txtAgent = IIf(Len(!MBMAgentName), !MBMAgentName, "")
            frmNewMBMOthers.txtAgntContact = IIf(Len(!MBMAgentPhoneNo), !MBMAgentPhoneNo, "")
               
            'If Len(!MBMSendingMode) Then
            frmNewMBMOthers.txtSendingMode = IIf(Len(!MBMSendingMode), !MBMSendingMode, "")
            'End If
            
            'If Len(!MBMCardDate) Then
            frmNewMBMOthers.txtCardSentDate = IIf(Len(!MBMCardDate), !MBMCardDate, "__/__/____")
            'End If
            
            txtMgmtNotes = IIf(Len(!MBMMmgtNotes), !MBMMmgtNotes, "")
            
            txtEndtDate = IIf(Len(!MBMENDtEffDate), !MBMENDtEffDate, "__/__/____")
            
            txtMBMAppName = IIf(Len(!MBMAppName), !MBMAppName, "")
        End With
End Function

Private Function showMBMOthers() As Integer
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
Dim TempMBMOthers As New ADODB.Recordset
Dim strMBMOthers As String
    
        With ADOrstMBM
            
            frmNewMBMOthers.Label4 = Trim(MBMNumber)
            'frmplandetails.txtgrpcompany = IIf(Len(Trim(!MBMGroupCompany)), Trim(!MBMGroupCompany), "")
            If Trim(!MBMCLMNonPayFr) <> "" Then
                frmNewMBMOthers.TxtClmNonPayFr = Format(!MBMCLMNonPayFr, "dd/mm/yyyy")
            End If
            If Trim(!MBMCLMNonPayTo) <> "" Then
                frmNewMBMOthers.TxtClmNonPayTo = Format(!MBMCLMNonPayTo, "dd/mm/yyyy")
            End If
            
            If Trim(!LabelStatus) = "Y" Then
                frmNewMBMOthers.Check1.Value = 1
            End If
            If Len(Trim(cboPAYCode)) <> 0 And Len(cboINSCode) <> 0 And txtAGECode <> "" And Len(cboHLTCode) <> 0 And txtMBMPayorEffDate <> "__/__/____" Then
                'txtMBMPremium = Format(IIf(IsNull(!MBMPremium) = True, Format(GetPremium(Trim(cboINSCode), Trim(cboPAYCode), Trim(txtAGECode), Trim(cboPLNCode), Trim(cboHLTCode), txtMBMPayorEffDate), "#,##0.00"), !MBMPremium), "#,##0.00")
                txtMBMPremium = IIf(IsNumeric(!MBMPremium), Format(!MBMPremium, "#,##0.00"), 0)
            End If
            If Len(Trim(cboPAYCode)) <> 0 And Len(cboINSCode) <> 0 And txtAGECode <> "" And Len(cboHLTCode) <> 0 And txtMBMPayorEffDate <> "__/__/____" Then
                txtMCOFees = Format(IIf(IsNull(!MBMMCOFee) = True, Format(GetMCO(Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild, txtMBMPayorEffDate, Mid(cboPLNCode, 1, 4)), "#,##0.00"), !MBMMCOFee), "#,##0.00")
            End If
            
            'If Len(cboSRVCode) <> 0 And Len(cboINSCode) <> 0 And Len(cboHLTCode) <> 0 And txtMBMAdultChild <> "" Then
            '    txtMBMProviderCharges = Format(IIf(IsNull(!MBMAccessFee) = True, Format(GetProviderCharges(Trim(Mid(cboSRVCode, 1, 2)), Trim(cboINSCode), Trim(cboHLTCode), Trim(txtMBMAdultChild)), "####0.00"), !MBMAccessFee), "#,##0.00")
            'End If
            'txtRemarks = IIf(Len(Trim(!MBMRemarks)), Trim(!MBMRemarks), "")
            txtMBMAgentCode = IIf(Len(Trim(!MBMAgentCode)), Trim(!MBMAgentCode), "")
            txtDeptName = IIf(Len(Trim(!MBMDepartment)), Trim(!MBMDepartment), "")
            txtMBMGroupCompany = IIf(Len(Trim(!MBMGroupCompany)), Trim(!MBMGroupCompany), "")
            txtMBMEmployeeNo = IIf(Len(Trim(!MBMEmployeeNo)), Trim(!MBMEmployeeNo), "")
            If Trim(!MBMInstallment) <> "" Then
                If Trim(!MBMInstallment) = "A" Then
                    CboInsMode.ListIndex = 0
                ElseIf Trim(!MBMInstallment) = "Q" Then
                    CboInsMode.ListIndex = 1
                ElseIf Trim(!MBMInstallment) = "S" Then
                    CboInsMode.ListIndex = 2
                ElseIf Trim(!MBMInstallment) = "M" Then
                    CboInsMode.ListIndex = 3
                End If
            End If
            If Trim(!MBMFirstJoinedDate) <> "" Then
                txtMBMDateJoin = Format(!MBMFirstJoinedDate, "dd/mm/yyyy")
            End If
            frmNewMBMOthers.TxtPOLSubNo = IIf(Len(Trim(!MBMPolSubNo1)), Trim(!MBMPolSubNo1), "")
            frmNewMBMOthers.txtMBMAllergic = IIf(Len(Trim(!MBMAllergic)), Trim(!MBMAllergic), "")
            frmNewMBMOthers.txtPLNCode = IIf(Len(Trim(!MBMNewPlanCode)), Trim(!MBMNewPlanCode), "")
            txtPLNCode = IIf(Len(Trim(!MBMNewPlanCode)), Trim(!MBMNewPlanCode), "")
            If Len(!MBMNewPolEffDate) Then
                frmNewMBMOthers.txtNewPolEffDate = Format(!MBMNewPolEffDate, "dd/mm/yyyy")
                txtNewPolEffDate = Format(!MBMNewPolEffDate, "dd/mm/yyyy")
            End If
            
            If Len(!MBMPolicyDisc) Then
                frmNewMBMOthers.txtPolicyDisc = !MBMPolicyDisc
            End If
            
            If Len(!MBMRenewalDisc) Then
                frmNewMBMOthers.txtRenewalDisc = !MBMRenewalDisc
            End If
            
            If Len(!MBMPolCondition) Then
                frmNewMBMOthers.cboPOLCondition = !MBMPolCondition
            End If
            
            If Len(!MBMUndExcess) Then
                frmNewMBMOthers.txtUndExcess = !MBMUndExcess
            End If
            frmNewMBMOthers.txtEndtRefNo = IIf(Len(!MBMENDtRefNo), !MBMENDtRefNo, "")

            If Len(!MBMGuaRenewal) Then
                If Trim(!MBMGuaRenewal) = "1" Then
                    CboGuaRenewal.ListIndex = 1
                ElseIf Trim(!MBMGuaRenewal) = "3" Then
                    CboGuaRenewal.ListIndex = 2
                End If
            End If
            tmpMBMFullCover = IIf(Len(!MBMFullCover), !MBMFullCover, "")
            txtCostCenter = IIf(Len(!MBMCostCenter), !MBMCostCenter, "")
            txtMBMPosition = IIf(Len(!MBMPosition), !MBMPosition, "")
            txtGrpCat = IIf(Len(!MBMGroupCategory), !MBMGroupCategory, "")
            txtDivision = IIf(Len(!MBMDivision), !MBMDivision, "")
            txtSubDivision = IIf(Len(!MBMSubDivision), !MBMSubDivision, "")
            If Trim(!MBMVIP) = "Y" Then
                tmpVIP = True
            Else
                tmpVIP = False
            End If
            
            
            txtSubsidiary = IIf(Len(!Subsidiary), !Subsidiary, "")
            
            txtMBMPymtDate = IIf(Len(!MBMPaymentDate), !MBMPaymentDate, "__/__/____")
            
        End With
End Function

Public Function showMBMOthersTwo() As Integer
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
Dim TempMBMOthers As New ADODB.Recordset
Dim strMBMOthers As String
    

    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    
        MBMNumber = txtMBMNumber
        With ADOrstMBM
            
            frmNewMBMOthers.Label4 = Trim(MBMNumber)
            
            If Len(!MBMPrevPLNCode) Then
                frmNewMBMOthers.txtMBMPrevPlan = !MBMPrevPLNCode
            End If
            
            If Len(!MBMRBAmt) Then
                frmNewMBMOthers.txtMBMRBAmt = !MBMRBAmt
            End If
            
            If Len(!MBMPAYRemarks) Then
                txtMBMPAYRemarks = !MBMPAYRemarks
            End If
                        
            If Len(!MBMPrevInsCom) Then
                frmNewMBMOthers.txtMBMPrevINSCom = !MBMPrevInsCom
            End If
        
            If Len(!MBMCoPayRB) Then
                frmNewMBMOthers.cboCOPayRB = !MBMCoPayRB
            End If
            
            If Len(!MBMGenWP) Then
                frmNewMBMOthers.TxtGenWP = !MBMGenWP
            End If
            
            If Len(!MBMPreExWP) Then
                frmNewMBMOthers.txtPREWP = !MBMPreExWP
            End If
            
            If Len(!MBMSpeIllWP) Then
                frmNewMBMOthers.TxtSpeWP = !MBMSpeIllWP
            End If
            
            If Len(!MBMStaffDec) Then
                frmNewMBMOthers.cboStaffDisc = !MBMStaffDec
            End If
            
            If Len(!MBMGenEndWP) Then
                frmNewMBMOthers.txtGENWPEnd = !MBMGenEndWP
            End If
            
            If Len(!MBMPreExEndWP) Then
                frmNewMBMOthers.txtPREWPEnd = !MBMPreExEndWP
            End If
            
            If Len(!MBMSpeIllEndWP) Then
                frmNewMBMOthers.txtSPEWPEnd = !MBMSpeIllEndWP
            End If
                            
            If Len(!MBMBTBG) Then
                frmNewMBMOthers.CboBTBG = !MBMBTBG
            End If
            
            If Len(!MBMOccupation) Then
                frmNewMBMOthers.cboMBMOccupation = !MBMOccupation
            End If
            
            If Len(!MBMWorkNature) Then
                frmNewMBMOthers.txtMBMWorkNature = !MBMWorkNature
            End If
            
            If Len(!MBMWeight) Then
                frmNewMBMOthers.txtMBMWeight = !MBMWeight
            End If
            
            If Len(!MBMHeight) Then
                frmNewMBMOthers.txtMBMHeight = !MBMHeight
            End If
                        
            If Len(!MBMBlood) Then
                frmNewMBMOthers.cboMBMBlood = !MBMBlood
            End If
                         
            If Len(!MBMSmoking) Then
                frmNewMBMOthers.cboMBMSmoking = !MBMSmoking
            End If
            
            If Len(!MBMAlcohol) Then
                frmNewMBMOthers.cboMBMAlcohol = !MBMAlcohol
            End If
            
            If Len(!MBMPayorPremGross) Then
                frmNewMBMOthers.txtPayPrem = !MBMPayorPremGross
            End If
            
            If Len(!MBMPayorPremNet) Then
                frmNewMBMOthers.txtPayPremNett = !MBMPayorPremNet
            End If
            
            If Len(!MBMPayorMCOGross) Then
                frmNewMBMOthers.txtPayMCOGross = !MBMPayorMCOGross
            End If
            
            If Len(!MBMPayorMCONet) Then
                frmNewMBMOthers.txtPayMCONett = !MBMPayorMCONet
            End If
            
            If Len(!MBMFamDiscAmt) Then
                frmNewMBMOthers.txtPolicyDiscAmt = !MBMFamDiscAmt
            End If
    
            If Len(!MBMRenewDiscAmt) Then
                frmNewMBMOthers.txtRenewalDiscAmt = !MBMRenewDiscAmt
            End If
    
            If Len(!MBMLoadingPrem) Then
                frmNewMBMOthers.txtLoadingPrem = !MBMLoadingPrem
            End If
            
            If Trim(!MBMExclusion) <> "" Then
                txtMBMExclusion = Trim(!MBMExclusion)
            End If
            
            If Len(!MBMBankACNo) Then
                frmNewMBMOthers.txtBankACNo = !MBMBankACNo
            End If
            
            'If Len(!MBMRemarks) Then
                'txtRemarks = !MBMRemarks
            'Else
                'txtRemarks = ""
            'End If
            
            txtPolIssDate = IIf(Len(!MBMPOLIssuedDate), !MBMPOLIssuedDate, "__/__/____")
            
            'frmNewMBMOthers.txtAlterType = IIf(Len(!MBMAlterationType), !MBMAlterationType, "")
            
            'frmNewMBMOthers.txtAlterDate = IIf(Len(!MBMAlterationEffDate), !MBMAlterationEffDate, "__/__/____")
            
            frmNewMBMOthers.txtReinstateNo = IIf(Len(!MBMReinstateOcc), !MBMReinstateOcc, "")
            
            txtAlterType = IIf(Len(!MBMAlterationType), !MBMAlterationType, "")
            
            txtAlterDate = IIf(Len(!MBMAlterationEffDate), !MBMAlterationEffDate, "__/__/____")
        End With
End Function


Private Sub funcword()
'On Error Resume Next
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim i As Integer
Dim ii As Integer
Dim RecordSelected As Integer
Dim MBMNote1 As New ADODB.Recordset
Dim strsql As String
Dim MBMNote2 As New ADODB.Recordset
Dim strsql2 As String
Dim MBMNote3 As New ADODB.Recordset
Dim strsql3 As String
Dim SuppRemark As New ADODB.Recordset
Dim strsqlSuppRemark As String
Dim SuppRemark2 As New ADODB.Recordset
Dim strsqlSuppRemark2 As String
Dim MBMBNF As New ADODB.Recordset
Dim strsql4 As String
Dim MBMCase1 As New ADODB.Recordset
Dim strsql5 As String
Dim MBMCase2 As New ADODB.Recordset
Dim strsql6 As String
Dim MBMCLM1 As New ADODB.Recordset
Dim strsqlCLM1 As String
Dim MBMCLM2 As New ADODB.Recordset
Dim strsqlCLM2 As String
Dim AppAmt As String
Dim APPAmt2 As String
Dim SuppRec As New ADODB.Recordset
Dim strsqlSuppRec As String
Dim SuppRec2 As New ADODB.Recordset
Dim strsqlSuppRec2 As String
Dim MBMRec As New ADODB.Recordset
Dim strsqlMBMRec As String
Dim MBMHistoryRec As New ADODB.Recordset
Dim strsqlMBMHistoryRec As String
Dim MBMAccHistoryRec As New ADODB.Recordset
Dim strsqlMBMAccHistoryRec As String
Dim CurrEffDate, PrevEffDate As String
Dim cnt As String


AppAmt = 0
APPAmt2 = 0
       
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
    
    'If Err.Number > 0 Then
    '    MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
    '    GoTo ExitFunction
    'End If
        
    On Error GoTo HANDLE_ERROR
    objExcel.Interactive = False
    
    
        If objExcel.Visible = False Then
            objExcel.Visible = True
        End If
    
        objExcel.WindowState = xlMaximized

        Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "MBM\MBMExclusion.xlt")
        'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMExclusion.xlt")
                
        
    '========================================================================================
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
           
                        
            'strsqlMBMRec = " SELECT MBMNumber, MBMPolicyNo, MBMIcBcPp, MBMName, " & _
                           " MBMAddress1, MBMAddress2, MBMAddress3, CTYCode, MBMPostCode, " & _
                           " MBMDOB, MBMSex, MBMPayorEffDate, MBMPayorExpDate, INSCode, PLNCode, " & _
                           " MBMRenewFlag, STACode, MBMTakeover, MBMPolicyYear, MBMDataStatus, " & _
                           " MBMStatus, MBMGroupCompany, MBMPremium, AnnLimit, MBMAvailableLimit, " & _
                           " MBMAgentCode, MBMOccupation, MBMInstallment , MBMBordxDate, MBMFirstJoinedDate, " & _
                           " MBMMemberType, MBMDepartment FROM VIEW_MBM_ENQUIRY Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
                
            'MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
            'If MBMRec.recordCount Then
            
            objWorksheet.Cells(6, "B") = txtMBMNumber '(MBMRec!MBMNumber)
            objWorksheet.Cells(7, "B") = txtMBMPolNo '(MBMRec!MBMPolicyNo)
            objWorksheet.Cells(8, "B") = txtMBMBirthDate 'Format(MBMRec!MBMDOB, "mm/dd/yyyy")
            objWorksheet.Cells(9, "B") = cboINSCode 'Trim(MBMRec!INSCode)
            objWorksheet.Cells(10, "B") = txtMBMPayorEffDate 'Format(MBMRec!MBMPayorEffDate, "mm/dd/yyyy")
            objWorksheet.Cells(11, "B") = txtMBMPayorExpDate 'Format(MBMRec!MBMPayorEXPDate, "mm/dd/yyyy")
            objWorksheet.Cells(12, "B") = txtMBMBordxDate 'Format(MBMRec!MBMBordxDate, "mm/dd/yyyy")
            objWorksheet.Cells(13, "B") = txtMBMDateJoin 'Format(MBMRec!MBMFirstJoinedDate, "mm/dd/yyyy")
            objWorksheet.Cells(14, "B") = txtMBMStatus 'MBMRec!MBMStatus
            objWorksheet.Cells(15, "B") = txtMBMAdd1 'MBMRec!MBMAddress1
            objWorksheet.Cells(16, "B") = txtMBMAdd2 'MBMRec!MBMAddress2
            objWorksheet.Cells(17, "B") = txtMBMAdd3 'MBMRec!MBMAddress3
            objWorksheet.Cells(18, "B") = txtMBMPostCode 'MBMRec!MBMPostCode
            objWorksheet.Cells(19, "B") = cboCTYCode 'MBMRec!CTYCode
            objWorksheet.Cells(20, "B") = cboSTACode 'MBMRec!STACode
            objWorksheet.Cells(23, "B") = txtMBMPremium 'Format(MBMRec!MBMPremium, "#,##0.00")
            objWorksheet.Cells(23, "E") = txtAnnualLimit 'Format(MBMRec!AnnLimit, "#,##0.00")
            objWorksheet.Cells(6, "G") = txtMBMName 'MBMRec!MBMName
            objWorksheet.Cells(7, "G") = txtMBMIcBcPP 'MBMRec!MBMIcBcPp
            objWorksheet.Cells(8, "G") = cboMBMSex 'MBMRec!MBMSex
            objWorksheet.Cells(9, "G") = cboPLNCode 'MBMRec!PLNCode
            objWorksheet.Cells(10, "G") = cboMBMTakeOver 'MBMRec!MBMTakeover
            objWorksheet.Cells(11, "G") = cboMBMRenewal 'MBMRec!MBMRenewFlag
            objWorksheet.Cells(12, "G") = CboInsMode 'MBMRec!MBMInstallment
            objWorksheet.Cells(13, "G") = "P" 'MBMRec!MBMMemberType
            objWorksheet.Cells(14, "G") = txtDataStatus 'MBMRec!MBMDataStatus
            objWorksheet.Cells(15, "G") = txtMBMGroupCompany 'MBMRec!MBMGroupCompany
            objWorksheet.Cells(16, "G") = txtDeptName 'MBMRec!MBMDepartment
            objWorksheet.Cells(17, "G") = txtMBMAgentCode 'MBMRec!MBMAgentCode
            objWorksheet.Cells(18, "G") = frmNewMBMOthers.cboMBMOccupation 'MBMRec!MBMOccupation
            objWorksheet.Cells(19, "G") = txtMBMPolicyYear 'MBMRec!MBMPolicyYear
            objWorksheet.Cells(23, "H") = txtAvailableAnnualLimit 'Format(MBMRec!MBMAvailableLimit, "#,##0.00")
            
            
            'If Trim(MBMRec!INSCode) = "F" Or Trim(MBMRec!INSCode) = "H" Then
            'If Trim(cboINSCode) = "F" Or Trim(cboINSCode) = "H" Then
        
            'strsqlSuppRec = "Select MBMCName, MBMCCoverID, MBMCICBCPPNo, MBMCDOB, MBMCSex, MBMCAdultChild, MBMCRELCode from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'" & _
                            " Order By MBMCCoverID "
            'SuppRec.Open strsqlSuppRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            'strsqlSuppRec2 = "Select MBMCPrevInsCom from MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "'" & _
                               " Order By MBMCCoverID "
            'SuppRec2.Open strsqlSuppRec2, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            'If SuppRec.recordCount Then
                'RecordSelected = SuppRec.recordCount
                RecordSelected = lstCovAdd.ListItems.Count
            
                For i = 1 To RecordSelected
                    
                    'objWorksheet.Cells(29 + CDbl(i), "A") = Trim(SuppRec!MBMCName)
                    'objWorksheet.Cells(29 + CDbl(i), "D") = Trim(SuppRec!MBMCCoverID)
                    'objWorksheet.Cells(29 + CDbl(i), "E") = Trim(SuppRec!MBMCICBCPPNo)
                    'objWorksheet.Cells(29 + CDbl(i), "G") = Format(SuppRec!MBMCDOB, "mm/dd/yyyy")
                    'objWorksheet.Cells(29 + CDbl(i), "H") = Trim(SuppRec!MBMCSex)
                    'objWorksheet.Cells(29 + CDbl(i), "I") = Trim(SuppRec!MBMCAdultChild)
                    'objWorksheet.Cells(29 + CDbl(i), "J") = Trim(SuppRec!MBMCRELCode)
                    
                    objWorksheet.Cells(29 + CDbl(i), "A") = lstCovAdd.ListItems(i).Text
                    objWorksheet.Cells(29 + CDbl(i), "D") = lstCovAdd.ListItems(i).SubItems(10)
                    objWorksheet.Cells(29 + CDbl(i), "E") = lstCovAdd.ListItems(i).SubItems(2)
                    objWorksheet.Cells(29 + CDbl(i), "G") = Format(lstCovAdd.ListItems(i).SubItems(3), "mm/dd/yyyy")
                    objWorksheet.Cells(29 + CDbl(i), "H") = lstCovAdd.ListItems(i).SubItems(4)
                    objWorksheet.Cells(29 + CDbl(i), "I") = lstCovAdd.ListItems(i).SubItems(5)
                    objWorksheet.Cells(29 + CDbl(i), "J") = lstCovAdd.ListItems(i).SubItems(6)
                    objWorksheet.Cells(29 + CDbl(i), "K") = lstCovAdd.ListItems(i).SubItems(22)
                                        
                'SuppRec.MoveNext
                    
                Next i
                
            'End If
            
            'If SuppRec2.recordCount Then
                'RecordSelected = SuppRec2.recordCount
            
                'For i = 1 To RecordSelected
                    
                    'objWorksheet.Cells(29 + CDbl(i), "K") = Trim(SuppRec2!MBMCPrevInsCom)
                                
                'SuppRec2.MoveNext
                    
                'Next i
                
            'End If
            
            'SuppRec.Close
            'Set SuppRec = Nothing
            'SuppRec2.Close
            'Set SuppRec2 = Nothing
        
        'End If
        
        
        
        Set objWorksheet = objWorkbook.Sheets(2)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    
        Set objWorksheet = objWorkbook.Worksheets.Item(2)

             If lstMBMHis.ListItems.Count = 0 Then
                Call MBMHISListing
             End If
             If lstMBMHis.ListItems.Count > 0 Then
                For i = 1 To lstMBMHis.ListItems.Count
                    objWorksheet.Cells(4 + CDbl(i), "A") = lstMBMHis.ListItems(i).Text & "-" & lstMBMHis.ListItems(i).SubItems(1)
                    objWorksheet.Cells(4 + CDbl(i), "C") = lstMBMHis.ListItems(i).SubItems(4)
                    objWorksheet.Cells(4 + CDbl(i), "F") = Format(lstMBMHis.ListItems(i).SubItems(5), "mm/dd/yyyy")
                    objWorksheet.Cells(4 + CDbl(i), "G") = Format(lstMBMHis.ListItems(i).SubItems(6), "mm/dd/yyyy")
                    objWorksheet.Cells(4 + CDbl(i), "H") = lstMBMHis.ListItems(i).SubItems(7)
                    objWorksheet.Cells(4 + CDbl(i), "I") = lstMBMHis.ListItems(i).SubItems(8)
                Next i
            End If
           
            
            strsqlMBMAccHistoryRec = " SELECT MBMNumber, CASDateAdmitted, CaseRegDate, CASNumber, ClaimVersion, " & _
                                     " DebitCreditDate, DebitCreditRefNo, CLMPayableByMedix2M " & _
                                     " FROM VIEW_Account_History WHERE MBMIcBcPp = '" & Trim(txtMBMIcBcPP) & "'"
    
            MBMAccHistoryRec.Open strsqlMBMAccHistoryRec, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            
            If MBMAccHistoryRec.recordCount Then
                RecordSelected = MBMAccHistoryRec.recordCount
            
                For i = 1 To RecordSelected
                    
                    objWorksheet.Cells(30 + CDbl(i), "A") = Trim(MBMAccHistoryRec!MBMNumber)
                    objWorksheet.Cells(30 + CDbl(i), "C") = Format(MBMAccHistoryRec!CASDateAdmitted, "mm/dd/yyyy")
                    objWorksheet.Cells(30 + CDbl(i), "D") = Format(MBMAccHistoryRec!CaseRegDate, "mm/dd/yyyy")
                    objWorksheet.Cells(30 + CDbl(i), "E") = MBMAccHistoryRec!CasNumber
                    objWorksheet.Cells(30 + CDbl(i), "F") = MBMAccHistoryRec!claimversion
                    objWorksheet.Cells(30 + CDbl(i), "G") = MBMAccHistoryRec!DebitCreditRefNo
                    objWorksheet.Cells(30 + CDbl(i), "H") = Format(MBMAccHistoryRec!DebitCreditDate, "mm/dd/yyyy")
                    objWorksheet.Cells(30 + CDbl(i), "I") = Format(MBMAccHistoryRec!CLMPayableByMedix2M, "#,##0.00")
                                                            
                MBMAccHistoryRec.MoveNext
                    
                Next i
                
            End If
            
            MBMAccHistoryRec.Close
            Set MBMAccHistoryRec = Nothing
            
                
    '==========================================================================================
        
        
        Set objWorksheet = objWorkbook.Sheets(3)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    
        Set objWorksheet = objWorkbook.Worksheets.Item(3)
    
            'With MBMNote1
                'objWorksheet.Cells(4, "A") = "Member No:" & " " & Trim(txtMBMNumber)
                objWorksheet.Cells(5, "A") = Trim(txtMBMName)
                objWorksheet.Cells(5, "B") = "00"
                objWorksheet.Cells(5, "C") = "P"
                objWorksheet.Cells(5, "D") = Trim(txtMBMExclusion)
                objWorksheet.Cells(5, "E") = Trim(txtRemarks)
                objWorksheet.Cells(5, "F") = txtAvailableAnnualLimit
                objWorksheet.Cells(5, "G") = txtAnnualLimit
        'If Trim(cboINSCode) = "F" Or Trim(cboINSCode) = "H" Then
        
            'strsqlSuppRemark = "Select MBMCExclusion, MBMCPAYRemarks from MBMCoveredPersonsTwo where MBMCNumber = '" & Trim(txtMBMNumber) & "'" & _
                               " Order By MBMCCoverID "
            'SuppRemark.Open strsqlSuppRemark, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            'strsqlSuppRemark2 = "Select MBMCName, MBMCCoverID, MBMCRELCode, MBMCAvailableLimit, MBMCAnnualLimit from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'" & _
                               " Order By MBMCCoverID "
            'SuppRemark2.Open strsqlSuppRemark2, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            'If SuppRemark2.recordCount Then
                'RecordSelected = SuppRemark2.recordCount
            
                'For i = 1 To RecordSelected
                    
                    'objWorksheet.Cells(5 + CDbl(i), "A") = Trim(SuppRemark2!MBMCName)
                    'objWorksheet.Cells(5 + CDbl(i), "B") = Trim(SuppRemark2!MBMCCoverID)
                    'objWorksheet.Cells(5 + CDbl(i), "C") = Trim(SuppRemark2!MBMCRELCode)
                    'objWorksheet.Cells(5 + CDbl(i), "F") = Format(SuppRemark2!MBMCAvailableLimit, "#,##0.00")
                    'objWorksheet.Cells(5 + CDbl(i), "G") = Format(SuppRemark2!MBMCAnnualLimit, "#,##0.00")
                    
                'SuppRemark2.MoveNext
                    
                'Next i
                
            'End If
            
            RecordSelected = lstCovAdd.ListItems.Count
            
            For i = 1 To RecordSelected
                
                objWorksheet.Cells(5 + CDbl(i), "A") = lstCovAdd.ListItems(i).Text
                objWorksheet.Cells(5 + CDbl(i), "B") = lstCovAdd.ListItems(i).SubItems(10)
                objWorksheet.Cells(5 + CDbl(i), "C") = lstCovAdd.ListItems(i).SubItems(6)
                objWorksheet.Cells(5 + CDbl(i), "D") = lstCovAdd.ListItems(i).SubItems(9)
                objWorksheet.Cells(5 + CDbl(i), "E") = lstCovAdd.ListItems(i).SubItems(21)
                objWorksheet.Cells(5 + CDbl(i), "F") = lstCovAdd.ListItems(i).SubItems(13)
                objWorksheet.Cells(5 + CDbl(i), "G") = lstCovAdd.ListItems(i).SubItems(12)
                
                                    
            'SuppRec.MoveNext
                
            Next i
            
            
            'If SuppRemark.recordCount Then
                'RecordSelected = SuppRemark.recordCount
            
                'For i = 1 To RecordSelected
                    
                    'objWorksheet.Cells(5 + CDbl(i), "D") = Trim(SuppRemark!MBMCExclusion)
                    'objWorksheet.Cells(5 + CDbl(i), "E") = Trim(SuppRemark!MBMCPAYRemarks)
                                
                'SuppRemark.MoveNext
                    
                'Next i
                
            'End If
            
            'SuppRemark.Close
            'Set SuppRemark = Nothing
            'SuppRemark2.Close
            'Set SuppRemark2 = Nothing
        
        'End If
        
        strsql = " Select PLNSpecialNote from PLN Where HLTCode = '" & Trim(cboHLTCode) & "' AND PLNCode = '" & Trim(cboPLNCode) & "'"
        MBMNote1.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBMNote1.recordCount Then
            objWorksheet.Cells(30, "A") = Trim(MBMNote1!PLNSpecialNote)
        End If
        MBMNote1.Close
        Set MBMNote1 = Nothing
        
        Set objWorksheet = objWorkbook.Sheets(4)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    
        Set objWorksheet = objWorkbook.Worksheets.Item(4)
        
        strsql3 = " Select * From VIEW_BNF Where PLNCode = '" & Trim(cboPLNCode) & "'" & _
                  " AND HLTCode = '" & Trim(cboHLTCode) & "'" & _
                  " AND INSCode = '" & Trim(cboINSCode) & "' ORDER BY BNFEffDate DESC"
        MBMBNF.Open strsql3, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBMBNF.recordCount Then
            RecordSelected = MBMBNF.recordCount
            cnt = 0
            For i = 1 To RecordSelected
                cnt = cnt + 1
                If cnt = 1 Then
                    PrevEffDate = MBMBNF!BNFEffDate
                End If
                CurrEffDate = MBMBNF!BNFEffDate
                If PrevEffDate <> CurrEffDate Then
                    Exit For
                End If

                With MBMBNF
                    objWorksheet.Cells(4 + CDbl(i), "A") = !BNFCode
                    objWorksheet.Cells(4 + CDbl(i), "B") = Trim(!BNFdescription)
                    objWorksheet.Cells(4 + CDbl(i), "C") = Trim(!BNFNoOfDays)
                    objWorksheet.Cells(4 + CDbl(i), "D") = Trim(!BNFClaimAbleAmount)
                    objWorksheet.Cells(4 + CDbl(i), "E") = Trim(!BNFRMperDis)
                End With
                
                MBMBNF.MoveNext
                    
            Next i
            
        End If
        
        MBMBNF.Close
        Set MBMBNF = Nothing
        
        Set objWorksheet = objWorkbook.Sheets(5)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    
        Set objWorksheet = objWorkbook.Worksheets.Item(5)
        
        strsql4 = " Select CasNumber, MBMNumber, MBMPatientCovID, MBMName, MBMCName, " & _
                  " MBMCRELCode, CASDateAdmitted, CASDischargeDate, CASClaimability, DIAFinalDiag1,CASOPChemo " & _
                  " From VIEW_MBMInquiry_CASE Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBMCase1.Open strsql4, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBMCase1.recordCount Then
            RecordSelected = MBMCase1.recordCount
        
            For i = 1 To RecordSelected
        
            With MBMCase1
                objWorksheet.Cells(4 + CDbl(i), "A") = !CasNumber
                objWorksheet.Cells(4 + CDbl(i), "B") = Trim(!MBMNumber)
                If !MBMPatientCovID = "00" Then
                    objWorksheet.Cells(4 + CDbl(i), "C") = Trim(!MBMName)
                Else
                    objWorksheet.Cells(4 + CDbl(i), "C") = Trim(!MBMCName)
                End If
                If !MBMPatientCovID = "00" Then
                    objWorksheet.Cells(4 + CDbl(i), "D") = "P"
                Else
                    objWorksheet.Cells(4 + CDbl(i), "D") = Trim(!MBMCRELCode)
                End If
                objWorksheet.Cells(4 + CDbl(i), "E") = Format(!CASDateAdmitted, "mm/dd/yyyy")
                objWorksheet.Cells(4 + CDbl(i), "F") = Format(!CASDischargeDate, "mm/dd/yyyy")
                objWorksheet.Cells(4 + CDbl(i), "G") = Trim(!CASClaimability)
                objWorksheet.Cells(4 + CDbl(i), "H") = Trim(!DIAFinalDiag1)

                strsqlCLM1 = " Select CASNumber, SUM(CLMTotalApproved) AS CLMTotalApproved From Claim Where CASNumber = '" & Trim(!CasNumber) & "'" & _
                             " GROUP BY CASNumber "
                MBMCLM1.Open strsqlCLM1, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                                
                If MBMCLM1.EOF = False Then
                    objWorksheet.Cells(4 + CDbl(i), "I") = Format(MBMCLM1!CLMTotalApproved, "#,##0.00")
                    AppAmt = AppAmt + MBMCLM1!CLMTotalApproved
                End If
                
                objWorksheet.Cells(4 + CDbl(i), "J") = IIf(Len(!CASOPChemo), !CASOPChemo, "")
                
                MBMCLM1.Close
                Set MBMCLM1 = Nothing
                
            End With
            
            MBMCase1.MoveNext
                
            Next i
            
        End If
        
        objWorksheet.Cells(5 + CDbl(i), "H") = "TOTAL"
        objWorksheet.Cells(5 + CDbl(i), "I") = Format(AppAmt, "#,##0.00")
        
        MBMCase1.Close
        Set MBMCase1 = Nothing
        
        'Set objWorksheet = objWorkbook.Sheets(4)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        'objExcel.DisplayAlerts = False
    
        'Set objWorksheet = objWorkbook.Worksheets.Item(4)
        If Trim(txtMBMNumber) <> "" And lstCASHistory.ListItems.Count = 0 Then
            Call ListALLCases
        End If
            RecordSelected = lstCASHistory.ListItems.Count
        
        'strsql5 = " Select CasNumber, MBMNumber, MBMPatientCovID, MBMName, MBMCName, " & _
        '          " MBMCRELCode, CASDateAdmitted, CASDischargeDate, CASClaimability, DIAFinalDiag1 " & _
        '          " From VIEW_MBMInquiry_CASE Where MBMIcBcPp = '" & Trim(txtMBMIcBcPP) & "'"
        'MBMCase2.Open strsql5, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        'If MBMCase2.recordCount Then
        '    RecordSelected = MBMCase2.recordCount
        
            For i = 1 To RecordSelected
        
            'With MBMCase2
                objWorksheet.Cells(37 + CDbl(i), "A") = lstCASHistory.ListItems(i).Text
                objWorksheet.Cells(37 + CDbl(i), "B") = lstCASHistory.ListItems(i).SubItems(1)
                'If !MBMPatientCovID = "00" Then
                '    objWorksheet.Cells(37 + CDbl(i), "C") = Trim(!MBMName)
                'Else
                objWorksheet.Cells(37 + CDbl(i), "C") = lstCASHistory.ListItems(i).SubItems(2)
                'End If
                'If !MBMPatientCovID = "00" Then
                '    objWorksheet.Cells(37 + CDbl(i), "D") = "P"
                'Else
                    objWorksheet.Cells(37 + CDbl(i), "D") = lstCASHistory.ListItems(i).SubItems(3)
                'End If
                objWorksheet.Cells(37 + CDbl(i), "E") = Format(lstCASHistory.ListItems(i).SubItems(4), "MM/DD/YYYY")
                objWorksheet.Cells(37 + CDbl(i), "F") = Format(lstCASHistory.ListItems(i).SubItems(5), "MM/DD/YYYY")
                objWorksheet.Cells(37 + CDbl(i), "G") = lstCASHistory.ListItems(i).SubItems(6)
                objWorksheet.Cells(37 + CDbl(i), "H") = lstCASHistory.ListItems(i).SubItems(7)
                
                strsqlCLM2 = " Select CASNumber, SUM(CLMTotalApproved) AS CLMTotalApproved From Claim Where CASNumber = '" & Trim(lstCASHistory.ListItems(i).Text) & "'" & _
                             " GROUP BY CASNumber "
                MBMCLM2.Open strsqlCLM2, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                               
                If MBMCLM2.EOF = False Then
                    objWorksheet.Cells(37 + CDbl(i), "I") = Format(MBMCLM2!CLMTotalApproved, "#,##0.00")
                    APPAmt2 = APPAmt2 + MBMCLM2!CLMTotalApproved
                End If
                
                MBMCLM2.Close
                Set MBMCLM2 = Nothing
                
            'End With
            
            'MBMCase2.MoveNext
                
            Next i
            
        'End If
        
        objWorksheet.Cells(38 + CDbl(i), "H") = "TOTAL"
        objWorksheet.Cells(38 + CDbl(i), "I") = Format(APPAmt2, "#,##0.00")
        
        'MBMCase2.Close
        'Set MBMCase2 = Nothing
        
        'End If
        
        'MBMRec.Close
        'Set MBMRec = Nothing
        
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

Private Sub MBMBenefitLimit(tmpMBMNumber As String, TmpCovID As String, tmpKidneyAnn As Double, tmpKidneyBal As Double, _
tmpCancerAnn As Double, tmpCancerBal As Double, tmpHomeNurseAnn As Double, tmpHomeNurseBal As Double)

Dim MBMBNFRec As New ADODB.Recordset
Dim StrBNFSql As String
    
    tmpKidneyAnn = 0
    tmpKidneyBal = 0
    tmpCancerAnn = 0
    tmpCancerBal = 0
    tmpHomeNurseAnn = 0
    tmpHomeNurseBal = 0
    
    StrBNFSql = "SELECT MBMCancerAnnLmt, MBMCancerBalLmt, MBMKidneyAnnLmt, MBMKidneyBalLmt, MBMHomeNCAnnLmt, MBMHomeNCBalLmt" & _
    " FROM MBMBnfLmt WHERE MBMNumber='" & Trim(tmpMBMNumber) & "' AND MBMCoveredID='" & Trim(TmpCovID) & "'"
    MBMBNFRec.Open StrBNFSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    Do While Not MBMBNFRec.EOF
        With MBMBNFRec
            tmpKidneyAnn = IIf(IsNumeric(!MBMKidneyAnnLmt), !MBMKidneyAnnLmt, 0)
            tmpKidneyBal = IIf(IsNumeric(!MBMKidneyBalLmt), !MBMKidneyBalLmt, 0)
            tmpCancerAnn = IIf(IsNumeric(!MBMCancerAnnLmt), !MBMCancerAnnLmt, 0)
            tmpCancerBal = IIf(IsNumeric(!MBMCancerBalLmt), !MBMCancerBalLmt, 0)
            tmpHomeNurseAnn = IIf(IsNumeric(!MBMHomeNCAnnLmt), !MBMHomeNCAnnLmt, 0)
            tmpHomeNurseBal = IIf(IsNumeric(!MBMHomeNCBalLmt), !MBMHomeNCBalLmt, 0)
        End With
        MBMBNFRec.MoveNext
    Loop
    
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
    lstPreUndIllness.ListItems.Clear
    If Not rst.EOF Then

        If SSTab1.Tab = 10 And lstPreUndIllness.ListItems.Count = 0 Then
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

Private Sub ListCases1(tmpHistory As Boolean, tmpMBMNo As String, tmpICNo As String, tmpSuppCas As Boolean)
Dim CASList As ListItem
Dim strsql As String
Dim CAS As New ADODB.Recordset
    If tmpHistory = False Then
        strsql = "Exec SearchCas '" & Trim(tmpMBMNo) & "'"
        lstCaseList.ListItems.Clear
    Else
        If tmpSuppCas = False Then
            strsql = "Exec SearchCASHistory '" & Trim(tmpICNo) & "'"
        Else
            strsql = "Exec SearchCasHistorySupp '" & Trim(tmpICNo) & "'"
        End If
    End If
    Set CAS = New ADODB.Recordset
    CAS.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly

    Screen.MousePointer = vbHourglass
        
    Do While Not CAS.EOF
        If tmpHistory = False Then
            Set CASList = lstCaseList.ListItems.Add(, , CAS!CasNumber)
        Else
            Set CASList = lstCASHistory.ListItems.Add(, , CAS!CasNumber)
        End If
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
        CASList.SubItems(8) = IIf(Len(CAS!CasStatus), CAS!CasStatus, "")
        CASList.SubItems(9) = IIf(Len(CAS!HosName), CAS!HosName, "")
        CASList.SubItems(10) = IIf(Len(CAS!PAYCode), CAS!PAYCode, "")
        cntCasesHisIC = CDbl(cntCasesHisIC) + 1
        CAS.MoveNext
        
    Loop
    
    CAS.Close
    Set CAS = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListWorksheet(tmpHistory As Boolean, tmpMBMNo As String, tmpICNo As String, tmpSuppCas As Boolean)
Dim CASList As ListItem
Dim strsql As String
Dim CAS As New ADODB.Recordset
   ' If tmpHistory = False Then
        strsql = "Exec SearchWorksheet '" & Trim(tmpMBMNo) & "'"
        lstWorksheet.ListItems.Clear
 '   Else
 '       If tmpSuppCas = False Then
 '           strsql = "Exec SearchCASHistory '" & Trim(tmpICNo) & "'"
 '       Else
 '           strsql = "Exec SearchCasHistorySupp '" & Trim(tmpICNo) & "'"
 '       End If
 '   End If
    Set CAS = New ADODB.Recordset
    CAS.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly

    Screen.MousePointer = vbHourglass
        
    Do While Not CAS.EOF
       ' If tmpHistory = False Then
            Set CASList = lstWorksheet.ListItems.Add(, , CAS!CasNumber)
        'Else
        '    Set CASList = lstCASHistory.ListItems.Add(, , CAS!CasNumber)
        'End If
        CASList.SubItems(1) = CAS!claimversion
        CASList.SubItems(2) = CAS!MBMNumber
        If CAS!MBMPatientCovID = "00" Then
            CASList.SubItems(3) = IIf(Len(Trim(CAS!MBMName)), Trim(CAS!MBMName), "")
            CASList.SubItems(4) = "P"
        Else
            CASList.SubItems(3) = IIf(Len(Trim(CAS!MBMCName)), Trim(CAS!MBMCName), "")
            CASList.SubItems(4) = IIf(Len(CAS!MBMCRELCode), CAS!MBMCRELCode, "")
        End If
        
        If Len(Trim(CAS!CASDateAdmitted)) Then
            CASList.SubItems(5) = CAS!CASDateAdmitted
        End If
        If Len(Trim(CAS!CASDischargeDate)) Then
            CASList.SubItems(6) = CAS!CASDischargeDate
        End If
        CASList.SubItems(7) = IIf(Len(CAS!CASClaimability), CAS!CASClaimability, "")
        CASList.SubItems(8) = IIf(Len(CAS!DIAFinalDiag1), CAS!DIAFinalDiag1, "")
        CASList.SubItems(9) = IIf(Len(CAS!CasStatus), CAS!CasStatus, "")
        CASList.SubItems(10) = IIf(Len(CAS!HosName), CAS!HosName, "")
        CASList.SubItems(11) = IIf(Len(CAS!PAYCode), CAS!PAYCode, "")
        CASList.SubItems(12) = IIf(Len(CAS!ClaimOpenDate), Format(CAS!ClaimOpenDate, "DD/MM/YYYY"), "")
        CASList.SubItems(13) = IIf(Len(CAS!CLMTotalActual), Format(CAS!CLMTotalActual, "#,##0.00"), "")
        CASList.SubItems(14) = IIf(Len(CAS!CLMBordxAmt), Format(CAS!CLMBordxAmt, "#,##0.00"), "")
        CASList.SubItems(15) = IIf(Len(CAS!CLMPrevMajorMed), Format(CAS!CLMPrevMajorMed, "#,##0.00"), "")
        CASList.SubItems(16) = IIf(Len(CAS!CLMMajorMedFee), Format(CAS!CLMMajorMedFee, "#,##0.00"), "")
        CASList.SubItems(17) = IIf(Len(CAS!BordxRefNo), CAS!BordxRefNo, "")
        CASList.SubItems(18) = IIf(Len(CAS!CLMBordxDate), Format(CAS!CLMBordxDate, "DD/MM/YYYY"), "")




        cntCasesHisIC = CDbl(cntCasesHisIC) + 1
        CAS.MoveNext
        
    Loop
    
    CAS.Close
    Set CAS = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListCasesTwo(tmpHistory As Boolean, tmpMBMNo As String, tmpName As String, tmpDOb As String, tmpSuppCas As Boolean)
Dim CASList As ListItem
Dim strsql As String
Dim CAS As New ADODB.Recordset
    strsql = ""
    If tmpHistory = False Then
        strsql = "Exec SearchCas '" & Trim(tmpMBMNo) & "'"
        lstCaseList.ListItems.Clear
    Else
        If tmpSuppCas = False Then
            strsql = "Exec SearchCASHistoryTwo '" & Trim(tmpName) & "', '" & Format(tmpDOb, "yyyy/mm/dd") & "'"
        Else
            strsql = "Exec SearchCasHistorySuppTwo '" & Trim(tmpName) & "', '" & Format(tmpDOb, "yyyy/mm/dd") & "'"
        End If
    End If
    Set CAS = New ADODB.Recordset
    CAS.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly

    Screen.MousePointer = vbHourglass
        
    Do While Not CAS.EOF
        If tmpHistory = False Then
            Set CASList = lstCaseList.ListItems.Add(, , CAS!CasNumber)
        Else
            Set CASList = lstCASHistoryTwo.ListItems.Add(, , CAS!CasNumber)
        End If
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
        CASList.SubItems(8) = IIf(Len(CAS!CasStatus), CAS!CasStatus, "")
        CASList.SubItems(9) = IIf(Len(CAS!HosName), CAS!HosName, "")
        CASList.SubItems(10) = IIf(Len(CAS!PAYCode), CAS!PAYCode, "")
        cntCasesHisDOB = CDbl(cntCasesHisDOB) + 1
        CAS.MoveNext
    Loop
    CAS.Close
    Set CAS = Nothing
    Screen.MousePointer = vbDefault
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
Dim prm7 As ADODB.Parameter
Dim prm8 As ADODB.Parameter

Dim rst5 As New ADODB.Recordset
Dim cmd5 As New ADODB.Command
Dim prm9 As ADODB.Parameter
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
                        Set prm9 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
                        cmd5.Parameters.Append prm9
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

Private Sub MBMHistory(TmpMBMType)
Dim TmpREc1 As New ADODB.Recordset
Dim HISList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpName As String
Dim strAppend As String
Dim tmpDOb As String

    tmpName = Trim(Replace(txtMBMName, "'", "''"))
    tmpDOb = Format(txtMBMBirthDate, "yyyy/mm/dd")
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
        Set HISList = lstMBMHis.ListItems.Add(, , TmpREc1!MBMNumber)
        HISList.SubItems(1) = IIf(Len(TmpREc1!MBMCOVID), TmpREc1!MBMCOVID, "")
        HISList.SubItems(2) = IIf(Len(TmpREc1!MBMTopupInd), TmpREc1!MBMTopupInd, "")
        HISList.SubItems(3) = IIf(Len(TmpREc1!MBMTopupNumber), TmpREc1!MBMTopupNumber, "")
        HISList.SubItems(4) = IIf(Len(TmpREc1!MBMPolicyNo), TmpREc1!MBMPolicyNo, "")
        HISList.SubItems(5) = IIf(TmpREc1!MBMPayorEffDate <> "", TmpREc1!MBMPayorEffDate, "")
        HISList.SubItems(6) = IIf(TmpREc1!MBMPayorEXPDate <> "", TmpREc1!MBMPayorEXPDate, "")
        HISList.SubItems(7) = IIf(Len(TmpREc1!MBMDataStatus), TmpREc1!MBMDataStatus, "")
        HISList.SubItems(8) = IIf(Len(TmpREc1!MBMStatus), TmpREc1!MBMStatus, "")
        HISList.SubItems(9) = IIf(Len(TmpREc1!MBMBordxDate), TmpREc1!MBMBordxDate, "")
        cntMBMHisIC = CDbl(cntMBMHisIC) + 1
        TmpREc1.MoveNext
    Loop
    TmpREc1.Close
    Set TmpREc1 = Nothing

End Sub

Private Sub MBMHistoryTwo(TmpMBMType)
Dim TmpREc1 As New ADODB.Recordset
Dim HISList As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpName As String
Dim strAppend As String
Dim tmpDOb As String

    tmpName = Trim(Replace(txtMBMName, "'", "''"))
    tmpDOb = Format(txtMBMBirthDate, "yyyy/mm/dd")
    If TmpMBMType = "Princ" Then
        strAppend = " And HISMaintenance.dbo.MBMCrossReference.MBMName = '" & tmpName
        strAppend = strAppend & "' And HISMaintenance.dbo.MBMCrossReference.MBMDOB = '" & Format(tmpDOb, "yyyy/mm/dd") & "'"
    Else
        strAppend = " And dbo.MBMCoveredPersons.MBMCName = '" & tmpName
        strAppend = strAppend & "' And dbo.MBMCoveredPersons.MBMCDOB= '" & Format(tmpDOb, "yyyy/mm/dd") & "'"
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
        Set HISList = lstMBMHisTwo.ListItems.Add(, , TmpREc1!MBMNumber)
        HISList.SubItems(1) = IIf(Len(TmpREc1!MBMCOVID), TmpREc1!MBMCOVID, "")
        HISList.SubItems(2) = IIf(Len(TmpREc1!MBMTopupInd), TmpREc1!MBMTopupInd, "")
        HISList.SubItems(3) = IIf(Len(TmpREc1!MBMTopupNumber), TmpREc1!MBMTopupNumber, "")
        HISList.SubItems(4) = IIf(Len(TmpREc1!MBMPolicyNo), TmpREc1!MBMPolicyNo, "")
        HISList.SubItems(5) = IIf(TmpREc1!MBMPayorEffDate <> "", TmpREc1!MBMPayorEffDate, "")
        HISList.SubItems(6) = IIf(TmpREc1!MBMPayorEXPDate <> "", TmpREc1!MBMPayorEXPDate, "")
        HISList.SubItems(7) = IIf(Len(TmpREc1!MBMDataStatus), TmpREc1!MBMDataStatus, "")
        HISList.SubItems(8) = IIf(Len(TmpREc1!MBMStatus), TmpREc1!MBMStatus, "")
        HISList.SubItems(9) = IIf(Len(TmpREc1!MBMBordxDate), TmpREc1!MBMBordxDate, "")
        cntMBMHisDOB = CDbl(cntMBMHisDOB) + 1
        
        TmpREc1.MoveNext
    Loop
    TmpREc1.Close
    Set TmpREc1 = Nothing

End Sub







