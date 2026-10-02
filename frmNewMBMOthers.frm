VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNewMBMOthers 
   Caption         =   "Member Others"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   6615
      Left            =   0
      TabIndex        =   45
      Top             =   240
      Width           =   11100
      Begin VB.TextBox txtAgntContact 
         Height          =   285
         Left            =   6360
         MaxLength       =   25
         TabIndex        =   115
         Top             =   5880
         Width           =   3615
      End
      Begin VB.TextBox txtAgent 
         Height          =   285
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   111
         Top             =   5760
         Width           =   3615
      End
      Begin VB.TextBox txtAlterType 
         Height          =   285
         Left            =   6360
         MaxLength       =   20
         TabIndex        =   104
         Top             =   5160
         Width           =   975
      End
      Begin VB.ComboBox cboMBMOccupation 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   120
         Width           =   3640
      End
      Begin VB.TextBox txtMBMAllergic 
         Height          =   2505
         Left            =   5040
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   43
         Top             =   2520
         Width           =   5940
      End
      Begin VB.Frame Frame12 
         Height          =   495
         Left            =   5040
         TabIndex        =   58
         Top             =   1680
         Width           =   3255
         Begin MSMask.MaskEdBox TxtClmNonPayFr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   250
            Left            =   555
            TabIndex        =   39
            Top             =   150
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtClmNonPayTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   40
            Top             =   150
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   1680
            TabIndex        =   60
            Top             =   200
            Width           =   285
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   80
            TabIndex        =   59
            Top             =   200
            Width           =   405
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1335
         Left            =   5040
         TabIndex        =   50
         Top             =   120
         Width           =   6015
         Begin VB.TextBox TxtGenWP 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   32
            Top             =   240
            Width           =   705
         End
         Begin VB.TextBox txtPREWP 
            Height          =   285
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   34
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox TxtSpeWP 
            Height          =   300
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   36
            Top             =   720
            Width           =   705
         End
         Begin VB.ComboBox cboStaffDisc 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   38
            Top             =   960
            Width           =   705
         End
         Begin MSMask.MaskEdBox txtGENWPEnd 
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
            Left            =   3240
            TabIndex        =   33
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPREWPEnd 
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
            Left            =   3240
            TabIndex        =   35
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSPEWPEnd 
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
            Left            =   3240
            TabIndex        =   37
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label99 
            Caption         =   "Staff Disc Req"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label100 
            Caption         =   "End of Period"
            Height          =   255
            Left            =   2160
            TabIndex        =   56
            Top             =   765
            Width           =   1005
         End
         Begin VB.Label Label98 
            Caption         =   "End of Period"
            Height          =   255
            Left            =   2160
            TabIndex        =   55
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label97 
            Caption         =   "End of Period"
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   54
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label96 
            Caption         =   "Gen W. P."
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label Label95 
            Caption         =   "Pre W. P."
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label Label94 
            Caption         =   "Spe. Ill. W. P."
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   2025
         End
      End
      Begin VB.ComboBox txtMBMClinical 
         Height          =   315
         Left            =   7440
         TabIndex        =   49
         Top             =   6600
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Frame Frame14 
         Height          =   495
         Left            =   8400
         TabIndex        =   46
         Top             =   1680
         Width           =   2535
         Begin VB.TextBox txtPLNCode 
            Height          =   225
            Left            =   480
            MaxLength       =   4
            TabIndex        =   41
            Top             =   150
            Width           =   615
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
            Height          =   270
            Left            =   1440
            TabIndex        =   42
            Top             =   150
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   476
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            Caption         =   "PLN"
            Height          =   255
            Left            =   80
            TabIndex        =   48
            Top             =   165
            Width           =   405
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            Caption         =   "Eff"
            Height          =   255
            Left            =   1150
            TabIndex        =   47
            Top             =   165
            Width           =   285
         End
      End
      Begin VB.ComboBox txtMBMWorkNature 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   405
         Width           =   3640
      End
      Begin VB.TextBox txtMBMWeight 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   2
         Top             =   690
         Width           =   1215
      End
      Begin VB.TextBox txtMBMHeight 
         Height          =   285
         Left            =   3785
         MaxLength       =   3
         TabIndex        =   3
         Top             =   690
         Width           =   1185
      End
      Begin VB.TextBox txtMBMEmail 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   4
         Top             =   950
         Width           =   3645
      End
      Begin VB.ComboBox cboLGNCode 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   1185
         Width           =   1215
      End
      Begin VB.ComboBox cboMBMSmoking 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   1470
         Width           =   1215
      End
      Begin VB.ComboBox cboMBMBlood 
         Height          =   315
         Left            =   3780
         TabIndex        =   6
         Top             =   1185
         Width           =   1185
      End
      Begin VB.ComboBox cboMBMAlcohol 
         Height          =   315
         Left            =   3780
         TabIndex        =   8
         Top             =   1470
         Width           =   1185
      End
      Begin VB.TextBox txtMBMTelNoH 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1755
         Width           =   1215
      End
      Begin VB.TextBox txtMBMTelNoM 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3765
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1755
         Width           =   1200
      End
      Begin VB.TextBox txtMBMTelNoO 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtMBMRBAmt 
         Height          =   330
         Left            =   3765
         MaxLength       =   8
         TabIndex        =   12
         Top             =   2040
         Width           =   1200
      End
      Begin VB.TextBox txtMBMPrevPlan 
         Height          =   330
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   13
         Top             =   2325
         Width           =   1215
      End
      Begin VB.ComboBox cboCOPayRB 
         Height          =   315
         ItemData        =   "frmNewMBMOthers.frx":0000
         Left            =   3765
         List            =   "frmNewMBMOthers.frx":0007
         TabIndex        =   14
         Top             =   2340
         Width           =   1200
      End
      Begin VB.TextBox txtMBMPrevINSCom 
         Height          =   330
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   15
         Top             =   2625
         Width           =   3645
      End
      Begin VB.ComboBox cboBTBG 
         Height          =   315
         ItemData        =   "frmNewMBMOthers.frx":000F
         Left            =   1320
         List            =   "frmNewMBMOthers.frx":0016
         TabIndex        =   16
         Top             =   2955
         Width           =   1215
      End
      Begin VB.TextBox TxtPOLSubNo 
         Height          =   300
         Left            =   3840
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   2955
         Width           =   1125
      End
      Begin VB.TextBox txtPayPrem 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   18
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtPayPremNett 
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   19
         Top             =   3225
         Width           =   1125
      End
      Begin VB.TextBox txtPayMCOGross 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   20
         Top             =   3495
         Width           =   1215
      End
      Begin VB.TextBox txtPolicyDisc 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   22
         Top             =   3750
         Width           =   1215
      End
      Begin VB.TextBox txtPayMCONett 
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   21
         Top             =   3480
         Width           =   1125
      End
      Begin VB.TextBox txtRenewalDisc 
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   23
         Top             =   3735
         Width           =   1125
      End
      Begin VB.TextBox txtPolicyDiscAmt 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   24
         Top             =   3975
         Width           =   1215
      End
      Begin VB.TextBox txtRenewalDiscAmt 
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   25
         Top             =   3960
         Width           =   1125
      End
      Begin VB.TextBox txtLoadingPrem 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   26
         Top             =   4230
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   3840
         TabIndex        =   27
         Top             =   4200
         Width           =   255
      End
      Begin VB.TextBox txtBankACNo 
         Height          =   285
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   28
         Top             =   4485
         Width           =   3615
      End
      Begin VB.ComboBox cboPOLCondition 
         Height          =   315
         ItemData        =   "frmNewMBMOthers.frx":001E
         Left            =   1320
         List            =   "frmNewMBMOthers.frx":0020
         TabIndex        =   29
         Top             =   4725
         Width           =   1215
      End
      Begin VB.TextBox txtUNDExcess 
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   30
         Top             =   4740
         Width           =   1100
      End
      Begin VB.TextBox txtEndtRefNo 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   31
         Top             =   4995
         Width           =   3615
      End
      Begin VB.TextBox txtSendingMode 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   100
         Top             =   5235
         Width           =   3585
      End
      Begin MSMask.MaskEdBox txtCardSentDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   101
         Top             =   5475
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtAlterDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   300
         Left            =   8490
         TabIndex        =   106
         Top             =   5160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtReinstateNo 
         Height          =   525
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   105
         Top             =   5400
         Width           =   3345
      End
      Begin VB.TextBox txtagentEmail 
         Height          =   285
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   113
         Top             =   6000
         Width           =   3615
      End
      Begin VB.Label Label22 
         Caption         =   "Agent Contact"
         Height          =   255
         Left            =   5160
         TabIndex        =   114
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label lblAgntEmail 
         Caption         =   "Agent Email"
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label lblAgent 
         Caption         =   "Agent Name"
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Alteration Type"
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   109
         Top             =   5180
         Width           =   1365
      End
      Begin VB.Label Label97 
         Caption         =   "Alteration Date"
         Height          =   255
         Index           =   2
         Left            =   7420
         TabIndex        =   108
         Top             =   5175
         Width           =   1365
      End
      Begin VB.Label Label96 
         Caption         =   "Reinstatement"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   107
         Top             =   5400
         Width           =   1620
      End
      Begin VB.Label Label96 
         Caption         =   "Sending Mode"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   103
         Top             =   5280
         Width           =   1620
      End
      Begin VB.Label Label97 
         Caption         =   "Card Sent Date"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   102
         Top             =   5520
         Width           =   1245
      End
      Begin VB.Label Label21 
         Caption         =   "Endt. Ref. No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   99
         Top             =   5055
         Width           =   1365
      End
      Begin VB.Label Label20 
         Caption         =   "Und. Excess"
         Height          =   255
         Left            =   2640
         TabIndex        =   98
         Top             =   4800
         Width           =   1485
      End
      Begin VB.Label Label19 
         Caption         =   "Back A/C No"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   4500
         Width           =   1365
      End
      Begin VB.Label Label18 
         Caption         =   "Pol Condition"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   4785
         Width           =   1140
      End
      Begin VB.Label Label17 
         Caption         =   "Loading Prem"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   4230
         Width           =   1365
      End
      Begin VB.Label Label15 
         Caption         =   "Email"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label31 
         Caption         =   "Occupation"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label14 
         Caption         =   "Work Nature"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   435
         Width           =   1140
      End
      Begin VB.Label Label35 
         Caption         =   "Height (cm)"
         Height          =   255
         Left            =   2640
         TabIndex        =   85
         Top             =   725
         Width           =   1005
      End
      Begin VB.Label Label34 
         Caption         =   "Weight (kgs)"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label105 
         Caption         =   "New Plan Code"
         Height          =   255
         Left            =   8400
         TabIndex        =   83
         Top             =   1500
         Width           =   1245
      End
      Begin VB.Label Label41 
         Caption         =   "Allergic"
         Height          =   255
         Left            =   5040
         TabIndex        =   82
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Claim Non-Pay Date"
         Height          =   255
         Left            =   5040
         TabIndex        =   81
         Top             =   1500
         Width           =   1725
      End
      Begin VB.Label Label12 
         Caption         =   "Alcohol (Y/N)"
         Height          =   255
         Left            =   2640
         TabIndex        =   80
         Top             =   1500
         Width           =   1275
      End
      Begin VB.Label Label11 
         Caption         =   "Smoking (Y/N)"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   1515
         Width           =   1275
      End
      Begin VB.Label Label37 
         Caption         =   "Blood Type"
         Height          =   255
         Left            =   2640
         TabIndex        =   78
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label Label10 
         Caption         =   "Pref. Language"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   1250
         Width           =   1140
      End
      Begin VB.Label Label9 
         Caption         =   "Ren. Disc. Rate"
         Height          =   255
         Left            =   2640
         TabIndex        =   76
         Top             =   3750
         Width           =   1485
      End
      Begin VB.Label Label8 
         Caption         =   "Pol. Disc. Rate"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   3780
         Width           =   1365
      End
      Begin VB.Label Label6 
         Caption         =   "Pay MCO Nett"
         Height          =   255
         Left            =   2640
         TabIndex        =   74
         Top             =   3465
         Width           =   1245
      End
      Begin VB.Label Label5 
         Caption         =   "Pay MCO Gross"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   3480
         Width           =   1365
      End
      Begin VB.Label Label112 
         Caption         =   "Pay Prem Nett"
         Height          =   255
         Left            =   2640
         TabIndex        =   72
         Top             =   3255
         Width           =   1245
      End
      Begin VB.Label Label110 
         Caption         =   "Pay Prem Gross"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   3225
         Width           =   1365
      End
      Begin VB.Label Label106 
         Caption         =   "BTB Guarantee"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   2955
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Pol Sub No"
         Height          =   255
         Left            =   2640
         TabIndex        =   69
         Top             =   3030
         Width           =   1125
      End
      Begin VB.Label Label93 
         Caption         =   "Co-Pay RB"
         Height          =   240
         Left            =   2640
         TabIndex        =   68
         Top             =   2355
         Width           =   1140
      End
      Begin VB.Label Label7 
         Caption         =   "Label Status"
         Height          =   255
         Left            =   2640
         TabIndex        =   67
         Top             =   4200
         Width           =   885
      End
      Begin VB.Label Label88 
         Caption         =   "T/O Pln Code"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2380
         Width           =   1125
      End
      Begin VB.Label Label87 
         Caption         =   "Prev Ins Com"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2670
         Width           =   1140
      End
      Begin VB.Label Label86 
         Caption         =   "R&&B Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   64
         Top             =   2130
         Width           =   780
      End
      Begin VB.Label Label28 
         Caption         =   "Tel Num (H)"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label32 
         Caption         =   "H/P Num"
         Height          =   255
         Left            =   2640
         TabIndex        =   62
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label33 
         Caption         =   "Tel Num (O)"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   2085
         Width           =   1005
      End
      Begin VB.Label Label16 
         Caption         =   "Pol. Disc. Amt"
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   4005
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Ren. Disc. Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   92
         Top             =   3975
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   89
      Top             =   6840
      Width           =   11100
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   9960
         TabIndex        =   44
         Top             =   160
         Width           =   735
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   91
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label46 
         Caption         =   "Member No:"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Member Others"
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
      TabIndex        =   95
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "frmNewMBMOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Dim txtMBMPayorEffDate As String
Dim HISDBComList As New ADODB.Connection
Dim MAintComList As New ADODB.Connection

Private Sub cmdOK_Click()
    If Source = 1 Then
        frmNewMBMAdjustment.PolDics = IIf(IsNumeric(txtPolicyDisc), txtPolicyDisc, 0)
        frmNewMBMAdjustment.RenDisc = IIf(IsNumeric(txtRenewalDisc), txtRenewalDisc, 0)
        frmNewMBMAdjustment.MemberCallFees
    ElseIf Source = 3 Then
        frmNewMBMReg05.PolDics = IIf(IsNumeric(txtPolicyDisc), txtPolicyDisc, 0)
        frmNewMBMReg05.RenDisc = IIf(IsNumeric(txtRenewalDisc), txtRenewalDisc, 0)
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        frmNewMBMReg05.MemberCallFees
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
    ElseIf Source = 5 Then
        FrmTempMBMReg.PolDics = IIf(IsNumeric(txtPolicyDisc), txtPolicyDisc, 0)
        FrmTempMBMReg.RenDisc = IIf(IsNumeric(txtRenewalDisc), txtRenewalDisc, 0)
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        FrmTempMBMReg.MemberCallFees
     '   Set medixmasterconn = Nothing
    ElseIf Source = 6 Then
        FrmTempMBMAdj.PolDics = IIf(IsNumeric(txtPolicyDisc), txtPolicyDisc, 0)
        FrmTempMBMAdj.RenDisc = IIf(IsNumeric(txtRenewalDisc), txtRenewalDisc, 0)
        FrmTempMBMAdj.MemberCallFees
    End If
    frmNewMBMOthers.Visible = False
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

Private Sub Form_Load()
    Call ComboListing
    If Source = 3 Then
        txtMBMPayorEffDate = frmNewMBMReg05.txtMBMPayorEffDate
    ElseIf Source = 1 Then
        txtMBMPayorEffDate = frmNewMBMAdjustment.txtMBMPayorEffDate
    ElseIf Source = 5 Then
        txtMBMPayorEffDate = FrmTempMBMReg.txtMBMPayorEffDate
    ElseIf Source = 6 Then
        txtMBMPayorEffDate = FrmTempMBMAdj.txtMBMPayorEffDate
    End If
End Sub

Private Sub TxtClmNonPayFr_LostFocus()
    If IsDate(TxtClmNonPayFr) Then
    Else
    If TxtClmNonPayFr <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        TxtClmNonPayFr.SetFocus
    End If
    End If
End Sub

Private Sub TxtClmNonPayTo_LostFocus()
    If IsDate(TxtClmNonPayTo) Then
    Else
    If TxtClmNonPayTo <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        TxtClmNonPayTo.SetFocus
    End If
    End If

End Sub

Private Sub TxtGenWP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtGenWP_LostFocus()
    If TxtGenWP <> "" Then
        If TxtGenWP > 0 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                txtGENWPEnd = Format(DateAdd("d", Trim(TxtGenWP), txtMBMPayorEffDate), "dd/mm/yyyy")
            End If
        Else
            txtGENWPEnd.mask = ""
            txtGENWPEnd = ""
            txtGENWPEnd.mask = "##/##/####"
        End If
    Else
        txtGENWPEnd.mask = ""
        txtGENWPEnd = ""
        txtGENWPEnd.mask = "##/##/####"
    End If

End Sub

Private Sub txtMBMAllergic_LostFocus()
    txtMBMAllergic = ChkStr(txtMBMAllergic)
End Sub

Private Sub txtMBMHeight_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMRBAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMTelNoH_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMTelNoM_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMTelNoO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMWeight_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtNewPolEffDate_LostFocus()
    If IsDate(txtNewPolEffDate) Then
    Else
    If txtNewPolEffDate <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtNewPolEffDate.SetFocus
    End If
    End If

End Sub

Private Sub txtPayMCOGross_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPayMCONett_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPayPrem_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPayPremNett_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPolicyDisc_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPolicyDiscAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtPOLSubNo_LostFocus()
    TxtPOLSubNo = ChkStr(TxtPOLSubNo)
End Sub

Private Sub txtPREWP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPREWP_LostFocus()
    If txtPREWP <> "" Then
        If txtPREWP > 0 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                txtPREWPEnd = Format(DateAdd("d", Trim(txtPREWP), txtMBMPayorEffDate), "dd/mm/yyyy")
            End If
        Else
            txtPREWPEnd.mask = ""
            txtPREWPEnd = ""
            txtPREWPEnd.mask = "##/##/####"
        End If
    Else
        txtPREWPEnd.mask = ""
        txtPREWPEnd = ""
        txtPREWPEnd.mask = "##/##/####"
    End If

End Sub

Private Sub txtRenewalDisc_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtRenewalDiscAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSpeWP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSpeWP_LostFocus()
    If TxtSpeWP <> "" Then
        If TxtSpeWP > 0 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                txtSPEWPEnd = Format(DateAdd("d", Trim(TxtSpeWP), txtMBMPayorEffDate), "dd/mm/yyyy")
            End If
        Else
            txtSPEWPEnd.mask = ""
            txtSPEWPEnd = ""
            txtSPEWPEnd.mask = "##/##/####"
        End If
    Else
        txtSPEWPEnd.mask = ""
        txtSPEWPEnd = ""
        txtSPEWPEnd.mask = "##/##/####"
    End If

End Sub

Private Sub ComboListing()
Dim SpName As String
Dim SelCriteria As String
Dim SelFields As String

    'HISDBComList.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'MAintComList.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    'SelFields = "SELECT OCPCode as tmpCode, OCPDescription as tmpName "
    'SelCriteria = " AND OCPCode is not null "
    'Call SPComboListing(SelFields, "OCP", SelCriteria, cboMBMOccupation)
    
    'SelFields = "SELECT LGNCode as tmpCode, LGNDescription as tmpName "
    'SelCriteria = " AND LGNCode is not null "
    'Call SPMaintComboListing(SelFields, "LGN", SelCriteria, cboLGNCode)

    'SelFields = "SELECT WorkID as tmpCode, WorkDescription as tmpName "
    'SelCriteria = " AND WorkID is not null "
    'Call SPMaintComboListing(SelFields, "WorkNature", SelCriteria, txtMBMWorkNature)
    'HISDBComList.Close
    'Set HISDBComList = Nothing
    'MAintComList.Close
    'Set MAintComList = Nothing

    cboMBMBlood.Clear
    cboMBMBlood.AddItem "A"
    cboMBMBlood.AddItem "B"
    cboMBMBlood.AddItem "AB"
    cboMBMBlood.AddItem "O"
    
    cboMBMSmoking.Clear
    cboMBMSmoking.AddItem "Y"
    cboMBMSmoking.AddItem "N"
    
    cboStaffDisc.Clear
    cboStaffDisc.AddItem "N - No"
    cboStaffDisc.AddItem "Y - Yes"
    
    cboMBMAlcohol.Clear
    cboMBMAlcohol.AddItem "Y"
    cboMBMAlcohol.AddItem "N"
    
    cboBTBG.Clear
    cboBTBG.AddItem "Y - Yes"
    cboBTBG.AddItem "N - No"
    
    cboPOLCondition.Clear
    cboPOLCondition.AddItem "X - Black Listed"
    cboPOLCondition.AddItem "O - Red Flagged"
End Sub

Private Sub txtUNDExcess_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub SPMaintComboListing(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, tmpComboBox As ComboBox)
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
    
    Set cmd.ActiveConnection = MAintComList
    cmd.CommandText = "SearchMaintTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    tmpComboBox.Clear
    Do Until TmpRec.EOF
        With TmpRec
            tmpComboBox.AddItem !tmpCode & " - " & !tmpName
            TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
End Sub


Private Sub SPComboListing(tmpSelFields As String, tmpTableName As String, tmpSelCriteria As String, tmpComboBox As ComboBox)
Dim TmpRec As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
    
    Set cmd1.ActiveConnection = HISDBComList
    cmd1.CommandText = "SearchHISTableRec"
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelFields)
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd1.Parameters.Append Prm2
    Set Prm3 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, tmpSelCriteria)
    cmd1.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd1.Execute
    tmpComboBox.Clear
    Do Until TmpRec.EOF
        With TmpRec
            tmpComboBox.AddItem !tmpCode & " - " & !tmpName
            TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
End Sub
