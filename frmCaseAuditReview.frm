VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseAuditReview 
   Caption         =   "Audit Review"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   53
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   529
      TabCaption(0)   =   "Audit Review"
      TabPicture(0)   =   "frmCaseAuditReview.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame22"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "General Info"
      TabPicture(1)   =   "frmCaseAuditReview.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstMember"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Case Details"
      TabPicture(2)   =   "frmCaseAuditReview.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   1530
         Left            =   -74880
         TabIndex        =   176
         Top             =   420
         Width           =   11415
         Begin VB.TextBox txtINSCode 
            BackColor       =   &H80000001&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   14400
            TabIndex        =   187
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtAdd1 
            Height          =   285
            Left            =   960
            MaxLength       =   50
            TabIndex        =   186
            Top             =   130
            Width           =   4815
         End
         Begin VB.TextBox txtAdd2 
            Height          =   285
            Left            =   960
            MaxLength       =   50
            TabIndex        =   185
            Top             =   390
            Width           =   4815
         End
         Begin VB.TextBox txtAdd3 
            Height          =   285
            Left            =   960
            MaxLength       =   50
            TabIndex        =   184
            Top             =   650
            Width           =   4815
         End
         Begin VB.TextBox txtPostCode 
            Height          =   285
            Left            =   960
            MaxLength       =   5
            TabIndex        =   183
            Top             =   900
            Width           =   1695
         End
         Begin VB.TextBox txtCity 
            Height          =   285
            Left            =   3480
            MaxLength       =   50
            TabIndex        =   182
            Top             =   900
            Width           =   2295
         End
         Begin VB.TextBox txtState 
            Height          =   285
            Left            =   960
            MaxLength       =   50
            TabIndex        =   181
            Top             =   1160
            Width           =   4815
         End
         Begin MSMask.MaskEdBox txtDOB 
            Height          =   285
            Left            =   7440
            TabIndex        =   188
            Top             =   130
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            TabIndex        =   180
            Top             =   390
            Width           =   1095
         End
         Begin VB.TextBox txtAvailableAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            TabIndex        =   179
            Top             =   650
            Width           =   1095
         End
         Begin VB.TextBox txtRecoveryOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            TabIndex        =   178
            Top             =   900
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            TabIndex        =   177
            Top             =   1160
            Width           =   1095
         End
         Begin VB.Label txtMorW 
            Alignment       =   2  'Center
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   9120
            TabIndex        =   200
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label lblBelow2000 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Below 2000"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   9120
            TabIndex        =   199
            Top             =   720
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label Label10 
            Caption         =   "Insured Type"
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
            Left            =   13200
            TabIndex        =   198
            Top             =   1245
            Width           =   1275
         End
         Begin VB.Label Label57 
            Caption         =   "Town"
            Height          =   255
            Left            =   2880
            TabIndex        =   197
            Top             =   960
            Width           =   540
         End
         Begin VB.Label Label5 
            Caption         =   "Avail. Limit"
            Height          =   255
            Left            =   6480
            TabIndex        =   196
            Top             =   700
            Width           =   885
         End
         Begin VB.Label Label72 
            Caption         =   "PCode"
            Height          =   255
            Left            =   240
            TabIndex        =   195
            Top             =   960
            Width           =   660
         End
         Begin VB.Label Label60 
            Caption         =   "State"
            Height          =   255
            Left            =   240
            TabIndex        =   194
            Top             =   1240
            Width           =   465
         End
         Begin VB.Label Label26 
            Caption         =   "Add"
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
            Left            =   240
            TabIndex        =   193
            Top             =   200
            Width           =   1140
         End
         Begin VB.Label Label14 
            Caption         =   "D.O.B"
            Height          =   255
            Left            =   6480
            TabIndex        =   192
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label22 
            Caption         =   "Installment"
            Height          =   245
            Left            =   6480
            TabIndex        =   191
            Top             =   1220
            Width           =   825
         End
         Begin VB.Label Label4 
            Caption         =   "Amt O/S"
            Height          =   255
            Left            =   6480
            TabIndex        =   190
            Top             =   960
            Width           =   1305
         End
         Begin VB.Label Label31 
            Caption         =   "Ann Limit"
            Height          =   255
            Left            =   6480
            TabIndex        =   189
            Top             =   390
            Width           =   885
         End
      End
      Begin VB.Frame Frame7 
         Height          =   5010
         Left            =   -74880
         TabIndex        =   74
         Top             =   360
         Width           =   11415
         Begin VB.Frame Frame18 
            Height          =   1335
            Left            =   5640
            TabIndex        =   125
            Top             =   3600
            Width           =   5415
            Begin VB.ComboBox cboREPCode1 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   285
               TabIndex        =   126
               Top             =   360
               Width           =   5085
            End
            Begin VB.ComboBox cboREPCode2 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   285
               TabIndex        =   127
               Top             =   640
               Width           =   5085
            End
            Begin VB.ComboBox cboREPCode3 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   285
               TabIndex        =   128
               Top             =   920
               Width           =   5085
            End
            Begin VB.Label Label74 
               Caption         =   "3"
               Height          =   255
               Left            =   120
               TabIndex        =   133
               Top             =   960
               Width           =   255
            End
            Begin VB.Label Label75 
               Caption         =   "2"
               Height          =   255
               Left            =   120
               TabIndex        =   132
               Top             =   680
               Width           =   255
            End
            Begin VB.Label Label77 
               Caption         =   "1"
               Height          =   255
               Left            =   120
               TabIndex        =   131
               Top             =   435
               Width           =   255
            End
            Begin VB.Label Label78 
               Caption         =   "Repudiation Reasons"
               Height          =   255
               Left            =   2520
               TabIndex        =   130
               Top             =   120
               Width           =   1695
            End
            Begin VB.Label Label80 
               Caption         =   "Code"
               Height          =   255
               Left            =   480
               TabIndex        =   129
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame17 
            Height          =   1220
            Left            =   5640
            TabIndex        =   109
            Top             =   2340
            Width           =   5415
            Begin VB.TextBox cboFDICDCode1 
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   111
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtFinalDiag1 
               Height          =   315
               Left            =   325
               MaxLength       =   150
               TabIndex        =   110
               Top             =   360
               Width           =   3405
            End
            Begin MSMask.MaskEdBox txtFDDOS1 
               Height          =   315
               Left            =   3720
               TabIndex        =   112
               Top             =   330
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
               Left            =   3720
               TabIndex        =   114
               Top             =   600
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
               Left            =   3720
               TabIndex        =   117
               Top             =   885
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
            Begin VB.TextBox txtFinalDiag2 
               Height          =   315
               Left            =   325
               MaxLength       =   150
               TabIndex        =   113
               Top             =   620
               Width           =   3405
            End
            Begin VB.TextBox txtFinalDiag3 
               Height          =   285
               Left            =   325
               MaxLength       =   150
               TabIndex        =   116
               Top             =   880
               Width           =   3405
            End
            Begin VB.TextBox cboFDICDCode2 
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   115
               Top             =   620
               Width           =   615
            End
            Begin VB.TextBox cboFDICDCode3 
               Height          =   285
               Left            =   4680
               MaxLength       =   3
               TabIndex        =   118
               Top             =   880
               Width           =   615
            End
            Begin VB.Label Label66 
               Caption         =   "ICD "
               Height          =   255
               Left            =   4680
               TabIndex        =   124
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label67 
               Caption         =   "Date Onset"
               Height          =   255
               Left            =   3720
               TabIndex        =   123
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label68 
               Caption         =   "Final Diag"
               Height          =   255
               Left            =   165
               TabIndex        =   122
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label69 
               Caption         =   "1"
               Height          =   255
               Left            =   120
               TabIndex        =   121
               Top             =   435
               Width           =   255
            End
            Begin VB.Label Label70 
               Caption         =   "2"
               Height          =   255
               Left            =   120
               TabIndex        =   120
               Top             =   680
               Width           =   255
            End
            Begin VB.Label Label71 
               Caption         =   "3"
               Height          =   255
               Left            =   120
               TabIndex        =   119
               Top             =   920
               Width           =   255
            End
         End
         Begin VB.Frame Frame16 
            Height          =   1215
            Left            =   120
            TabIndex        =   93
            Top             =   3600
            Width           =   5175
            Begin VB.TextBox txtUD1 
               Height          =   285
               Left            =   340
               MaxLength       =   150
               TabIndex        =   95
               Top             =   360
               Width           =   3180
            End
            Begin VB.TextBox cboUDICDCode1 
               Height          =   285
               Left            =   4440
               MaxLength       =   3
               TabIndex        =   94
               Top             =   360
               Width           =   615
            End
            Begin MSMask.MaskEdBox txtUDDOS1 
               Height          =   300
               Left            =   3480
               TabIndex        =   96
               Top             =   345
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
            Begin VB.TextBox txtUD2 
               Height          =   285
               Left            =   340
               MaxLength       =   150
               TabIndex        =   97
               Top             =   600
               Width           =   3180
            End
            Begin VB.TextBox txtUD3 
               Height          =   285
               Left            =   360
               MaxLength       =   150
               TabIndex        =   98
               Top             =   840
               Width           =   3180
            End
            Begin MSMask.MaskEdBox txtUDDOS2 
               Height          =   300
               Left            =   3480
               TabIndex        =   99
               Top             =   600
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
               Left            =   3480
               TabIndex        =   100
               Top             =   855
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
            Begin VB.TextBox cboUDICDCode2 
               Height          =   285
               Left            =   4440
               MaxLength       =   3
               TabIndex        =   101
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox cboUDICDCode3 
               Height          =   285
               Left            =   4440
               MaxLength       =   3
               TabIndex        =   102
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label56 
               Caption         =   "ICD "
               Height          =   255
               Left            =   4440
               TabIndex        =   108
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label58 
               Caption         =   "Date Onset"
               Height          =   255
               Left            =   3480
               TabIndex        =   107
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label59 
               Caption         =   "Underlying Disease"
               Height          =   255
               Left            =   165
               TabIndex        =   106
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label61 
               Caption         =   "1"
               Height          =   255
               Left            =   120
               TabIndex        =   105
               Top             =   390
               Width           =   255
            End
            Begin VB.Label Label62 
               Caption         =   "2"
               Height          =   255
               Left            =   120
               TabIndex        =   104
               Top             =   680
               Width           =   255
            End
            Begin VB.Label Label65 
               Caption         =   "3"
               Height          =   255
               Left            =   120
               TabIndex        =   103
               Top             =   950
               Width           =   255
            End
         End
         Begin VB.Frame Frame15 
            Height          =   1220
            Left            =   120
            TabIndex        =   77
            Top             =   2340
            Width           =   5175
            Begin VB.TextBox txtPrelimDiag1 
               Height          =   285
               Left            =   280
               MaxLength       =   150
               TabIndex        =   80
               Top             =   340
               Width           =   3210
            End
            Begin VB.TextBox cboICDCode1 
               Height          =   285
               Left            =   4440
               MaxLength       =   3
               TabIndex        =   78
               Top             =   340
               Width           =   615
            End
            Begin MSMask.MaskEdBox txtPDDOS1 
               Height          =   300
               Left            =   3480
               TabIndex        =   79
               Top             =   345
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
            Begin VB.TextBox txtPrelimDiag2 
               Height          =   285
               Left            =   280
               MaxLength       =   150
               TabIndex        =   81
               Top             =   600
               Width           =   3210
            End
            Begin VB.TextBox txtPrelimDiag3 
               Height          =   285
               Left            =   280
               MaxLength       =   150
               TabIndex        =   82
               Top             =   860
               Width           =   3210
            End
            Begin MSMask.MaskEdBox txtPDDOS2 
               Height          =   300
               Left            =   3480
               TabIndex        =   83
               Top             =   600
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
            Begin MSMask.MaskEdBox txtPDDOS3 
               Height          =   300
               Left            =   3480
               TabIndex        =   84
               Top             =   855
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
            Begin VB.TextBox cboICDCode2 
               Height          =   285
               Left            =   4440
               MaxLength       =   3
               TabIndex        =   85
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox cboICDCode3 
               Height          =   285
               Left            =   4440
               MaxLength       =   3
               TabIndex        =   86
               Top             =   860
               Width           =   615
            End
            Begin VB.Label Label52 
               Caption         =   "1"
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "2"
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   640
               Width           =   255
            End
            Begin VB.Label Label54 
               Caption         =   "3"
               Height          =   250
               Left            =   120
               TabIndex        =   90
               Top             =   895
               Width           =   255
            End
            Begin VB.Label Label48 
               Caption         =   "Date Onset"
               Height          =   255
               Left            =   3480
               TabIndex        =   89
               Top             =   135
               Width           =   855
            End
            Begin VB.Label Label51 
               Caption         =   "Prelim. Diag"
               Height          =   255
               Left            =   240
               TabIndex        =   88
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label47 
               Caption         =   "ICD "
               Height          =   255
               Left            =   4440
               TabIndex        =   87
               Top             =   120
               Width           =   615
            End
         End
         Begin VB.TextBox txtPatientIC 
            Enabled         =   0   'False
            Height          =   330
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   76
            Top             =   140
            Width           =   3495
         End
         Begin VB.ComboBox cboPatientList 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
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
            TabIndex        =   75
            Top             =   120
            Width           =   5160
         End
         Begin MSMask.MaskEdBox txtFirstCall 
            Height          =   315
            Left            =   7800
            TabIndex        =   134
            Top             =   390
            Width           =   1215
            _ExtentX        =   2143
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
            Left            =   7800
            TabIndex        =   135
            Top             =   670
            Width           =   1215
            _ExtentX        =   2143
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
            Height          =   315
            Left            =   7800
            TabIndex        =   136
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSMask.MaskEdBox txtDateDischarge 
            Height          =   315
            Left            =   7800
            TabIndex        =   137
            Top             =   1245
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSMask.MaskEdBox txtFCTime 
            Height          =   255
            Left            =   9720
            TabIndex        =   138
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox txtHOSCode 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmCaseAuditReview.frx":0054
            Left            =   1200
            List            =   "frmCaseAuditReview.frx":0056
            Sorted          =   -1  'True
            TabIndex        =   139
            Top             =   400
            Width           =   5160
         End
         Begin VB.TextBox txtReason 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   1200
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   140
            Top             =   680
            Width           =   5160
         End
         Begin VB.TextBox txtATTDoc 
            Height          =   285
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   141
            Top             =   1230
            Width           =   5160
         End
         Begin VB.TextBox txtREFBy 
            Height          =   285
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   142
            Top             =   1480
            Width           =   5160
         End
         Begin VB.ComboBox cboNature 
            Height          =   315
            Left            =   1200
            TabIndex        =   143
            Top             =   1740
            Width           =   1935
         End
         Begin VB.ComboBox cboModality 
            Height          =   315
            Left            =   1200
            TabIndex        =   144
            Top             =   2030
            Width           =   1935
         End
         Begin VB.ComboBox cboStayType 
            Height          =   315
            Left            =   4440
            TabIndex        =   145
            Top             =   1740
            Width           =   1935
         End
         Begin VB.ComboBox CboPayType 
            Height          =   315
            ItemData        =   "frmCaseAuditReview.frx":0058
            Left            =   4440
            List            =   "frmCaseAuditReview.frx":005A
            TabIndex        =   146
            Top             =   2020
            Width           =   1935
         End
         Begin MSMask.MaskEdBox txtFTTime 
            Height          =   255
            Left            =   9720
            TabIndex        =   147
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDOATime 
            Height          =   255
            Left            =   9720
            TabIndex        =   148
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDODTime 
            Height          =   255
            Left            =   9720
            TabIndex        =   149
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtCASReserveAmount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7800
            TabIndex        =   151
            Top             =   1520
            Width           =   1215
         End
         Begin VB.TextBox txtCASLGAmt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9720
            TabIndex        =   150
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox cboCondwillOcc 
            Height          =   315
            Left            =   7800
            TabIndex        =   152
            Top             =   1760
            Width           =   3495
         End
         Begin VB.ComboBox cboFLWTreatment 
            Height          =   315
            Left            =   7800
            TabIndex        =   153
            Top             =   2040
            Width           =   3495
         End
         Begin VB.Label Label40 
            Caption         =   "Time"
            Height          =   255
            Left            =   9240
            TabIndex        =   171
            Top             =   1275
            UseMnemonic     =   0   'False
            Width           =   495
         End
         Begin VB.Label Label42 
            Caption         =   "Cash Type"
            Height          =   255
            Left            =   3480
            TabIndex        =   175
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Fol-up req"
            Height          =   255
            Left            =   6840
            TabIndex        =   174
            Top             =   2100
            Width           =   855
         End
         Begin VB.Label Label38 
            Caption         =   "Nat Of Clm"
            Height          =   255
            Left            =   120
            TabIndex        =   173
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Reserve Amt"
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
            Left            =   6840
            TabIndex        =   172
            Top             =   1580
            Width           =   1275
         End
         Begin VB.Label Label35 
            Caption         =   "Time"
            Height          =   255
            Left            =   9240
            TabIndex        =   170
            Top             =   1005
            Width           =   495
         End
         Begin VB.Label Label30 
            Caption         =   "Time"
            Height          =   255
            Left            =   9240
            TabIndex        =   169
            Top             =   735
            Width           =   495
         End
         Begin VB.Label Label29 
            Caption         =   "F. Treated"
            Height          =   255
            Left            =   6840
            TabIndex        =   168
            Top             =   780
            Width           =   1230
         End
         Begin VB.Label Label27 
            Caption         =   "Time"
            Height          =   255
            Left            =   9240
            TabIndex        =   167
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label25 
            Caption         =   "F. Call"
            Height          =   255
            Left            =   6840
            TabIndex        =   166
            Top             =   500
            Width           =   615
         End
         Begin VB.Label Label23 
            Caption         =   "IC Num"
            Height          =   255
            Left            =   6840
            TabIndex        =   165
            Top             =   160
            Width           =   1230
         End
         Begin VB.Label Label21 
            Caption         =   "Referred by"
            Height          =   255
            Left            =   120
            TabIndex        =   164
            Top             =   1500
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Attending Doc"
            Height          =   255
            Left            =   120
            TabIndex        =   163
            Top             =   1220
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Reason"
            Height          =   255
            Left            =   120
            TabIndex        =   162
            Top             =   800
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "D.O.D"
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
            Left            =   6840
            TabIndex        =   161
            Top             =   1320
            Width           =   675
         End
         Begin VB.Label Label8 
            Caption         =   "D.O.A"
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
            Left            =   6840
            TabIndex        =   160
            Top             =   1035
            Width           =   795
         End
         Begin VB.Label Label7 
            Caption         =   "Hospitals"
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
            Left            =   120
            TabIndex        =   159
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label39 
            Caption         =   "Treat. Mode"
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label44 
            Caption         =   "Stay Type"
            Height          =   255
            Left            =   3480
            TabIndex        =   157
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label83 
            Caption         =   "Recur"
            Height          =   255
            Left            =   6840
            TabIndex        =   156
            Top             =   1820
            Width           =   735
         End
         Begin VB.Label Label82 
            Caption         =   "Patient"
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
            Left            =   120
            TabIndex        =   155
            Top             =   165
            Width           =   1140
         End
         Begin VB.Label Label45 
            Caption         =   "LG Amt"
            Height          =   255
            Left            =   9165
            TabIndex        =   154
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   615
         End
      End
      Begin VB.Frame Frame22 
         Height          =   3705
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   11415
         Begin VB.TextBox txtINVNo 
            Height          =   285
            Left            =   1845
            MaxLength       =   50
            TabIndex        =   15
            Top             =   240
            Width           =   2010
         End
         Begin MSMask.MaskEdBox txtINVDate 
            Height          =   300
            Left            =   1845
            TabIndex        =   16
            Top             =   480
            Width           =   2010
            _ExtentX        =   3545
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
         Begin VB.TextBox txtINVAmt 
            Height          =   285
            Left            =   1845
            MaxLength       =   15
            TabIndex        =   17
            Top             =   735
            Width           =   2010
         End
         Begin VB.TextBox txtDISRec 
            Height          =   285
            Left            =   1845
            MaxLength       =   15
            TabIndex        =   18
            Top             =   990
            Width           =   2010
         End
         Begin MSMask.MaskEdBox txtDLS 
            Height          =   300
            Left            =   1845
            TabIndex        =   19
            Top             =   1245
            Width           =   2010
            _ExtentX        =   3545
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
         Begin MSMask.MaskEdBox txtDLR 
            Height          =   300
            Left            =   1845
            TabIndex        =   20
            Top             =   1515
            Width           =   2010
            _ExtentX        =   3545
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
         Begin VB.TextBox txtAuditReview 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   4920
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   240
            Width           =   6375
         End
         Begin VB.TextBox TxtARAnswer 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   4920
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   1040
            Width           =   6375
         End
         Begin VB.ComboBox CboARStatus 
            Height          =   315
            Left            =   1845
            TabIndex        =   21
            Top             =   1790
            Width           =   2010
         End
         Begin MSMask.MaskEdBox TxtDateOpen 
            Height          =   300
            Left            =   1845
            TabIndex        =   22
            Top             =   2040
            Width           =   2010
            _ExtentX        =   3545
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
         Begin MSMask.MaskEdBox TxtDRTC 
            Height          =   300
            Left            =   1845
            TabIndex        =   23
            Top             =   2300
            Width           =   2010
            _ExtentX        =   3545
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
         Begin VB.TextBox txtARRemarks 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   4920
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   1790
            Width           =   6375
         End
         Begin VB.TextBox TxtPVNo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   28
            Top             =   2865
            Width           =   6375
         End
         Begin VB.TextBox TxtPVDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   29
            Top             =   3120
            Width           =   6375
         End
         Begin VB.TextBox TxtAmtPaid 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   212
            Top             =   3360
            Width           =   6375
         End
         Begin VB.ComboBox cboDiscType 
            Height          =   315
            Left            =   1845
            TabIndex        =   24
            Top             =   2570
            Width           =   2010
         End
         Begin VB.Label Label53 
            Caption         =   "Discount Type"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   214
            Top             =   2655
            Width           =   1575
         End
         Begin VB.Label Label88 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   3960
            TabIndex        =   211
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label87 
            Caption         =   "Date Open"
            Height          =   255
            Left            =   240
            TabIndex        =   210
            Top             =   2080
            Width           =   1455
         End
         Begin VB.Label Label86 
            Caption         =   "Date to Claims"
            Height          =   375
            Left            =   240
            TabIndex        =   208
            Top             =   2350
            Width           =   1815
         End
         Begin VB.Label Label85 
            Caption         =   "Amt Paid"
            Height          =   255
            Left            =   3960
            TabIndex        =   207
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label84 
            Caption         =   "PV Date"
            Height          =   255
            Left            =   3960
            TabIndex        =   206
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label81 
            Caption         =   "PV Number"
            Height          =   255
            Left            =   3960
            TabIndex        =   205
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label79 
            Caption         =   "Answer"
            Height          =   255
            Left            =   3960
            TabIndex        =   204
            Top             =   1050
            Width           =   615
         End
         Begin VB.Label Label53 
            Caption         =   "Audit Review Status"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   203
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label46 
            Caption         =   "Query"
            Height          =   255
            Left            =   3960
            TabIndex        =   73
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Discount Received"
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   1035
            Width           =   1575
         End
         Begin VB.Label Label15 
            Caption         =   "Inv Amount"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   780
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Inv Number"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   255
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Date Letter Reply"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label32 
            Caption         =   "Date Letter Sent"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   1290
            Width           =   1335
         End
         Begin VB.Label Label49 
            Caption         =   "Invoice Date"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   525
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Hospital Visit"
         Height          =   975
         Left            =   120
         TabIndex        =   61
         Top             =   4080
         Width           =   11415
         Begin VB.TextBox txtHospVisit 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   4560
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   6735
         End
         Begin VB.TextBox txtOfficer 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   32
            Top             =   600
            Width           =   2730
         End
         Begin MSMask.MaskEdBox txtDV 
            Height          =   315
            Left            =   1200
            TabIndex        =   30
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSMask.MaskEdBox txtDOVTime 
            Height          =   255
            Left            =   3000
            TabIndex        =   31
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label37 
            Caption         =   "D.O.Visit"
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
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label41 
            Caption         =   "Time"
            Height          =   255
            Left            =   2520
            TabIndex        =   64
            Top             =   195
            Width           =   495
         End
         Begin VB.Label Label43 
            Caption         =   "Action"
            Height          =   255
            Left            =   4080
            TabIndex        =   63
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label50 
            Caption         =   "Off-in-Charge"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   600
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView lstBenefitDetails 
         Height          =   5115
         Left            =   -74880
         TabIndex        =   54
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
         Height          =   3495
         Left            =   -74880
         TabIndex        =   201
         Top             =   1980
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6165
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   375
         Left            =   840
         TabIndex        =   209
         Top             =   4440
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Case No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PV No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "PV Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amt Paid"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame13 
      Height          =   495
      Left            =   75
      TabIndex        =   35
      Top             =   240
      Width           =   11655
      Begin VB.CommandButton cmdNextRevCase 
         Caption         =   "Next &Rev Case"
         Height          =   285
         Left            =   4680
         TabIndex        =   213
         Top             =   160
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   300
         Left            =   2640
         TabIndex        =   1
         Top             =   160
         Width           =   570
      End
      Begin VB.TextBox txtCaseNo 
         BeginProperty Font 
            Name            =   "MS Serif"
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
      Begin VB.ComboBox cboStatus1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8760
         TabIndex        =   4
         Top             =   135
         Width           =   855
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Prev"
         Height          =   285
         Left            =   3240
         TabIndex        =   2
         Top             =   160
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         Top             =   160
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.ComboBox cboClaimability 
         Enabled         =   0   'False
         Height          =   315
         Left            =   10560
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
      Begin VB.Label txtPolStatus 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7320
         TabIndex        =   37
         Top             =   165
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "Case Num"
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
         Left            =   120
         TabIndex        =   40
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label55 
         Caption         =   "Pol Status"
         Height          =   255
         Left            =   6480
         TabIndex        =   39
         Top             =   165
         Width           =   915
      End
      Begin VB.Label Label34 
         Caption         =   "Case Status"
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
         Left            =   7800
         TabIndex        =   38
         Top             =   165
         Width           =   1140
      End
      Begin VB.Label Label36 
         Caption         =   "Claimability"
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
         Left            =   9720
         TabIndex        =   36
         Top             =   165
         Width           =   1140
      End
   End
   Begin VB.Frame Frame9 
      Height          =   720
      Left            =   75
      TabIndex        =   41
      Top             =   720
      Width           =   11655
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1050
         MaxLength       =   18
         TabIndex        =   6
         Top             =   130
         Width           =   2175
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   8
         Top             =   130
         Width           =   2415
      End
      Begin VB.TextBox txtRenewal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         TabIndex        =   11
         Top             =   130
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtMBMEffDate 
         Height          =   285
         Left            =   8760
         TabIndex        =   13
         Top             =   130
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtMBMName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   7
         Top             =   390
         Width           =   2175
      End
      Begin VB.TextBox txtPayCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   9
         Top             =   390
         Width           =   465
      End
      Begin VB.TextBox txtMBMPolNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         MaxLength       =   25
         TabIndex        =   10
         Top             =   390
         Width           =   1335
      End
      Begin VB.TextBox txtTakeOver 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         TabIndex        =   12
         Top             =   390
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtMBMExpiryDate 
         Height          =   285
         Left            =   8760
         TabIndex        =   14
         Top             =   390
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label73 
         Caption         =   "Pol Num"
         Height          =   255
         Left            =   4680
         TabIndex        =   44
         Top             =   420
         Width           =   780
      End
      Begin VB.Label Label64 
         Caption         =   "IC Num"
         Height          =   255
         Left            =   3600
         TabIndex        =   52
         Top             =   165
         Width           =   1230
      End
      Begin VB.Label Label76 
         Caption         =   "Mem Num"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   165
         Width           =   795
      End
      Begin VB.Label Label33 
         Caption         =   "Payor"
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
         Left            =   3600
         TabIndex        =   50
         Top             =   405
         Width           =   1275
      End
      Begin VB.Label Label16 
         Caption         =   "Mem Name"
         Height          =   230
         Left            =   120
         TabIndex        =   49
         Top             =   440
         Width           =   900
      End
      Begin VB.Label Label63 
         Caption         =   "T/Over"
         Height          =   255
         Left            =   6720
         TabIndex        =   48
         Top             =   435
         Width           =   825
      End
      Begin VB.Label Label12 
         Caption         =   "Renewal"
         Height          =   255
         Left            =   6720
         TabIndex        =   47
         Top             =   160
         Width           =   825
      End
      Begin VB.Label Label28 
         Caption         =   "Eff Date"
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
         Left            =   8040
         TabIndex        =   46
         Top             =   160
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "Exp Date"
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
         Left            =   8040
         TabIndex        =   45
         Top             =   405
         Width           =   1275
      End
      Begin VB.Label lblWaitingPeriod 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Waiting Period"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   9960
         TabIndex        =   43
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblExpired 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Expired"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   9960
         TabIndex        =   42
         Top             =   405
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame Frame10 
      Height          =   615
      Left            =   120
      TabIndex        =   55
      Top             =   7080
      Width           =   11610
      Begin VB.CommandButton cmdUpdate1 
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9600
         TabIndex        =   34
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton cmdExit1 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10680
         TabIndex        =   60
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox txtINSCode1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   59
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtPLNCode 
         Height          =   285
         Left            =   3720
         TabIndex        =   58
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtHLTCOde 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtSRVPath 
         Height          =   285
         Left            =   3840
         TabIndex        =   56
         Text            =   "Text2"
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Audit Review"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   202
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmCaseAuditReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intLastPoz2 As Integer
Dim FirstDiag As Integer
Dim FinalDiag As Integer
Dim SaveOthers As Integer
Dim LGNumber As String
Dim EXCNumber As String
Dim MBMNumber As String
Dim TempCASNo As String
Dim RB As String
Dim patient As String
Dim ReprintLG As Boolean
Dim bTabDone As Boolean
Dim ScanCode As String

Private Sub CboARStatus_Change()
    If CboARStatus.Text = "" Then
        If InStr(CboARStatus.Text, "null") Then
           CboARStatus.Enabled = False
        End If
    End If
End Sub

Private Sub CboARStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
        
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(CboARStatus.Text, CboARStatus.SelStart) + Chr(KeyAscii) + Mid(CboARStatus.Text, CboARStatus.SelStart + CboARStatus.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboARStatus.ListCount - 1
            
            If NewText = LCase(Left(CboARStatus.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboARStatus.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                    CboARStatus.Text = ValidValue
                    CboARStatus.SelStart = 0
                    CboARStatus.SelLength = Len(ValidValue)
                 End If
                 
            End If
    
End Sub

Private Sub cboDiscType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
        
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboDiscType.Text, cboDiscType.SelStart) + Chr(KeyAscii) + Mid(cboDiscType.Text, cboDiscType.SelStart + cboDiscType.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboDiscType.ListCount - 1
            
            If NewText = LCase(Left(cboDiscType.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboDiscType.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
                 If ValidCount = 1 Then
                    cboDiscType.Text = ValidValue
                    cboDiscType.SelStart = 0
                    cboDiscType.SelLength = Len(ValidValue)
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

    If txtMBMNumber <> "" And IsNumeric(Len(Mid(cboPatientList, 1, 1))) = True Then
    
        'StrSql = "Select * from MBM where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        'MBM.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "1"
        strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_AuditReview"
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
        
        'strsqlCOV = "Select * from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 1) & "'"
        'MBMCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "2"
        strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Mid(cboPatientList, 1, 1) & "'"
        Set cnn.ActiveConnection = medixmasterconn
        cnn.CommandText = "Case_AuditReview"
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
        
        If MBM.EOF Then
            MsgBox "Sorry, No Record Found", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            txtPatientIC = MBM!MBMIcBcPp
        End If
    
        If Not MBMCov.EOF Then
            txtPatientIC = MBMCov!MBMCICBCPPNo
        End If
       
        MBM.Close
        MBMCov.Close
        Set MBM = Nothing
        Set MBMCov = Nothing
    End If

End Sub

Private Sub cmdExit1_Click()
        Unload Me
        Set frmCaseAuditReview = Nothing

End Sub

Private Sub cmdNextRevCase_Click()
Dim CAS As New ADODB.Recordset
Dim sql As String
Dim CASStatus As Boolean
Dim CasNumber As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strOrderby As String
    
    CmdNext.Enabled = False
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'sql = " Select CASNumber from CASOthers where CASNumber > '" & RemStr(txtCaseNo) & "' And CASLetterSent IS Not Null" & _
        " Order by CASnumber "
    CAS.MaxRecords = 1
    'CAS.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "3"
    strAppend = " AND CASNumber > '" & RemStr(txtCaseNo) & "'"
    strOrderby = " Order by CASnumber"
    strAppend = strAppend + strOrderby
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
    If CASStatus = True Then
        txtCaseNo.Text = CasNumber
        cmdShow_Click
    Else
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
    End If
    CmdNext.Enabled = True

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

Private Sub CmdNext_Click()
Dim CAS As New ADODB.Recordset
Dim sql As String
Dim CASStatus As Boolean
Dim CasNumber As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strOrderby As String
    
    CmdNext.Enabled = False
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'sql = " select CASNumber from CAS where CASNumber > '" & RemStr(txtCaseNo) & "' " & _
        " Order by CASnumber "
    CAS.MaxRecords = 1
    'CAS.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "4"
    strAppend = " AND CASNumber > '" & RemStr(txtCaseNo) & "'"
    strOrderby = " Order by CASnumber"
    strAppend = strAppend + strOrderby
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
    If CASStatus = True Then
        txtCaseNo.Text = CasNumber
        cmdShow_Click
    Else
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
    End If
    CmdNext.Enabled = True
End Sub

Private Sub cmdPrev_Click()
Dim CAS As New ADODB.Recordset
Dim sql As String
Dim CASStatus As Boolean
Dim CasNumber As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strOrderby As String
    
    cmdPrev.Enabled = False
    If Trim(txtCaseNo) <> "" Then
        
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

        'sql = " select CASNumber from CAS where CASnumber < '" & RemStr(txtCaseNo) & "' " & _
            "Order by CASNumber DESC "
        CAS.MaxRecords = 1
        'CAS.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "4"
        strAppend = " AND CASnumber < '" & RemStr(txtCaseNo) & "'"
        strOrderby = " Order by CASnumber DESC"
        strAppend = strAppend + strOrderby
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_AuditReview"
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
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
        
        If CASStatus = True Then
            txtCaseNo.Text = CasNumber
            cmdShow_Click
        Else
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Case Number Format! ", vbCritical + vbOKOnly
        End If
    End If
    cmdUpdate1.Visible = True
    cmdPrev.Enabled = True
End Sub

Public Sub cmdShow_Click()
Dim getRecord As Integer
    
    cmdShow.Enabled = False
    Frame2.Enabled = False
    Frame7.Enabled = False
    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
        txtCaseNo.SetFocus
    Else
        ScanCode = Trim(txtCaseNo)
        If Mid(ScanCode, 1, 2) = "LO" Then
            txtCaseNo = ""
            Call SearchForm(ScanCode)
        Else
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            'medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
            TempCASNo = txtCaseNo
            patient = cboPatientList
            Call ClearEntries
            cboPatientList.Clear
            txtCaseNo = TempCASNo
            getRecord = showCAS
            Screen.MousePointer = vbHourglass
            Call showMBM
            MBMNumber = Trim(txtMBMNumber)
            Call ShowFirstDiag
            Call ShowFinalDiag
            Call ShowUD
            Call ShowRepudiation
            Call CallLimit
            Call showOther
            Call showMBMListing
            Call showMBMCovListing
            Call showPaymentRecord
            If txtDOB <> "__/__/____" Then
                txtMorW = Year(Date) - Year(txtDOB)
                txtMorW = txtMorW & " Years Old"
            End If
            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
            'medixmasterconnPayment.Close
            'Set medixmasterconnPayment = Nothing
        End If
        Screen.MousePointer = Default
    End If
    cmdShow.Enabled = True
End Sub

Private Function showCAS()
Dim sqlstr As String
Dim CAS As New ADODB.Recordset
Dim CoverID As String
Dim HOS As New ADODB.Recordset
Dim strsqlHOS As String
Dim cmd As New ADODB.Command
Dim cnn As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    'sqlstr = "Select CASNumber, MBMNumber, MBMPatientCovID, " & _
            " SRVCode, CASStatus, CASDischargeDate, CASAdmissionReason, CASLGAmt, CASStayType, CASPayType, " & _
            " CASERegDate, CaseDateOfLoss, CASDateAdmitted, CASHOSCode, CASReserveAmount, CASReviewAudit, " & _
            " CASClaimability, MBMName, MBMCName, AttendDoctor, ReferedBy, PayTypeCode, ClaimNatureCode, ModalityCode, " & _
            " CASInvNo, CASInvAmt, InvDate, Officer, CASLetterSent, CASHospReply, CASDisRec, CASARStatus "
    'sqlstr = sqlstr & " From VIEW_CAS_AuditRev where CASNumber = '" & Trim(txtCaseNo) & "'"
    'CAS.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "5"
    strAppend = " AND CASNumber = '" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
        
    If CAS.EOF Then
       MsgBox "No record found!", vbCritical + vbOKOnly
    Else
        showCAS = 1
        With CAS
            txtCaseNo = CAS!CasNumber
            txtMBMNumber = Trim(CAS!MBMNumber)
            If CAS!MBMPatientCovID = "00" Then
                cboPatientList = Trim(CAS!MBMPatientCovID) & " - " & Trim(CAS!MBMName)
            Else
                cboPatientList = Trim(CAS!MBMPatientCovID) & " - " & Trim(CAS!MBMCName)

            End If
            
            If Len(!CASDateAdmitted) Then
                txtDateAdmitted.mask = ""
                txtDateAdmitted = Format(!CASDateAdmitted, "dd/mm/yyyy")
                txtDateAdmitted.mask = "##/##/####"
            End If
            

            If Len(CAS!CASDischargeDate) Then
                txtDateDischarge.mask = ""
                txtDateDischarge = Format(CAS!CASDischargeDate, "dd/mm/yyyy")
                txtDateDischarge.mask = "##/##/####"
            End If
                       
       'strsqlHOS = "Select * from HOS where HOSCode = '" & Trim(CAS!CASHOSCode) & "'"
       'HOs.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "6"
        strAppend = " AND HOSCode = '" & Trim(CAS!CASHOSCode) & "'"
        Set cnn.ActiveConnection = medixmasterconn
        cnn.CommandText = "Case_AuditReview"
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
        
            If Not HOS.EOF Then
                 If Len(CAS!CASHOSCode) Then
                     txtHOSCode = HOS!HosName
                 End If
            End If
        HOS.Close
        Set HOS = Nothing

            If Len(CAS!CASStayType) Then
                cboStayType = !CASStayType
            End If

            If Len(CAS!CASPAYType) Then
                CboPayType = !CASPAYType
            End If
            
            If Len(!CASAdmissionReason) Then
                txtReason = !CASAdmissionReason
            End If
            
            If Len(CAS!ClaimNatureCode) Then
                cboNature = !ClaimNatureCode
            End If
            If Len(CAS!ModalityCode) Then
                cboModality = !ModalityCode
            End If
            If Len(CAS!CASAttendDoctor1) Then
                txtATTDoc = !CASAttendDoctor1
            End If
            If Len(CAS!ReferedBy) Then
                txtREFBy = !ReferedBy
            End If
            If Len(!CASReserveAmount) Then
                txtCASReserveAmount = !CASReserveAmount
            End If
            If Len(!CASLGAmt) Then
                txtCASLGAmt = !CASLGAmt
            End If
            If Len(!CASStatus) Then
                cboStatus1.Text = !CASStatus
            End If
       'Sai Ling
            If Len(!CASClaimability) Then
                cboClaimability = !CASClaimability
            End If
        End With
    End If
   
    CAS.Close
    Set CAS = Nothing
End Function

Private Function showMBM()
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim sqlstr1 As String
Dim MBMCov As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim cnn As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    'sqlstr = " Select MBMNumber, HLTCode, MBMPolicyNo,  MBMName ," & _
            " MBMIcBcPp,  MBMMemberType , " & _
            " MBMStatus, PAYCode, AGECode , " & _
            " INSCode, PLNCode, " & _
            " MBMAddress1, MBMAddress2, " & _
            " MBMAddress3, MBMPostCode, CTYCode, STACode, " & _
            " MBMDOB , BRNCode, MBMTakeOver , MBMRenewFlag,  MBMAvailableLimit, " & _
            " MBMPayorEffDate, MBMPayorExpDate, MBMENDtEffDate  "
    'sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "7"
    strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
        
    If Not MBM.EOF Then
        If Len(MBM!MBMNumber) Then
            txtMBMNumber = Trim(MBM!MBMNumber)
        End If
        If Len(MBM!MBMName) Then
            txtMBMName = Trim(MBM!MBMName)
        End If
        If Len(MBM!MBMPolicyNo) Then
            txtMBMPolNo = Trim(MBM!MBMPolicyNo)
        End If
        If Len((MBM!MBMIcBcPp)) Then
            txtMBMIcBcPP = Trim(MBM!MBMIcBcPp)
        End If
        If Len(MBM!MBMAddress1) Then
            txtAdd1 = Trim(MBM!MBMAddress1)
        End If
        If Len(MBM!MBMAddress2) Then
            txtAdd2 = Trim(MBM!MBMAddress2)
        End If
        If Len(MBM!MBMAddress3) Then
            txtAdd3 = Trim(MBM!MBMAddress3)
        End If
        If Len(MBM!PAYCode) Then
            txtPayCode = Trim(MBM!PAYCode)
        End If
        If Len(Format(MBM!MBMPayorEffDate, "dd/mm/yyyy")) Then
            txtMBMEffDate.mask = ""
            txtMBMEffDate.Text = Trim(Format(MBM!MBMPayorEffDate, "dd/mm/yyyy"))
            txtMBMEffDate.mask = "##/##/####"
        End If
        If Len(MBM!HLTCode) Then
            txtHLTCOde = Trim(MBM!HLTCode)
        End If
        If Len(Format(MBM!MBMPayorEXPDate, "dd/mm/yyyy")) Then
            txtMBMExpiryDate.mask = ""
            txtMBMExpiryDate = Trim(Format(MBM!MBMPayorEXPDate, "dd/mm/yyyy"))
            txtMBMExpiryDate.mask = "##/##/####"
        End If
        If Len(MBM!PLNCode) Then
            txtPLNCode = Trim(MBM!PLNCode)
        End If
        If Len(MBM!INSCode) Then
            txtINSCode1 = Trim(MBM!INSCode)
        End If
        If Len(Format(MBM!MBMDOB, "dd/mm/yyyy")) Then
            txtDOB.mask = ""
            txtDOB = Trim(Format(MBM!MBMDOB, "dd/mm/yyyy"))
            txtDOB.mask = "##/##/####"
        End If
        If Len(MBM!CTYCode) Then
            txtCity = Trim(MBM!CTYCode)
        End If
        If Len(MBM!MBMPostCode) Then
            txtPostCode = Trim(MBM!MBMPostCode)
        End If
            If Len(Trim(MBM!STACode)) Then
                txtState = Trim(MBM!STACode)
            End If
            If Trim(MBM!MBMTakeover) <> "" Then
                txtTakeOver = Trim(MBM!MBMTakeover)
            End If
            If Len(MBM!MBMRenewFlag) Then
                txtRenewal = MBM!MBMRenewFlag
            End If

            txtAvailableAnnualLimit = Format(Trim(MBM!MBMAvailableLimit), "#,##0.00")
            txtPolStatus = Trim(MBM!MBMStatus)
            'sqlstr1 = " select MBMCNumber, MBMCCoverID, MBMCName, " & _
             "  MBMCAnnualLimit " & _
             " from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
            'MBMCov.Open sqlstr1, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppendCase = "8"
            strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "'"
            Set cnn.ActiveConnection = medixmasterconn
            cnn.CommandText = "Case_AuditReview"
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
        
            cboPatientList.AddItem "00 -" & txtMBMName
            
            If Not MBMCov.EOF Then
                Do While Not MBMCov.EOF
                    cboPatientList.AddItem MBMCov!MBMCCoverID & "-" & MBMCov!MBMCName
                    MBMCov.MoveNext
                Loop
            End If
            MBMCov.Close
            Set MBMCov = Nothing
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
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    'sqlstr = "Select PLNCode, INSCode, HLTCode, MBMPayorEffDate from MBM Where MBMNumber = '" & RemStr(txtMBMNumber) & "'"
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "9"
    strAppend = " AND MBMNumber = '" & RemStr(txtMBMNumber) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
        
    If Not MBM.EOF Then
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

Private Function showOther()
Dim sqlstr As String
Dim Other As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    'sqlstr = "Select * From CASOthers where CASNumber = '" & RemStr(txtCaseNo) & "'"
    'Other.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "10"
    strAppend = " AND CASNumber = '" & RemStr(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
    
    If Not Other.EOF Then
        Do Until Other.EOF
        With Other
            If Len(Trim(!CASOCondWillRecur)) Then
                 cboCondwillOcc = !CASOCondWillRecur
            End If
            If Len(Trim(!CASOFollowUpRequired)) Then
                 cboFLWTreatment = !CASOFollowUpRequired
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
            If Len(!CASFCallTime) Then
                txtFCTime = Format(!CASFCallTime, "hh:mm")
            End If
            
            If Len(!CASFTreatedTime) Then
                 txtFTTime = Format(!CASFTreatedTime, "hh:mm")
            End If
            
            If Len(!CASDOATime) Then
                 txtDOATime = Format(!CASDOATime, "hh:mm")
            End If
            
            If Len(!CASDODTime) Then
                 txtDODTime = Format(!CASDODTime, "hh:mm")
            End If
                        
            If Len(!CASDOV) Then
                 txtDV = Format(!CASDOV, "dd/mm/yyyy")
            End If
            
            If Len(!CASDOVTime) Then
                 txtDOVTime = Format(!CASDOVTime, "hh:mm")
            End If
            
            If Len(!CASVRemarks) Then
                txtHospVisit = !CASVRemarks
            End If
            
            If Len(!CASARAnswer) Then
                TxtARAnswer = !CASARAnswer
            End If
            
            If Len(!CASDTC) Then
                 TxtDRTC = Format(!CASDTC, "dd/mm/yyyy")
            End If
            
            If Len(!CASAROpenDate) Then
                 TxtDateOpen = Format(!CASAROpenDate, "dd/mm/yyyy")
            End If
            
            If Len(!CASARRemarks) Then
                 txtARRemarks = !CASARRemarks
            End If
            
            If Len(!CASReviewAudit) Then
                txtAuditReview = !CASReviewAudit
            End If
            
            If Len(!CASInvNo) Then
                txtINVNo = !CASInvNo
            End If
            
            If Len(!InvDate) Then
                txtINVDate.mask = ""
                txtINVDate = Format(!InvDate, "dd/mm/yyyy")
                txtINVDate.mask = "##/##/####"
            End If
            
            If Len(!CASInvAmt) Then
                txtINVAmt = !CASInvAmt
            End If
                        
            If Len(!Officer) Then
                txtOfficer = !Officer
            End If
                        
            If Len(!CASLetterSent) Then
                txtDLS.mask = ""
                txtDLS = Format(!CASLetterSent, "dd/mm/yyyy")
                txtDLS.mask = "##/##/####"
            End If

            If Len(!CASHospReply) Then
                txtDLR.mask = ""
                txtDLR = Format(!CASHospReply, "dd/mm/yyyy")
                txtDLR.mask = "##/##/####"
            End If

            If Len(!CASDisRec) Then
                txtDISRec = !CASDisRec
            End If
            
            If Len(!CASARStatus) Then
                CboARStatus = !CASARStatus
            End If
            
            If Len(!CASDiscType) Then
                cboDiscType = !CASDiscType
            End If
            
        End With
        Other.MoveNext
        Loop
    End If
    Other.Close
    Set Other = Nothing
End Function

Private Sub cmdUpdate1_Click()
Dim RecordSaved
Dim cmd As New ADODB.Command
Dim rst2 As New ADODB.Recordset
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strsqlCASAuditReview, strsqlCASRef As String
Dim InvDate, CASLetterSent, CASHospReply, CASDOV, CASDOVTime As String
Dim CASDTC, CASAROpenDate, CASInvAmt, CASDisRec As String
Dim AuditReview, answer, Remarks, HospVisit
Dim tempdate As String
    
    cmdUpdate1.Enabled = False
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If txtCaseNo = "" Then
        MsgBox "Please Enter Case No!", vbCritical + vbOKOnly, DialogBoxTitle
        txtCaseNo.SetFocus
    ElseIf cboDiscType = "" Then
        MsgBox "Please Enter Discount Type!", vbCritical + vbOKOnly, DialogBoxTitle
        txtCaseNo.SetFocus
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            
            strAppend = " AND CASNumber = '" & txtCaseNo & "'"
            strAppendCase = "19"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Case_AuditReview"
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
            'updatecase = "Select CASReviewAudit,CASARStatus,CASNumber from CAS Where CASNumber ='" & Trim(tempcasNo) & "' And MBMNumber ='" & Trim(MBMNumber) & " '"
            'CAS.Open updatecase, medixmasterconn, adOpenKeyset, adLockOptimistic
    
            'With CAS
                'If Len(txtAuditReview) Then
                '    !CASReviewAudit = ChkStr(txtAuditReview)
                'Else
                '    !CASReviewAudit = Null
                'End If
                
                'If Len(CboARStatus) Then
                '    !CASARStatus = Mid(CboARStatus, 1, 1)
                'Else
                '    !CASARStatus = Null
                'End If
                
            'End With
            'CAS.Update
            'CAS.Close
            'Set CAS = Nothing
        
        '---------------------------------------------------------------------------------------------------------------------------------------
        'Update the Case Others Fields
            
            'updatecothers = "Select * from CASOthers Where CASNumber ='" & Trim(tempcasNo) & "'"
            'CASOthers.Open updatecothers, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
            'If CASOthers.EOF Then
            '    CASOthers.AddNew
            '    CASOthers!CasNumber = tempcasNo
            'End If
            'With CASOthers
            '    If Len(txtINVNo) Then
            '        !CASInvNo = ChkStr(txtINVNo)
            '    Else
            '        !CASInvNo = Null
            '    End If
                
                If txtINVDate <> "__/__/____" Then
                    InvDate = Format(txtINVDate, "mm/dd/yyyy")
                Else
                    InvDate = Empty
                End If
                
                If Len(txtINVAmt) Then
                    CASInvAmt = txtINVAmt
                Else
                    CASInvAmt = 0
                End If
                
                If txtDLS <> "__/__/____" Then
                    CASLetterSent = Format(txtDLS, "mm/dd/yyyy")
                Else
                    CASLetterSent = Empty
                End If
                
                If txtDLR <> "__/__/____" Then
                    CASHospReply = Format(txtDLR, "mm/dd/yyyy")
                Else
                    CASHospReply = Empty
                End If
              
                If txtDV <> "__/__/____" Then
                     CASDOV = Format(txtDV, "mm/dd/yyyy")
                Else
                     CASDOV = Null
                End If
                 
                If txtDOVTime <> "__:__" Then
                    CASDOVTime = Format(txtDOVTime, "hh:mm")
                End If
                 
             '   If Len(txtOfficer) Then
             '       !Officer = ChkStr(txtOfficer)
             '   Else
             '       !Officer = Null
             '   End If
               
             '   If Len(txtHospVisit) Then
             '       !CASVRemarks = ChkStr(txtHospVisit)
             '   Else
             '       !CASVRemarks = Null
             '   End If
              
                If Len(txtDISRec) Then
                    CASDisRec = ChkStr(txtDISRec)
                Else
                    CASDisRec = 0
                End If
                
                If TxtDRTC <> "__/__/____" Then
                     CASDTC = Format(TxtDRTC, "mm/dd/yyyy")
                Else
                     CASDTC = Empty
                End If
                
                If TxtDateOpen <> "__/__/____" Then
                     CASAROpenDate = Format(TxtDateOpen, "mm/dd/yyyy")
                Else
                    CASAROpenDate = Empty
                End If
                
              '  If Len(TxtARAnswer) Then
              '      !CASARAnswer = ChkStr(TxtARAnswer)
              '  Else
              '      !CASARAnswer = Null
              '  End If
                
              '  If Len(txtARRemarks) Then
              '      !CASARRemarks = txtARRemarks
              '  Else
              '      !CASARRemarks = Null
              '  End If
            'End With
            'CASOthers.Update
            tempdate = Format(Date, "yyyy/mm/dd")
            answer = ChkStr(Replace(TxtARAnswer, "'", "''"))
            Remarks = ChkStr(Replace(txtARRemarks, "'", "''"))
            HospVisit = ChkStr(Replace(txtHospVisit, "'", "''"))
            AuditReview = ChkStr(Replace(txtAuditReview, "'", "''"))
            
            strsqlCASAuditReview = "'" & ChkStr(Trim(TempCASNo)) & "', '" & ChkStr(txtINVNo) & "', " & _
                            "'" & InvDate & "', " & CASInvAmt & ", '" & CASLetterSent & "', '" & CASHospReply & "', " & _
                            "'" & CASDOV & "', '" & CASDOVTime & "', '" & ChkStr(txtOfficer) & "', '" & HospVisit & "', " & _
                            "" & CASDisRec & ", '" & CASDTC & "', '" & CASAROpenDate & "', " & _
                            "'" & answer & "', '" & Remarks & "', '" & AuditReview & "', '" & Mid(CboARStatus, 1, 1) & "', " & _
                            "'" & Logon & "', '" & tempdate & "', '" & Mid(cboDiscType, 1, 1) & "'"
            
            If rst2.EOF Then
                strsqlCASAuditReview = "EXEC Case_AuditInsert " + strsqlCASAuditReview
            Else
                strsqlCASAuditReview = "EXEC Case_AuditUpdate " + strsqlCASAuditReview
            End If
            rst2.Close
            Set rst2 = Nothing
            
            medixmasterconn.Execute strsqlCASAuditReview
            
            
        '---------------------------------------------------------------------------------------------------------------------------------------
        'Update the Cas Maintenance Fields
        
        'updatecaseMaint = "Select CASNumber, MBMNumber, CASARStatus from CASCrossReference Where CASNumber ='" & Trim(tempcasNo) & "' And MBMNumber ='" & Trim(MBMNumber) & " '"
        'CASMaint.Open updatecaseMaint, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
           ' With CASMaint
           '     If Len(CboARStatus) Then
           '         !CASARStatus = Mid(CboARStatus, 1, 1)
           '     Else
           '         !CASARStatus = Null
           '     End If
                
           ' End With
           ' CASMaint.Update
            strsqlCASRef = "EXEC Case_AuditUpdate '" & Trim(TempCASNo) & "', '" & Trim(MBMNumber) & "', '" & Mid(CboARStatus, 1, 1) & "'"
            medixmasterconnMaint.Execute strsqlCASRef
           ' CASMaint.Close
           ' Set CASMaint = Nothing
            
            MsgBox "Record Updated Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
        End If
    End If
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    cmdUpdate1.Enabled = True
End Sub

Private Sub txtSearch_GotFocus()
    Me.KeyPreview = False
End Sub

Private Sub Form_Load()
    FirstDiag = 0
    FinalDiag = 0
    CboARStatus.AddItem "C - Close"
    CboARStatus.AddItem "O - Open"
    cboDiscType.AddItem "A - Discount in main bordx"
    cboDiscType.AddItem "B - Discount in reverse bordx"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
    
End Sub

Private Sub txtSearch_LostFocus()
    Me.KeyPreview = True
End Sub

Private Sub ShowRepudiation()
Dim REP As New ADODB.Recordset
Dim CAS As New ADODB.Recordset
Dim strsqlREP As String
Dim strsqlCAS As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
     
    'strsqlREP = "select * from CASRepudiation where CASNumber = '" & txtCaseNo & "'"
    'REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "13"
    strAppend = " AND CASNumber = '" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
            If Len(!CASREPCode1) Then
                 cboREPCode1 = !CASREPCode1
            End If
            If Len(!CASREPCode2) Then
                 cboREPCode2 = !CASREPCode2
            End If
            If Len(!CASREPCode3) Then
                 cboREPCode3 = !CASREPCode3
            End If
        End With

        End If
    REP.Close
    Set REP = Nothing
End Sub

Public Sub ClearEntries(Optional KIV As Boolean)
   On Error Resume Next
   Dim ctl As Control

     For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
'            If KIV = True Then
 '               ctl.Text = ""
  '          Else
   '             If ctl.Enabled = True Then
                    ctl.Text = ""
    '            End If
     '       End If
        ElseIf TypeOf ctl Is ComboBox Then
            If ctl.Enabled = True Then
                ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is MaskEdBox Then
'            If ctl.Enabled = True Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
 '           End If
        End If
    Next ctl
        txtFCTime.mask = ""
        txtFCTime.Text = ""
        txtFCTime.mask = "##:##"
        txtFTTime.mask = ""
        txtFTTime.Text = ""
        txtFTTime.mask = "##:##"
        txtDOATime.mask = ""
        txtDOATime.Text = ""
        txtDOATime.mask = "##:##"
        txtDODTime.mask = ""
        txtDODTime.Text = ""
        txtDODTime.mask = "##:##"
        txtDOVTime.mask = ""
        txtDOVTime.Text = ""
        txtDOVTime.mask = "##:##"
        cboStatus1 = "O - Open"
        cboClaimability = "A - Accepted"
        lstMember.ListItems.Clear
        txtCASLGAmt = 2500
        txtCASReserveAmount = 1600
        cboNature = "ILL - ILLNess"
        cboModality = "S - Surgical"
End Sub

Private Sub showMBMListing()
Dim MBM As New ADODB.Recordset
Dim sqlstr As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
                
    'sqlstr = " Select MBMNumber, HLTCode, MBMPolicyNo,  MBMName ," & _
            " MBMIcBcPp,  MBMMemberType , " & _
            " MBMStatus, PAYCode, AGECode , " & _
            " INSCode, PLNCode, " & _
            " MBMAddress1, MBMAddress2, " & _
            " MBMAddress3, MBMPostCode, CTYCode, STACode,  " & _
            " MBMDOB , BRNCode, MBMTakeOver ,  MBMAvailableLimit, " & _
            " MBMPayorEffDate, MBMPayorExpDate, MBMENDtEffDate  "
    'sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "14"
    strAppend = " AND MBMNumber = '" & Trim(txtMBMNumber) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
    
    lstMember.ListItems.Clear
    If Not MBM.EOF Then
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
            End With
    End If
    MBM.Close
    Set MBM = Nothing
End Sub
' * End

' * Start

Private Sub showMBMCovListing()
Dim MBMCov As New ADODB.Recordset
Dim sqlstr As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
                
    'sqlstr = "select MBMCNumber, MBMCCoverID, MBMCName," & _
        " MBMCICBCPPNo , MBMCDOB, MBMCSex, MBMCAdultChild, MBMCAgeband, " & _
        " MBMCRELCode,  " & _
        " MBMCStatus , MBMCAnnualLimit " & _
        " from  MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    'MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "15"
    strAppend = " AND MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
    
    If Not MBMCov.EOF Then
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

Private Sub ShowUD()
Dim CASUD As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppend As String
Dim strAppendCase As String
            
        strAppendCase = "18"
        strAppend = " AND CASNumber = '" & txtCaseNo & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "Case_AuditReview"
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
    
        If Not CASUD.EOF Then
           ' Do Until CASUD.EOF
            With CASUD
                If Len(!CASUDCode1) Then
                    cboUDICDCode1 = !CASUDCode1
                End If
                If Len(!CASUDDOS1) Then
                    txtUDDOS1.mask = ""
                    txtUDDOS1 = Format(!CASUDDOS1, "dd/mm/yyyy")
                    txtUDDOS1.mask = "##/##/####"
                End If
                If Len(!CASUDReasons1) Then
                    txtUD1 = !CASUDReasons1
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
    strAppendCase = "16"
    strAppend = " AND CASNumber = '" & txtCaseNo & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
            'Do Until diafinal.EOF
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
                        txtFinalDiag1 = !DIAFinalDiag1
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
                        txtFinalDiag2 = !DIAFinalDiag2
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
                        txtFinalDiag3 = !DIAFinalDiag3
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
Dim strAppend As String
    
    'strsql = "Select * from DIAFirst where CasNumber = '" & Trim(txtCaseNo) & "'"
    'diafirst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "17"
    strAppend = " AND CasNumber = '" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReview"
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
    
        If Not diafirst.EOF Then
            Do Until diafirst.EOF
            With diafirst
                If Len(!ICDCode1) Then
                    cboICDCode1 = !ICDCode1
                End If
                If Len(!DIADOS1) Then
                    txtPDDOS1.mask = ""
                    txtPDDOS1 = Format(!DIADOS1, "dd/mm/yyyy")
                    txtPDDOS1.mask = "##/##/####"
                End If
                If Len(!DIAPrelimDiag1) Then
                    txtPrelimDiag1 = !DIAPrelimDiag1
                End If
                If Len(!ICDCode2) Then
                    cboICDCode2 = !ICDCode2
                End If
                If Len(!DIADOS2) Then
                    txtPDDOS2.mask = ""
                    txtPDDOS2 = Format(!DIADOS2, "dd/mm/yyyy")
                    txtPDDOS2.mask = "##/##/####"
                End If
                If Len(!DIAPrelimDiag2) Then
                    txtPrelimDiag2 = !DIAPrelimDiag2
                End If
                If Len(!ICDCode3) Then
                    cboICDCode3 = !ICDCode3
                End If
                If Len(!DIADOS3) Then
                    txtPDDOS3.mask = ""
                    txtPDDOS3 = Format(!DIADOS3, "dd/mm/yyyy")
                    txtPDDOS3.mask = "##/##/####"
                End If
                If Len(!DIAPrelimDiag3) Then
                    txtPrelimDiag3 = !DIAPrelimDiag3
                End If
            End With
            diafirst.MoveNext
            Loop
            End If
    diafirst.Close
    Set diafirst = Nothing
End Sub

Private Sub txtDISRec_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtDLR_LostFocus()
    If IsDate(txtDLR) Then
    Else
    If txtDLR <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtDLR.SetFocus
    End If
    End If

End Sub

Private Sub txtDLS_LostFocus()
    
    If IsDate(txtDLS) Then
    Else
    If txtDLS <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtDLS.SetFocus
    End If
    End If

End Sub

Private Sub txtDOVTime_LostFocus()
    If txtDOVTime <> "__:__" Then
        If isTime(txtDOVTime) = False Then
            MsgBox "Invalid Time Format! ", vbCritical
            txtDOVTime.SetFocus
        End If
    End If

End Sub

Private Sub txtDV_LostFocus()
    If IsDate(txtDV) Then
    Else
        If txtDV <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDV.SetFocus
        End If
    End If

End Sub

Private Sub TxtDRTC_LostFocus()
    If IsDate(TxtDRTC) Then
    Else
        If TxtDRTC <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDRTC.SetFocus
        End If
    End If

End Sub

Private Sub txtINVAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtInvDate_LostFocus()

    If IsDate(txtINVDate) Then
    Else
        If txtINVDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtINVDate.SetFocus
        End If
    End If

End Sub

Private Function showPaymentRecord()
Dim ForLoop, ListCount As Integer
Dim INVAmount As Double
Dim CheckPayment As New ADODB.Recordset
Dim strsql As String
Dim Listing As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
    
    TxtPVNo = ""
    TxtPVDate = ""
    INVAmount = 0
    ListView1.ListItems.Clear
    
    'strsql = "Select CASNumber, PVHOSNo, PVHOSDate, HOSPaidAmt From Payment2HOS where CASNumber = '" & Trim(txtCaseNo) & "'"
    'CheckPayment.Open strsql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    strAppend = " AND CASNumber = '" & Trim(txtCaseNo) & "'"
    Set cmd.ActiveConnection = medixmasterconnPayment
    cmd.CommandText = "Case_AuditReview"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    CheckPayment.CursorLocation = adUseClient
    CheckPayment.CursorType = adOpenForwardOnly
    CheckPayment.CacheSize = 100
    Set CheckPayment = cmd.Execute
    
    If CheckPayment.EOF = False Then
    
    'Else
       Do Until CheckPayment.EOF
            Set Listing = ListView1.ListItems.Add(, , "00")
            With CheckPayment
                'Set Listing = ListView1.listitems.Add(, !CasNumber)
                If Len(Trim(!CasNumber)) Then
                    Listing.SubItems(1) = !CasNumber
                End If
                If Len(Trim(!PVHOSNo)) Then
                    Listing.SubItems(2) = !PVHOSNo
                End If
                If Len(Trim(!PVHOSDate)) Then
                    Listing.SubItems(3) = !PVHOSDate
                End If
                If Len(Trim(!HOSPaidAmt)) Then
                    Listing.SubItems(4) = !HOSPaidAmt
                End If
        
                .MoveNext
            End With
        Loop
    End If
    CheckPayment.Close
    Set CheckPayment = Nothing
    
    ListCount = ListView1.ListItems.Count
    
    For ForLoop = 1 To ListCount
        With ListView1.ListItems(ForLoop)
                  
            If Trim(.SubItems(2)) <> "" Then
                If Trim(TxtPVNo) <> "" Then
                    TxtPVNo = TxtPVNo & ", " & .SubItems(2)
                Else
                    TxtPVNo = .SubItems(2)
                End If
            End If
                                
            If Trim(.SubItems(3)) <> "" Then
                If Trim(TxtPVDate) <> "" Then
                    TxtPVDate = TxtPVDate & "," & .SubItems(3)
                Else
                    TxtPVDate = .SubItems(3)
                End If
            End If
            
            INVAmount = CDbl(INVAmount) + CDbl(IIf(IsNumeric(.SubItems(4)), .SubItems(4), 0))
            TxtAmtPaid = INVAmount
            
        End With
    Next
    
End Function
