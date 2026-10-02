VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimAdjustment 
   Caption         =   "Worksheet Adjustment"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   10065
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      TabIndex        =   231
      Top             =   0
      Width           =   9975
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   232
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Claim Adjustment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   180
         TabIndex        =   233
         Top             =   60
         Width           =   6615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   13361
      _Version        =   393216
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmClaimAdjustment.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GeneralFrame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Others"
      TabPicture(1)   =   "frmClaimAdjustment.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame16"
      Tab(1).Control(1)=   "CmdInvoice"
      Tab(1).Control(2)=   "cmdShowHistory"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "WorkSheet"
      TabPicture(2)   =   "frmClaimAdjustment.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).ControlCount=   2
      Begin VB.Frame GeneralFrame 
         Height          =   6915
         Left            =   240
         TabIndex        =   193
         Top             =   480
         Width           =   8295
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
            Left            =   1320
            TabIndex        =   212
            Top             =   240
            Width           =   3105
         End
         Begin VB.TextBox txtMBMPolNo 
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
            Left            =   6000
            TabIndex        =   211
            Top             =   960
            Width           =   1860
         End
         Begin VB.TextBox txtINSCode 
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
            Left            =   6000
            TabIndex        =   210
            Top             =   600
            Width           =   1305
         End
         Begin VB.TextBox txtPatientName 
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
            TabIndex        =   209
            Top             =   600
            Width           =   3105
         End
         Begin VB.TextBox txtPLNCode 
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
            Left            =   6000
            TabIndex        =   208
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtHLTCode 
            Height          =   285
            Left            =   120
            TabIndex        =   206
            Top             =   3720
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtRelCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   205
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox txtDisChargeDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6000
            TabIndex        =   204
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtHOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   203
            Top             =   1320
            Width           =   3135
         End
         Begin VB.TextBox txtDateOfLoss 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   202
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtAdmissionDate 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   201
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtCASRemark 
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
            Height          =   645
            Left            =   1320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   200
            Top             =   3480
            Width           =   6180
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   199
            Top             =   2400
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   198
            Top             =   960
            Width           =   3105
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6000
            TabIndex        =   197
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   196
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6000
            TabIndex        =   195
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6000
            TabIndex        =   194
            Top             =   1680
            Width           =   615
         End
         Begin MSMask.MaskEdBox txtPolEffDate 
            Height          =   285
            Left            =   6000
            TabIndex        =   207
            Top             =   1320
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Caption         =   "Member Name"
            BeginProperty Font 
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
            TabIndex        =   230
            Top             =   285
            Width           =   1275
         End
         Begin VB.Label Label27 
            Caption         =   "Policy No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   229
            Top             =   960
            Width           =   1275
         End
         Begin VB.Label Label9 
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
            Left            =   4560
            TabIndex        =   228
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label8 
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
            Left            =   120
            TabIndex        =   227
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label5 
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
            Left            =   4560
            TabIndex        =   226
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label47 
            Caption         =   "Policy Eff. Date"
            Height          =   375
            Left            =   4560
            TabIndex        =   225
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label48 
            Caption         =   "Relationship"
            Height          =   375
            Left            =   120
            TabIndex        =   224
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label50 
            Caption         =   "Discharge Date"
            Height          =   255
            Left            =   4560
            TabIndex        =   223
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label51 
            Caption         =   "Hospital"
            Height          =   255
            Left            =   120
            TabIndex        =   222
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label52 
            Caption         =   "Date Of Loss"
            Height          =   375
            Left            =   120
            TabIndex        =   221
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label49 
            Caption         =   "Admission Date"
            Height          =   255
            Left            =   120
            TabIndex        =   220
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label60 
            Caption         =   "Remark"
            Height          =   375
            Left            =   120
            TabIndex        =   219
            Top             =   3480
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Nature Of Claim"
            Height          =   255
            Left            =   120
            TabIndex        =   218
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Patient ICBc"
            Height          =   255
            Left            =   120
            TabIndex        =   217
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Treatment Modality "
            Height          =   255
            Left            =   4560
            TabIndex        =   216
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label25 
            Caption         =   "Payment Mode"
            Height          =   255
            Left            =   120
            TabIndex        =   215
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label44 
            Caption         =   "Payment Type"
            Height          =   255
            Left            =   4560
            TabIndex        =   214
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label53 
            Caption         =   "In-Patient Days"
            Height          =   375
            Left            =   4560
            TabIndex        =   213
            Top             =   1680
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pre/Post Treatement"
         Height          =   2490
         Left            =   -74520
         TabIndex        =   5
         Top             =   600
         Width           =   7455
         Begin MSComctlLib.ListView lstHosList 
            Height          =   1905
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   3360
            View            =   3
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
               Text            =   "Hospital Code"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Hospital Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Admission Date"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Discharge Date"
               Object.Width           =   1587
            EndProperty
         End
      End
      Begin VB.CommandButton cmdShowHistory 
         Caption         =   "Claim History"
         Height          =   375
         Left            =   -74400
         TabIndex        =   4
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton CmdInvoice 
         Caption         =   "Invoices"
         Height          =   375
         Left            =   -72960
         TabIndex        =   3
         Top             =   5520
         Width           =   975
      End
      Begin VB.Frame Frame16 
         Caption         =   "Final Diagnosis"
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
         Left            =   -74520
         TabIndex        =   1
         Top             =   3360
         Width           =   7395
         Begin MSComctlLib.ListView lstFinalDiag 
            Height          =   1695
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   6945
            _ExtentX        =   12250
            _ExtentY        =   2990
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Diag. Code"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Diag. By"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Diag. Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Status"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame9 
         Height          =   615
         Left            =   -74880
         TabIndex        =   183
         Top             =   360
         Width           =   9495
         Begin VB.CommandButton cmdApprove 
            Caption         =   "Approve"
            Height          =   255
            Left            =   960
            TabIndex        =   187
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdCalc 
            Caption         =   "Calculate"
            Height          =   255
            Left            =   120
            TabIndex        =   186
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            BeginProperty Font 
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
            TabIndex        =   185
            Top             =   240
            Width           =   855
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
            Height          =   255
            Left            =   2640
            TabIndex        =   184
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label43 
            Caption         =   "Due To/(From) Member"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8280
            TabIndex        =   188
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label Label39 
            Caption         =   "Limit"
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
            Left            =   4440
            TabIndex        =   192
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label40 
            Caption         =   "Actual"
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
            Left            =   5400
            TabIndex        =   191
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label41 
            Caption         =   "Approved"
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
            Left            =   6240
            TabIndex        =   190
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label42 
            Caption         =   "Payable to Hospital"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7320
            TabIndex        =   189
            Top             =   120
            Width           =   810
         End
      End
      Begin VB.Frame Frame4 
         Height          =   7095
         Left            =   -74760
         TabIndex        =   7
         Top             =   960
         Width           =   9615
         Begin VB.Frame Frame11 
            Height          =   7215
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   9045
            Begin VB.TextBox ICU1 
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
               Height          =   240
               Left            =   2385
               TabIndex        =   145
               Text            =   "0"
               Top             =   405
               Width           =   435
            End
            Begin VB.TextBox LG1 
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
               Height          =   240
               Left            =   2400
               TabIndex        =   144
               Text            =   "0"
               Top             =   4320
               Width           =   435
            End
            Begin VB.TextBox DB1 
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
               Height          =   240
               Left            =   2400
               TabIndex        =   143
               Text            =   "0"
               Top             =   180
               Width           =   435
            End
            Begin VB.TextBox M20 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   142
               Text            =   "0.00"
               Top             =   5040
               Width           =   915
            End
            Begin VB.TextBox H20 
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
               Left            =   7110
               TabIndex        =   141
               Text            =   "0.00"
               Top             =   5040
               Width           =   915
            End
            Begin VB.TextBox P20 
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
               Left            =   4185
               TabIndex        =   140
               Text            =   "0.00"
               Top             =   5040
               Width           =   915
            End
            Begin VB.TextBox A20 
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
               Left            =   5160
               TabIndex        =   139
               Text            =   "0.00"
               Top             =   5040
               Width           =   915
            End
            Begin VB.TextBox TotalGrossHos 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   138
               Text            =   "0.00"
               Top             =   6720
               Width           =   915
            End
            Begin VB.TextBox TotalGrossMember 
               Alignment       =   1  'Right Justify
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
               Left            =   8085
               TabIndex        =   137
               Text            =   "0.00"
               Top             =   6720
               Width           =   915
            End
            Begin VB.TextBox TotalDiscountHos 
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
               Left            =   7110
               TabIndex        =   136
               Text            =   "0.00"
               Top             =   6450
               Width           =   915
            End
            Begin VB.TextBox TotalDiscountMember 
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
               Left            =   8085
               TabIndex        =   135
               Text            =   "0.00"
               Top             =   6450
               Width           =   915
            End
            Begin VB.TextBox TotalPaidMedixHos 
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
               Left            =   7110
               TabIndex        =   134
               Text            =   "0.00"
               Top             =   6240
               Width           =   915
            End
            Begin VB.TextBox TotalPaidMedixMember 
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
               Left            =   8085
               TabIndex        =   133
               Text            =   "0.00"
               Top             =   6225
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.TextBox TotalPaidByMemberHos 
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
               Left            =   7110
               TabIndex        =   132
               Text            =   "0.00"
               Top             =   6000
               Width           =   915
            End
            Begin VB.TextBox TotalPaidByMemberMember 
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
               Left            =   8085
               TabIndex        =   131
               Text            =   "0.00"
               Top             =   6000
               Width           =   915
            End
            Begin VB.TextBox PAnnualLimit 
               Alignment       =   1  'Right Justify
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
               Left            =   4185
               TabIndex        =   130
               Text            =   "0.00"
               Top             =   5760
               Width           =   915
            End
            Begin VB.TextBox TotalHospital 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   129
               Text            =   "0.00"
               Top             =   5715
               Width           =   915
            End
            Begin VB.TextBox TotalMember 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00;(#,##0.00)"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   128
               Text            =   "0.00"
               Top             =   5715
               Width           =   915
            End
            Begin VB.TextBox A19 
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
               Left            =   5160
               TabIndex        =   127
               Text            =   "0.00"
               Top             =   4770
               Width           =   915
            End
            Begin VB.TextBox P19 
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
               Left            =   4185
               TabIndex        =   126
               Text            =   "0.00"
               Top             =   4770
               Width           =   915
            End
            Begin VB.TextBox H19 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   125
               Text            =   "0.00"
               Top             =   4770
               Width           =   915
            End
            Begin VB.TextBox M19 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   124
               Text            =   "0.00"
               Top             =   4770
               Width           =   915
            End
            Begin VB.TextBox A18 
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
               Left            =   5160
               TabIndex        =   123
               Text            =   "0.00"
               Top             =   4545
               Width           =   915
            End
            Begin VB.TextBox P18 
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
               Left            =   4185
               TabIndex        =   122
               Text            =   "0.00"
               Top             =   4545
               Width           =   915
            End
            Begin VB.TextBox H18 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   121
               Text            =   "0.00"
               Top             =   4545
               Width           =   915
            End
            Begin VB.TextBox M18 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   120
               Text            =   "0.00"
               Top             =   4545
               Width           =   915
            End
            Begin VB.TextBox A17 
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
               Left            =   5160
               TabIndex        =   119
               Text            =   "0.00"
               Top             =   4320
               Width           =   915
            End
            Begin VB.TextBox P17 
               Alignment       =   1  'Right Justify
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
               Left            =   4200
               TabIndex        =   118
               Text            =   "0.00"
               Top             =   4320
               Width           =   915
            End
            Begin VB.TextBox A16 
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
               Left            =   5160
               TabIndex        =   117
               Text            =   "0.00"
               Top             =   4095
               Width           =   915
            End
            Begin VB.TextBox P16 
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
               Left            =   4185
               TabIndex        =   116
               Text            =   "0.00"
               Top             =   4095
               Width           =   915
            End
            Begin VB.TextBox A15 
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
               Left            =   5160
               TabIndex        =   115
               Text            =   "0.00"
               Top             =   3870
               Width           =   915
            End
            Begin VB.TextBox P15 
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
               Left            =   4185
               TabIndex        =   114
               Text            =   "0.00"
               Top             =   3870
               Width           =   915
            End
            Begin VB.TextBox H17 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   113
               Text            =   "0.00"
               Top             =   4320
               Width           =   915
            End
            Begin VB.TextBox H16 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   112
               Text            =   "0.00"
               Top             =   4095
               Width           =   915
            End
            Begin VB.TextBox H15 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   111
               Text            =   "0.00"
               Top             =   3870
               Width           =   915
            End
            Begin VB.TextBox M17 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   110
               Text            =   "0.00"
               Top             =   4320
               Width           =   915
            End
            Begin VB.TextBox M16 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   109
               Text            =   "0.00"
               Top             =   4095
               Width           =   915
            End
            Begin VB.TextBox M15 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   108
               Text            =   "0.00"
               Top             =   3870
               Width           =   915
            End
            Begin VB.TextBox A14 
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
               Left            =   5160
               TabIndex        =   107
               Text            =   "0.00"
               Top             =   3480
               Width           =   915
            End
            Begin VB.TextBox P14 
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
               Left            =   4185
               TabIndex        =   106
               Text            =   "0.00"
               Top             =   3480
               Width           =   915
            End
            Begin VB.TextBox H14 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   105
               Text            =   "0.00"
               Top             =   3480
               Width           =   915
            End
            Begin VB.TextBox M14 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   104
               Text            =   "0.00"
               Top             =   3510
               Width           =   915
            End
            Begin VB.TextBox A13 
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
               Left            =   5160
               TabIndex        =   103
               Text            =   "0.00"
               Top             =   3285
               Width           =   915
            End
            Begin VB.TextBox P13 
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
               Left            =   4200
               TabIndex        =   102
               Text            =   "0.00"
               Top             =   3285
               Width           =   915
            End
            Begin VB.TextBox A12 
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
               Left            =   5160
               TabIndex        =   101
               Text            =   "0.00"
               Top             =   3060
               Width           =   915
            End
            Begin VB.TextBox P12 
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
               Left            =   4185
               TabIndex        =   100
               Text            =   "0.00"
               Top             =   3060
               Width           =   915
            End
            Begin VB.TextBox A11 
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
               Left            =   5160
               TabIndex        =   99
               Text            =   "0.00"
               Top             =   2835
               Width           =   915
            End
            Begin VB.TextBox P11 
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
               Left            =   4185
               TabIndex        =   98
               Text            =   "0.00"
               Top             =   2835
               Width           =   915
            End
            Begin VB.TextBox H13 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   97
               Text            =   "0.00"
               Top             =   3285
               Width           =   915
            End
            Begin VB.TextBox H12 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   96
               Text            =   "0.00"
               Top             =   3060
               Width           =   915
            End
            Begin VB.TextBox H11 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   95
               Text            =   "0.00"
               Top             =   2835
               Width           =   915
            End
            Begin VB.TextBox M13 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   94
               Text            =   "0.00"
               Top             =   3285
               Width           =   915
            End
            Begin VB.TextBox M12 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   93
               Text            =   "0.00"
               Top             =   3060
               Width           =   915
            End
            Begin VB.TextBox M11 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   92
               Text            =   "0.00"
               Top             =   2835
               Width           =   915
            End
            Begin VB.TextBox A10 
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
               Left            =   5160
               TabIndex        =   91
               Text            =   "0.00"
               Top             =   2475
               Width           =   915
            End
            Begin VB.TextBox P10 
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
               Left            =   4200
               TabIndex        =   90
               Text            =   "0.00"
               Top             =   2475
               Width           =   915
            End
            Begin VB.TextBox A9 
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
               Left            =   5160
               TabIndex        =   89
               Text            =   "0.00"
               Top             =   2250
               Width           =   915
            End
            Begin VB.TextBox P9 
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
               Left            =   4185
               TabIndex        =   88
               Text            =   "0.00"
               Top             =   2250
               Width           =   915
            End
            Begin VB.TextBox A8 
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
               Left            =   5160
               TabIndex        =   87
               Text            =   "0.00"
               Top             =   2025
               Width           =   915
            End
            Begin VB.TextBox P8 
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
               Left            =   4185
               TabIndex        =   86
               Text            =   "0.00"
               Top             =   2025
               Width           =   915
            End
            Begin VB.TextBox H10 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   85
               Text            =   "0.00"
               Top             =   2475
               Width           =   915
            End
            Begin VB.TextBox H9 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   84
               Text            =   "0.00"
               Top             =   2250
               Width           =   915
            End
            Begin VB.TextBox H8 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   83
               Text            =   "0.00"
               Top             =   2040
               Width           =   915
            End
            Begin VB.TextBox M10 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   82
               Text            =   "0.00"
               Top             =   2475
               Width           =   915
            End
            Begin VB.TextBox M9 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   81
               Text            =   "0.00"
               Top             =   2250
               Width           =   915
            End
            Begin VB.TextBox M8 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   80
               Text            =   "0.00"
               Top             =   2025
               Width           =   915
            End
            Begin VB.TextBox M5 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   79
               Text            =   "0.00"
               Top             =   1215
               Width           =   915
            End
            Begin VB.TextBox M6 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   78
               Text            =   "0.00"
               Top             =   1440
               Width           =   915
            End
            Begin VB.TextBox M7 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   77
               Text            =   "0.00"
               Top             =   1665
               Width           =   915
            End
            Begin VB.TextBox H5 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   76
               Text            =   "0.00"
               Top             =   1200
               Width           =   915
            End
            Begin VB.TextBox H6 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   75
               Text            =   "0.00"
               Top             =   1440
               Width           =   915
            End
            Begin VB.TextBox H7 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   74
               Text            =   "0.00"
               Top             =   1665
               Width           =   915
            End
            Begin VB.TextBox P5 
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
               Left            =   4185
               TabIndex        =   73
               Text            =   "0.00"
               Top             =   1215
               Width           =   915
            End
            Begin VB.TextBox A5 
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
               Left            =   5160
               TabIndex        =   72
               Text            =   "0.00"
               Top             =   1215
               Width           =   915
            End
            Begin VB.TextBox P6 
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
               Left            =   4185
               TabIndex        =   71
               Text            =   "0.00"
               Top             =   1440
               Width           =   915
            End
            Begin VB.TextBox A6 
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
               Left            =   5160
               TabIndex        =   70
               Text            =   "0.00"
               Top             =   1440
               Width           =   915
            End
            Begin VB.TextBox P7 
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
               Left            =   4185
               TabIndex        =   69
               Text            =   "0.00"
               Top             =   1665
               Width           =   915
            End
            Begin VB.TextBox A7 
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
               Left            =   5160
               TabIndex        =   68
               Text            =   "0.00"
               Top             =   1665
               Width           =   915
            End
            Begin VB.TextBox M1 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   67
               Text            =   "0.00"
               Top             =   180
               Width           =   915
            End
            Begin VB.TextBox M2 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   66
               Text            =   "0.00"
               Top             =   405
               Width           =   915
            End
            Begin VB.TextBox M3 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   65
               Text            =   "0.00"
               Top             =   630
               Width           =   915
            End
            Begin VB.TextBox M4 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8085
               TabIndex        =   64
               Text            =   "0.00"
               Top             =   855
               Width           =   915
            End
            Begin VB.TextBox H1 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   63
               Text            =   "0.00"
               Top             =   180
               Width           =   915
            End
            Begin VB.TextBox H2 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   62
               Text            =   "0.00"
               Top             =   405
               Width           =   915
            End
            Begin VB.TextBox H3 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   7110
               TabIndex        =   61
               Text            =   "0.00"
               Top             =   630
               Width           =   915
            End
            Begin VB.TextBox H4 
               Alignment       =   1  'Right Justify
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
               Left            =   7110
               TabIndex        =   60
               Text            =   "0.00"
               Top             =   855
               Width           =   915
            End
            Begin VB.TextBox A4 
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
               Left            =   5160
               TabIndex        =   59
               Text            =   "0.00"
               Top             =   855
               Width           =   915
            End
            Begin VB.TextBox P4 
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
               Left            =   4200
               TabIndex        =   58
               Text            =   "0.00"
               Top             =   840
               Width           =   915
            End
            Begin VB.TextBox A3 
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
               Left            =   5160
               TabIndex        =   57
               Text            =   "0.00"
               Top             =   630
               Width           =   915
            End
            Begin VB.TextBox P3 
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
               Left            =   4185
               TabIndex        =   56
               Text            =   "0.00"
               Top             =   600
               Width           =   915
            End
            Begin VB.TextBox A2 
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
               Left            =   5160
               TabIndex        =   55
               Text            =   "0.00"
               Top             =   405
               Width           =   915
            End
            Begin VB.TextBox P2 
               Alignment       =   1  'Right Justify
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
               Left            =   4185
               TabIndex        =   54
               Text            =   "0.00"
               Top             =   405
               Width           =   915
            End
            Begin VB.TextBox A1 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   5160
               TabIndex        =   53
               Text            =   "0.00"
               Top             =   180
               Width           =   915
            End
            Begin VB.TextBox P1 
               Alignment       =   1  'Right Justify
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
               Height          =   240
               Left            =   4185
               TabIndex        =   52
               Text            =   "0.00"
               Top             =   180
               Width           =   915
            End
            Begin VB.TextBox TotalDiscount 
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
               Left            =   5160
               TabIndex        =   51
               Text            =   "0.00"
               Top             =   6450
               Width           =   915
            End
            Begin VB.TextBox TotalGross 
               Alignment       =   1  'Right Justify
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
               Left            =   5160
               TabIndex        =   50
               Text            =   "0.00"
               Top             =   6720
               Width           =   915
            End
            Begin VB.TextBox AP20 
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
               Left            =   6135
               TabIndex        =   49
               Text            =   "0.00"
               Top             =   5040
               Width           =   915
            End
            Begin VB.TextBox TotalApproved 
               Alignment       =   1  'Right Justify
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
               Left            =   6135
               TabIndex        =   48
               Text            =   "0.00"
               Top             =   5760
               Width           =   915
            End
            Begin VB.TextBox TotalActual 
               Alignment       =   1  'Right Justify
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
               Left            =   5160
               TabIndex        =   47
               Text            =   "0.00"
               Top             =   5760
               Width           =   915
            End
            Begin VB.TextBox AP19 
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
               Left            =   6135
               TabIndex        =   46
               Text            =   "0.00"
               Top             =   4770
               Width           =   915
            End
            Begin VB.TextBox AP18 
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
               Left            =   6135
               TabIndex        =   45
               Text            =   "0.00"
               Top             =   4545
               Width           =   915
            End
            Begin VB.TextBox AP17 
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
               Left            =   6135
               TabIndex        =   44
               Text            =   "0.00"
               Top             =   4320
               Width           =   915
            End
            Begin VB.TextBox AP16 
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
               Left            =   6135
               TabIndex        =   43
               Text            =   "0.00"
               Top             =   4095
               Width           =   915
            End
            Begin VB.TextBox AP15 
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
               Left            =   6135
               TabIndex        =   42
               Text            =   "0.00"
               Top             =   3870
               Width           =   915
            End
            Begin VB.TextBox AP14 
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
               Left            =   6135
               TabIndex        =   41
               Text            =   "0.00"
               Top             =   3510
               Width           =   915
            End
            Begin VB.TextBox AP13 
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
               Left            =   6135
               TabIndex        =   40
               Text            =   "0.00"
               Top             =   3285
               Width           =   915
            End
            Begin VB.TextBox AP12 
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
               Left            =   6135
               TabIndex        =   39
               Text            =   "0.00"
               Top             =   3060
               Width           =   915
            End
            Begin VB.TextBox AP11 
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
               Left            =   6135
               TabIndex        =   38
               Text            =   "0.00"
               Top             =   2835
               Width           =   915
            End
            Begin VB.TextBox AP10 
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
               Left            =   6135
               TabIndex        =   37
               Text            =   "0.00"
               Top             =   2475
               Width           =   915
            End
            Begin VB.TextBox AP9 
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
               Left            =   6135
               TabIndex        =   36
               Text            =   "0.00"
               Top             =   2250
               Width           =   915
            End
            Begin VB.TextBox AP8 
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
               Left            =   6135
               TabIndex        =   35
               Text            =   "0.00"
               Top             =   2025
               Width           =   915
            End
            Begin VB.TextBox AP5 
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
               Left            =   6135
               TabIndex        =   34
               Text            =   "0.00"
               Top             =   1215
               Width           =   915
            End
            Begin VB.TextBox AP6 
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
               Left            =   6135
               TabIndex        =   33
               Text            =   "0.00"
               Top             =   1440
               Width           =   915
            End
            Begin VB.TextBox AP7 
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
               Left            =   6120
               TabIndex        =   32
               Text            =   "0.00"
               Top             =   1665
               Width           =   915
            End
            Begin VB.TextBox AP4 
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
               Left            =   6120
               TabIndex        =   31
               Text            =   "0.00"
               Top             =   855
               Width           =   915
            End
            Begin VB.TextBox AP3 
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
               Left            =   6135
               TabIndex        =   30
               Text            =   "0.00"
               Top             =   630
               Width           =   915
            End
            Begin VB.TextBox AP2 
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
               Left            =   6120
               TabIndex        =   29
               Text            =   "0.00"
               Top             =   405
               Width           =   915
            End
            Begin VB.TextBox AP1 
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
               Left            =   6120
               TabIndex        =   28
               Text            =   "0.00"
               Top             =   180
               Width           =   915
            End
            Begin VB.TextBox AP21 
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
               Left            =   6150
               TabIndex        =   27
               Text            =   "0.00"
               Top             =   5280
               Width           =   915
            End
            Begin VB.TextBox A21 
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
               Left            =   5175
               TabIndex        =   26
               Text            =   "0.00"
               Top             =   5280
               Width           =   915
            End
            Begin VB.TextBox P21 
               Alignment       =   1  'Right Justify
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
               Left            =   4200
               TabIndex        =   25
               Text            =   "0.00"
               Top             =   5280
               Width           =   915
            End
            Begin VB.TextBox H21 
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
               Left            =   7110
               TabIndex        =   24
               Text            =   "0.00"
               Top             =   5280
               Width           =   915
            End
            Begin VB.TextBox M21 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
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
               Left            =   8100
               TabIndex        =   23
               Text            =   "0.00"
               Top             =   5280
               Width           =   915
            End
            Begin VB.TextBox txtRMCash 
               Height          =   285
               Left            =   3720
               TabIndex        =   22
               Text            =   "0"
               Top             =   5280
               Width           =   435
            End
            Begin VB.TextBox txtDayCash 
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
               Left            =   2400
               TabIndex        =   21
               Text            =   "0"
               Top             =   5280
               Width           =   435
            End
            Begin VB.TextBox txtLGAmount 
               Alignment       =   1  'Right Justify
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
               TabIndex        =   20
               Text            =   "0.00"
               Top             =   6000
               Width           =   915
            End
            Begin VB.TextBox txtAnnualLimit 
               Alignment       =   1  'Right Justify
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
               TabIndex        =   19
               Text            =   "0.00"
               Top             =   6000
               Width           =   915
            End
            Begin VB.TextBox LodgerAmount 
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
               Height          =   240
               Left            =   3720
               TabIndex        =   18
               Text            =   "0"
               Top             =   4320
               Width           =   435
            End
            Begin VB.TextBox txtOthers 
               Height          =   285
               Left            =   2040
               TabIndex        =   17
               Top             =   5000
               Width           =   2055
            End
            Begin VB.TextBox ICUAmount 
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
               Left            =   3720
               TabIndex        =   16
               Text            =   "0"
               Top             =   405
               Width           =   435
            End
            Begin VB.TextBox RoomAmount 
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
               Left            =   3720
               TabIndex        =   15
               Text            =   "0"
               Top             =   180
               Width           =   435
            End
            Begin VB.TextBox P1A 
               Height          =   285
               Left            =   3360
               TabIndex        =   14
               Top             =   840
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox P2A 
               Height          =   285
               Left            =   2640
               TabIndex        =   13
               Top             =   840
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.TextBox P17A 
               Height          =   285
               Left            =   2760
               TabIndex        =   12
               Top             =   3960
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtRMCashA 
               Height          =   285
               Left            =   2400
               TabIndex        =   11
               Top             =   4680
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox TotalGrossApproved 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   6120
               TabIndex        =   10
               Text            =   "0.00"
               Top             =   6720
               Width           =   915
            End
            Begin VB.Line Line1 
               BorderWidth     =   2
               X1              =   90
               X2              =   9450
               Y1              =   1170
               Y2              =   1170
            End
            Begin VB.Label Label61 
               Caption         =   "Hospital"
               Height          =   375
               Left            =   120
               TabIndex        =   182
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label62 
               Caption         =   "Surgical"
               Height          =   495
               Left            =   120
               TabIndex        =   181
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Label63 
               Caption         =   "Non- Surgical"
               Height          =   495
               Left            =   120
               TabIndex        =   180
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Label64 
               Caption         =   "Out- Patient Benefits"
               Height          =   735
               Left            =   120
               TabIndex        =   179
               Top             =   2880
               Width           =   615
            End
            Begin VB.Label Label65 
               Caption         =   "Others"
               Height          =   615
               Left            =   120
               TabIndex        =   178
               Top             =   4200
               Width           =   495
            End
            Begin VB.Label Label45 
               Caption         =   "Days @ RM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2880
               TabIndex        =   177
               Top             =   4320
               Width           =   825
            End
            Begin VB.Label Label33 
               Caption         =   "Days @ RM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2880
               TabIndex        =   176
               Top             =   405
               Width           =   870
            End
            Begin VB.Label Label28 
               Caption         =   "Days @ RM"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2880
               TabIndex        =   175
               Top             =   180
               Width           =   825
            End
            Begin VB.Label Label38 
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
               Height          =   255
               Left            =   900
               TabIndex        =   174
               Top             =   5040
               Width           =   1635
            End
            Begin VB.Label Label32 
               Caption         =   "Total"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2880
               TabIndex        =   173
               Top             =   5760
               Width           =   1635
            End
            Begin VB.Line Line5 
               BorderWidth     =   2
               X1              =   0
               X2              =   12240
               Y1              =   5715
               Y2              =   5715
            End
            Begin VB.Label Label31 
               Caption         =   "Tax/Others"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   172
               Top             =   4770
               Width           =   1635
            End
            Begin VB.Label Label30 
               Caption         =   "Admin/Registration"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   171
               Top             =   4545
               Width           =   1635
            End
            Begin VB.Line Line4 
               BorderWidth     =   2
               X1              =   90
               X2              =   9450
               Y1              =   3825
               Y2              =   3825
            End
            Begin VB.Label Label29 
               Caption         =   "Lodger"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   170
               Top             =   4320
               Width           =   1635
            End
            Begin VB.Label Label24 
               Caption         =   "Telephone Charges"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   169
               Top             =   4095
               Width           =   2085
            End
            Begin VB.Label Label23 
               Caption         =   "Medical Report"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   168
               Top             =   3870
               Width           =   1770
            End
            Begin VB.Label Label22 
               Caption         =   "Monthly Out-Patient Kidney Dialysis && Cancer"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   167
               Top             =   3510
               Width           =   3975
            End
            Begin VB.Line Line3 
               BorderWidth     =   2
               X1              =   90
               X2              =   9450
               Y1              =   2790
               Y2              =   2790
            End
            Begin VB.Label Label21 
               Caption         =   "Outpatient Physiotherapy Treatment"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   166
               Top             =   3285
               Width           =   2805
            End
            Begin VB.Label Label20 
               Caption         =   "Ambulance Fee"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   165
               Top             =   3060
               Width           =   2085
            End
            Begin VB.Label Label19 
               Caption         =   "Emergency Out-Patient (Accident)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   164
               Top             =   2835
               Width           =   2670
            End
            Begin VB.Line Line2 
               BorderWidth     =   2
               X1              =   -150
               X2              =   9210
               Y1              =   2000
               Y2              =   2000
            End
            Begin VB.Label Label17 
               Caption         =   "In-Hospital Physician Visit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   163
               Top             =   2280
               Width           =   2085
            End
            Begin VB.Label Label13 
               Caption         =   "Pre-Hospitalisation Services"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   162
               Top             =   2040
               Width           =   2580
            End
            Begin VB.Label Label16 
               Caption         =   "Pre-Surgical Diagnostic Services"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   161
               Top             =   1215
               Width           =   2670
            End
            Begin VB.Label Label15 
               Caption         =   "Surgical Fees"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   160
               Top             =   1440
               Width           =   2085
            End
            Begin VB.Label Label14 
               Caption         =   "Anaesthetic's Fees"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   159
               Top             =   1665
               Width           =   1635
            End
            Begin VB.Label Label12 
               Caption         =   "Operating Theatre"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   158
               Top             =   855
               Width           =   1590
            End
            Begin VB.Label Label11 
               Caption         =   "Hospital Supplies && Services"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   157
               Top             =   615
               Width           =   2445
            End
            Begin VB.Label Label10 
               Caption         =   "Intensive Care Unit"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   156
               Top             =   405
               Width           =   1485
            End
            Begin VB.Label Label1 
               Caption         =   "Daily Room && Board"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   155
               Top             =   120
               Width           =   1515
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               Caption         =   "Paid By Member"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2880
               TabIndex        =   154
               Top             =   6000
               Width           =   1230
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "Payable by MEDIX"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2880
               TabIndex        =   153
               Top             =   6225
               Width           =   1380
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               Caption         =   "Discount"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2880
               TabIndex        =   152
               Top             =   6450
               Width           =   675
            End
            Begin VB.Label Label37 
               Caption         =   "Gross Payable/(Recoverable)"
               BeginProperty Font 
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
               TabIndex        =   151
               Top             =   6720
               Width           =   2310
            End
            Begin VB.Label Label56 
               Caption         =   "Daily cash allowance at Govt hospital"
               Height          =   375
               Left            =   900
               TabIndex        =   150
               Top             =   5280
               Width           =   1575
            End
            Begin VB.Label Label57 
               Caption         =   "Days @RM"
               Height          =   255
               Left            =   2880
               TabIndex        =   149
               Top             =   5280
               Width           =   975
            End
            Begin VB.Label Label58 
               Caption         =   "LG Amount"
               Height          =   255
               Left            =   120
               TabIndex        =   148
               Top             =   5760
               Width           =   975
            End
            Begin VB.Label Label59 
               Caption         =   "Avail.Annual Limit"
               Height          =   255
               Left            =   1200
               TabIndex        =   147
               Top             =   5760
               Width           =   1335
            End
            Begin VB.Label Label18 
               Caption         =   "Post Hospitalisaion Treatment"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   900
               TabIndex        =   146
               Top             =   2520
               Width           =   2805
            End
         End
         Begin VB.VScrollBar VScroll2 
            Height          =   5415
            Left            =   9240
            TabIndex        =   8
            Top             =   240
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   234
      Top             =   360
      Width           =   9855
      Begin VB.TextBox txtCASNumber 
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
         TabIndex        =   238
         Top             =   240
         Width           =   1740
      End
      Begin VB.CommandButton cmdSearchCase 
         Caption         =   "Search"
         Height          =   255
         Left            =   3480
         TabIndex        =   237
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "##"
         Default         =   -1  'True
         Height          =   255
         Left            =   3120
         TabIndex        =   236
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   5640
         TabIndex        =   235
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label46 
         Caption         =   "Claim No"
         BeginProperty Font 
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
         TabIndex        =   240
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label26 
         Caption         =   "Member Number"
         Height          =   255
         Left            =   4320
         TabIndex        =   239
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmClaimAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intLastPoz As Integer
Dim intLastPoz2 As Integer
Dim VersionNo As Integer

Private Sub CmdInvoice_Click()
    frmClaimInvoice.show
End Sub

Private Sub Form_Load()
    Dim intFullHeight As Integer
    Dim intDisplayHeight As Integer
    
    VersionNo = 1
    
    intFullHeight = 5000
    intDisplayHeight = 2535
    intLastPoz = 0
    intLastPoz2 = 0
    
    Me.Height = intDisplayHeight
    
    With VScroll2
        '.Height = Me.ScaleHeight
        .Min = 0
        .Max = intFullHeight - intDisplayHeight
        .SmallChange = Screen.TwipsPerPixelX * 5
        .LargeChange = .SmallChange
    End With
End Sub

Private Sub cmdExit_Click()
    Dim RecordExit
    RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        Frame11.Top = 120
        Unload Me
    End If
End Sub

Private Sub cmdShowHistory_Click()
    If Trim(txtCASNumber.Text) = "" And Trim(txtMBMNumber.Text) = "" Then
        MsgBox "Please key in Case Number or Member Number first!", vbCritical + vbOKOnly
    Else
        With frmClaimHistory
            .txtCASNumber = txtCASNumber.Text
            .txtMBMNumber = txtMBMNumber.Text
            Call .ListCases("O", txtMBMNumber, txtCASNumber)
            Call .ListCases("C", txtMBMNumber)
            Call .ListCases("R", txtMBMNumber)
            .show
        End With
    End If
End Sub

' ############ Scroll Bar ##################
Public Sub ScrollForm2()
    Frame11.Top = Frame11.Top + intLastPoz2 - VScroll2.value
    intLastPoz2 = VScroll2.value
End Sub



Private Sub VScroll2_Change()
    ScrollForm2
End Sub

Private Sub VScroll2_Scroll()
    ScrollForm2
End Sub

' ### calculation  functions
' Section  Limit
'Section  ON Enter ==== LIMIT
Private Sub DB1_lostfocus()
    If Trim(DB1) = "" Then
        MsgBox "Please Enter No of Days for Room And Board", vbOKOnly + vbCritical, DialogBoxTitle
        DB1 = 0
        DB1.SetFocus
    ElseIf Trim(RoomAmount) = "" Then
        MsgBox "Please Enter Room Amount for Room And Board", vbOKOnly + vbCritical, DialogBoxTitle
        RoomAmount = 0
        RoomAmount.SetFocus
    Else
        Call CalculateRoomLimit
        Call CalculateTax
    End If
End Sub

Private Sub RoomAmount_LOSTFOCUS()
        If Trim(DB1) = "" Then
            MsgBox "Please Enter No of Days for Room And Board", vbOKOnly + vbCritical, DialogBoxTitle
            DB1 = 0
            DB1.SetFocus
        ElseIf Trim(RoomAmount) = "" Then
            MsgBox "Please Enter Room Amount for Room And Board", vbOKOnly + vbCritical, DialogBoxTitle
            RoomAmount = 0
            RoomAmount.SetFocus
        Else
            Call CalculateRoomLimit
            Call CalculateTax
        End If
End Sub

Private Sub ICU1_lostfocus()
    If ICU1 = "" Then
        MsgBox "Please Enter No of Days for Intensive Care Unit", vbOKOnly + vbCritical, DialogBoxTitle
        ICU1 = 0
        ICU1.SetFocus
    ElseIf Trim(ICUAmount) = "" Then
        MsgBox "Please Enter Room Amount for Intensive Care Unit", vbOKOnly + vbCritical, DialogBoxTitle
        ICUAmount = 0
        ICUAmount.SetFocus
    Else
        Call CalculateICULimit
        Call CalculateTax
    End If
End Sub

Private Sub ICUAmount_lostfocus()
    If Trim(ICU1) = "" Then
        MsgBox "Please Enter No of Days for Intensive Care Unit", vbOKOnly + vbCritical, DialogBoxTitle
        ICU1 = 0
        ICU1.SetFocus
    ElseIf Trim(ICUAmount) = "" Then
        MsgBox "Please Enter Room Amount for Intensive Care Unit", vbOKOnly + vbCritical, DialogBoxTitle
        ICUAmount = 0
        ICUAmount.SetFocus
    Else
        Call CalculateICULimit
        Call CalculateTax
    End If
End Sub

Private Sub LG1_lostfocus()
    If KeyAscii = 13 Then
        If Trim(LG1) = "" Then
            MsgBox "Please Enter No of Days for Lodger", vbOKOnly + vbCritical, DialogBoxTitle
            LG1 = 0
            LG1.SetFocus
        ElseIf Trim(LodgerAmount) = "" Then
            MsgBox "Please Enter Lodger Amount for Lodger", vbOKOnly + vbCritical, DialogBoxTitle
            LodgerAmount = 0
            LodgerAmount.SetFocus
        Else
            Call CalculateLodgerLimit
        End If
    End If
End Sub

Private Sub LodgerAmount_lostfocus()
    If LG1 = "" Then
        MsgBox "Please Enter No of Days for Lodger", vbOKOnly + vbCritical, DialogBoxTitle
        LG1 = 0
        LG1.SetFocus
    ElseIf Trim(LodgerAmount) = "" Then
        MsgBox "Please Enter Lodger Amount for Lodger", vbOKOnly + vbCritical, DialogBoxTitle
        LodgerAmount = 0
        LodgerAmount.SetFocus
    Else
        Call CalculateLodgerLimit
    End If
End Sub


Private Sub txtDayCash_lostfocus()
    If Trim(txtDayCash) = "" Then
        MsgBox "Please Enter No of Days for Cash Allowance", vbOKOnly + vbCritical, DialogBoxTitle
        txtDayCash = 0
        txtDayCash.SetFocus
    ElseIf Trim(txtRMCash) = "" Then
        MsgBox "Please Enter the Amount for Cash Allowance", vbOKOnly + vbCritical, DialogBoxTitle
        txtRMCash = 0
        txtRMCash.SetFocus
    Else
        If IsNumeric(txtDayCash) = True And IsNumeric(txtRMCash) Then
            P21 = Format(CDbl(txtDayCash) * CDbl(txtRMCash), "#,#0.00")
            txtDayCash.SetFocus
        End If
    End If
End Sub

Private Sub txtRMCash_LostFocus()
        If Trim(txtDayCash) = "" Then
            MsgBox "Please Enter No of Days for Cash Allowance", vbOKOnly + vbCritical, DialogBoxTitle
            txtDayCash = 0
            txtDayCash.SetFocus
        ElseIf Trim(txtRMCash) = "" Then
            MsgBox "Please Enter the Amount for Cash Allowance", vbOKOnly + vbCritical, DialogBoxTitle
            txtRMCash = 0
            txtRMCash.SetFocus
        Else
            If IsNumeric(txtDayCash) = True And IsNumeric(txtRMCash) Then
                P21 = Format(CDbl(txtDayCash) * CDbl(txtRMCash), "#,#0.00")
                txtDayCash.SetFocus
            End If
        End If
End Sub


' $########### Calculation for limit
Private Sub CalculateRoomLimit()
    Dim LimitAmount As Double
    If IsNumeric(RoomAmount) Then
        LimitAmount = DB1 * RoomAmount
    '    If LimitAmount > CDbl(P1A) Then
     '       P1 = P1A
      '  Else
         P1 = Format(LimitAmount, "#,#0.00")
       ' End If
    End If
End Sub


Private Sub CalculateICULimit()
    Dim ICULimitAmount
    ICULimitAmount = ICU1 * ICUAmount
'    If ICULimitAmount > CDbl(P2A) Then
 '       P2 = P2A
  '  Else
        P2 = Format(ICULimitAmount, "#,#0.00")
   ' End If
End Sub


Private Sub CalculateLodgerLimit()
    Dim LodgerLimitAmount As Double
    
    LodgerLimitAmount = LG1 * LodgerAmount
   '     If LodgerLimitAmount > CDbl(P17A) Then
    '        P17 = P17A
     '   Else
            P17 = Format(LodgerLimitAmount, "#,#0.00")
      '  End If
End Sub

Private Sub CalculateTax()
'On Error Resume Next
    If IsNumeric(P1) = True And IsNumeric(P2) = True Then
        P19 = Format((CDbl(P1) + CDbl(P2)) * 0.05, "#,#0.00")
        Call SumActual
    End If

End Sub

' Section Approved Amount
' ON lost focus for Actual Value
Private Sub A1_LostFocus()
    If IsNumeric(A1) Then
        Call ApprovedRoom
    End If
    M1 = Format(CDbl(AP1) - CDbl(H1), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A2_LostFocus()
    If IsNumeric(A2) Then
        Call ApprovedICU
    End If
    M2 = Format(CDbl(AP2) - CDbl(H2), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A3_LostFocus()
    If IsNumeric(A3) Then
        AP3 = Format(A3, "#,#0.00")
        H3 = Format(A3, "#,#0.00")
    End If
    M3 = Format(CDbl(AP3) - CDbl(H3), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A4_LostFocus()
    If IsNumeric(A4) Then
        AP4 = Format(A4, "#,#0.00")
        H4 = Format(A4, "#,#0.00")
    End If
    M4 = Format(CDbl(AP4) - CDbl(H4), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A5_LostFocus()
    If IsNumeric(A5) Then
        AP5 = Format(A5, "#,#0.00")
        H5 = Format(A5, "#,#0.00")
    End If
    M5 = Format(CDbl(AP5) - CDbl(H5), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A6_LostFocus()
    If IsNumeric(A6) Then
        Call ApprovedSurgical
    End If
    M6 = Format(CDbl(AP6) - CDbl(H6), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A7_LostFocus()
    If IsNumeric(A7) Then
        Call ApprovedAna
    End If
    M7 = Format(CDbl(AP7) - CDbl(H7), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A8_LostFocus()
    If IsNumeric(A8) Then
        AP8 = Format(A8, "#,#0.00")
        H8 = Format(A8, "#,#0.00")
    End If
    M8 = Format(CDbl(AP8) - CDbl(H8), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A9_LostFocus()
    If IsNumeric(A9) Then
        AP9 = Format(A9, "#,#0.00")
        H9 = Format(A9, "#,#0.00")
    End If
    M9 = Format(CDbl(AP9) - CDbl(H9), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A10_LostFocus()
    If IsNumeric(A10) Then
        AP10 = Format(A10, "#,#0.00")
        H10 = Format(A10, "#,#0.00")
    End If
    M10 = Format(CDbl(AP10) - CDbl(H10), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A11_LostFocus()
    If IsNumeric(A11) Then
        AP11 = Format(A11, "#,#0.00")
        H11 = Format(A11, "#,#0.00")
    End If
    M11 = Format(CDbl(AP11) - CDbl(H11), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A12_LostFocus()
    If IsNumeric(A12) Then
        AP12 = Format(A12, "#,#0.00")
        H12 = Format(A12, "#,#0.00")
    End If
    M12 = Format(CDbl(AP12) - CDbl(H12), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A13_LostFocus()
    If IsNumeric(A13) Then
        H13 = Format(A13, "#,#0.00")
    End If
    M13 = Format(CDbl(AP13) - CDbl(H13), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A14_LostFocus()
    If IsNumeric(A14) Then
        Call ApprovedOut
    End If
    M14 = Format(CDbl(AP14) - CDbl(H14), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A15_LostFocus()
    If IsNumeric(A15) Then
        AP15 = Format(A15, "#,#0.00")
        H15 = Format(A15, "#,#0.00")
    End If
    M15 = Format(CDbl(AP15) - CDbl(H15), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A16_LostFocus()
    If IsNumeric(A16) Then
        AP16 = Format(A16, "#,#0.00")
        H16 = Format(A16, "#,#0.00")
    End If
    M16 = Format(CDbl(AP16) - CDbl(H16), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A17_LostFocus()
    If IsNumeric(A17) Then
        Call ApprovedLodger
    End If
    M17 = Format(CDbl(AP17) - CDbl(H17), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A18_LostFocus()
    If IsNumeric(A18) Then
        H18 = Format(A18, "#,#0.00")
        AP18 = Format(A18, "#,#0.00")
    End If
    M18 = Format(CDbl(AP18) - CDbl(H18), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A19_LostFocus()
    If IsNumeric(A19) Then
        Call ApprovedTax
    End If
    M19 = Format(CDbl(AP19) - CDbl(H19), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A20_LostFocus()
    If IsNumeric(A20) Then
        Call ApprovedOthers
    End If
    M20 = Format(CDbl(AP20) - CDbl(H20), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

Private Sub A21_LostFocus()
    If IsNumeric(A21) Then
        AP21 = Format(A21, "#,#0.00")
    End If
    H21 = 0
    M21 = Format(CDbl(AP21) - CDbl(H21), "#,#0.00;(#,#0.00)")
    Call SumActual
End Sub

' End Lost Focus FOR LIMIT SECTION
' ###################################################3
' --------------------- START APPROVED SECTION --------------

' Start Approved Functions
Private Sub ApprovedRoom()
'On Error Resume Next
    If CDbl(A1) > P1 Then
        AP1 = Format(P1, "#,#0.00")
    Else
        AP1 = Format(A1, "#,#0.00")
    End If
    H1 = Format(A1, "#,#0.00")
End Sub


Private Sub ApprovedICU()
'On Error Resume Next
    If CDbl(A2) > P2 Then
        AP2 = Format(P2, "#,#0.00")
    Else
        AP2 = Format(A2, "#,#0.00")
    End If
    H2 = Format(A2, "#,#0.00")
End Sub

Private Sub ApprovedSurgical()
    If CDbl(A6) > CDbl(P6) Then
        AP6 = Format(P6, "#,#0.00")
    Else
        AP6 = Format(A6, "#,#0.00")
    End If
    H6 = Format(A6, "#,#0.00")
End Sub

Private Sub ApprovedAna()
    If CDbl(A7) > CDbl(P7) Then
        AP7 = Format(P7, "#,#0.00")
    Else
        AP7 = Format(A7, "#,#0.00")
    End If
    H7 = Format(A7, "#,#0.00")
End Sub

Private Sub ApprovedOut()
'On Error Resume Next
    If CDbl(A14) > CDbl(P14) Then
        AP14 = Format(P14, "#,#0.00")
    Else
        AP14 = Format(A14, "#,#0.00")
    End If
    H14 = Format(A14, "#,#0.00")
End Sub

Private Sub ApprovedLodger()
    If CDbl(A17) > P17 Then
        AP17 = Format(P17, "#,#0.00")
    Else
        AP17 = Format(A17, "#,#0.00")
    End If
    H17 = Format(A17, "#,#0.00")
End Sub

Private Sub ApprovedTax()
    If CDbl(A19) > CDbl(P19) Then
        AP19 = Format(P19, "#,#0.00")
    Else
        AP19 = Format(A19, "#,#0.00")
    End If
    H19 = Format(A19, "#,#0.00")
End Sub

Private Sub ApprovedOthers()
    If CDbl(A20) > CDbl(P20) Then
        AP20 = Format(P20, "#,#0.00")
    Else
        AP20 = Format(A20, "#,#0.00")
    End If
    H20 = Format(A20, "#,#0.00")
End Sub
' ----------------- Start Lost focus to get the due to/from Member ------------
Private Sub AP1_lostFocus()
    M1 = Format(CDbl(AP1) - CDbl(H1), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP2_lostFocus()
    M2 = Format(CDbl(AP2) - CDbl(H2), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP3_lostFocus()
    M3 = Format(CDbl(AP3) - CDbl(H3), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP4_lostFocus()
    M4 = Format(CDbl(AP4) - CDbl(H4), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP5_lostFocus()
    M5 = Format(CDbl(AP5) - CDbl(H5), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP6_lostFocus()
    M6 = Format(CDbl(AP6) - CDbl(H6), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP7_lostFocus()
    M7 = Format(CDbl(AP7) - CDbl(H7), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP8_lostFocus()
    M8 = Format(CDbl(AP8) - CDbl(H8), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP9_lostFocus()
    M9 = Format(CDbl(AP9) - CDbl(H9), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP10_lostFocus()
    M10 = Format(CDbl(AP10) - CDbl(H10), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP11_lostFocus()
    M11 = Format(CDbl(AP11) - CDbl(H11), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP12_lostFocus()
    M12 = Format(CDbl(AP12) - CDbl(H12), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP13_lostFocus()
    M13 = Format(CDbl(AP13) - CDbl(H13), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP14_lostFocus()
    M14 = Format(CDbl(AP14) - CDbl(H14), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP15_lostFocus()
    M15 = Format(CDbl(AP15) - CDbl(H15), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP16_lostFocus()
    M16 = Format(CDbl(AP16) - CDbl(H16), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP17_lostFocus()
    M17 = Format(CDbl(AP17) - CDbl(H17), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP18_lostFocus()
    M18 = Format(CDbl(AP18) - CDbl(H18), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP19_lostFocus()
    M19 = Format(CDbl(AP19) - CDbl(H19), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP20_lostFocus()
    M20 = Format(CDbl(AP20) - CDbl(H20), "#,#0.00;(#,#0.00)")
End Sub

Private Sub AP21_lostFocus()
    M21 = Format(CDbl(AP21) - CDbl(H21), "#,#0.00;(#,#0.00)")
End Sub


'Section TOTAL
' '''''''''''''''''''''''''''''''''''

Private Sub TotalDiscountMember_lostfocus()
    TotalGrossMember = Format(CDbl(TotalMember) + CDbl(TotalPaidByMemberMember) + CDbl(TotalPaidMedixMember) + CDbl(TotalDiscountMember), "#,#0.00;(#,#0.00)")
End Sub

Private Sub TotalDiscount_LOSTFOCUS()
    'Call CalculateTotalGross
    TotalGross = Format(CDbl(TotalActual) + CDbl(TotalDiscount), "#,#0.00;(#,#0.00)")
End Sub

Private Sub TotalPaidByMemberHos_LostFocus()
    TotalPaidByMemberMember = Format(CDbl(TotalPaidByMemberHos), "#,#0.00;(#,#0.00)")
    TotalPaidMedixHos = IIf(CDbl(TotalHospital) - CDbl(TotalPaidByMemberHos) > CDbl(TotalPaidMedixHos) + CDbl(txtLGAmount), Format(txtLGAmount, "#,#0.00;(#,#0.00)"), Format(CDbl(TotalHospital) - CDbl(TotalPaidByMemberHos), "#,#0.00;(#,#0.00)"))
    TotalGrossHos = Format(CDbl(TotalPaidMedixHos) + CDbl(TotalDiscountHos), "#,#0.00;(#,#0.00)")
    TotalGrossMember = Format(CDbl(TotalMember) + CDbl(TotalPaidByMemberMember) + CDbl(TotalPaidMedixMember) + CDbl(TotalDiscountMember), "#,#0.00;(#,#0.00)")
End Sub

Private Sub TotalDiscountHos_Lostfocus()
    Call CalculateTotalGross
End Sub

Private Sub CalculateTotalGross()
'On Error Resume Next
    'TotalPaidByMemberHos = Format(CDbl(TotalPaidByMember), "#,#0.00;(#,#0.00)")
    'TotalPaidMedix = Format(IIf((CDbl(TotalActual) - CDbl(TotalPaidByMember)) > P21, P21, CDbl(TotalActual) - CDbl(TotalPaidByMember)), "#,#0.00;(#,#0.00)")
    'TotalPaidMedixMember = Format(CDbl(TotalPaidMedix), "#,#0.00;(#,#0.00)")
    TotalGrossApproved = Format(IIf(CDbl(TotalApproved) < CDbl(PAnnualLimit), TotalApproved, PAnnualLimit), "#,#0.00;(#,#0.00)")
    TotalPaidMedixHos = IIf(CDbl(TotalHospital) - CDbl(TotalPaidByMemberHos) > CDbl(TotalPaidMedixHos) + CDbl(txtLGAmount), Format(txtLGAmount, "#,#0.00;(#,#0.00)"), Format(CDbl(TotalHospital) - CDbl(TotalPaidByMemberHos), "#,#0.00;(#,#0.00)"))
    TotalPaidByMemberMember = Format(CDbl(TotalPaidByMemberHos), "#,#0.00;(#,#0.00)")
    TotalDiscountMember = IIf(CDbl(TotalGrossApproved) < CDbl(TotalApproved), Format(CDbl(TotalGrossApproved) - CDbl(TotalApproved), "#,#0.00;(#,#0.00)"), 0)
    
    TotalGross = Format(CDbl(TotalActual) + CDbl(TotalDiscount), "#,#0.00;(#,#0.00)")
    TotalGrossHos = Format(CDbl(TotalPaidMedixHos) + CDbl(TotalDiscountHos), "#,#0.00;(#,#0.00)")
    TotalGrossMember = Format(CDbl(TotalMember) + CDbl(TotalPaidByMemberMember) + CDbl(TotalPaidMedixMember) + CDbl(TotalDiscountMember), "#,#0.00;(#,#0.00)")
End Sub

Private Sub SumActual()
'On Error Resume Next
    TotalActual = Format(CDbl(A1) + CDbl(A2) + CDbl(A3) + CDbl(A4) + CDbl(A5) + CDbl(A6) + CDbl(A7) + CDbl(A8) + CDbl(A9) + CDbl(A10) + CDbl(A11) + CDbl(A12) + CDbl(A13) + CDbl(A14) + CDbl(A15) + CDbl(A16) + CDbl(A17) + CDbl(A18) + CDbl(A19) + CDbl(A20) + CDbl(A21), "#,#0.00")
    TotalApproved = Format(CDbl(AP1) + CDbl(AP2) + CDbl(AP3) + CDbl(AP4) + CDbl(AP5) + CDbl(AP6) + CDbl(AP7) + CDbl(AP8) + CDbl(AP9) + CDbl(AP10) + CDbl(AP11) + CDbl(AP12) + CDbl(AP13) + CDbl(AP14) + CDbl(AP15) + CDbl(AP16) + CDbl(AP17) + CDbl(AP18) + CDbl(AP19) + CDbl(AP20) + CDbl(AP21), "#,#0.00")
    TotalHospital = Format(CDbl(H1) + CDbl(H2) + CDbl(H3) + CDbl(H4) + CDbl(H5) + CDbl(H6) + CDbl(H7) + CDbl(H8) + CDbl(H9) + CDbl(H10) + CDbl(H11) + CDbl(H12) + CDbl(H13) + CDbl(H14) + CDbl(H15) + CDbl(H16) + CDbl(H17) + CDbl(H18) + CDbl(H19) + CDbl(H20) + CDbl(H21), "#,#0.00")
    TotalMember = Format(CDbl(M1) + CDbl(M2) + CDbl(M3) + CDbl(M4) + CDbl(M5) + CDbl(M6) + CDbl(M7) + CDbl(M8) + CDbl(M9) + CDbl(M10) + CDbl(M11) + CDbl(M12) + CDbl(M13) + CDbl(M14) + CDbl(M15) + CDbl(M16) + CDbl(M17) + CDbl(M18) + CDbl(M19) + CDbl(M20) + CDbl(M21), "#,#0.00;(#,#0.00)")
    
End Sub


Private Sub CmdCalc_Click()
   
    M1 = Format(CDbl(AP1) - CDbl(H1), "#,#0.00;(#,#0.00)")
    M2 = Format(CDbl(AP2) - CDbl(H2), "#,#0.00;(#,#0.00)")
    M3 = Format(CDbl(AP3) - CDbl(H3), "#,#0.00;(#,#0.00)")
    M4 = Format(CDbl(AP4) - CDbl(H4), "#,#0.00;(#,#0.00)")
    M5 = Format(CDbl(AP5) - CDbl(H5), "#,#0.00;(#,#0.00)")
    M6 = Format(CDbl(AP6) - CDbl(H6), "#,#0.00;(#,#0.00)")
    M7 = Format(CDbl(AP7) - CDbl(H7), "#,#0.00;(#,#0.00)")
    M8 = Format(CDbl(AP8) - CDbl(H8), "#,#0.00;(#,#0.00)")
    M9 = Format(CDbl(AP9) - CDbl(H9), "#,#0.00;(#,#0.00)")
    M10 = Format(CDbl(AP10) - CDbl(H10), "#,#0.00;(#,#0.00)")
    M11 = Format(CDbl(AP11) - CDbl(H11), "#,#0.00;(#,#0.00)")
    M12 = Format(CDbl(AP12) - CDbl(H12), "#,#0.00;(#,#0.00)")
    M13 = Format(CDbl(AP13) - CDbl(H13), "#,#0.00;(#,#0.00)")
    M14 = Format(CDbl(AP14) - CDbl(H14), "#,#0.00;(#,#0.00)")
    M15 = Format(CDbl(AP15) - CDbl(H15), "#,#0.00;(#,#0.00)")
    M16 = Format(CDbl(AP16) - CDbl(H16), "#,#0.00;(#,#0.00)")
    M17 = Format(CDbl(AP17) - CDbl(H17), "#,#0.00;(#,#0.00)")
    M18 = Format(CDbl(AP18) - CDbl(H18), "#,#0.00;(#,#0.00)")
    M19 = Format(CDbl(AP19) - CDbl(H19), "#,#0.00;(#,#0.00)")
    M20 = Format(CDbl(AP20) - CDbl(H20), "#,#0.00;(#,#0.00)")
    M21 = Format(CDbl(AP21) - CDbl(H21), "#,#0.00;(#,#0.00)")
    
    Call SumActual
    Call CalculateTotalGross
End Sub



' #################### Search Funtion
Private Sub cmdSearchCase_Click()
    frmCASList.Source = 5
    frmCASList.show
End Sub

Public Sub cmdDisplay_Click()
On Error Resume Next
    Dim CAS As New ADODB.Recordset
    Dim strsql As String
    '    cboHOSCode.Clear
 '   lstInvoice.ListItems.Clear
    strsql = "SELECT CASNumber, MBMNumber, " & _
        " MBMPatientCovID, MBMPolicyNo, MBMName, " & _
        " INSCode, PLNCode, CASStatus, " & _
        " CASRemark, CaseDateOfLoss , MBMPayorEffDate , HLTCode ," & _
        " CASHOSCode,    CASDateAdmitted, CASDischargeDate , MBMAvailableLimit, CASReserveAmount"
    strsql = strsql & " from VIEW_Cas_MBM_Claims where 0=0"
    
    If Trim(txtCASNumber) <> "" Then
        strsql = strsql & " And CASNumber = '" & Trim(txtCASNumber) & "'"
        CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If CAS.recordCount Then
                
'                Call ListCases("O", CAS!MBMNumber, txtCASNumber)
 '               Call ListCases("C", CAS!MBMNumber)
  '              Call ListCases("R", CAS!MBMNumber)
                Call display_coverPersons(CAS!MBMPatientCovID, CAS!MBMNumber, CAS!MBMName)
                Call Display_Diagnosis("F", txtCASNumber)
                Call Display_Hospital(txtCASNumber)
                With frmClaimAdmission
                    .txtCASNumber = CAS!CASNumber
                    .txtMBMNumber = CAS!MBMNumber
                    .txtMBMName = CAS!MBMName
                    .txtMBMPolNo = CAS!MBMPolicyNo
                    .txtPLNCode = CAS!PLNCode
                    .txtINSCode = CAS!InsCode
                    '.txtPolEffDate.Mask = ""
                    '.txtPolEffDate = CAS!MBMPayorEffDate
                    If Len(Trim(CAS!CASRemark)) Then
                        .txtCASRemark = CAS!CASRemark
                    End If
                    .txtHLTCode = CAS!HLTCode
                    .txtAdmissionDate = CAS!CASDateAdmitted
                    If Len(Trim(CAS!CASDischargeDate)) <> 0 Then
                        .txtDisChargeDate = CAS!CASDischargeDate
                    End If
                    If Len(Trim(CAS!CASHOSCode)) <> 0 Then
                        .txtHOS = CAS!CASHOSCode
                    End If
                    .txtAnnualLimit = CAS!MBMAvailableLimit
                    '.PAnnualLimit = CAS!MBMAvailableLimit
                    '.txtLGAmount = CAS!CASReserveAmount
                    Call display_benefit
                  '  cboHOSCode.AddItem .txtHOS & "-" & FindHosName(.txtHOS)
                 End With
            Else
                MsgBox "Record not found! "
            End If
    Else
        MsgBox "Case Number cannot be empty! "
    End If
End Sub


' ############## Claims History ###############################
Public Sub ListCases(CaseStatus As String, MBMNumber As String, Optional CaseNumber As String)
    Dim CASList As ListItem
    Dim strsql As String
    Dim CAS As New ADODB.Recordset
    
    strsql = "SELECT CASNumber, MBMNumber, " & _
        " MBMPatientCovID, MBMPolicyNo, MBMName, " & _
        " INSCode, PLNCode, CASStatus, " & _
        " CASRemark, CaseDateOfLoss , MBMPayorEffDate , HLTCode"
    strsql = strsql & " from VIEW_Cas_MBM_Claims where 0=0"
    
    If Len(Trim(CaseStatus)) Then
        strsql = strsql & " AND CASStatus ='" & Trim(CaseStatus) & "'"
    End If
    
    If Trim(MBMNumber) <> "" Then
        strsql = strsql & " And MBMNumber = '" & Trim(MBMNumber) & "'"
    End If
    
    If Trim(CaseNumber) <> "" Then
        strsql = strsql & " And CASNumber <> '" & Trim(CaseNumber) & "'"
    End If
    
    
    CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'If ChkStr(CaseStatus) = "O" Then
     '    lstClaimsList.ListItems.Clear
    'ElseIf ChkStr(CaseStatus) = "C" Then
     '    lstClosedClaims.ListItems.Clear
    'Else
     '    lstRejectedClaims.ListItems.Clear
    'End If
    If CAS.recordCount <> 0 Then
        Screen.MousePointer = vbHourglass
        Do
         Do Until CAS.EOF
     '       If ChkStr(CaseStatus) = "O" Then
      '          Set CASList = lstClaimsList.ListItems.Add(, , CAS!CASNumber)
       '     ElseIf ChkStr(CaseStatus) = "C" Then
        '        Set CASList = lstClosedClaims.ListItems.Add(, , CAS!CASNumber)
         '   Else
          '      Set CASList = lstRejectedClaims.ListItems.Add(, , CAS!CASNumber)
           ' End If
                CASList.SubItems(1) = CAS!MBMNumber
                CASList.SubItems(2) = CAS!MBMName
                If Len(Trim(CAS!CaseDateOfLoss)) Then
                    CASList.SubItems(3) = CAS!CaseDateOfLoss
                End If
                CASList.SubItems(4) = CAS!CASStatus
                CASList.SubItems(5) = CAS!MBMPatientCovID
                If Len(Trim(CAS!CASRemark)) Then
                    CASList.SubItems(6) = CAS!CASRemark
                End If
                CAS.MoveNext
                        
            DoEvents
            If intExit = 1 Then
                check = False
                Exit Do
            End If
         Loop
        Loop Until check = False
        Screen.MousePointer = vbDefault
    End If
    intExit = 0
    CAS.Close
    Set CAS = Nothing
End Sub

'  ############## Display data
' ******* Cover Persons  *******
Public Sub display_coverPersons(CoverID As String, MBMNumber As String, Optional MBMName As String)
    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String

    If ChkStr(CoverID) <> "SELF" Then
        sqlstr = "select MBMCNumber, MBMCCoverID, MBMCName, " & _
                " MBMCRELCode,  MBMCStatus , MBMCAnnualLimit " & _
                " from  MBMCoveredPersons where MBMCNumber = '" & Trim(MBMNumber) & "'" & _
                " AND MBMCCoverID = '" & CoverID & "'"
        MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBMCov.recordCount Then
            txtPatientName = MBMCov!MBMCName
            txtRELcode = MBMCov!MBMCRELCode
        End If
    Else
        txtPatientName = MBMName
        txtRELcode = "Self"
    End If
End Sub

Public Sub Display_Diagnosis(DIAStatus As String, CASNumber As String)
    Dim DiagnosisSql As String
    Dim lstDiag As ListItem
    Dim rst3 As New ADODB.Recordset
    
    DiagnosisSql = "select DIAIndex, CASNumber , ICDCode , " & _
                        " DIADiagnoseBy, DIADiagDate "
    If ChkStr(DIAStatus) = "I" Then
        DiagnosisSql = DiagnosisSql & " from VIEW_DiaFirst_ICD where CASNUMBer ='" & Trim(CASNumber) & "'"
    Else
        DiagnosisSql = DiagnosisSql & " from View_DiagFinal_ICD where CASNUMBer ='" & Trim(CASNumber) & "'"
    End If
    
    rst3.Open DiagnosisSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    ' For Final
    lstFinalDiag.listitems.Clear

     If rst3.recordCount Then
'    If Len(Trim(rst3!DiaIndex)) <> 0 Then
        Do Until rst3.EOF
                Set lstDiag = lstFinalDiag.listitems.Add(, , rst3!DiaIndex)
                    With lstDiag
                    If Len(Trim(rst3!ICDDescription)) Then
                        .SubItems(1) = Trim(rst3!ICDDescription)
                    End If
                        .SubItems(2) = rst3!DIADiagnoseBy
                        .SubItems(3) = rst3!DIADiagDate
                    End With
                rst3.MoveNext
        Loop
 '   End If
    End If
    
    rst3.Close
    Set rst3 = Nothing
End Sub

Public Function FindHosName(HosCode As String) As String
    Dim HOSSql As String
    Dim rst3 As New ADODB.Recordset
    HOSSql = " select HOSCode, HOSNAme  from HOS where HOSCOde='" & Trim(HosCode) & "'"
    
    rst3.Open HOSSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If rst3.recordCount Then
        FindHosName = rst3!HOSName
    Else
        MsgBox "Hospital Name not found!", vbCritical + vbOKOnly
    End If
    rst3.Close
    Set rst3 = Nothing
End Function


Public Sub Display_Hospital(CASNumber As String)
    Dim HOSSql As String
    Dim lstHos As ListItem
    Dim rst3 As New ADODB.Recordset
    HOSSql = " select HOSCode, CaseCode, " & _
             " DateAdmit, DateDischarge ,HOSName"
    HOSSql = HOSSql & " from VIEW_CAS_HOS_HOS where CASeCode ='" & Trim(CASNumber) & "'"
    
    rst3.Open HOSSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    lstHosList.listitems.Clear
    If rst3.recordCount Then
        Do Until rst3.EOF
            Set lstHos = lstHosList.listitems.Add(, , Trim(rst3!HosCode))
                With lstHos
                    .SubItems(1) = Trim(rst3!HOSName)
                    If Len(rst3!DateAdmit) Then
                        .SubItems(2) = CDate(rst3!DateAdmit)
                    End If
                    If Len(rst3!DateDischarge) Then
                        .SubItems(3) = CDate(rst3!DateDischarge)
                    End If
                End With
            'cboHOSCode.AddItem "Trim(rst3!HOSCode)" & "-" & Trim(rst3!HOSName)
            rst3.MoveNext
        Loop
    End If
    rst3.Close
    Set rst3 = Nothing
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
' Benefits

Public Sub display_benefit()
    Dim benefitstr  As String
    Dim BNF As New ADODB.Recordset
     
    benefitstr = "Select BNFCode , BNFClaimableAmount from BNF Where INSCode ='" & Trim(txtINSCode) & "' And PLNCode ='" & Trim(txtPLNCode) & "' And HLTCode ='" & Trim(txtHLTCode) & "'"
    BNF.Open benefitstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    If BNF.recordCount Then
        Do Until BNF.EOF
            With BNF
                If InStr(1, !BNFCode, "-01") Then
                    P1A = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-02") Then
                    P2A = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-03") Then
                    P3 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-04") Then
                    P4 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-05") Then
                    P5 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-06") Then
                    P6 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-07") Then
                    P7 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-08") Then
                    P8 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-09") Then
                    P9 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-10") Then
                    P10 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-11") Then
                    P11 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-12") Then
                    P12 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-13") Then
                    P13 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-14") Then
                    P14 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-15") Then
                    P15 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-16") Then
                    P16 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-17") Then
                    P17A = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-18") Then
                    P18 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-19") Then
                    P19 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-20") Then
                    P20 = !BNFClaimAbleAmount
                ElseIf InStr(1, !BNFCode, "-21") Then
                    txtRMCashA = !BNFClaimAbleAmount
                End If
            End With
            BNF.MoveNext
        Loop
        If IsNumeric(P1A) Then
            RoomAmount = CInt(P1A) / 150
        End If
        If IsNumeric(P2A) Then
            ICUAmount = CInt(P2A) / 75
        End If
        If IsNumeric(P17A) Then
            LodgerAmount = CInt(P17A) / 60
        End If
        If IsNumeric(txtRMCashA) Then
            txtRMCash = CInt(txtRMCashA) / 150
        End If
        
    Else
        MsgBox "Error! Please check on your benefit table!", vbCritical + vbOKOnly
    End If
    BNF.Close
    Set BNF = Nothing
End Sub


' ########### Invoices  ##############
Private Sub cmdAddHos_Click()
    Dim HOSMaster As ListItem
    Dim RecordSaved
    If txtInvoicedDate = "__/__/____" Then
        MsgBox "Please enter Invoice Date ", vbOKOnly + vbCritical
    ElseIf cboHosCode <> "" Then
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Set HOSMaster = lstInvoice.listitems.Add(, , Mid(cboHosCode, 1, InStr(1, cboHosCode, "-") - 1))
            HOSMaster.SubItems(1) = Trim(txtInvoiceNo)
            HOSMaster.SubItems(2) = Trim(txtInvoicedDate)
            HOSMaster.SubItems(3) = Trim(txtAmount)
            If Trim(txtPaidHos) <> "" Then
                HOSMaster.SubItems(4) = Trim(txtPaidHos)
            End If
            If Trim(txtPaidMBM) <> "" Then
                HOSMaster.SubItems(5) = Trim(txtPaidMBM)
            End If

            If Trim(txtPaidMBM) <> "" Then
                HOSMaster.SubItems(6) = Trim(txtByMBM)
            End If
            If Trim(txtRemark) <> "" Then
                HOSMaster.SubItems(7) = Trim(txtRemark)
            End If
        
            txtInvoiceNo = ""
            cboHosCode = ""
            txtRemark = ""
            txtByMBM = ""
            txtPaidMBM = ""
            txtPaidHos = ""
            txtAmount = ""
            txtInvoicedDate.Mask = ""
            txtInvoicedDate.Text = ""
            txtInvoicedDate.Mask = "##/##/####"
        End If
    End If
   
End Sub


Private Sub cmdRemoveHos_Click()
'On Error Resume Next
    Dim ForLoop As Integer
    Dim ListCount As Integer
    Dim RecordSaved
    
    ListCount = lstInvoice.listitems.count
    If ListCount > 0 Then
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            For ForLoop = 1 To ListCount
                If lstInvoice.listitems(ForLoop).Selected = True Then
                    lstInvoice.listitems.Remove (ForLoop)
                    Exit For
                End If
            Next ForLoop
        End If
    End If
End Sub

' ########### SAVE TO DB #############################

Private Sub cmdSave_Click()
    On Error GoTo errSave
    Dim RecordSaved
    Dim startTrans As Boolean
    Dim strsql As String
    
    Dim CheckClaim As New ADODB.Recordset
    strsql = "select CASNumber from Claim where CAsNUmber ='" & Trim(txtCASNumber) & "'"
    CheckClaim.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If CheckClaim.recordCount Then
      MsgBox "Worksheet for this case has been opened before!" & vbNewLine & _
             " Please use Claim adjustment for further editting!", vbCritical + vbOKOnly
    Else
        If Len(Trim(txtCASNumber)) = 0 Then
            MsgBox "Please enter Case Number!", vbOKOnly + vbCritical, DialogBoxTitle
            txtCASNumber.SetFocus
        ElseIf Len(Trim(txtMBMNumber)) = 0 Then
            MsgBox "Please Member Number !", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMNumber.SetFocus
        ElseIf lstInvoice.listitems.count = 0 Then
            MsgBox "Worksheet cannot be saved without any invoice!", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 1
        Else
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
               Screen.MousePointer = vbHourglass
               startTrans = True
               medixmasterconn.beginTrans
               
               Call saveClaim
               Call saveLimit
               Call saveActual
               Call saveApprove
               Call saveDueMember
               Call saveHospital
               Call saveInvoice
               
               medixmasterconn.CommitTrans
               Screen.MousePointer = Default
               MsgBox "Record Save", vbOKOnly + vbInformation, DialogBoxTitle
               cmdSave.Enabled = False
             End If
        End If
    End If
    CheckClaim.Close
    Set CheckClaim = Nothing
    Exit Sub
errSave:
    MsgBox ERR.Description
    If startTrans = True Then
        medixmasterconn.RollbackTrans
    End If
End Sub

Private Sub saveClaim()
    Dim Claim As New ADODB.Recordset
    Claim.Open "Claim", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Claim.AddNew
    With Claim
         !CASNumber = txtCASNumber
         !claimversion = VersionNo
         !MBMNumber = txtMBMNumber
         !MBMCoveredID = txtCoveredID
         !ClaimStatus = "O"
         !ClaimOpenDate = Date
         !ClaimOpenBy = ChkStr(Logon)
         !ClaimOthers = txtOthers
         !ClaimBoardDay = IIf(IsNumeric(DB1), DB1, 0)
         !ClaimICUDay = IIf(IsNumeric(ICU1), ICU1, 0)
         !ClaimLodgerDay = IIf(IsNumeric(LG1), LG1, 0)
         !ClaimCashAllowanceDay = IIf(IsNumeric(txtDayCash), txtDayCash, 0)
         !ClaimRoomAmount = IIf(IsNumeric(RoomAmount), RoomAmount, 0)
         !ClaimICUAmount = IIf(IsNumeric(ICUAmount), ICUAmount, 0)
         !ClaimLodgerAmount = IIf(IsNumeric(LodgerAmount), LodgerAmount, 0)
         !ClaimCashAllowanceAmount = IIf(IsNumeric(txtRMCash), txtRMCash, 0)
         !ClaimLastUpdateUser = ChkStr(Logon)
         !ClaimLastUpdateDate = Date
    End With
    Claim.Update
    
    Claim.Close
    Set Claim = Nothing
End Sub

Private Sub saveActual()
    Dim ClaimActual  As New ADODB.Recordset
    ClaimActual.Open "ClaimActual", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    ClaimActual.AddNew
    With ClaimActual
         !CASNumber = txtCASNumber
         !CLMVersion = VersionNo
         !CLMBoard = A1
         !CLMIntensive = A2
         !CLMServices = A3
         !CLMTheatre = A4
         !CLMPre = A5
         !CLMSurgical = A6
         !CLMAna = A7
         !CLMHospitalisation = A8
         !CLMPhysician = A9
         !CLMPost = A10
         !CLMEmergency = A11
         !CLMAmbulance = A12
         !CLMOutpatient = A13
         !CLMMonthly = A14
         !CLMMedical = A15
         !CLMPhone = A16
         !CLMLodger = A17
         !CLMAdmin = A18
         !CLMTax = A19
         !CLMOthers = A20
         !CLMCash = A21
         !CLMDiscount = TotalDiscount
         !CLMGross = TotalGross
         !CLMTotal = TotalActual
    End With
    ClaimActual.Update
    
    ClaimActual.Close
    Set ClaimActual = Nothing
End Sub

Private Sub saveLimit()
    Dim ClaimLimit  As New ADODB.Recordset
    ClaimLimit.Open "ClaimLimit", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    ClaimLimit.AddNew
    With ClaimLimit
         !CASNumber = txtCASNumber
         !CLMLVersion = VersionNo
         !CLMLBoard = A1
         !CLMLIntensive = A2
         !CLMLServices = A3
         !CLMLTheatre = A4
         !CLMLPre = A5
         !CLMLSurgical = A6
         !CLMLAna = A7
         !CLMLHospitalisation = A8
         !CLMLPhysician = A9
         !CLMLPost = A10
         !CLMLEmergency = A11
         !CLMLAmbulance = A12
         !CLMLOutpatient = A13
         !CLMLMonthly = A14
         !CLMLMedical = A15
         !CLMLPhone = A16
         !CLMLLodger = A17
         !CLMLAdmin = A18
         !CLMLTax = A19
         !CLMLOthers = A20
    End With
    ClaimLimit.Update
    
    ClaimLimit.Close
    Set ClaimLimit = Nothing
End Sub


Private Sub saveApprove()
    Dim ClaimApprove As New ADODB.Recordset
    ClaimApprove.Open "ClaimApprove", medixmasterconn, adOpenKeyset, adLockOptimistic

    ClaimApprove.AddNew
    With ClaimApprove
         !CASNumber = txtCASNumber
         !CLAVersion = VersionNo
         !CLABoard = AP1
         !CLAIntensive = AP2
         !CLAServices = AP3
         !CLATheatre = AP4
         !CLAPre = AP5
         !CLASurgical = AP6
         !CLAAna = AP7
         !CLAHospitalisation = AP8
         !CLAPhysician = AP9
         !CLAPost = AP10
         !CLAEmergency = AP11
         !CLAAmbulance = AP12
         !CLAOutpatient = AP13
         !CLAMonthly = AP14
         !CLAMedical = AP15
         !CLAPhone = AP16
         !CLALodger = AP17
         !CLAAdmin = AP18
         !CLATax = AP19
         !CLAOthers = AP20
         !CLACash = AP21
         !CLATotal = TotalApproved
         !CLAGross = TotalGrossApproved
    End With
    ClaimApprove.Update
    
    ClaimApprove.Close
    Set ClaimApprove = Nothing

End Sub


Private Sub saveDueMember()
    Dim ClaimDueMember As New ADODB.Recordset

    ClaimDueMember.Open "ClaimDueMember", medixmasterconn, adOpenKeyset, adLockOptimistic
    ClaimDueMember.AddNew
    With ClaimDueMember
         !CASNumber = txtCASNumber
         !CLDVersion = VersionNo
         !CLDBoard = M1
         !CLDIntensive = M2
         !CLDServices = M3
         !CLDTheatre = M4
         !CLDPre = M5
         !CLDSurgical = M6
         !CLDAna = M7
         !CLDHospitalisation = M8
         !CLDPhysician = M9
         !CLDPost = M10
         !CLDEmergency = M11
         !CLDAmbulance = M12
         !CLDOutpatient = M13
         !CLDMonthly = M14
         !CLDMedical = M15
         !CLDPhone = M16
         !CLDLodger = M17
         !CLDAdmin = M18
         !CLDTax = M19
         !CLDOthers = M20
         !CLDCash = M21
         !CLDPaidByMember = TotalPaidByMemberMember
         !CLDPayableByMedix = TotalPaidMedixMember
         !CLDDiscount = TotalDiscountMember
         !CLDGross = TotalGrossMember
         !CLDTotal = TotalMember
    End With
    ClaimDueMember.Update

    ClaimDueMember.Close
    Set ClaimDueMember = Nothing
End Sub

Private Sub saveHospital()
    Dim ClaimHospital As New ADODB.Recordset
    ClaimHospital.Open "ClaimHospital", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    ClaimHospital.AddNew
    With ClaimHospital
         !CASNumber = txtCASNumber
         !CLHVersion = VersionNo
         !CLHBoard = H1
         !CLHIntensive = H2
         !CLHServices = H3
         !CLHTheatre = H4
         !CLHPre = H5
         !CLHSurgical = H6
         !CLHAna = H7
         !CLHHospitalisation = H8
         !CLHPhysician = H9
         !CLHPost = H10
         !CLHEmergency = H11
         !CLHAmbulance = H12
         !CLHOutpatient = H13
         !CLHMonthly = H14
         !CLHMedical = H15
         !CLHPhone = H16
         !CLHLodger = H17
         !CLHAdmin = H18
         !CLHTax = H19
         !CLHOthers = H20
         !CLHCash = H21
         !CLHPaidByMember = TotalPaidByMemberHos
         !CLHPayableByMedix = TotalPaidMedixHos
         !CLHDiscount = TotalDiscountHos
         !CLHGross = TotalGrossHos
         !CLHTotal = TotalHospital
    End With
    ClaimHospital.Update

    ClaimHospital.Close
    Set ClaimHospital = Nothing
End Sub

Private Sub saveInvoice()
    Dim ClaimInvoice As New ADODB.Recordset
    Dim ForLoop, ListCount As Integer
    
    ListCount = lstInvoice.listitems.count
    
    ClaimInvoice.Open "ClaimInvoice", medixmasterconn, adOpenKeyset, adLockOptimistic

    With ClaimInvoice
            For ForLoop = 1 To ListCount
                .AddNew
                !HosCode = lstInvoice.listitems.Item(ForLoop).Text
                !INVNO = lstInvoice.listitems.Item(ForLoop).SubItems(1)
                !InvDate = lstInvoice.listitems.Item(ForLoop).SubItems(2)
                !INVAmount = IIf(IsNumeric(lstInvoice.listitems.Item(ForLoop).SubItems(3)), lstInvoice.listitems.Item(ForLoop).SubItems(3), 0)
                !PaidToHos = IIf(IsNumeric(lstInvoice.listitems.Item(ForLoop).SubItems(4)), lstInvoice.listitems.Item(ForLoop).SubItems(4), 0)
                !PaidtoMBM = IIf(IsNumeric(lstInvoice.listitems.Item(ForLoop).SubItems(5)), lstInvoice.listitems.Item(ForLoop).SubItems(5), 0)
                !PaidByMBM = IIf(IsNumeric(lstInvoice.listitems.Item(ForLoop).SubItems(6)), lstInvoice.listitems.Item(ForLoop).SubItems(6), 0)
                !INVRemark = lstInvoice.listitems.Item(ForLoop).SubItems(7)
            Next
    End With
    ClaimInvoice.UpdateBatch
    
    ClaimInvoice.Close
    Set ClaimInvoice = Nothing
End Sub

' Key Press
Private Sub cboHOSCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHosCode.Text, cboHosCode.SelStart) + Chr(KeyAscii) + Mid(cboHosCode.Text, cboHosCode.SelStart + cboHosCode.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboHosCode.ListCount - 1
        
        If NewText = LCase(Left(cboHosCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHosCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
            If ValidCount = 1 Then
                cboHosCode.Text = ValidValue
                cboHosCode.SelStart = 0
                cboHosCode.SelLength = Len(ValidValue)
            End If
        
        End If

End Sub



