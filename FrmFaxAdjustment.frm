VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFaxAdjustment 
   Caption         =   "Fax Received Adjustment"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleMode       =   0  'User
   ScaleWidth      =   32920.87
   Begin VB.Frame Frame6 
      Height          =   2655
      Left            =   120
      TabIndex        =   71
      Top             =   5400
      Width           =   12135
      Begin TabDlg.SSTab SSTab1 
         Height          =   2175
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   3836
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Fax Listing"
         TabPicture(0)   =   "FrmFaxAdjustment.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstFaxAdj"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "MBM Listing"
         TabPicture(1)   =   "FrmFaxAdjustment.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstMbm"
         Tab(1).ControlCount=   1
         Begin MSComctlLib.ListView lstFaxAdj 
            Height          =   1695
            Left            =   120
            TabIndex        =   80
            Top             =   360
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   2990
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   33
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Fax ref. No."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Doc Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Sender Category"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Sender name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Date Rec"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Time Rec"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Date Passout"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Time Passout"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Rec. By"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Hos"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Claim Assessor"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "DOA"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "DOD"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Bill Amount"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "Case No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "Membership No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Text            =   "Cover Id"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   17
               Text            =   "Reg By"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   18
               Text            =   "Reg Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   19
               Text            =   "Update By"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   20
               Text            =   "Update Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   21
               Text            =   "Reg Time"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   22
               Text            =   "Inpatient / Outpatient"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   23
               Text            =   "Patient name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   24
               Text            =   "Patient IC"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   25
               Text            =   "Policy No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   26
               Text            =   "payor"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   27
               Text            =   "Policy Status"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   28
               Text            =   "Exp. Status"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   29
               Text            =   "Eff date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   30
               Text            =   "Exp date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   31
               Text            =   "Comp"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   32
               Text            =   "Remarks"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lstMbm 
            Height          =   1695
            Left            =   -74880
            TabIndex        =   81
            Top             =   360
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   2990
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
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
               Text            =   "Mem No."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Coverid"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Inpatient / Outpatient"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Patient name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Patient IC"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Policy No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "payor"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Policy Status"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Exp. Status"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Eff date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Exp date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Comp"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   45
      TabIndex        =   1
      Top             =   480
      Width           =   12135
      Begin VB.Frame Frame2 
         Height          =   3855
         Left            =   7080
         TabIndex        =   43
         Top             =   120
         Width           =   4935
         Begin VB.TextBox txtMBMNoGo 
            Height          =   315
            Left            =   960
            TabIndex        =   68
            Top             =   210
            Width           =   1935
         End
         Begin VB.CommandButton cmdGo 
            Caption         =   "Go"
            Height          =   315
            Left            =   3840
            TabIndex        =   67
            Top             =   210
            Width           =   735
         End
         Begin VB.Frame Frame3 
            Caption         =   "Member Details"
            Height          =   3015
            Left            =   120
            TabIndex        =   45
            Top             =   600
            Width           =   4695
            Begin VB.TextBox txtPatientName 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   54
               Top             =   360
               Width           =   3255
            End
            Begin VB.TextBox txtPatientIc 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   53
               Top             =   675
               Width           =   1575
            End
            Begin VB.TextBox txtMemNo 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   52
               Top             =   990
               Width           =   2055
            End
            Begin VB.TextBox txtPolicyNo 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   51
               Top             =   1305
               Width           =   3255
            End
            Begin VB.TextBox txtPayor 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   1620
               Width           =   3255
            End
            Begin VB.TextBox txtPolicyStatus 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   1935
               Width           =   1170
            End
            Begin VB.TextBox txtExpStatus 
               Height          =   315
               Left            =   3285
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   1920
               Width           =   1170
            End
            Begin VB.TextBox txtcoverid 
               Height          =   315
               Left            =   3600
               MaxLength       =   2
               TabIndex        =   47
               Top             =   990
               Width           =   855
            End
            Begin VB.TextBox txtComp 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   2565
               Width           =   3255
            End
            Begin MSMask.MaskEdBox MaskEdBox7 
               Height          =   315
               Left            =   1200
               TabIndex        =   76
               Top             =   2250
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox8 
               Height          =   315
               Left            =   3285
               TabIndex        =   77
               Top             =   2250
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label13 
               Caption         =   "Patient Name"
               Height          =   315
               Left            =   120
               TabIndex        =   66
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label14 
               Caption         =   "Patient Ic"
               Height          =   315
               Left            =   120
               TabIndex        =   65
               Top             =   675
               Width           =   975
            End
            Begin VB.Label Label15 
               Caption         =   "Mem. No."
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   990
               Width           =   975
            End
            Begin VB.Label Label16 
               Caption         =   "Policy No"
               Height          =   315
               Left            =   120
               TabIndex        =   63
               Top             =   1305
               Width           =   975
            End
            Begin VB.Label Label17 
               Caption         =   "Payor"
               Height          =   315
               Left            =   120
               TabIndex        =   62
               Top             =   1620
               Width           =   975
            End
            Begin VB.Label Label18 
               Caption         =   "Policy Status"
               Height          =   315
               Left            =   120
               TabIndex        =   61
               Top             =   1935
               Width           =   975
            End
            Begin VB.Label Label22 
               Caption         =   "Exp. Status"
               Height          =   315
               Left            =   2420
               TabIndex        =   60
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label23 
               Caption         =   "Eff. Date"
               Height          =   315
               Left            =   120
               TabIndex        =   59
               Top             =   2250
               Width           =   1095
            End
            Begin VB.Label Label24 
               Caption         =   "Exp. date"
               Height          =   315
               Left            =   2420
               TabIndex        =   58
               Top             =   2250
               Width           =   975
            End
            Begin VB.Label Label26 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3360
               TabIndex        =   57
               Top             =   990
               Width           =   135
            End
            Begin VB.Label Label28 
               Caption         =   "Comp"
               Height          =   315
               Left            =   120
               TabIndex        =   56
               Top             =   2565
               Width           =   1095
            End
            Begin VB.Label lblINorOut 
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   3120
               TabIndex        =   55
               Top             =   705
               Width           =   1335
            End
         End
         Begin VB.TextBox txtcoveridgo 
            Height          =   315
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   44
            Top             =   210
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Mem No"
            Height          =   315
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   69
            Top             =   150
            Width           =   135
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4575
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   6855
         Begin VB.TextBox txtamount 
            Height          =   315
            Left            =   1800
            TabIndex        =   16
            Top             =   3075
            Width           =   1500
         End
         Begin VB.ComboBox cboClaimAssessor 
            Height          =   315
            Left            =   1800
            TabIndex        =   15
            Top             =   2445
            Width           =   4930
         End
         Begin VB.ComboBox cboAdmHos 
            Height          =   315
            Left            =   1800
            TabIndex        =   14
            Top             =   2130
            Width           =   4930
         End
         Begin VB.ComboBox cboRecBy 
            Height          =   315
            Left            =   1800
            TabIndex        =   13
            Top             =   1800
            Width           =   4930
         End
         Begin VB.ComboBox cboSenderCategory 
            Height          =   315
            Left            =   1800
            TabIndex        =   12
            Top             =   555
            Width           =   4930
         End
         Begin VB.ComboBox cboDocType 
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
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   4930
         End
         Begin VB.TextBox txtCaseNo 
            Height          =   315
            Left            =   5100
            TabIndex        =   10
            Top             =   3075
            Width           =   1260
         End
         Begin VB.TextBox txtRemarks 
            Height          =   615
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   3390
            Width           =   4930
         End
         Begin VB.ComboBox cbosendername 
            Height          =   315
            Left            =   1800
            TabIndex        =   8
            Top             =   870
            Width           =   4930
         End
         Begin VB.TextBox txtFaxRefNo 
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
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   4005
            Width           =   1500
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   315
            Left            =   1800
            TabIndex        =   17
            Top             =   1500
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   315
            Left            =   1800
            TabIndex        =   18
            Top             =   1185
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker DTPassoutDate 
            Height          =   315
            Left            =   6430
            TabIndex        =   19
            Top             =   1200
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66846721
            CurrentDate     =   40525
         End
         Begin MSComCtl2.DTPicker DTRecdate 
            Height          =   315
            Left            =   3360
            TabIndex        =   20
            Top             =   1200
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66846721
            CurrentDate     =   40525
         End
         Begin MSComCtl2.DTPicker DTDoa 
            Height          =   315
            Left            =   3360
            TabIndex        =   21
            Top             =   2760
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66846721
            CurrentDate     =   40525
         End
         Begin MSComCtl2.DTPicker DTDod 
            Height          =   315
            Left            =   6430
            TabIndex        =   22
            Top             =   2760
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66846721
            CurrentDate     =   40525
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   315
            Left            =   5100
            TabIndex        =   23
            Top             =   1200
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   315
            Left            =   1800
            TabIndex        =   24
            Top             =   2760
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   315
            Left            =   5100
            TabIndex        =   25
            Top             =   2760
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   315
            Left            =   5100
            TabIndex        =   26
            Top             =   1500
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label32 
            Caption         =   "24 Hrs"
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
            Left            =   6120
            TabIndex        =   75
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label31 
            Caption         =   "24 Hrs"
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
            Left            =   3000
            TabIndex        =   74
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Bill Amount"
            Height          =   315
            Left            =   120
            TabIndex        =   42
            Top             =   3075
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Time Passout"
            Height          =   315
            Left            =   3990
            TabIndex        =   41
            Top             =   1500
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Date Passout"
            Height          =   315
            Left            =   3990
            TabIndex        =   40
            Top             =   1185
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Claim Assessor"
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   2445
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Admit. Hospital"
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   2130
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Rec. By"
            Height          =   315
            Left            =   120
            TabIndex        =   37
            Top             =   1815
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Time"
            Height          =   315
            Left            =   120
            TabIndex        =   36
            Top             =   1500
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Date"
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   1185
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Sender Name"
            Height          =   315
            Left            =   120
            TabIndex        =   34
            Top             =   870
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Sender Category"
            Height          =   315
            Left            =   120
            TabIndex        =   33
            Top             =   555
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Doc. Type"
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label19 
            Caption         =   "DOA"
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "DOD"
            Height          =   315
            Left            =   3990
            TabIndex        =   30
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Case No"
            Height          =   315
            Left            =   3990
            TabIndex        =   29
            Top             =   3075
            Width           =   1095
         End
         Begin VB.Label Label25 
            Caption         =   "Remarks"
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   3390
            Width           =   1335
         End
         Begin VB.Label Label29 
            Caption         =   "Fax Ref. No."
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   4005
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   7080
         TabIndex        =   2
         Top             =   3960
         Width           =   4935
         Begin VB.CommandButton cmdcancel 
            Caption         =   "Cancel"
            Height          =   300
            Left            =   1020
            TabIndex        =   78
            Top             =   250
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   300
            Left            =   1920
            TabIndex        =   73
            Top             =   250
            Width           =   800
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            Height          =   300
            Left            =   1020
            TabIndex        =   72
            Top             =   250
            Width           =   800
         End
         Begin VB.CommandButton cmdReg 
            Caption         =   "Search"
            Height          =   300
            Left            =   120
            TabIndex        =   5
            Top             =   250
            Width           =   800
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   300
            Left            =   2920
            TabIndex        =   4
            Top             =   250
            Width           =   800
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   300
            Left            =   3820
            TabIndex        =   3
            Top             =   250
            Width           =   800
         End
      End
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Fax Received Adjustment"
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
      Width           =   12495
   End
End
Attribute VB_Name = "FrmFaxAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpTime As String
Dim tmpDateTime As String
Dim tmpTimeformat As String
Dim tmpTimeonly As String
Dim tmpDateonly As String
Dim tmpWeekDay As String
Dim tmp24time As String


Private Sub ServerTime()
    Dim lpbuffer As Long
    Dim t_struct As TIME_OF_DAY_INFO
    Dim ret As Long
    Dim bServer() As Byte
    
    If Trim(SvrDBPath) = "" Then
        'Local machine
        ret = NetRemoteTOD(vbNullString, lpbuffer)
    Else
        'Check the syntax of the ServerName string
        If InStr(SvrDBPath, "\\") = 1 Then
            bServer = SvrDBPath & vbNullChar
        Else
            bServer = SvrDBPath & vbNullChar
        End If
        ret = NetRemoteTOD(bServer(0), lpbuffer)
    End If
    CopyMem t_struct, ByVal lpbuffer, Len(t_struct)
    If lpbuffer Then
        Call NetApiBufferFree(lpbuffer)
    End If
    
    tmpDateTime = DateSerial(t_struct.tod_year, t_struct.tod_month, t_struct.tod_day) + TimeSerial(t_struct.tod_hours, t_struct.tod_mins - t_struct.tod_timezone, t_struct.tod_secs)
    tmpTime = Format(tmpDateTime, "hh:mm")
    tmpTimeonly = Format(tmpTime, "hh:mm")
    tmpTimeformat = Format(tmpTime, "AM/PM")
    tmpDateonly = Format(tmpDateTime, "dd/mm/yyyy")
    tmpWeekDay = WeekdayName(Weekday(tmpDateonly))
    'tmpDateTime = "16/12/2010 10:40:40 PM"
    tmp24time = Format(tmpDateTime, "hmm")
End Sub


Private Sub ClearEntries()
Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Clear
        'ElseIf TypeOf ctl Is Label Then
        '    ctl.Caption = ""
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            'ctl.mask = "##/##/####"
        End If
    Next ctl
    
lstFaxAdj.ListItems.Clear
lblINorOut.Caption = ""
    
End Sub

Private Sub ClearEntries2()
Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        'ElseIf TypeOf ctl Is ListView Then
        '    ctl.ListItems.Clear
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Clear
        'ElseIf TypeOf ctl Is Label Then
        '    ctl.Caption = ""
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            'ctl.mask = "##/##/####"
        End If
    Next ctl
    
'lstFaxAdj.ListItems.Clear
lblINorOut.Caption = ""
    
End Sub

Private Sub EnabledEntries()
Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is ListView Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = False
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Enabled = False
        End If
    Next ctl
    
    DTRecdate.Enabled = False
    DTPassoutDate.Enabled = False
    DTDoa.Enabled = False
    DTDod.Enabled = False
    
End Sub
Private Sub UEnabledEntries()
Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is ListView Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Enabled = True
        End If
    Next ctl
    
    DTRecdate.Enabled = True
    DTPassoutDate.Enabled = True
    DTDoa.Enabled = True
    DTDod.Enabled = True
    
End Sub


Private Sub cboAdmHos_Change()
cboAdmHos = ChkStr(cboAdmHos)
End Sub

Private Sub cboAdmHos_LostFocus()
cboAdmHos = ChkStr(cboAdmHos)
End Sub

Private Sub cboClaimAssessor_Change()
cboClaimAssessor = ChkStr(cboClaimAssessor)
End Sub

Private Sub cboClaimAssessor_LostFocus()
cboClaimAssessor = ChkStr(cboClaimAssessor)
End Sub

Private Sub cboDocType_Change()
cboDocType = ChkStr(cboDocType)
End Sub

Private Sub cboDocType_LostFocus()
cboDocType = ChkStr(cboDocType)
End Sub

Private Sub cboRecBy_Change()
cboRecBy = ChkStr(cboRecBy)
End Sub

Private Sub cboRecBy_LostFocus()
cboRecBy = ChkStr(cboRecBy)
End Sub

Private Sub cboSenderCategory_Change()
cboSenderCategory = ChkStr(cboSenderCategory)
End Sub

Private Sub cboSenderCategory_LostFocus()
cboSenderCategory = ChkStr(cboSenderCategory)
End Sub

Private Sub cbosendername_Change()
cbosendername = ChkStr(cbosendername)
End Sub

Private Sub cbosendername_LostFocus()
cbosendername = ChkStr(cbosendername)
End Sub

Private Sub cmdcancel_Click()
    Call ClearEntries2
    Call ComboListing
    Call avalue
    lstMbm.ListItems.Clear
    Frame2.Enabled = False
    Frame4.Enabled = False

    'Frame6.Enabled = True
    lstFaxAdj.Enabled = True
    lstMbm.Enabled = False

    cmdReg.Enabled = False
    cmdEdit.Enabled = True
    cmdEdit.Visible = True
    cmdcancel.Visible = False
    cmdSave.Enabled = False
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    
        SSTab1.Tab = 0
End Sub

Private Sub CmdClear_Click()
Call UEnabledEntries
 cmdReg.Enabled = True
 Call ClearEntries
 Call ComboListing
 Call avalue
Frame2.Enabled = True
Frame4.Enabled = True

'Frame6.Enabled = False
    lstFaxAdj.Enabled = False
    lstMbm.Enabled = False
    
        SSTab1.Tab = 0

    cmdReg.Enabled = True
    cmdEdit.Enabled = False
    cmdEdit.Visible = True
    cmdcancel.Visible = False
    cmdSave.Enabled = False
    cmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
Private Sub avalue()
    MaskEdBox1.mask = "##/##/####"
    MaskEdBox2.mask = "##/##/####"
    MaskEdBox3.mask = "##/##/####"
    MaskEdBox4.mask = "##/##/####"
    
    MaskEdBox5.mask = "##:##"
    MaskEdBox6.mask = "##:##"
    
    MaskEdBox7.mask = "##/##/####"
    MaskEdBox8.mask = "##/##/####"
    
    'MaskEdBox1.Text = tmpDateonly
    'MaskEdBox5.Text = tmpTimeonly
    'txtRegBy.Text = Trim(Logon)
End Sub

Private Sub cmdEdit_Click()
    cmdReg.Enabled = False
    cmdEdit.Enabled = False
    cmdEdit.Visible = False
    cmdcancel.Visible = True
    cmdSave.Enabled = True
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    
    Frame2.Enabled = True
    Frame4.Enabled = True
    
    'Frame6.Enabled = False
        lstFaxAdj.Enabled = False
    lstMbm.Enabled = False
    
        SSTab1.Tab = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set FrmFaxEnq = Nothing
End Sub

Private Sub cmdGo_Click()
Dim mbminfo As New ADODB.Recordset
lstMbm.ListItems.Clear

    txtMemNo.Text = ""
    txtcoverid.Text = ""
    txtPatientName.Text = ""
    txtPatientIc.Text = ""
    txtPolicyNo.Text = ""
    txtPayor.Text = ""
    txtPolicyStatus.Text = ""
    txtExpStatus.Text = ""
    MaskEdBox7.mask = "##/##/####"
    MaskEdBox8.mask = "##/##/####"
    lblINorOut.Caption = ""
    txtComp.Text = ""

    If Not medixmasterconn Is Nothing And Not medixmasterconn = "" Then
        medixmasterconn.Close
    End If
    
  
    If Trim(txtMBMNoGo) = "" Then
        MsgBox "Please select a Member Number ! ", vbCritical
        txtMBMNoGo.SetFocus
    ElseIf Trim(txtcoveridgo) = "" Then
        MsgBox "Please select a Member cover id ! ", vbCritical
        txtcoveridgo.SetFocus
    Else
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Screen.MousePointer = vbHourglass
        Call MBMInquiryList
        Screen.MousePointer = vbDefault
        medixmasterconn.Close
        Set medixmasterconn = Nothing
    End If
End Sub

Private Sub MBMInquiryList()

Dim rst1 As New ADODB.Recordset
Dim Record As ListItem
Dim strCriteria As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim countrecord As Integer
countrecord = 0
'strCriteria = ""
'strCriteria = strCriteria & " AND MBMTopupNumber =  '" & RemStr(tmpTopupNumber) & "' order by MBMPayorEffDate"
        
Set cmd.ActiveConnection = medixmasterconn
cmd.CommandText = "Search_fax_MBM2"
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 300
Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, Trim(txtMBMNoGo.Text))
cmd.Parameters.Append Prm1
Set Prm2 = cmd.CreateParameter("coverid", adVarChar, adParamInput, 2, Trim(txtcoveridgo.Text))
cmd.Parameters.Append Prm2
rst1.CursorLocation = adUseClient
rst1.CursorType = adOpenForwardOnly
rst1.CacheSize = 100
Set rst1 = cmd.Execute

    
    Do Until rst1.EOF
        With rst1
            countrecord = countrecord + 1
            Set Record = lstMbm.ListItems.Add(, , !MBMNumber)
            Record.SubItems(1) = IIf(Len(!CoverID), !CoverID, "")
            Record.SubItems(2) = IIf(Len(!status), Trim(!status), "")
            Record.SubItems(3) = IIf(Len(!MBMName), Trim(!MBMName), "")
            Record.SubItems(4) = IIf(Len(!MBMIcBcPp), Trim(!MBMIcBcPp), "")
            
            Record.SubItems(5) = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
            Record.SubItems(6) = IIf(Len(!PAYCompanyName), Trim(!PAYCompanyName), "")
            Record.SubItems(7) = IIf(Len(!MBMStatus), Trim(!MBMStatus), "")
            Record.SubItems(8) = IIf(Len(!MBMExpiredStatus), Trim(!MBMExpiredStatus), "")
            
            If Len(!MBMPayorEffDate) Then
                Record.SubItems(9) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
            End If
            If Len(!MBMPayorEXPDate) Then
                Record.SubItems(10) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
            End If
            
            Record.SubItems(11) = IIf(Len(!MBMGroupCompany), Trim(!MBMGroupCompany), "")

        rst1.MoveNext
        End With
    Loop
    rst1.Close
    Set rst1 = Nothing
    
    If countrecord = 0 Then
        MsgBox "No Record Found ! ", vbCritical
    End If
    
    

     If lstMbm.ListItems.Count >= 1 Then
        lstFaxAdj.Enabled = False
        lstMbm.Enabled = True
        SSTab1.Tab = 1
     End If
    
    If lstMbm.ListItems.Count = 1 Then
      
        txtMemNo.Text = lstMbm.SelectedItem.Text
        txtcoverid.Text = lstMbm.SelectedItem.SubItems(1)
        lblINorOut.Caption = lstMbm.SelectedItem.SubItems(2)
        txtPatientName.Text = lstMbm.SelectedItem.SubItems(3)
        txtPatientIc.Text = lstMbm.SelectedItem.SubItems(4)
        txtPolicyNo.Text = lstMbm.SelectedItem.SubItems(5)
        txtPayor.Text = lstMbm.SelectedItem.SubItems(6)
        txtPolicyStatus.Text = lstMbm.SelectedItem.SubItems(7)
        txtExpStatus.Text = lstMbm.SelectedItem.SubItems(8)
        MaskEdBox7.Text = lstMbm.SelectedItem.SubItems(9)
        MaskEdBox8.Text = lstMbm.SelectedItem.SubItems(10)
        txtComp.Text = lstMbm.SelectedItem.SubItems(11)

    End If
            
End Sub

Private Sub cmdReg_Click()
    SSTab1.Tab = 0

    If Not medixmasterconnMaint Is Nothing And Not medixmasterconnMaint = "" Then
        medixmasterconnMaint.Close
    End If
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Call SearchRecord
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
End Sub

Private Sub ComboListing()
Dim DocType As New ADODB.Recordset
Dim senderCategory As New ADODB.Recordset
Dim RecBy As New ADODB.Recordset
Dim HosName As New ADODB.Recordset
Dim ClaimAssessor As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim SenderName As New ADODB.Recordset


    
    If Not medixmasterconn Is Nothing And Not medixmasterconn = "" Then
        medixmasterconn.Close
    End If
    
    If Not medixmasterconnMaint Is Nothing And Not medixmasterconnMaint = "" Then
        medixmasterconnMaint.Close
    End If
    
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    'DocType
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchFaxDocType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    DocType.CursorLocation = adUseClient
    DocType.CursorType = adOpenForwardOnly
    DocType.CacheSize = 100
    Set DocType = cmd.Execute
 
    Do Until DocType.EOF
        With DocType
            'cboDocType.AddItem Trim(!DocType) & " - " & Trim(!DocTypeDesc)
            cboDocType.AddItem Trim(!DocType) '& " - " & Trim(!DocTypeDesc)
            DocType.MoveNext
        End With
    Loop
    DocType.Close
    Set DocType = Nothing
    
    'senderCategory
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchFaxSenderCategory"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    senderCategory.CursorLocation = adUseClient
    senderCategory.CursorType = adOpenForwardOnly
    senderCategory.CacheSize = 100
    Set senderCategory = cmd.Execute
 
    Do Until senderCategory.EOF
        With senderCategory
            'cboSenderCategory.AddItem Trim(!senderCategory) & " - " & Trim(!SenderCategoryDesc)
            cboSenderCategory.AddItem Trim(!senderCategory) '& " - " & Trim(!SenderCategoryDesc)
            senderCategory.MoveNext
        End With
    Loop
    senderCategory.Close
    Set senderCategory = Nothing
    
    
        'RecBy
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "FaxSearchUSR"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    RecBy.CursorLocation = adUseClient
    RecBy.CursorType = adOpenForwardOnly
    RecBy.CacheSize = 100
    Set RecBy = cmd.Execute
 
    Do Until RecBy.EOF
        With RecBy
            'cboRecBy.AddItem Trim(!USRCode) & " - " & Trim(!USRName)
            cboRecBy.AddItem Trim(!USRCode) '& " - " & Trim(!USRName)
            RecBy.MoveNext
        End With
    Loop
    RecBy.Close
    Set RecBy = Nothing
    
    'cbosendername
    
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "FaxSearchHOSPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    SenderName.CursorLocation = adUseClient
    SenderName.CursorType = adOpenForwardOnly
    SenderName.CacheSize = 100
    Set SenderName = cmd.Execute
 
    Do Until SenderName.EOF
        With SenderName
            'cboClaimAssessor.AddItem Trim(!USRCode) & " - " & Trim(!USRName)
            cbosendername.AddItem Trim(!FullName) '& " - " & Trim(!Code)
            SenderName.MoveNext
        End With
    Loop
    SenderName.Close
    Set SenderName = Nothing
    
                'HosName
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "FaxSearchHOS"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HosName.CursorLocation = adUseClient
    HosName.CursorType = adOpenForwardOnly
    HosName.CacheSize = 100
    Set HosName = cmd.Execute
 
    Do Until HosName.EOF
        With HosName
            'cboClaimAssessor.AddItem Trim(!USRCode) & " - " & Trim(!USRName)
            cboAdmHos.AddItem Trim(!HosName) '& " - " & Trim(!HOSCode)
            HosName.MoveNext
        End With
    Loop
    HosName.Close
    Set HosName = Nothing
    
    
            'RecBy
    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "FaxSearchUSR"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    ClaimAssessor.CursorLocation = adUseClient
    ClaimAssessor.CursorType = adOpenForwardOnly
    ClaimAssessor.CacheSize = 100
    Set ClaimAssessor = cmd.Execute
 
    Do Until ClaimAssessor.EOF
        With ClaimAssessor
            'cboClaimAssessor.AddItem Trim(!USRCode) & " - " & Trim(!USRName)
            cboClaimAssessor.AddItem Trim(!USRCode) '& " - " & Trim(!USRName)
            ClaimAssessor.MoveNext
        End With
    Loop
    ClaimAssessor.Close
    Set ClaimAssessor = Nothing
    
    
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdSave_Click()
Dim Updatesql As String
Dim RecordSaved
Updatesql = ""
Dim fax As New ADODB.Recordset

Dim tmpDateRec As String
Dim tmpDatePassOut As String
Dim tmpdoa As String
Dim tmpdod As String
Dim Effdate As String
Dim ExpDate As String

    If Not medixmasterconnMaint Is Nothing And Not medixmasterconnMaint = "" Then
        medixmasterconnMaint.Close
    End If
    
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then
    
        If MaskEdBox1.Text = "__/__/____" Then
            tmpDateRec = ""
        Else
            tmpDateRec = Format(MaskEdBox1.Text, "mm/dd/yyyy")
        End If
        
        If MaskEdBox2.Text = "__/__/____" Then
            tmpDatePassOut = ""
        Else
            tmpDatePassOut = Format(MaskEdBox2.Text, "mm/dd/yyyy")
        End If
        
        If MaskEdBox3.Text = "__/__/____" Then
            tmpdoa = ""
        Else
            tmpdoa = Format(MaskEdBox3.Text, "mm/dd/yyyy")
        End If
        
        If MaskEdBox4.Text = "__/__/____" Then
            tmpdod = ""
        Else
            tmpdod = Format(MaskEdBox4.Text, "mm/dd/yyyy")
        End If
        
        If MaskEdBox7.Text = "__/__/____" Then
            Effdate = ""
        Else
            Effdate = Format(MaskEdBox7.Text, "mm/dd/yyyy")
        End If
        
        If MaskEdBox8.Text = "__/__/____" Then
            ExpDate = ""
        Else
            ExpDate = Format(MaskEdBox8.Text, "mm/dd/yyyy")
        End If
     
        Updatesql = "Select * FROM FaxDetails Where FaxRefNo ='" & Trim(txtFaxRefNo) & "'"
        fax.Open Updatesql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        With fax
        
               
            !DocType = ChkStr(Trim(cboDocType.Text))
            !senderCategory = ChkStr(Trim(cboSenderCategory.Text))
            !SenderName = ChkStr(Trim(cbosendername.Text))
            !DateREC = tmpDateRec
            !TimeRec = MaskEdBox5.Text
            If Trim(tmpDatePassOut) <> "" Then
                !DatePassOut = tmpDatePassOut
            End If
            !TimePassOut = MaskEdBox6.Text
            !HOS = ChkStr(Trim(cboAdmHos.Text))
            !ClaimAssessor = ChkStr(Trim(cboClaimAssessor.Text))
            If Trim(tmpdoa) <> "" Then
                !DOA = tmpdoa
            End If
            If Trim(tmpdod) <> "" Then
                !DOD = tmpdod
            End If
            If Trim(txtamount.Text) <> "" Then
                !BillAmount = txtamount.Text
            End If
            '!BillAmount = Trim(txtamount.Text) '''
            !CaseNo = ChkStr(Trim(txtCaseNo.Text))
            !Remarks = ChkStr(Trim(txtRemarks.Text))
            !MBMNo = ChkStr(Trim(txtMemNo.Text))
            !CoverID = Trim(txtcoverid.Text)
            !UpdateBy = ChkStr(Logon)
            !UpdateDate = Format(tmpDateonly, "mm/dd/yyyy")
            !RecBy = ChkStr(Trim(cboRecBy.Text))
            !InorOut = Trim(lblINorOut)
            !PatientIc = ChkStr(Trim(txtPatientIc.Text))
            !PatientName = ChkStr(Trim(txtPatientName.Text))
            !Policyno = ChkStr(Trim(txtPolicyNo.Text))
            !Payor = ChkStr(Trim(txtPayor.Text))
            !comp = ChkStr(Trim(txtComp.Text))
            !PolicyStatus = ChkStr(Trim(txtPolicyStatus.Text))
            !ExpStatus = ChkStr(Trim(txtExpStatus.Text))
            If Trim(Effdate) <> "" Then
                !Effdate = Effdate
            End If
            If Trim(ExpDate) <> "" Then
                !ExpDate = ExpDate
            End If
        
        End With
        fax.Update
        fax.Close
        Set fax = Nothing
        MsgBox "Record Updated Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
        
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
        Call UEnabledEntries
        cmdReg.Enabled = True
        Call ClearEntries
        Call ComboListing
        Call avalue
        Frame2.Enabled = True
        Frame4.Enabled = True

        'Frame6.Enabled = False
        lstFaxAdj.Enabled = False
        lstMbm.Enabled = False
    
        SSTab1.Tab = 0

        cmdReg.Enabled = True
        cmdEdit.Enabled = False
        cmdEdit.Visible = True
        cmdcancel.Visible = False
        cmdSave.Enabled = False
        cmdClear.Enabled = True
        cmdExit.Enabled = True
    Else
        Exit Sub
    End If

    
   
End Sub

Private Sub DTDoa_Change()
MaskEdBox3.Text = DTDoa.Value
End Sub

Private Sub DTDod_Change()
MaskEdBox4.Text = DTDod.Value
End Sub

Private Sub DTPassoutDate_Change()
MaskEdBox2.Text = DTPassoutDate.Value
End Sub

Private Sub DTPassoutDate_Click()
MaskEdBox2.Text = DTPassoutDate.Value
End Sub

Private Sub DTRecdate_Change()
MaskEdBox1.Text = DTRecdate.Value
End Sub

Private Sub DTRecdate_Click()
MaskEdBox1.Text = DTRecdate.Value
End Sub

Private Sub Form_Load()
    Call UEnabledEntries
    Call ServerTime
    Call ClearEntries
    Call ComboListing
    Call avalue
    Frame2.Enabled = True
    Frame4.Enabled = True
    
    lstFaxAdj.Enabled = False
    lstMbm.Enabled = False
    
    SSTab1.Tab = 0
    
    cmdReg.Enabled = True
    cmdEdit.Enabled = False
    cmdEdit.Visible = True
    cmdcancel.Visible = False
    cmdSave.Enabled = False
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    
End Sub

Private Function SearchRecord()
Dim strAppend As String

    If Len(Trim(cboDocType.Text)) Then
        strAppend = strAppend + " AND DocType = '" & ChkStr(Trim(cboDocType.Text)) & "'"
    End If
    
    If Len(Trim(cboSenderCategory.Text)) Then
        strAppend = strAppend + " AND SenderCategory = '" & ChkStr(Trim(cboSenderCategory.Text)) & "'"
    End If
    
    If Len(Trim(cbosendername.Text)) Then
        strAppend = strAppend + " AND SenderName = '" & ChkStr(Trim(cbosendername.Text)) & "'"
    End If
    
    If Len(Trim(MaskEdBox1.Text)) And Trim(MaskEdBox1.Text) <> "__/__/____" Then
        strAppend = strAppend + " AND DateRec = '" & Format(MaskEdBox1.Text, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(MaskEdBox2.Text)) And Trim(MaskEdBox2.Text) <> "__/__/____" Then
        strAppend = strAppend + " AND DatePassOut = '" & Format(MaskEdBox2.Text, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(MaskEdBox5.Text)) And Trim(MaskEdBox5.Text) <> "__:__" Then
        strAppend = strAppend + " AND TimeRec = '" & Trim(MaskEdBox5.Text) & "'"
    End If
    
    If Len(Trim(MaskEdBox6.Text)) And Trim(MaskEdBox6.Text) <> "__:__" Then
        strAppend = strAppend + " AND TimePassOut = '" & Trim(MaskEdBox6.Text) & "'"
    End If
    
    
    If Len(Trim(cboRecBy.Text)) Then
        strAppend = strAppend + " AND RecBy = '" & ChkStr(Trim(cboRecBy.Text)) & "'"
    End If
    
    If Len(Trim(cboAdmHos.Text)) Then
        strAppend = strAppend + " AND Hos = '" & ChkStr(Trim(cboAdmHos.Text)) & "'"
    End If
    
    If Len(Trim(cboClaimAssessor.Text)) Then
        strAppend = strAppend + " AND ClaimAssessor = '" & ChkStr(Trim(cboClaimAssessor.Text)) & "'"
    End If
    
    If Len(Trim(MaskEdBox3.Text)) And Trim(MaskEdBox3.Text) <> "__/__/____" Then
        strAppend = strAppend + " AND DOA = '" & Format(MaskEdBox3.Text, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(MaskEdBox4.Text)) And Trim(MaskEdBox4.Text) <> "__/__/____" Then
        strAppend = strAppend + " AND DOD = '" & Format(MaskEdBox4.Text, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtamount.Text)) Then
        strAppend = strAppend + " AND BillAmount = '" & Trim(txtamount.Text) & "'"
    End If
    
    If Len(Trim(txtCaseNo.Text)) Then
        strAppend = strAppend + " AND CaseNo = '" & ChkStr(Trim(txtCaseNo.Text)) & "'"
    End If
    
    If Len(Trim(txtRemarks.Text)) Then
        strAppend = strAppend + " AND Remarks = '" & ChkStr(Trim(txtRemarks.Text)) & "'"
    End If
    
    If Len(Trim(txtFaxRefNo.Text)) Then
        strAppend = strAppend + " AND FaxRefNo = '" & ChkStr(Trim(txtFaxRefNo.Text)) & "'"
    End If
    
   '''
    If Len(Trim(txtPatientName.Text)) Then
        strAppend = strAppend + " AND Patientname = '" & ChkStr(Trim(txtPatientName.Text)) & "'"
    End If
    
    If Len(Trim(txtPatientIc.Text)) Then
        strAppend = strAppend + " AND PatientIc = '" & ChkStr(Trim(txtPatientIc.Text)) & "'"
    End If
    
    If Len(Trim(txtMemNo.Text)) Then
        strAppend = strAppend + " AND MbmNo = '" & ChkStr(Trim(txtMemNo.Text)) & "'"
    End If

    If Len(Trim(txtcoverid.Text)) Then
        strAppend = strAppend + " AND CoverId = '" & Trim(txtcoverid.Text) & "'"
    End If
    
    If Len(Trim(txtPolicyNo.Text)) Then
        strAppend = strAppend + " AND PolicyNo = '" & ChkStr(Trim(txtPolicyNo.Text)) & "'"
    End If
    
    If Len(Trim(txtPayor.Text)) Then
        strAppend = strAppend + " AND Payor = '" & ChkStr(Trim(txtPayor.Text)) & "'"
    End If
    
    If Len(Trim(txtPolicyStatus.Text)) Then
        strAppend = strAppend + " AND PolicyStatus = '" & ChkStr(Trim(txtPolicyStatus.Text)) & "'"
    End If
    
    If Len(Trim(txtExpStatus.Text)) Then
        strAppend = strAppend + " AND ExpStatus = '" & ChkStr(Trim(txtExpStatus.Text)) & "'"
    End If
    
    If Len(Trim(txtComp.Text)) Then
        strAppend = strAppend + " AND comp = '" & ChkStr(Trim(txtComp.Text)) & "'"
    End If
    
    If Len(Trim(MaskEdBox7.Text)) And Trim(MaskEdBox7.Text) <> "__/__/____" Then
        strAppend = strAppend + " AND Effdate = '" & Format(MaskEdBox7.Text, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(MaskEdBox8.Text)) And Trim(MaskEdBox8.Text) <> "__/__/____" Then
        strAppend = strAppend + " AND ExpDate = '" & Format(MaskEdBox8.Text, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(strAppend)) Then
        Call FaxTrackingListing(strAppend)
    Else
        MsgBox "No Record Found ! ", vbCritical
    End If
    
End Function

Private Sub FaxTrackingListing(strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String

lstFaxAdj.ListItems.Clear

       sql = " SELECT [FaxRefNo],[DocType],[SenderCategory],[SenderName],[DateRec],[TimeRec], " & _
       " [DatePassOut],[TimePassOut],[Hos], " & _
       " [ClaimAssessor] , [DOA], [DOD], [BillAmount], [CaseNo], [Remarks], [MBMNo], " & _
       " [CoverID] , [RegBy], [RegDate], [UpdateBy], [UpdateDate], [RecBy], [RegTime], " & _
       " [InorOut],[PatientIc],[Patientname],[PolicyNo],[Payor],[comp], " & _
       " [PolicyStatus] , [ExpStatus], [Effdate], [ExpDate] " & _
       " From [HISMaintenance].[dbo].[FaxDetails] where  0=0 "
       
       sql = sql + strAppend
       
       rst2.Open sql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
       
        If rst2.EOF = True Then
    
            Frame2.Enabled = True
            Frame4.Enabled = True
            
            'Frame6.Enabled = False
            lstFaxAdj.Enabled = False
            lstMbm.Enabled = False
            MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
                Frame2.Enabled = False
                Frame4.Enabled = False
                
                'Frame6.Enabled = True
                lstFaxAdj.Enabled = True
                lstMbm.Enabled = False
                
            'Call EnabledEntries
            'lstFaxAdj.Enabled = True
            cmdReg.Enabled = False
            Do Until rst2.EOF
                With rst2
                    'countrecord = countrecord + 1
                    Set Record = lstFaxAdj.ListItems.Add(, , !FaxRefNo)
                    Record.SubItems(1) = IIf(Len(!DocType), !DocType, "")
                    Record.SubItems(2) = IIf(Len(!senderCategory), Trim(!senderCategory), "")
                    Record.SubItems(3) = IIf(Len(!SenderName), Trim(!SenderName), "")
                    Record.SubItems(4) = IIf(Len(!DateREC), Trim(!DateREC), "")
        
                    Record.SubItems(5) = IIf(Len(!TimeRec), !TimeRec, "")
                    Record.SubItems(6) = IIf(Len(!DatePassOut), Trim(!DatePassOut), "")
                    Record.SubItems(7) = IIf(Len(!TimePassOut), Trim(!TimePassOut), "")
                    Record.SubItems(8) = IIf(Len(!RecBy), Trim(!RecBy), "")
                    Record.SubItems(9) = IIf(Len(!HOS), Trim(!HOS), "")
                    
                    Record.SubItems(10) = IIf(Len(!ClaimAssessor), !ClaimAssessor, "")
                    Record.SubItems(11) = IIf(Len(!DOA), Trim(!DOA), "")
                    Record.SubItems(12) = IIf(Len(!DOD), Trim(!DOD), "")
                    Record.SubItems(13) = IIf(Len(!BillAmount), Trim(!BillAmount), "")
            
                    Record.SubItems(14) = IIf(Len(!CaseNo), !CaseNo, "")
                    Record.SubItems(15) = IIf(Len(!MBMNo), Trim(!MBMNo), "")
                    Record.SubItems(16) = IIf(Len(!CoverID), Trim(!CoverID), "")
                    Record.SubItems(17) = IIf(Len(!RegBy), Trim(!RegBy), "")
                    
                    Record.SubItems(18) = IIf(Len(!RegDate), Trim(!RegDate), "")
                    Record.SubItems(19) = IIf(Len(!UpdateBy), !UpdateBy, "")
                    Record.SubItems(20) = IIf(Len(!UpdateDate), Trim(!UpdateDate), "")
                    Record.SubItems(21) = IIf(Len(!RegTime), Trim(!RegTime), "")
                    Record.SubItems(22) = IIf(Len(!InorOut), Trim(!InorOut), "")
                    
                    Record.SubItems(23) = IIf(Len(!PatientName), !PatientName, "")
                    Record.SubItems(24) = IIf(Len(!PatientIc), Trim(!PatientIc), "")
                    Record.SubItems(25) = IIf(Len(!Policyno), Trim(!Policyno), "")
                    Record.SubItems(26) = IIf(Len(!Payor), Trim(!Payor), "")
                    
                    Record.SubItems(27) = IIf(Len(!PolicyStatus), !PolicyStatus, "")
                    Record.SubItems(28) = IIf(Len(!ExpStatus), Trim(!ExpStatus), "")
                    Record.SubItems(29) = IIf(Len(!Effdate), Trim(!Effdate), "")
                    Record.SubItems(30) = IIf(Len(!ExpDate), Trim(!ExpDate), "")
                    Record.SubItems(31) = IIf(Len(!comp), Trim(!comp), "")
                    Record.SubItems(32) = IIf(Len(!Remarks), Trim(!Remarks), "")
                           
               '     If Len(!MBMPayorEffDate) Then
               '         Record.SubItems(9) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
               '     End If
               '     If Len(!MBMPayorEXPDate) Then
               '         Record.SubItems(10) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
               '     End If
            
               '     Record.SubItems(11) = IIf(Len(!MBMGroupCompany), Trim(!MBMGroupCompany), "")

                    rst2.MoveNext
                End With
            Loop
            rst2.Close
            Set rst2 = Nothing
        End If
    

End Sub

Private Sub lstFaxAdj_Click()
'Call UEnabledEntries
 'cmdReg.Enabled = True
 Call ClearEntries2
 Call ComboListing
 Call avalue
 
If lstFaxAdj.ListItems.Count > 0 Then

Frame2.Enabled = False
Frame4.Enabled = False

'Frame6.Enabled = True
    lstFaxAdj.Enabled = True
    lstMbm.Enabled = True

    cmdReg.Enabled = False
    cmdEdit.Enabled = True
    cmdEdit.Visible = True
    cmdcancel.Visible = False
    cmdSave.Enabled = False
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    
    txtFaxRefNo.Text = lstFaxAdj.SelectedItem.Text
    cboDocType.Text = lstFaxAdj.SelectedItem.SubItems(1)
    cboSenderCategory.Text = lstFaxAdj.SelectedItem.SubItems(2)
    cbosendername.Text = lstFaxAdj.SelectedItem.SubItems(3)
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(4)) <> "" Then
        MaskEdBox1.Text = lstFaxAdj.SelectedItem.SubItems(4)
    End If
    
    'MaskEdBox1.Text = lstFaxAdj.SelectedItem.SubItems(4)
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(5)) <> "" Then
        'MaskEdBox5.Text = Mid(lstFaxAdj.SelectedItem.SubItems(5), 1, 5)
        MaskEdBox5.Text = lstFaxAdj.SelectedItem.SubItems(5)
        'cborectime.Text = Mid(lstFaxAdj.SelectedItem.SubItems(5), 6, 2)
    End If
    

    If Trim(lstFaxAdj.SelectedItem.SubItems(6)) <> "" Then
        MaskEdBox2.Text = lstFaxAdj.SelectedItem.SubItems(6)
    End If
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(7)) <> "" Then
        'MaskEdBox6.Text = Mid(lstFaxAdj.SelectedItem.SubItems(7), 1, 5)
        MaskEdBox6.Text = lstFaxAdj.SelectedItem.SubItems(7)
        'cbopassouttime.Text = Mid(lstFaxAdj.SelectedItem.SubItems(7), 6, 2)
    End If
   
    
    cboRecBy.Text = lstFaxAdj.SelectedItem.SubItems(8)
    
    cboAdmHos.Text = lstFaxAdj.SelectedItem.SubItems(9)
    cboClaimAssessor.Text = lstFaxAdj.SelectedItem.SubItems(10)
    If Trim(lstFaxAdj.SelectedItem.SubItems(11)) <> "" Then
        MaskEdBox3.Text = lstFaxAdj.SelectedItem.SubItems(11)
    End If
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(12)) <> "" Then
        MaskEdBox4.Text = lstFaxAdj.SelectedItem.SubItems(12)
    End If
    'MaskEdBox3.Text = lstFaxAdj.SelectedItem.SubItems(11)
    'MaskEdBox4.Text = lstFaxAdj.SelectedItem.SubItems(12)
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(13)) = "" Then
        txtamount.Text = ""
    ElseIf IsNumeric(Trim(lstFaxAdj.SelectedItem.SubItems(13))) = True Then
        txtamount.Text = Format(Trim(lstFaxAdj.SelectedItem.SubItems(13)), "#,##0.00")
    Else
        txtamount.Text = ""
    End If
    
    txtCaseNo.Text = lstFaxAdj.SelectedItem.SubItems(14)
    
    txtRemarks.Text = lstFaxAdj.SelectedItem.SubItems(32)
    
 
    
    txtMemNo.Text = lstFaxAdj.SelectedItem.SubItems(15)
    txtcoverid.Text = lstFaxAdj.SelectedItem.SubItems(16)
    lblINorOut.Caption = lstFaxAdj.SelectedItem.SubItems(22)
    txtPatientName.Text = lstFaxAdj.SelectedItem.SubItems(23)
    txtPatientIc.Text = lstFaxAdj.SelectedItem.SubItems(24)
    txtPolicyNo.Text = lstFaxAdj.SelectedItem.SubItems(25)
    txtPayor.Text = lstFaxAdj.SelectedItem.SubItems(26)
    txtPolicyStatus.Text = lstFaxAdj.SelectedItem.SubItems(27)
    txtExpStatus.Text = lstFaxAdj.SelectedItem.SubItems(28)
    If Trim(lstFaxAdj.SelectedItem.SubItems(29)) <> "" Then
        MaskEdBox7.Text = lstFaxAdj.SelectedItem.SubItems(29)
    End If
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(30)) <> "" Then
        MaskEdBox8.Text = lstFaxAdj.SelectedItem.SubItems(30)
    End If
    
    txtComp.Text = lstFaxAdj.SelectedItem.SubItems(31)
 End If
End Sub

Private Sub lstFaxAdj_DblClick()
'Call UEnabledEntries
 'cmdReg.Enabled = True
 Call ClearEntries2
 Call ComboListing
 Call avalue
 
If lstFaxAdj.ListItems.Count > 0 Then

Frame2.Enabled = False
Frame4.Enabled = False

'Frame6.Enabled = True
    lstFaxAdj.Enabled = True
    lstMbm.Enabled = True

    cmdReg.Enabled = False
    cmdEdit.Enabled = True
    cmdEdit.Visible = True
    cmdcancel.Visible = False
    cmdSave.Enabled = False
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    
    
    txtFaxRefNo.Text = lstFaxAdj.SelectedItem.Text
    cboDocType.Text = lstFaxAdj.SelectedItem.SubItems(1)
    cboSenderCategory.Text = lstFaxAdj.SelectedItem.SubItems(2)
    cbosendername.Text = lstFaxAdj.SelectedItem.SubItems(3)
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(4)) <> "" Then
        MaskEdBox1.Text = lstFaxAdj.SelectedItem.SubItems(4)
    End If
    
    'MaskEdBox1.Text = lstFaxAdj.SelectedItem.SubItems(4)
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(5)) <> "" Then
        'MaskEdBox5.Text = Mid(lstFaxAdj.SelectedItem.SubItems(5), 1, 5)
        MaskEdBox5.Text = lstFaxAdj.SelectedItem.SubItems(5)
        'cborectime.Text = Mid(lstFaxAdj.SelectedItem.SubItems(5), 6, 2)
    End If
    

    If Trim(lstFaxAdj.SelectedItem.SubItems(6)) <> "" Then
        MaskEdBox2.Text = lstFaxAdj.SelectedItem.SubItems(6)
    End If
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(7)) <> "" Then
        'MaskEdBox6.Text = Mid(lstFaxAdj.SelectedItem.SubItems(7), 1, 5)
        MaskEdBox6.Text = lstFaxAdj.SelectedItem.SubItems(7)
        'cbopassouttime.Text = Mid(lstFaxAdj.SelectedItem.SubItems(7), 6, 2)
    End If
   
    
    cboRecBy.Text = lstFaxAdj.SelectedItem.SubItems(8)
    
    cboAdmHos.Text = lstFaxAdj.SelectedItem.SubItems(9)
    cboClaimAssessor.Text = lstFaxAdj.SelectedItem.SubItems(10)
    If Trim(lstFaxAdj.SelectedItem.SubItems(11)) <> "" Then
        MaskEdBox3.Text = lstFaxAdj.SelectedItem.SubItems(11)
    End If
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(12)) <> "" Then
        MaskEdBox4.Text = lstFaxAdj.SelectedItem.SubItems(12)
    End If
    'MaskEdBox3.Text = lstFaxAdj.SelectedItem.SubItems(11)
    'MaskEdBox4.Text = lstFaxAdj.SelectedItem.SubItems(12)
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(13)) = "" Then
        txtamount.Text = ""
    ElseIf IsNumeric(Trim(lstFaxAdj.SelectedItem.SubItems(13))) = True Then
        txtamount.Text = Format(Trim(lstFaxAdj.SelectedItem.SubItems(13)), "#,##0.00")
    Else
        txtamount.Text = ""
    End If
    
    txtCaseNo.Text = lstFaxAdj.SelectedItem.SubItems(14)
    
    txtRemarks.Text = lstFaxAdj.SelectedItem.SubItems(32)
    
 
    
    txtMemNo.Text = lstFaxAdj.SelectedItem.SubItems(15)
    txtcoverid.Text = lstFaxAdj.SelectedItem.SubItems(16)
    lblINorOut.Caption = lstFaxAdj.SelectedItem.SubItems(22)
    txtPatientName.Text = lstFaxAdj.SelectedItem.SubItems(23)
    txtPatientIc.Text = lstFaxAdj.SelectedItem.SubItems(24)
    txtPolicyNo.Text = lstFaxAdj.SelectedItem.SubItems(25)
    txtPayor.Text = lstFaxAdj.SelectedItem.SubItems(26)
    txtPolicyStatus.Text = lstFaxAdj.SelectedItem.SubItems(27)
    txtExpStatus.Text = lstFaxAdj.SelectedItem.SubItems(28)
    If Trim(lstFaxAdj.SelectedItem.SubItems(29)) <> "" Then
        MaskEdBox7.Text = lstFaxAdj.SelectedItem.SubItems(29)
    End If
    
    If Trim(lstFaxAdj.SelectedItem.SubItems(30)) <> "" Then
        MaskEdBox8.Text = lstFaxAdj.SelectedItem.SubItems(30)
    End If
    
    txtComp.Text = lstFaxAdj.SelectedItem.SubItems(31)
 End If
End Sub

Private Sub lstMbm_DblClick()
If lstMbm.ListItems.Count > 0 Then
    txtMemNo.Text = lstMbm.SelectedItem.Text
    txtcoverid.Text = lstMbm.SelectedItem.SubItems(1)
    lblINorOut.Caption = lstMbm.SelectedItem.SubItems(2)
    txtPatientName.Text = lstMbm.SelectedItem.SubItems(3)
    txtPatientIc.Text = lstMbm.SelectedItem.SubItems(4)
    txtPolicyNo.Text = lstMbm.SelectedItem.SubItems(5)
    txtPayor.Text = lstMbm.SelectedItem.SubItems(6)
    txtPolicyStatus.Text = lstMbm.SelectedItem.SubItems(7)
    txtExpStatus.Text = lstMbm.SelectedItem.SubItems(8)
    MaskEdBox7.Text = lstMbm.SelectedItem.SubItems(9)
    MaskEdBox8.Text = lstMbm.SelectedItem.SubItems(10)
    txtComp.Text = lstMbm.SelectedItem.SubItems(11)
 End If
End Sub

Private Sub MaskEdBox5_LostFocus()
If Trim(MaskEdBox5.Text) = "" Then
ElseIf Trim(MaskEdBox5.Text) = "__:__" Then
ElseIf Mid(MaskEdBox5.Text, 1, 2) <= 11 And Mid(MaskEdBox5.Text, 1, 2) >= 1 Then

ElseIf Mid(MaskEdBox5.Text, 1, 2) = 12 Then

ElseIf Mid(MaskEdBox5.Text, 1, 2) = 0 Then
    'If MsgBox("Press YES if Rec Time is  " & Trim(MaskEdBox5.Text) & "AM.", vbExclamation + vbYesNo, DialogBoxTitle) = vbNo Then
    '    MsgBox "Invalid Rec. Time ! ", vbCritical
    '    txtcoverid.Text = ""
    '    txtcoverid.SetFocus
    'Else
    '
    'End If
ElseIf Mid(MaskEdBox5.Text, 1, 2) <= 24 And Mid(MaskEdBox5.Text, 1, 2) >= 13 Then
   
Else
    MsgBox "Invalid Rec. Time ! ", vbCritical
    MaskEdBox5.mask = "##:##"
End If
End Sub

Private Sub MaskEdBox6_LostFocus()
If Trim(MaskEdBox6.Text) = "" Then
ElseIf Trim(MaskEdBox6.Text) = "__:__" Then
ElseIf Mid(MaskEdBox6.Text, 1, 2) <= 11 And Mid(MaskEdBox6.Text, 1, 2) >= 1 Then

ElseIf Mid(MaskEdBox6.Text, 1, 2) = 12 Then

ElseIf Mid(MaskEdBox6.Text, 1, 2) = 0 Then
    'If MsgBox("Press YES if Rec Time is  " & Trim(MaskEdBox5.Text) & "AM.", vbExclamation + vbYesNo, DialogBoxTitle) = vbNo Then
    '    MsgBox "Invalid Rec. Time ! ", vbCritical
    '    txtcoverid.Text = ""
    '    txtcoverid.SetFocus
    'Else
    '
    'End If
ElseIf Mid(MaskEdBox6.Text, 1, 2) <= 24 And Mid(MaskEdBox6.Text, 1, 2) >= 13 Then
   
Else
    MsgBox "Invalid Pass Out. Time ! ", vbCritical
    MaskEdBox6.mask = "##:##"
    'MaskEdBox6.SetFocus
End If
End Sub

Private Sub txtamount_LostFocus()
If Trim(txtamount.Text) = "" Then

ElseIf IsNumeric(txtamount.Text) = True Then
    txtamount.Text = Format(txtamount, "#,##0.00")
Else
    MsgBox "Invalid amount ! ", vbCritical
    txtamount.Text = ""
    txtamount.SetFocus
End If
End Sub

Private Sub txtCaseNo_LostFocus()
txtCaseNo = ChkStr(txtCaseNo)
End Sub

Private Sub txtComp_LostFocus()
txtComp = ChkStr(txtComp)
End Sub

Private Sub txtcoverid_LostFocus()
If Trim(txtcoverid.Text) = "" Then

ElseIf IsNumeric(txtcoverid.Text) = True Then
   
Else
    MsgBox "Invalid Cover ID ! ", vbCritical
    txtcoverid.Text = ""
    txtcoverid.SetFocus
End If
End Sub

Private Sub txtExpStatus_LostFocus()
txtExpStatus = ChkStr(txtExpStatus)
End Sub

Private Sub txtFaxRefNo_LostFocus()
txtFaxRefNo = ChkStr(txtFaxRefNo)
End Sub

Private Sub txtMemNo_LostFocus()
txtMemNo = ChkStr(txtMemNo)
End Sub

Private Sub txtPatientIc_LostFocus()
txtPatientIc = ChkStr(txtPatientIc)
End Sub

Private Sub txtPatientName_LostFocus()
txtPatientName = ChkStr(txtPatientName)
End Sub

Private Sub txtPayor_LostFocus()
txtPayor = ChkStr(txtPayor)
End Sub

Private Sub txtPolicyNo_LostFocus()
txtPolicyNo = ChkStr(txtPolicyNo)
End Sub

Private Sub txtPolicyStatus_LostFocus()
txtPolicyStatus = ChkStr(txtPolicyStatus)
End Sub

Private Sub txtRemarks_LostFocus()
txtRemarks = ChkStr(txtRemarks)
End Sub


