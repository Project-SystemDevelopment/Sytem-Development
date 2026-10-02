VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFaxReg 
   Caption         =   "Fax Received Registration"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   12390
   Begin VB.Frame Frame6 
      Height          =   2655
      Left            =   120
      TabIndex        =   70
      Top             =   5400
      Width           =   12135
      Begin MSComctlLib.ListView lstMbm 
         Height          =   2295
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4048
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
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12135
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   7080
         TabIndex        =   66
         Top             =   3960
         Width           =   4935
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   300
            Left            =   3480
            TabIndex        =   69
            Top             =   250
            Width           =   1000
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   300
            Left            =   1980
            TabIndex        =   68
            Top             =   250
            Width           =   1000
         End
         Begin VB.CommandButton cmdReg 
            Caption         =   "Register"
            Height          =   300
            Left            =   480
            TabIndex        =   67
            Top             =   250
            Width           =   1000
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4575
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   6855
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
            TabIndex        =   39
            Top             =   4005
            Width           =   1500
         End
         Begin VB.ComboBox cbosendername 
            Height          =   315
            Left            =   1800
            TabIndex        =   38
            Top             =   870
            Width           =   4930
         End
         Begin VB.TextBox txtRemarks 
            Height          =   615
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            Top             =   3390
            Width           =   4930
         End
         Begin VB.TextBox txtCaseNo 
            Height          =   315
            Left            =   5100
            TabIndex        =   36
            Top             =   3075
            Width           =   1260
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
            TabIndex        =   35
            Top             =   240
            Width           =   4930
         End
         Begin VB.ComboBox cboSenderCategory 
            Height          =   315
            Left            =   1800
            TabIndex        =   34
            Top             =   555
            Width           =   4930
         End
         Begin VB.ComboBox cboRecBy 
            Height          =   315
            Left            =   1800
            TabIndex        =   33
            Top             =   1800
            Width           =   4930
         End
         Begin VB.ComboBox cboAdmHos 
            Height          =   315
            Left            =   1800
            TabIndex        =   32
            Top             =   2130
            Width           =   4930
         End
         Begin VB.ComboBox cboClaimAssessor 
            Height          =   315
            Left            =   1800
            TabIndex        =   31
            Top             =   2445
            Width           =   4930
         End
         Begin VB.TextBox txtamount 
            Height          =   315
            Left            =   1800
            TabIndex        =   30
            Top             =   3075
            Width           =   1500
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   315
            Left            =   1800
            TabIndex        =   40
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
            TabIndex        =   41
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
            TabIndex        =   42
            Top             =   1200
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66387969
            CurrentDate     =   40525
         End
         Begin MSComCtl2.DTPicker DTRecdate 
            Height          =   315
            Left            =   3360
            TabIndex        =   43
            Top             =   1200
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66387969
            CurrentDate     =   40525
         End
         Begin MSComCtl2.DTPicker DTDoa 
            Height          =   315
            Left            =   3360
            TabIndex        =   44
            Top             =   2760
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66387969
            CurrentDate     =   40525
         End
         Begin MSComCtl2.DTPicker DTDod 
            Height          =   315
            Left            =   6430
            TabIndex        =   45
            Top             =   2760
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            _Version        =   393216
            Format          =   66387969
            CurrentDate     =   40525
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   315
            Left            =   5100
            TabIndex        =   46
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
            TabIndex        =   47
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
            TabIndex        =   48
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
            TabIndex        =   49
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
            TabIndex        =   74
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
            TabIndex        =   73
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "Fax Ref. No."
            Height          =   315
            Left            =   120
            TabIndex        =   65
            Top             =   4005
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Remarks"
            Height          =   315
            Left            =   120
            TabIndex        =   64
            Top             =   3390
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Case No"
            Height          =   315
            Left            =   3990
            TabIndex        =   63
            Top             =   3075
            Width           =   1095
         End
         Begin VB.Label Label20 
            Caption         =   "DOD"
            Height          =   315
            Left            =   3990
            TabIndex        =   62
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "DOA"
            Height          =   315
            Left            =   120
            TabIndex        =   61
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Doc. Type"
            Height          =   315
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Sender Category"
            Height          =   315
            Left            =   120
            TabIndex        =   59
            Top             =   555
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Sender Name"
            Height          =   315
            Left            =   120
            TabIndex        =   58
            Top             =   870
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Date"
            Height          =   315
            Left            =   120
            TabIndex        =   57
            Top             =   1185
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Time"
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   1500
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Rec. By"
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   1815
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Admit. Hospital"
            Height          =   315
            Left            =   120
            TabIndex        =   54
            Top             =   2130
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Claim Assessor"
            Height          =   315
            Left            =   120
            TabIndex        =   53
            Top             =   2445
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Date Passout"
            Height          =   315
            Left            =   3990
            TabIndex        =   52
            Top             =   1185
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Time Passout"
            Height          =   315
            Left            =   3990
            TabIndex        =   51
            Top             =   1500
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Bill Amount"
            Height          =   315
            Left            =   120
            TabIndex        =   50
            Top             =   3075
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3855
         Left            =   7080
         TabIndex        =   1
         Top             =   120
         Width           =   4935
         Begin VB.TextBox txtcoveridgo 
            Height          =   315
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   26
            Top             =   210
            Width           =   495
         End
         Begin VB.Frame Frame3 
            Caption         =   "Member Details"
            Height          =   3015
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   4695
            Begin VB.TextBox txtComp 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   2565
               Width           =   3255
            End
            Begin VB.TextBox txtcoverid 
               Height          =   315
               Left            =   3600
               MaxLength       =   2
               TabIndex        =   12
               Top             =   990
               Width           =   855
            End
            Begin VB.TextBox txtExpStatus 
               Height          =   315
               Left            =   3285
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   1920
               Width           =   1170
            End
            Begin VB.TextBox txtPolicyStatus 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   1935
               Width           =   1170
            End
            Begin VB.TextBox txtPayor 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   1620
               Width           =   3255
            End
            Begin VB.TextBox txtPolicyNo 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   1305
               Width           =   3255
            End
            Begin VB.TextBox txtMemNo 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   990
               Width           =   2055
            End
            Begin VB.TextBox txtPatientIc 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   675
               Width           =   1575
            End
            Begin VB.TextBox txtPatientName 
               Height          =   315
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   360
               Width           =   3255
            End
            Begin MSMask.MaskEdBox MaskEdBox7 
               Height          =   315
               Left            =   1200
               TabIndex        =   75
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
            Begin VB.Label lblINorOut 
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   3120
               TabIndex        =   25
               Top             =   705
               Width           =   1335
            End
            Begin VB.Label Label28 
               Caption         =   "Comp"
               Height          =   315
               Left            =   120
               TabIndex        =   24
               Top             =   2565
               Width           =   1095
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
               TabIndex        =   23
               Top             =   990
               Width           =   135
            End
            Begin VB.Label Label24 
               Caption         =   "Exp. date"
               Height          =   315
               Left            =   2400
               TabIndex        =   22
               Top             =   2250
               Width           =   975
            End
            Begin VB.Label Label23 
               Caption         =   "Eff. Date"
               Height          =   315
               Left            =   120
               TabIndex        =   21
               Top             =   2250
               Width           =   1095
            End
            Begin VB.Label Label22 
               Caption         =   "Exp. Status"
               Height          =   315
               Left            =   2400
               TabIndex        =   20
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label18 
               Caption         =   "Policy Status"
               Height          =   315
               Left            =   120
               TabIndex        =   19
               Top             =   1935
               Width           =   975
            End
            Begin VB.Label Label17 
               Caption         =   "Payor"
               Height          =   315
               Left            =   120
               TabIndex        =   18
               Top             =   1620
               Width           =   975
            End
            Begin VB.Label Label16 
               Caption         =   "Policy No"
               Height          =   315
               Left            =   120
               TabIndex        =   17
               Top             =   1305
               Width           =   975
            End
            Begin VB.Label Label15 
               Caption         =   "Mem. No."
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   990
               Width           =   975
            End
            Begin VB.Label Label14 
               Caption         =   "Patient Ic"
               Height          =   315
               Left            =   120
               TabIndex        =   15
               Top             =   675
               Width           =   975
            End
            Begin VB.Label Label13 
               Caption         =   "Patient Name"
               Height          =   315
               Left            =   120
               TabIndex        =   14
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdGo 
            Caption         =   "Go"
            Height          =   315
            Left            =   3840
            TabIndex        =   3
            Top             =   210
            Width           =   735
         End
         Begin VB.TextBox txtMBMNoGo 
            Height          =   315
            Left            =   960
            TabIndex        =   2
            Top             =   210
            Width           =   1935
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
            TabIndex        =   28
            Top             =   150
            Width           =   135
         End
         Begin VB.Label Label12 
            Caption         =   "Mem No"
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Fax Received Registration"
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
      TabIndex        =   72
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "FrmFaxReg"
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

Private Sub avalue()
    MaskEdBox1.mask = "##/##/####"
    MaskEdBox2.mask = "##/##/####"
    MaskEdBox3.mask = "##/##/####"
    MaskEdBox4.mask = "##/##/####"
    
    MaskEdBox5.mask = "##:##"
    MaskEdBox6.mask = "##:##"
    
    MaskEdBox1.Text = tmpDateonly
    MaskEdBox5.Text = tmpTimeonly
    'txtRegBy.Text = Trim(Logon)
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
    
lstMbm.ListItems.Clear
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

Private Sub cboAdmHos_Click()
'cboAdmHos = ChkStr(cboAdmHos)
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

Private Sub CmdClear_Click()
Call UEnabledEntries
 cmdReg.Enabled = True
 cmdGo.Enabled = True
 Call ClearEntries
 Call ComboListing
 Call avalue
     Frame4.Enabled = True
    Frame2.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set FrmFaxReg = Nothing
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

Private Sub Newfaxreg()
Dim faxNumber As String
Dim RecordSaved
faxNumber = ""
Dim startTrans As Boolean
Dim strsqlfax As String

Dim tmpDateRec As String
Dim tmpDatePassOut As String
Dim tmpdoa As String
Dim tmpdod As String
Dim Effdate As String
Dim ExpDate As String

    RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then
    
        'If Not medixmasterconnMaint Is Nothing And Not medixmasterconnMaint = "" Then
         '   medixmasterconnMaint.Close
        'End If

        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        'Begin Transaction
        medixmasterconnMaint.beginTrans
        startTrans = True
        
        faxNumber = GenerateFaxRegNo()
        
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
        
        
        strsqlfax = "EXEC Fax_Insert '" & ChkStr(faxNumber) & "', '" & ChkStr(Trim(cboDocType.Text)) & "', '" & ChkStr(Trim(cboSenderCategory.Text)) & "','" & ChkStr(Trim(cbosendername.Text)) & "','" & tmpDateRec & "','" & MaskEdBox5.Text & "', " & _
                    "'" & tmpDatePassOut & "', '" & MaskEdBox6.Text & "', '" & ChkStr(Trim(cboAdmHos.Text)) & "', " & _
                    "'" & ChkStr(Trim(cboClaimAssessor.Text)) & "', '" & tmpdoa & "', '" & tmpdod & "', '" & Trim(txtamount.Text) & "','" & ChkStr(Trim(txtCaseNo.Text)) & "'," & _
                    "'" & ChkStr(Trim(txtRemarks.Text)) & "', '" & ChkStr(Trim(txtMemNo.Text)) & "', '" & Trim(txtcoverid.Text) & "' , '" & ChkStr(Logon) & "', '" & Format(tmpDateonly, "mm/dd/yyyy") & "', " & _
                    "'" & ChkStr(Trim(cboRecBy.Text)) & "','" & tmp24time & "','" & Trim(lblINorOut) & "', " & _
                    "'" & ChkStr(Trim(txtPatientIc.Text)) & "','" & ChkStr(Trim(txtPatientName.Text)) & "','" & ChkStr(Trim(txtPolicyNo.Text)) & "','" & ChkStr(Trim(txtPayor.Text)) & "','" & ChkStr(Trim(txtComp.Text)) & "', " & _
                    "'" & ChkStr(Trim(txtPolicyStatus.Text)) & "','" & ChkStr(Trim(txtExpStatus.Text)) & "','" & Effdate & "','" & ExpDate & "'"
                        
    'medixmasterconnMaint.Execute strsqlfax
        'medixmasterconnMaint.Close
        'Set medixmasterconnMaint = Nothing
        
        If startTrans = True Then
            medixmasterconnMaint.commitTrans
            medixmasterconnMaint.Execute strsqlfax
            MsgBox "Record Saved, Your Fax Ref. No.is : " & faxNumber, vbOKOnly + vbInformation, DialogBoxTitle
            'Call EnabledEntries
            Frame4.Enabled = False
            Frame2.Enabled = False
            cmdReg.Enabled = False
            cmdGo.Enabled = False
        End If
     End If
    
    Exit Sub
ErrorSave:
    MsgBox Err.Description
    If startTrans = True Then
            Frame4.Enabled = True
            Frame2.Enabled = True
        medixmasterconn.RollbackTrans

    End If
End Sub

Private Sub cmdReg_Click()

    If Not medixmasterconnMaint Is Nothing And Not medixmasterconnMaint = "" Then
        medixmasterconnMaint.Close
    End If
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Call Newfaxreg
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub DTDoa_Change()
MaskEdBox3.Text = DTDoa.Value
End Sub

Private Sub DTDoa_Click()
'MaskEdBox3.Text = DTDoa.Value
End Sub

Private Sub DTDod_Change()
MaskEdBox4.Text = DTDod.Value
End Sub

Private Sub DTDod_Click()
'MaskEdBox4.Text = DTDod.Value
End Sub

Private Sub DTPassoutDate_Change()
MaskEdBox2.Text = DTPassoutDate.Value
End Sub

Private Sub DTPassoutDate_Click()
'MaskEdBox2.Text = DTPassoutDate.Value
End Sub

Private Sub DTRecdate_Change()
MaskEdBox1.Text = DTRecdate.Value
End Sub

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

Private Sub DTRecdate_Click()
'MaskEdBox1.Text = DTRecdate.Value
End Sub

Private Sub Form_Load()
    Call UEnabledEntries
    Call ServerTime
    Call ClearEntries
    Call ComboListing
    Call avalue
    Frame4.Enabled = True
    Frame2.Enabled = True
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

Private Function GenerateFaxRegNo()
    Dim fax As New ADODB.Recordset
    Dim strsql As String
    Dim faxno As String
    Dim prefix As String
    prefix = "F"

    strsql = "Select * from FaxSeq Where faxPrefix = '" & prefix & "'"
    fax.Open strsql, medixmasterconnMaint, adOpenStatic, adLockOptimistic
    
    If fax.recordCount = 0 Then
        With fax
            .AddNew
            !faxPrefix = prefix
            !faxSeqNo = "00001"
            .Update
            faxno = "00001"
        End With
    Else
        fax!faxSeqNo = Format(fax!faxSeqNo + 1, "00000")
        faxno = fax!faxSeqNo
        fax.Update
    End If

    'GenerateCaseNo = PayorCode & PlanCode & Mid(BusinessYear, 3, 2) & CAQNumber
    GenerateFaxRegNo = prefix & faxno
    fax.Close
    Set fax = Nothing
End Function




Private Sub txtCaseNo_Change()
'txtCaseNo = ChkStr(txtCaseNo)
End Sub

Private Sub txtCaseNo_LostFocus()
txtCaseNo = ChkStr(txtCaseNo)
End Sub

Private Sub txtComp_Change()
'txtComp = ChkStr(txtComp)
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

Private Sub txtcoveridgo_LostFocus()
If Trim(txtcoveridgo.Text) = "" Then

ElseIf IsNumeric(txtcoveridgo.Text) = True Then
   
Else
    MsgBox "Invalid Cover ID ! ", vbCritical
    txtcoveridgo.Text = ""
    txtcoveridgo.SetFocus
End If
End Sub

Private Sub txtEffDate_Change()

End Sub

Private Sub txtExpDate_Change()

End Sub

Private Sub txtExpStatus_Change()
'txtExpStatus = ChkStr(txtExpStatus)
End Sub

Private Sub txtExpStatus_LostFocus()
txtExpStatus = ChkStr(txtExpStatus)
End Sub

Private Sub txtFaxRefNo_Change()
'txtFaxRefNo = ChkStr(txtFaxRefNo)
End Sub


Private Sub txtFaxRefNo_LostFocus()
txtFaxRefNo = ChkStr(txtFaxRefNo)
End Sub

Private Sub txtMBMNoGo_Change()
'txtMBMNoGo = ChkStr(txtMBMNoGo)
End Sub

Private Sub txtMBMNoGo_LostFocus()
txtMBMNoGo = ChkStr(txtMBMNoGo)
End Sub

Private Sub txtMemNo_Change()
'txtMemNo = ChkStr(txtMemNo)
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

Private Sub txtPolicyNo_Change()
'txtPolicyNo = ChkStr(txtPolicyNo)
End Sub

Private Sub txtPolicyNo_LostFocus()
txtPolicyNo = ChkStr(txtPolicyNo)
End Sub

Private Sub txtPolicyStatus_Change()
'txtPolicyStatus = ChkStr(txtPolicyStatus)
End Sub

Private Sub txtPolicyStatus_LostFocus()
txtPolicyStatus = ChkStr(txtPolicyStatus)
End Sub

Private Sub txtRemarks_Change()
'txtRemarks = ChkStr(txtRemarks)
End Sub

Private Sub txtRemarks_LostFocus()
txtRemarks = ChkStr(txtRemarks)
End Sub
