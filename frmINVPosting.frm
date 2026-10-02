VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmINVPosting 
   Caption         =   "Invoice Posting"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   570
      Left            =   0
      TabIndex        =   13
      Top             =   240
      Width           =   9255
      Begin VB.TextBox TxtINVNoTo 
         Height          =   285
         Left            =   3120
         TabIndex        =   15
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox TxtINVNoFr 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   180
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtINVRegDateTo 
         Height          =   285
         Left            =   7320
         TabIndex        =   16
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtINVRegDateFr 
         Height          =   285
         Left            =   5880
         TabIndex        =   17
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   7080
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "INV Reg Date"
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Inv No from"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   9255
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Height          =   375
         Left            =   6240
         TabIndex        =   23
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   8520
         TabIndex        =   5
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   7920
         TabIndex        =   4
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lstInvPosting 
      Height          =   4215
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
         Text            =   "A/C Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Based Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Debit/Credit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Trans Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Journal Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Trans Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Desc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "T1 Clinic"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "T2 Plan No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "T3 Pay Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "T4  InvDORMth"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Grp Com"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "T6 Bordx Ref"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Invoice Posting"
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
      TabIndex        =   22
      Top             =   0
      Width           =   9645
   End
End
Attribute VB_Name = "frmINVPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub txtINVRegDateFr_LostFocus()
    If txtInvDate <> "__/__/____" Then
        If IsDate(txtInvDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtInvDate.SetFocus
        ElseIf CDate(txtInvDate) > CDate(Date) Then
            MsgBox "Invoice Date cannot greater than system date ! ", vbCritical
            txtInvDate.SetFocus
        End If
    End If
End Sub


Private Sub txtINVRegDateTo_LostFocus()
    If txtInvDate <> "__/__/____" Then
        If IsDate(txtInvDate) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtInvDate.SetFocus
        ElseIf CDate(txtInvDate) > CDate(Date) Then
            MsgBox "Invoice Date cannot greater than system date ! ", vbCritical
            txtInvDate.SetFocus
        End If
    End If
End Sub
