VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMembershipCardControl 
   Caption         =   "Membership Card Control"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   10140
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkCheckAll 
      Caption         =   "Check All"
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
      TabIndex        =   40
      Top             =   1480
      Width           =   1095
   End
   Begin VB.Frame Frame8 
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   240
      Width           =   9975
      Begin VB.TextBox txtMBMNumberFR 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtMBMNumberTO 
         Height          =   285
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtBordxDateFR 
         Height          =   285
         Left            =   6600
         TabIndex        =   4
         Top             =   240
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
      Begin MSMask.MaskEdBox txtBordxDateTO 
         Height          =   285
         Left            =   8280
         TabIndex        =   5
         Top             =   240
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
      Begin VB.ComboBox cboPayor 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtBatchNo1 
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
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtBatchNo2 
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
         MaxLength       =   3
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cboSendCat 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   750
         Width           =   3975
      End
      Begin MSMask.MaskEdBox txtSentDateFr 
         Height          =   285
         Left            =   6600
         TabIndex        =   8
         Top             =   720
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
      Begin MSMask.MaskEdBox txtSentDateTo 
         Height          =   285
         Left            =   8280
         TabIndex        =   9
         Top             =   720
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
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8040
         TabIndex        =   33
         Top             =   765
         Width           =   315
      End
      Begin VB.Label Label26 
         Caption         =   "Sent Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   32
         Top             =   765
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "Sending Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   765
         Width           =   1245
      End
      Begin VB.Label Label5 
         Caption         =   "Batch No."
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
         Left            =   5640
         TabIndex        =   30
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "To"
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
         Left            =   8040
         TabIndex        =   29
         Top             =   495
         Width           =   495
      End
      Begin VB.Label Label24 
         Caption         =   "Payor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   520
         Width           =   1245
      End
      Begin VB.Label Label24 
         Caption         =   "MBM Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label26 
         Caption         =   "Bordx Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   22
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   8040
         TabIndex        =   21
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   20
         Top             =   240
         Width           =   315
      End
   End
   Begin MSComctlLib.ListView lstMemberListing 
      Height          =   3735
      Left            =   120
      TabIndex        =   18
      Top             =   1725
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6588
      SortKey         =   8
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MBMNumber"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cover ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Member Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Bordx Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Batch No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Sending Mode"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Sent Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Batch No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Grp Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Cost Center"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Branch"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Doc Ref"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   5520
      Width           =   9975
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   13
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         TabIndex        =   16
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9120
         TabIndex        =   17
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   12
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         TabIndex        =   14
         Top             =   210
         Width           =   720
      End
      Begin VB.ComboBox cbosendingmode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7680
         TabIndex        =   15
         Top             =   210
         Width           =   720
      End
      Begin MSMask.MaskEdBox txtsentdate 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.TextBox txtDocRef 
         Height          =   285
         Left            =   3840
         TabIndex        =   41
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label24 
         Caption         =   "Doc Reference"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   2640
         TabIndex        =   42
         Top             =   200
         Width           =   1245
      End
      Begin VB.Label Label24 
         Caption         =   "Sent Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   500
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "Send Mode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   200
         Width           =   1155
      End
   End
   Begin VB.Label lblTotalPrincipal 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   39
      Top             =   6480
      Width           =   1035
   End
   Begin VB.Label Label24 
      Caption         =   "Total Principal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   38
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Label lblTotalSupp 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   37
      Top             =   6480
      Width           =   1275
   End
   Begin VB.Label Label24 
      Caption         =   "Total Supp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   2640
      TabIndex        =   36
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Label lblTotalMember 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   35
      Top             =   6480
      Width           =   1755
   End
   Begin VB.Label Label24 
      Caption         =   "Total Member"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   5160
      TabIndex        =   34
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Membership Card Control"
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
      TabIndex        =   26
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmMembershipCardControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpedit As Boolean
Public intExit As Integer
Private m_blnDirection(1 To 9) As Boolean

Private Sub cboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayor.Text, cboPayor.SelStart) + Chr(KeyAscii) + Mid(cboPayor.Text, cboPayor.SelStart + cboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayor.ListCount - 1
 
        If NewText = LCase(Left(cboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayor.Text = ValidValue
                cboPayor.SelStart = 0
                cboPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboSendCat_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboSendCat.Text, cboSendCat.SelStart) + Chr(KeyAscii) + Mid(cboSendCat.Text, cboSendCat.SelStart + cboSendCat.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboSendCat.ListCount - 1
 
        If NewText = LCase(Left(cboSendCat.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboSendCat.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboSendCat.Text = ValidValue
                cboSendCat.SelStart = 0
                cboSendCat.SelLength = Len(ValidValue)
            End If
        End If
End Sub


Private Sub cbosendingmode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cbosendingmode.Text, cbosendingmode.SelStart) + Chr(KeyAscii) + Mid(cbosendingmode.Text, cbosendingmode.SelStart + cbosendingmode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cbosendingmode.ListCount - 1
 
        If NewText = LCase(Left(cbosendingmode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cbosendingmode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cbosendingmode.Text = ValidValue
                cbosendingmode.SelStart = 0
                cbosendingmode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cmdAdd_Click()
    cbosendingmode = ""
    txtsentdate.mask = ""
    txtsentdate.Text = ""
    txtsentdate.mask = "##/##/####"
    cbosendingmode.Enabled = True
    txtsentdate.Enabled = True
    tmpedit = False
End Sub

Private Sub chkCheckAll_Click()
    If chkCheckAll.Value = "1" Then
        SetCheckAllItems True
    Else
        SetCheckAllItems False
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lstMemberListing.ListItems.Clear
    lblTotalPrincipal = 0
    lblTotalSupp = 0
    lblTotalMember = 0
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    Screen.MousePointer = vbHourglass
    
    If lstMemberListing.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstMemberListing)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
    

End Sub

Private Sub CmdSearch_Click()
    cmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    cmdPrint.Enabled = False
    cmdClear.Enabled = False
    CmdUpdate.Enabled = False
    CmdExit.Enabled = False
    cbosendingmode = ""
    txtsentdate = "__/__/____"
    lblTotalPrincipal = 0
    lblTotalSupp = 0
    lblTotalMember = 0
    
    Call MBMInquiryList
    
    cmdPrint.Enabled = True
    cmdClear.Enabled = True
    CmdUpdate.Enabled = True
    CmdExit.Enabled = True

 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    Screen.MousePointer = Default
    cmdSearch.Enabled = True

End Sub

Private Sub MBMInquiryList()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sqlstr As String
Dim strCriteria As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim strsqlCOV As String
Dim RstCov As New ADODB.Recordset
Dim Check As Boolean
Dim rstCard As New ADODB.Recordset
Dim strsqlCard As String

    strCriteria = ""
    If cboPayor <> "" Then
        strCriteria = strCriteria & " AND PAYCode =  '" & Mid(cboPayor, 1, 2) & "'"
    End If
    If txtMBMNumberFR <> "" Then
        strCriteria = strCriteria & " AND MBMNumber >=  '" & Trim(txtMBMNumberFR) & "'"
    End If
    If txtMBMNumberTO <> "" Then
        strCriteria = strCriteria & " AND MBMNumber <=  '" & (txtMBMNumberTO) & "'"
    End If
    If txtBordxDateFR <> "__/__/____" Then
        strCriteria = strCriteria & " AND MBMBordxDate >= '" & Format(txtBordxDateFR, "MM/DD/YYYY") & "'"
    End If
    If txtBordxDateTO <> "__/__/____" Then
        strCriteria = strCriteria & " AND MBMBordxDate <= '" & Format(txtBordxDateTO, "MM/DD/YYYY") & "'"
    End If
    If txtBatchNo1 <> "" Then
        strCriteria = strCriteria & " AND MBMBatchNo >=  '" & (txtBatchNo1) & "'"
    End If
    If txtBatchNo2 <> "" Then
        strCriteria = strCriteria & " AND MBMBatchNo <=  '" & (txtBatchNo2) & "'"
    End If
    If txtSentDateFr <> "__/__/____" Then
        strCriteria = strCriteria & " AND MBMCardDate >= '" & Format(txtSentDateFr, "MM/DD/YYYY") & "'"
    End If
    If txtSentDateTo <> "__/__/____" Then
        strCriteria = strCriteria & " AND MBMCardDate <= '" & Format(txtSentDateTo, "MM/DD/YYYY") & "'"
    End If
    If cboSendCat <> "" Then
        strCriteria = strCriteria & " AND MBMSendingMode <=  '" & (cboSendCat) & "'"
    End If
    
    lstMemberListing.ListItems.Clear
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Search_MBM_CardSent"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
    cmd.Parameters.Append Prm1
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    Do
        Do Until rst2.EOF
            With rst2
                Set Record = lstMemberListing.ListItems.Add(, , "")
                Record.SubItems(1) = !MBMNumber
                Record.SubItems(2) = "00"
                Record.SubItems(3) = !MBMName
                Record.SubItems(4) = IIf(Len(!MBMBordxDate), Trim(!MBMBordxDate), "")
                Record.SubItems(5) = IIf(Len(!MBMBatchNo), Trim(!MBMBatchNo), "")
                Record.SubItems(6) = IIf(Len(!MBMSendingMode), Trim(!MBMSendingMode), "")
                Record.SubItems(7) = IIf(Len(!MBMCardDate), Format(!MBMCardDate, "dd/mm/yyyy"), "")
                Record.SubItems(8) = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                Record.SubItems(9) = IIf(Len(!MBMBatchNo), Trim(!MBMBatchNo), "")
                Record.SubItems(10) = IIf(Len(!MBMSendingMode), Trim(!MBMSendingMode), "")
                Record.SubItems(11) = IIf(Len(!MBMCardDate), Format(!MBMCardDate, "dd/mm/yyyy"), "")
                Record.SubItems(12) = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                Record.SubItems(13) = IIf(Len(!MBMDocRef), !MBMDocRef, "")
' dbo.MBMTwo.MBMDOCRef,
 '                     dbo.MBMOthers.MBMGroupCompany , dbo.MBMOthers.MBMDepartment, dbo.MBMOthers.MBMCostCenter, dbo.MBMOthers.MBMBranch
                
                
                lblTotalPrincipal = CDbl(lblTotalPrincipal) + 1

                If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                    strsqlCOV = ""
                    'strsqlCOV = "Select MBMCNumber,MBMCCoverID,MBMSendingMode,MBMCardsentDate From MBMCoveredPersonsTwo where MBCMNumber = '" & Trim(!MBMNumber) & "'"
                    'RstCov.Open strsqlCOV, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                    'strsqlCOV = "SELECT MBMCNumber,MBMCCoverID,MBMSendingMode,MBMCardDate From MBMCoveredPersonsTwo  Where MBMCNumber = '" & Trim(!MBMNumber) & "'"
                    'RstCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
                    strsqlCOV = ""
                    strsqlCOV = "SELECT MBMCoveredPersons.MBMCNumber, " & _
                                "MBMCoveredPersons.MBMCCoverID, MBMCoveredPersons.MBMCName," & _
                                "MBMCoveredPersonsTwo.MBMSendingMode,MBMCoveredPersonsTwo.MBMDocRef,  MBMCoveredPersonsTwo.MBMCardDate " & _
                                "FROM MBMCoveredPersons INNER JOIN  MBMCoveredPersonsTwo ON " & _
                                "MBMCoveredPersons.MBMCNumber = MBMCoveredPersonsTwo.MBMCNumber AND " & _
                                "MBMCoveredPersons.MBMCCoverID = MBMCoveredPersonsTwo.MBMCCoverID  Where MBMCoveredPersons.MBMCNumber = '" & Trim(!MBMNumber) & "'"
                    RstCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                    'strsqlCOV = ""
                    'strsqlCOV = "Select MBMCNumber,MBMCCoverID,MBMSendingMode,MBMCardsentDate From MBMCoveredPersonsTwo Where MBMCNumber = '" & Trim(!MBMNumber) & "'"
                    'RstCov.Open strsqlCOV, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                    
                    Do While Not RstCov.EOF
                        Set Record = lstMemberListing.ListItems.Add(, , "")
                        Record.SubItems(1) = IIf(Len(!MBMNumber), Trim(!MBMNumber), "")
                        Record.SubItems(2) = IIf(Len(RstCov!MBMCCoverID), Trim(RstCov!MBMCCoverID), "")
                        Record.SubItems(3) = IIf(Len(RstCov!MBMCName), Trim(RstCov!MBMCName), "")
                        Record.SubItems(4) = IIf(Len(!MBMBordxDate), !MBMBordxDate, "")
                        Record.SubItems(5) = IIf(Len(!MBMBatchNo), Trim(!MBMBatchNo), "")
                        Record.SubItems(6) = IIf(Len(RstCov!MBMSendingMode), Trim(RstCov!MBMSendingMode), "")
                        Record.SubItems(7) = IIf(Len(RstCov!MBMCardDate), Format(RstCov!MBMCardDate, "dd/mm/yyyy"), "")
                        Record.SubItems(8) = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                        Record.SubItems(9) = IIf(Len(!MBMBatchNo), Trim(!MBMBatchNo), "")
                        Record.SubItems(10) = IIf(Len(!MBMSendingMode), Trim(!MBMSendingMode), "")
                        Record.SubItems(11) = IIf(Len(!MBMCardDate), Format(!MBMCardDate, "dd/mm/yyyy"), "")
                        Record.SubItems(12) = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                        Record.SubItems(12) = IIf(Len(!MBMDocRef), !MBMDocRef, "")
                        
                        lblTotalSupp = CDbl(lblTotalSupp) + 1
                    RstCov.MoveNext
                    Loop
                    RstCov.Close
                    Set RstCov = Nothing
                End If
                                                            
                                                            
            End With
            rst2.MoveNext
            DoEvents
            If intExit = 1 Then
               Check = False
               Exit Do
            End If
        Loop
    Loop Until Check = False
    lblTotalMember = CDbl(lblTotalSupp) + CDbl(lblTotalPrincipal)
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    If lstMemberListing.ListItems.Count = 0 Then
        MsgBox "Record Not Found!", vbCritical + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Sub cmdStop_Click()
    intExit = 1
End Sub

Private Sub CmdUpdate_Click()
On Error GoTo errSave
Dim RecordSaved
Dim RstMBM As New ADODB.Recordset
Dim strsql As String
Dim i As Integer
Dim startTrans As Boolean
Dim tmpsentdate As String

        startTrans = False
        'ItemChecked = False
        tmpsentdate = ""
        If Trim(cbosendingmode) = "" Then
            MsgBox "Please Enter Card Sending Mode!", vbCritical + vbOKOnly, DialogBoxTitle
            cbosendingmode.SetFocus
        ElseIf txtsentdate = "__/__/____" Then
            MsgBox "Please Enter Card Sent Date!", vbCritical + vbOKOnly, DialogBoxTitle
            txtsentdate.SetFocus
        Else
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                CmdUpdate.Enabled = False
                CmdExit.Enabled = False
                cmdSearch.Enabled = False
                tmpsentdate = Format(txtsentdate, "yyyy/mm/dd")
               ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
               
                Screen.MousePointer = vbHourglass
                startTrans = True
                medixmasterconn.beginTrans
                    For i = 1 To lstMemberListing.ListItems.Count
                        If lstMemberListing.ListItems(i).Checked = True Then
                            strsql = ""
                            If lstMemberListing.ListItems(i).SubItems(2) = "00" Then
                                strsql = "Update MBMTwo set MBMSendingMode = '" & Trim(cbosendingmode) & "', MBMCardDate = '" & Format(tmpsentdate, "YYYY/MM/DD") & "', MBMDOCRef = '" & Trim(txtDocRef) & "' Where MBMNumber = '" & lstMemberListing.ListItems(i).SubItems(1) & "'"
                            Else
                                strsql = "Update MBMCoveredPersonsTwo set MBMSendingMode = '" & Trim(cbosendingmode) & "', MBMCardDate = '" & Format(tmpsentdate, "YYYY/MM/DD") & "', MBMDOCRef = '" & Trim(txtDocRef) & "'  Where MBMCNumber = '" & lstMemberListing.ListItems(i).SubItems(1) & "' And MBMCCoverID = '" & lstMemberListing.ListItems(i).SubItems(2) & "'"
                            End If
                            medixmasterconn.Execute strsql
                        End If
                    Next i
                medixmasterconn.commitTrans
                MsgBox "Record updated !", vbOKOnly
                
                lstMemberListing.ListItems.Clear
                Call MBMInquiryList
               ' medixmasterconn.Close
              '  Set medixmasterconn = Nothing
                cmdPrint.Enabled = True
                cmdClear.Enabled = True
                CmdUpdate.Enabled = True
                CmdExit.Enabled = True
                cmdSearch.Enabled = True
                chkCheckAll.Value = 1
            Screen.MousePointer = Default
            
        End If
    End If
    Exit Sub
errSave:
    MsgBox Err.Description
    If startTrans = True Then
        medixmasterconn.RollbackTrans
    End If
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
    cmdPrint.Enabled = True
    cmdClear.Enabled = True
    CmdUpdate.Enabled = True
    CmdExit.Enabled = True
    cmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub Form_Load()
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    lblTotalPrincipal = 0
    lblTotalSupp = 0
    lblTotalMember = 0
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
   
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute

    Do While Not PAY.EOF
        With PAY
            cboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    
    PAY.Close
    Set PAY = Nothing
    
    cbosendingmode.AddItem "By Hand"
    cbosendingmode.AddItem "By Post"
    cbosendingmode.AddItem "Courier"
    
    cboSendCat.AddItem "By Hand"
    cboSendCat.AddItem "By Post"
    cboSendCat.AddItem "Courier"
    
End Sub


Private Sub lstMemberListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstMemberListing, ColumnHeader.Index, ldtstring, m_blnDirection(1)
        Case 2
            SortListView lstMemberListing, ColumnHeader.Index, ldtstring, m_blnDirection(2)
        Case 3
            SortListView lstMemberListing, ColumnHeader.Index, ldtstring, m_blnDirection(3)
        Case 4
            SortListView lstMemberListing, ColumnHeader.Index, ldtstring, m_blnDirection(4)
        Case 5
            SortListView lstMemberListing, ColumnHeader.Index, ldtDatetime, m_blnDirection(5)
        Case 6
            SortListView lstMemberListing, ColumnHeader.Index, ldtstring, m_blnDirection(6)
        Case 7
            SortListView lstMemberListing, ColumnHeader.Index, ldtstring, m_blnDirection(7)
        Case 8
            SortListView lstMemberListing, ColumnHeader.Index, ldtDatetime, m_blnDirection(8)
        Case 9
            SortListView lstMemberListing, ColumnHeader.Index, ldtstring, m_blnDirection(9)
            
        End Select

End Sub

Private Sub lstMemberListing_KeyDown(KeyCode As Integer, Shift As Integer)
    'If lstMemberListing.ListItems.Count > 0 Then
    '    cbosendingmode = lstMemberListing.SelectedItem.SubItems(5)
    '    txtsentdate = Format(lstMemberListing.SelectedItem.SubItems(6), "dd/mm/yyyy")
    'End If
End Sub

Private Sub lstMemberListing_KeyUp(KeyCode As Integer, Shift As Integer)
    'If lstMemberListing.ListItems.Count > 0 Then
    '    cbosendingmode = lstMemberListing.SelectedItem.SubItems(5)
    '    txtsentdate = Format(lstMemberListing.SelectedItem.SubItems(6), "dd/mm/yyyy")
    'End If
End Sub

Private Sub smdEdit_Click()
    tmpedit = True
End Sub

Private Sub txtBordxDateFR_LostFocus()
    If IsDate(txtBordxDateFR) Then
    Else
        If txtBordxDateFR <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDateFR.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDateTO_LostFocus()
    If IsDate(txtBordxDateTO) Then
    Else
        If txtBordxDateTO <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDateTO.SetFocus
        End If
    End If
End Sub

Private Sub txtsentdate_LostFocus()
    If IsDate(txtsentdate) Then
    Else
        If txtsentdate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtsentdate.SetFocus
        End If
    End If

End Sub

Private Sub txtSentDateFr_LostFocus()
    If IsDate(txtSentDateFr) Then
    Else
        If txtSentDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtSentDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtSentDateTo_LostFocus()
    If IsDate(txtSentDateTo) Then
    Else
        If txtSentDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtSentDateTo.SetFocus
        End If
    End If
End Sub


Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstMemberListing.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstMemberListing.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub



