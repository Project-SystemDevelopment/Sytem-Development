VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReceiptTracking 
   Caption         =   "Receipt Tracking"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstReceiptTracking 
      Height          =   3375
      Left            =   120
      TabIndex        =   49
      Top             =   3960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5953
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Receipt No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Receipt Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Payor"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Member Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Claim No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "D/N No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "A/C Period"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3660
      Left            =   120
      TabIndex        =   38
      Top             =   240
      Width           =   6135
      Begin VB.ComboBox CboHLTCode 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   3615
      End
      Begin VB.OptionButton OptPayor 
         Caption         =   "Payor"
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   3615
      End
      Begin VB.OptionButton OptMBM 
         Caption         =   "Member"
         Height          =   255
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptMNRB 
         Caption         =   "MNRB"
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtACPeriodTo 
         Height          =   255
         Left            =   5040
         TabIndex        =   53
         Top             =   6960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##/##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtMBMNo 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   5
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox TxtMBMName 
         Height          =   285
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   6
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox TxtClaimNo1 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox TxtClaimNo2 
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtDebitFr 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtDebitTo 
         Height          =   285
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtRecNoTo 
         Height          =   285
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2160
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtRecDate2 
         Height          =   300
         Left            =   3840
         TabIndex        =   16
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtAmtTo 
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   18
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmReceiptTracking.frx":0000
         Left            =   3840
         List            =   "FrmReceiptTracking.frx":0002
         TabIndex        =   20
         Top             =   2880
         Width           =   1575
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmReceiptTracking.frx":0004
         Left            =   3840
         List            =   "FrmReceiptTracking.frx":0006
         TabIndex        =   22
         Top             =   3150
         Width           =   1575
      End
      Begin VB.TextBox txtRecNoFr 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2160
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtRecDate1 
         Height          =   300
         Left            =   1800
         TabIndex        =   15
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtAmtFr 
         Height          =   285
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   17
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmReceiptTracking.frx":0008
         Left            =   1800
         List            =   "FrmReceiptTracking.frx":000A
         TabIndex        =   19
         Top             =   2880
         Width           =   1575
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmReceiptTracking.frx":000C
         Left            =   1800
         List            =   "FrmReceiptTracking.frx":000E
         TabIndex        =   21
         Top             =   3150
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Health Type"
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   555
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "CLM No"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   61
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   59
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "A/C Period Yr Fr"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   3375
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   56
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   55
         Top             =   3090
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "A/C Period Mth Fr"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   3090
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Track By"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Mem Name"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Mem No"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   46
         Top             =   2565
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Receipt Date"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   44
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Debit Note"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   42
         Top             =   2310
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Receipt No"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   40
         Top             =   2835
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2760
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   7320
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   4080
         TabIndex        =   66
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   8520
         TabIndex        =   29
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   5400
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7680
         TabIndex        =   28
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6840
         TabIndex        =   27
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9480
         TabIndex        =   31
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10320
         TabIndex        =   32
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3240
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1320
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1080
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1680
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstTrackMNRB 
      Height          =   3375
      Left            =   120
      TabIndex        =   51
      Top             =   3960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5953
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Receipt No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Receipt Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Rec't  Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Member Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "WKS No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "D/N No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "A/C Period"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Reports Type"
      Height          =   3660
      Left            =   6360
      TabIndex        =   58
      Top             =   240
      Width           =   5415
      Begin VB.OptionButton Option4 
         Caption         =   "By Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "By Receipt No"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "By Claim"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Bordx"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Receipt Tracking"
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
      TabIndex        =   63
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmReceiptTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Public ReceiptTrackingCriteria As String
Private AA As Integer
Private m_blnDirection(1 To 8) As Boolean


Private Sub cboACPeriodMthFr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 48 And KeyAscii <> 57 Then
    NewText = LCase(Left(cboACPeriodMthFr.Text, cboACPeriodMthFr.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodMthFr.Text, cboACPeriodMthFr.SelStart + cboACPeriodMthFr.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodMthFr.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodMthFr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodMthFr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodMthFr.Text = ValidValue
                cboACPeriodMthFr.SelStart = 0
                cboACPeriodMthFr.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboACPeriodMthTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 48 And KeyAscii <> 57 Then
    NewText = LCase(Left(cboACPeriodMthTo.Text, cboACPeriodMthTo.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodMthTo.Text, cboACPeriodMthTo.SelStart + cboACPeriodMthTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodMthTo.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodMthTo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodMthTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodMthTo.Text = ValidValue
                cboACPeriodMthTo.SelStart = 0
                cboACPeriodMthTo.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboACPeriodYRFr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 48 And KeyAscii <> 57 Then
    NewText = LCase(Left(cboACPeriodYRFr.Text, cboACPeriodYRFr.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodYRFr.Text, cboACPeriodYRFr.SelStart + cboACPeriodYRFr.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodYRFr.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodYRFr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodYRFr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodYRFr.Text = ValidValue
                cboACPeriodYRFr.SelStart = 0
                cboACPeriodYRFr.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboACPeriodYRTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 48 And KeyAscii <> 57 Then
    NewText = LCase(Left(cboACPeriodYRTo.Text, cboACPeriodYRTo.SelStart) + Chr(KeyAscii) + Mid(cboACPeriodYRTo.Text, cboACPeriodYRTo.SelStart + cboACPeriodYRTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboACPeriodYRTo.ListCount - 1
 
        If NewText = LCase(Left(cboACPeriodYRTo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboACPeriodYRTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboACPeriodYRTo.Text = ValidValue
                cboACPeriodYRTo.SelStart = 0
                cboACPeriodYRTo.SelLength = Len(ValidValue)
            End If
        End If
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

Private Sub cmdEXPtoExcel_Click()
    'Screen.MousePointer = vbHourglass
        
        'If OptMBM.value = True Then
            'Call ExportToExcelReceipt(lstReceiptTracking)
        'Else
            'Call ExportToExcelReceipt(lstTrackMNRB)
        'End If
    'Screen.MousePointer = vbDefault
    
    Screen.MousePointer = vbHourglass
        
    If OptMBM.Value = True Then
        If lstReceiptTracking.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(lstReceiptTracking)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
        
    Else
    
        If lstTrackMNRB.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(lstTrackMNRB)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
    
    End If
        
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub CmdSearch_Click()
On Error GoTo errRun
  '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    
    Screen.MousePointer = vbHourglass
        CmdClear.Enabled = False
        CmdExit.Enabled = False
        cmdEXPtoExcel.Enabled = False
        cmdPrint.Enabled = False
        CmdSearch.Enabled = False
            Call SearchRecord
    
    Screen.MousePointer = Default
    
   ' medixmasterconnPayment.Close
    'Set medixmasterconnPayment = Nothing
    Exit Sub
    
errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstReceiptTracking.ListItems.Clear
    lblcurrent = ""
    cmdPrint.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    cmdEXPtoExcel.Enabled = True
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
End Sub

Private Sub Form_Load()
    intExit = 0
    lblcurrent = 0
    lbltotal = 0
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdClear_Click()
    lstReceiptTracking.ListItems.Clear
    lstTrackMNRB.ListItems.Clear
    TxtMBMNo.Text = ""
    txtMBMName.Text = ""
    CboPayor.Text = ""
    cboHLTCode.Text = ""
    txtRECNoFr.Text = ""
    txtRECNoTo.Text = ""
    txtRecDate1.Text = "__/__/____"
    txtRecDate2.Text = "__/__/____"
    txtDebitFr.Text = ""
    txtDebitTo.Text = ""
    txtAmtFr.Text = ""
    TxtAmtTo.Text = ""
    lblcurrent = 0
    lbltotal = 0
    lbltotalAmt = 0
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub

'******************* Start Combo Listing ********************************

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim monthstart
Dim yearstart

    'PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    For monthstart = 1 To 12
        With cboACPeriodMthFr
            .AddItem monthstart
        End With
        With cboACPeriodMthTo
            .AddItem monthstart
        End With
    Next monthstart
    
    For yearstart = 1999 To 3000
        With cboACPeriodYRFr
            .AddItem yearstart
        End With
        With cboACPeriodYRTo
            .AddItem yearstart
        End With
    Next yearstart
    
    cboHLTCode.AddItem "S - Sihat"
    cboHLTCode.AddItem "N - Non-Sihat"
    
End Sub
'******************* End Combo Listing ********************************


Private Function SearchRecord()
Dim strsql As String
Dim StrAppend As String
Dim Sql As String
    
    StrAppend = ""
    
    If Option2.Value = True Then
    
        If Len(Trim(cboHLTCode)) Then
            StrAppend = StrAppend + " And HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & "'"
        End If
        
        If Len(Trim(TxtMBMNo)) Then
            StrAppend = StrAppend + " And MBMNumber like '%" & ChkStr(TxtMBMNo) & "%'"
        End If
        
        If Len(Trim(CboPayor)) Then
            StrAppend = StrAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
        End If
        
        If Len(Trim(txtMBMName)) Then
            StrAppend = StrAppend + " And MBMName like '%" & ChkStr(txtMBMName) & "%'"
        End If
        
        If Len(Trim(TxtClaimNo1)) Then
            StrAppend = StrAppend + " And CASNumber >= '" & Trim(TxtClaimNo1) & "'"
        End If
        
        If Len(Trim(TxtClaimNo2)) Then
            StrAppend = StrAppend + " And CASNumber <= '" & Trim(TxtClaimNo2) & "'"
        End If
        
        If Len(Trim(TxtBordxNo1)) Then
            StrAppend = StrAppend + " And BordxNo >= '" & Trim(TxtBordxNo1) & "'"
        End If
        
        If Len(Trim(TxtBordxNo2)) Then
            StrAppend = StrAppend + " And BordxNo <= '" & Trim(TxtBordxNo2) & "'"
        End If
        
        If Len(Trim(txtRECNoFr)) Then
            StrAppend = StrAppend + " AND ReceiptNo >= '" & RemStr(Trim(txtRECNoFr)) & "'"
        End If
               
        If Len(Trim(txtRECNoTo)) Then
            StrAppend = StrAppend + " AND ReceiptNo <= '" & RemStr(Trim(txtRECNoTo)) & "'"
        End If
        
        If (txtRecDate1) <> "__/__/____" And IsDate(txtRecDate1) = True _
            And (txtRecDate2) <> "__/__/____" And IsDate(txtRecDate2) = True Then
            Sql = ""
            Sql = " AND ReceiptDate >= '" & Format(Trim(txtRecDate1), "mm/dd/yyyy") & _
                  "' And ReceiptDate <= '" & Format(Trim(txtRecDate2), "mm/dd/yyyy") & "'"
            StrAppend = StrAppend + Sql
        End If
        
        If OptMBM.Value = True Then
        
            If Len(Trim(cboACPeriodMthFr)) Then
                StrAppend = StrAppend + " And RECMBMAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
            End If
            If Len(Trim(cboACPeriodMthTo)) Then
                StrAppend = StrAppend + " And RECMBMAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
            End If
            
            If Len(Trim(cboACPeriodYRFr)) Then
                StrAppend = StrAppend + " And RECMBMAccPeriodYR >= '" & Trim(cboACPeriodYRFr) & "'"
            End If
            If Len(Trim(cboACPeriodYRTo)) Then
                StrAppend = StrAppend + " And RECMBMAccPeriodYR <= '" & Trim(cboACPeriodYRTo) & "'"
            End If
        Else
            If Len(Trim(cboACPeriodMthFr)) Then
            StrAppend = StrAppend + " And RECAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
            End If
            If Len(Trim(cboACPeriodMthTo)) Then
                StrAppend = StrAppend + " And RECAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
            End If
            
            If Len(Trim(cboACPeriodYRFr)) Then
                StrAppend = StrAppend + " And RECAccPeriodYR >= '" & Trim(cboACPeriodYRFr) & "'"
            End If
            If Len(Trim(cboACPeriodYRTo)) Then
                StrAppend = StrAppend + " And RECAccPeriodYR <= '" & Trim(cboACPeriodYRTo) & "'"
            End If
        End If
        
        If Len(Trim(txtRecDate1)) > 0 And IsDate(txtRecDate1) = True Then
            StrAppend = StrAppend + " And ReceiptDate >= '" & Format(txtRecDate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtRecDate2)) > 0 And IsDate(txtRecDate2) = True Then
            StrAppend = StrAppend + " And ReceiptDate <= '" & Format(txtRecDate2, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDebitFr)) Then
            StrAppend = StrAppend + " AND DebitCreditRefNo >= '" & RemStr(Trim(txtDebitFr)) & "'"
        End If
        
        If Len(Trim(txtDebitTo)) Then
            StrAppend = StrAppend + " AND DebitCreditRefNo <= '" & RemStr(Trim(txtDebitTo)) & "'"
        End If
               
        If Len(Trim(txtAmtFr)) Then
            If IsNumeric(txtAmtFr) Then
                StrAppend = StrAppend + " AND ReceiptAmt >= " & Trim(txtAmtFr) & ""
            Else
                MsgBox "Amount From is not Integer! It would be bypassed!", vbCritical
            End If
        End If
        
        If IsNumeric(TxtAmtTo) Then
           If IsNumeric(TxtAmtTo) Then
                StrAppend = StrAppend + " AND ReceiptAmt <= " & Trim(TxtAmtTo) & ""
           Else
                MsgBox "Amount From is not Integer! It would be bypassed!", vbCritical
           End If
        End If
        
    ElseIf Option1.Value = True Or Option3.Value = True Or Option4.Value = True Then
    
        If Len(Trim(cboHLTCode)) Then
            StrAppend = StrAppend + " And HLTCode = '" & Trim(Mid(cboHLTCode, 1, 1)) & "'"
        End If
        
        If Len(Trim(CboPayor)) Then
            StrAppend = StrAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
        End If
        
        If Len(Trim(txtRecDate1)) > 0 And IsDate(txtRecDate1) = True Then
            StrAppend = StrAppend + " And ReceiptDate >= '" & Format(txtRecDate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtRecDate2)) > 0 And IsDate(txtRecDate2) = True Then
            StrAppend = StrAppend + " And ReceiptDate <= '" & Format(txtRecDate2, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(TxtBordxNo1)) Then
            StrAppend = StrAppend + " And BordxNo >= '" & Trim(TxtBordxNo1) & "'"
        End If
        
        If Len(Trim(TxtBordxNo2)) Then
            StrAppend = StrAppend + " And BordxNo <= '" & Trim(TxtBordxNo2) & "'"
        End If
        
        If Len(Trim(txtRECNoFr)) Then
            StrAppend = StrAppend + " AND ReceiptNo >= '" & RemStr(Trim(txtRECNoFr)) & "'"
        End If
               
        If Len(Trim(txtRECNoTo)) Then
            StrAppend = StrAppend + " AND ReceiptNo <= '" & RemStr(Trim(txtRECNoTo)) & "'"
        End If
    
    End If

        
    If OptMBM.Value = True And Option1.Value = True Then
        Call ReceiptMemberTrackListing2(StrAppend)
    ElseIf OptPayor = True And Option1.Value = True Then
        Call ReceiptPayorTrackListing2(StrAppend)
    ElseIf OptMNRB = True And Option1.Value = True Then
        Call ReceiptMNRBTrackListing2(StrAppend)
    ElseIf OptMBM.Value = True And Option2.Value = True Then
        Call ReceiptMemberTrackListing(StrAppend)
    ElseIf OptPayor = True And Option2.Value = True Then
        Call ReceiptPayorTrackListing(StrAppend)
    ElseIf OptMNRB = True And Option2.Value = True Then
        Call ReceiptMNRBTrackListing(StrAppend)
    ElseIf OptMBM.Value = True And Option3.Value = True Then
        Call ReceiptMemberTrackListing3(StrAppend)
    ElseIf OptPayor = True And Option3.Value = True Then
        Call ReceiptPayorTrackListing3(StrAppend)
    ElseIf OptMNRB = True And Option3.Value = True Then
        Call ReceiptMNRBTrackListing3(StrAppend)
    ElseIf OptMBM.Value = True And Option4.Value = True Then
        Call ReceiptMemberByPayor(StrAppend)
    ElseIf OptPayor = True And Option4.Value = True Then
        Call ReceiptPayorByPayor(StrAppend)
    ElseIf OptMNRB = True And Option4.Value = True Then
        Call ReceiptMNRBByPayor(StrAppend)
    End If
   
End Function
'******************* End Search Record*******************************

'******************* Start ReceiptMBM Tracking Listing *******************************
Private Sub ReceiptMemberTrackListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstReceiptTracking.ListItems.Clear
            
    'Sql = " Select CASNumber, ClaimVersion, MBMNumber, MBMName, ReceiptNo, PAYCode, " & _
          " DebitCreditRefNo, ReceiptDate, ReceiptAmt, ReceiptRemarks, " & _
          " RECMBMAccPeriodMth, RECMBMAccPeriodYr " & _
          " From VIEW_ReceiptfrMBMTracking where 0 = 0 "
     
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
        
            With rst2
                Set Record = lstReceiptTracking.ListItems.Add(, , rst2!ReceiptNo)
                If Len(rst2!ReceiptDate) Then
                   Record.SubItems(1) = Format(rst2!ReceiptDate, "dd/mm/yyyy")
                End If
                If Len(rst2!ReceiptAmt) Then
                   Record.SubItems(2) = Format(rst2!ReceiptAmt, "#,##0.00")
                End If
                If Len(rst2!PAYCode) Then
                    Record.SubItems(3) = rst2!PAYCode
                End If
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(4) = rst2!MBMNumber
                End If
                If Len(rst2!MBMName) Then
                    Record.SubItems(5) = Trim(rst2!MBMName)
                End If
                If Len(!CasNumber & "-" & !claimversion) Then
                    Record.SubItems(6) = (!CasNumber & "-" & !claimversion)
                End If
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(7) = rst2!DebitCreditRefNo
                End If
                If Len(rst2!ReceiptRemarks) Then
                    Record.SubItems(8) = Trim(rst2!ReceiptRemarks)
                End If
                
                If Len(rst2!RECMBMAccPeriodMth & "/" & rst2!RECMBMAccPeriodYr) Then
                    Record.SubItems(9) = (rst2!RECMBMAccPeriodMth & "/" & rst2!RECMBMAccPeriodYr)
                End If
                                
                Current = Current + 1
                lblcurrent = Current
                If Len(rst2!ReceiptAmt) Then
                    TotalAmt = TotalAmt + rst2!ReceiptAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub ReceiptMemberTrackListing2(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstReceiptTracking.ListItems.Clear
            
    'Sql = " Select BordxNo, BordxDate, COUNT(BordxNo) AS NoofCase, SUM(ReceiptMBMAmt) AS ReceiptMBMAmt From VIEW_ReceiptfrMBMTracking2 where 0 = 0 "
    
    'sql2 = "Group By BordxNo, BordxDate "
         
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend + sql2
    'Else
    '    Sql = Sql + sql2
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "2"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
                
            With rst2
                If Len(!BordxNo) Then
                    Set Record = lstReceiptTracking.ListItems.Add(, , !BordxNo)
                Else
                    Set Record = lstReceiptTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(1) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = Trim(!NoofCase)
                Else
                    Record.SubItems(2) = "0"
                End If
                If Len(!ReceiptMBMAmt) Then
                    Record.SubItems(3) = Format(!ReceiptMBMAmt, "#,##0.00")
                End If
                                                
                Current = Current + 1
                lblcurrent = Current
                If Len(!ReceiptMBMAmt) Then
                    TotalAmt = TotalAmt + !ReceiptMBMAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub ReceiptMemberTrackListing3(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstReceiptTracking.ListItems.Clear
            
    'Sql = " Select ReceiptNo, ReceiptDate, COUNT(ReceiptNo) AS NoofCase, SUM(ReceiptMBMAmt) AS ReceiptMBMAmt From VIEW_ReceiptfrMBMTracking2 where 0 = 0 "
    
    'sql2 = "Group By ReceiptNo, ReceiptDate "
     
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend + sql2
    'Else
    '    Sql = Sql + sql2
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "3"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
        
            With rst2
                If Len(!ReceiptNo) Then
                    Set Record = lstReceiptTracking.ListItems.Add(, , !ReceiptNo)
                Else
                    Set Record = lstReceiptTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!ReceiptDate) Then
                    Record.SubItems(1) = Format(!ReceiptDate, "dd/mm/yyyy")
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = Trim(!NoofCase)
                Else
                    Record.SubItems(2) = "0"
                End If
                If Len(!ReceiptMBMAmt) Then
                   Record.SubItems(3) = Format(!ReceiptMBMAmt, "#,##0.00")
                End If
                                                
                Current = Current + 1
                lblcurrent = Current
                If Len(!ReceiptMBMAmt) Then
                    TotalAmt = TotalAmt + !ReceiptMBMAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub
'******************* End Receipt MBM Tracking Listing *******************************

'******************* Start ReceiptPayor Tracking Listing *******************************

Private Sub ReceiptPayorTrackListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstTrackMNRB.ListItems.Clear
            
    'Sql = " Select CASNumber, ClaimVersion, MBMNumber, MBMName, " & _
          " ReceiptNo, DebitCreditRefNo, ReceiptDate, " & _
          " RECAccPeriodYr, RECAccPeriodMth, " & _
          " BordxNo, ReceiptAmt, Remark, PAYCode, RecPAYAmt, CheckAmt " & _
          " From VIEW_ReceiptfrPayorTracking where 0 = 0 "
     
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "4"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
        'Total = rst2.recordCount
        'lblTotal = Total
        
            With rst2
                Set Record = lstTrackMNRB.ListItems.Add(, , rst2!ReceiptNo)
                If Len(rst2!ReceiptDate) Then
                   Record.SubItems(1) = Format(rst2!ReceiptDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CheckAmt) Then
                   Record.SubItems(2) = Format(rst2!CheckAmt, "#,##0.00")
                End If
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(3) = rst2!MBMNumber
                End If
                If Len(rst2!MBMName) Then
                    Record.SubItems(4) = Trim(rst2!MBMName)
                End If
                If Len(rst2!BordxNo) Then
                    Record.SubItems(5) = rst2!BordxNo
                End If
                If Len(!CasNumber & "-" & !claimversion) Then
                    Record.SubItems(6) = (!CasNumber & "-" & !claimversion)
                End If
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(7) = rst2!DebitCreditRefNo
                End If
                If Len(rst2!Remark) Then
                    Record.SubItems(8) = Trim(rst2!Remark)
                End If
                
                If Len(rst2!RECAccPeriodMth & "/" & rst2!RECAccPeriodYR) Then
                    Record.SubItems(9) = (rst2!RECAccPeriodMth & "/" & rst2!RECAccPeriodYR)
                End If
                                
                Current = Current + 1
                lblcurrent = Current
                If Len(rst2!CheckAmt) Then
                    TotalAmt = TotalAmt + rst2!CheckAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub ReceiptPayorTrackListing2(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstTrackMNRB.ListItems.Clear
            
    'Sql = " Select BordxNo, BordxDate, COUNT(BordxNo) AS NoofCase, SUM(ReceiptPAYAmt) AS ReceiptPAYAmt From VIEW_ReceiptfrPayorTracking2 where 0 = 0 "
    
    'sql2 = "Group By BordxNo, BordxDate "
     
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend + sql2
    'Else
    '    Sql = Sql + sql2
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "5"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
        
            With rst2
                If Len(!BordxNo) Then
                    Set Record = lstTrackMNRB.ListItems.Add(, , !BordxNo)
                Else
                    Set Record = lstTrackMNRB.ListItems.Add(, , "Empty")
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(1) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = Trim(!NoofCase)
                Else
                    Record.SubItems(2) = "0"
                End If
                If Len(!ReceiptPAYAmt) Then
                   Record.SubItems(3) = Format(!ReceiptPAYAmt, "#,##0.00")
                End If
                                                
                Current = Current + 1
                lblcurrent = Current
                If Len(!ReceiptPAYAmt) Then
                    TotalAmt = TotalAmt + !ReceiptPAYAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub ReceiptPayorTrackListing3(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstTrackMNRB.ListItems.Clear
            
    'Sql = " Select ReceiptNo, ReceiptDate, COUNT(ReceiptNo) AS NoofCase, SUM(ReceiptPAYAmt) AS ReceiptPAYAmt From VIEW_ReceiptfrPayorTracking2 where 0 = 0 "
    
    'sql2 = "Group By ReceiptNo, ReceiptDate "
     
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend + sql2
    'Else
    '    Sql = Sql + sql2
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "6"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
        
            With rst2
                If Len(!ReceiptNo) Then
                    Set Record = lstTrackMNRB.ListItems.Add(, , !ReceiptNo)
                Else
                    Set Record = lstTrackMNRB.ListItems.Add(, , "Empty")
                End If
                If Len(!ReceiptDate) Then
                    Record.SubItems(1) = Format(!ReceiptDate, "dd/mm/yyyy")
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = Trim(!NoofCase)
                Else
                    Record.SubItems(2) = "0"
                End If
                If Len(!ReceiptPAYAmt) Then
                   Record.SubItems(3) = Format(!ReceiptPAYAmt, "#,##0.00")
                End If
                                                
                Current = Current + 1
                lblcurrent = Current
                If Len(!ReceiptPAYAmt) Then
                    TotalAmt = TotalAmt + !ReceiptPAYAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub
'******************* End ReceiptPayor Tracking Listing *******************************


'******************* Start ReceiptMNRB Tracking Listing *******************************

Private Sub ReceiptMNRBTrackListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstTrackMNRB.ListItems.Clear
            
    'Sql = " Select CASNumber, ClaimVersion, MBMNumber, MBMName, " & _
          " ReceiptNo, DebitCreditRefNo, ReceiptDate, " & _
          " RECAccPeriodYr, RECAccPeriodMth, " & _
          " BordxNo, ReceiptAmt, Remark, PAYCode, CheckAmt " & _
          " From VIEW_ReceiptfrMNRBTracking where 0 = 0 "
     
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "7"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
        'Total = rst2.recordCount
        'lblTotal = Total
        
            With rst2
            
                If Len(rst2!ReceiptNo) Then
                    Set Record = lstTrackMNRB.ListItems.Add(, , rst2!ReceiptNo)
                Else
                    Set Record = lstTrackMNRB.ListItems.Add(, , "Empty")
                End If
                If Len(rst2!ReceiptDate) Then
                   Record.SubItems(1) = Format(rst2!ReceiptDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CheckAmt) Then
                   Record.SubItems(2) = Format(rst2!CheckAmt, "#,##0.00")
                End If
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(3) = rst2!MBMNumber
                End If
                If Len(rst2!MBMName) Then
                    Record.SubItems(4) = Trim(rst2!MBMName)
                End If
                If Len(rst2!BordxNo) Then
                    Record.SubItems(5) = rst2!BordxNo
                End If
                If Len(!CasNumber & "-" & !claimversion) Then
                    Record.SubItems(6) = (!CasNumber & "-" & !claimversion)
                End If
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(7) = rst2!DebitCreditRefNo
                End If
                If Len(rst2!Remark) Then
                    Record.SubItems(8) = Trim(rst2!Remark)
                End If
                
                If Len(rst2!RECAccPeriodMth & "/" & rst2!RECAccPeriodYR) Then
                    Record.SubItems(9) = (rst2!RECAccPeriodMth & "/" & rst2!RECAccPeriodYR)
                End If
                                
                Current = Current + 1
                lblcurrent = Current
                If Len(rst2!CheckAmt) Then
                    TotalAmt = TotalAmt + rst2!CheckAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub ReceiptMNRBTrackListing2(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstTrackMNRB.ListItems.Clear
            
    'Sql = " Select BordxNo, BordxDate, COUNT(BordxNo) AS NoofCase, SUM(ReceiptMNRBAmt) AS ReceiptMNRBAmt From VIEW_ReceiptfrMNRBTracking2 where 0 = 0 "
    
    'sql2 = " Group By BordxNo, BordxDate "
     
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend + sql2
    'Else
    '    Sql = Sql + sql2
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "8"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
                
            With rst2
                If Len(!BordxNo) Then
                    Set Record = lstTrackMNRB.ListItems.Add(, , !BordxNo)
                Else
                    Set Record = lstTrackMNRB.ListItems.Add(, , "Empty")
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(1) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = Trim(!NoofCase)
                Else
                    Record.SubItems(2) = "0"
                End If
                If Len(!ReceiptMNRBAmt) Then
                   Record.SubItems(3) = Format(!ReceiptMNRBAmt, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                                                
                Current = Current + 1
                lblcurrent = Current
                If Len(!ReceiptMNRBAmt) Then
                    TotalAmt = TotalAmt + !ReceiptMNRBAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub ReceiptMNRBTrackListing3(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstTrackMNRB.ListItems.Clear
            
    'Sql = " Select ReceiptNo, ReceiptDate, COUNT(ReceiptNo) AS NoofCase, SUM(ReceiptMNRBAmt) AS ReceiptMNRBAmt From VIEW_ReceiptfrMNRBTracking2 where 0 = 0 "
    
    'sql2 = " Group By ReceiptNo, ReceiptDate "
     
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend + sql2
    'Else
    '    Sql = Sql + sql2
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "9"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
                
            With rst2
                If Len(!ReceiptNo) Then
                    Set Record = lstTrackMNRB.ListItems.Add(, , !ReceiptNo)
                Else
                    Set Record = lstTrackMNRB.ListItems.Add(, , "Empty")
                End If
                If Len(!ReceiptDate) Then
                    Record.SubItems(1) = Format(!ReceiptDate, "dd/mm/yyyy")
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(2) = Trim(!NoofCase)
                Else
                    Record.SubItems(2) = "0"
                End If
                If Len(!ReceiptMNRBAmt) Then
                   Record.SubItems(3) = Format(!ReceiptMNRBAmt, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                                                
                Current = Current + 1
                lblcurrent = Current
                If Len(!ReceiptMNRBAmt) Then
                    TotalAmt = TotalAmt + !ReceiptMNRBAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub
'******************* End ReceiptMNRB Tracking Listing *******************************

'Private Sub cmdPrint_Click()
    'Screen.MousePointer = vbHourglass
    'If lstReceiptTracking.listitems.count <> 0 Or lstTrackMNRB.listitems.count <> 0 Then
        
        'If OptMBM.value = True Then
            'RptReceiptTracking.show
        'ElseIf OptMNRB.value = True Then
            'RptMNRBTracking.show
        'End If
    'Else
        'MsgBox "Report cannot be printed without any record.", vbInformation
    'End If
    'Screen.MousePointer = Default
    
    
'End Sub


'**********************Start Sorted ListView *******************************
Private Sub lstReceiptTracking_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
If Option1.Value = True Then
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
    Case 3
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtString, m_blnDirection(8)
    End Select

ElseIf (Option2.Value = True Or Option3.Value = True) Then
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
    Case 3
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    End Select
    
ElseIf Option4.Value = True Then
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
    Case 3
        SortListView lstReceiptTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    End Select

End If

End Sub

Private Sub Option1_Click()
    
    TxtMBMNo.Text = ""
    TxtMBMNo.Enabled = False
    txtMBMName.Text = ""
    txtMBMName.Enabled = False
    TxtClaimNo1.Text = ""
    TxtClaimNo1.Enabled = False
    TxtClaimNo2.Text = ""
    TxtClaimNo2.Enabled = False
    txtDebitFr.Text = ""
    txtDebitFr.Enabled = False
    txtDebitTo.Text = ""
    txtDebitTo.Enabled = False
    txtRECNoFr.Text = ""
    txtRECNoFr.Enabled = True
    txtRECNoTo.Text = ""
    txtRECNoTo.Enabled = True
    txtAmtFr.Text = ""
    txtAmtFr.Enabled = False
    TxtAmtTo.Text = ""
    TxtAmtTo.Enabled = False
    cboACPeriodMthFr.Text = ""
    cboACPeriodMthFr.Enabled = False
    cboACPeriodMthTo.Text = ""
    cboACPeriodMthTo.Enabled = False
    cboACPeriodYRFr.Text = ""
    cboACPeriodYRFr.Enabled = False
    cboACPeriodYRTo.Text = ""
    cboACPeriodYRTo.Enabled = False
        
    If OptMBM.Value = True Then
        lstReceiptTracking.Visible = True
        lstTrackMNRB.ListItems.Clear
        lstTrackMNRB.Visible = False
        
        lstReceiptTracking.ColumnHeaders(1).Text = "Bordx No"
        lstReceiptTracking.ColumnHeaders(2).Text = "Bordx Date"
        lstReceiptTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstReceiptTracking.ColumnHeaders(3).Text = "No of Case"
        lstReceiptTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstReceiptTracking.ColumnHeaders(4).Text = "Amt Receipt Fr MBM"
        lstReceiptTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstReceiptTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf OptPayor.Value = True Then
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Bordx No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Bordx Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "No of Case"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Amt Receipt Fr Payor"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf OptMNRB.Value = True Then
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Bordx No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Bordx Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "No of Case"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Amt Receipt Fr MNRB"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
    End If
End Sub

Private Sub Option2_Click()
    
    TxtMBMNo.Enabled = True
    txtMBMName.Enabled = True
    TxtClaimNo1.Enabled = True
    TxtClaimNo2.Enabled = True
    txtDebitFr.Enabled = True
    txtDebitTo.Enabled = True
    txtRECNoFr.Enabled = True
    txtRECNoTo.Enabled = True
    txtAmtFr.Enabled = True
    TxtAmtTo.Enabled = True
    cboACPeriodMthFr.Enabled = True
    cboACPeriodMthTo.Enabled = True
    cboACPeriodYRFr.Enabled = True
    cboACPeriodYRTo.Enabled = True
    
    If OptMBM.Value = True Then
        lstReceiptTracking.Visible = True
        lstTrackMNRB.ListItems.Clear
        lstTrackMNRB.Visible = False
        
        lstReceiptTracking.ColumnHeaders(1).Text = "Receipt No"
        lstReceiptTracking.ColumnHeaders(2).Text = "Receipt Date"
        lstReceiptTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstReceiptTracking.ColumnHeaders(3).Text = "Amount"
        lstReceiptTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstReceiptTracking.ColumnHeaders(4).Text = "Payor"
        lstReceiptTracking.ColumnHeaders(4).Alignment = lvwColumnCenter
        lstReceiptTracking.ColumnHeaders(5).Text = "Member No"
        lstReceiptTracking.ColumnHeaders(5).Alignment = lvwColumnLeft
        lstReceiptTracking.ColumnHeaders(6).Text = "Member Name"
        lstReceiptTracking.ColumnHeaders(6).Alignment = lvwColumnLeft
        lstReceiptTracking.ColumnHeaders(7).Text = "Claim No"
        lstReceiptTracking.ColumnHeaders(7).Alignment = lvwColumnLeft
        lstReceiptTracking.ColumnHeaders(8).Text = "D/N No"
        lstReceiptTracking.ColumnHeaders(8).Alignment = lvwColumnLeft
        lstReceiptTracking.ColumnHeaders(9).Text = "Remarks"
        lstReceiptTracking.ColumnHeaders(9).Alignment = lvwColumnLeft
        lstReceiptTracking.ColumnHeaders(10).Text = "A/C Period"
        lstReceiptTracking.ColumnHeaders(10).Alignment = lvwColumnLeft
        
    ElseIf OptPayor.Value = True Then
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Receipt No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Receipt Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "Rec't Amt"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Member No"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(5).Text = "Member Name"
        lstTrackMNRB.ColumnHeaders(5).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(6).Text = "Bordx No"
        lstTrackMNRB.ColumnHeaders(6).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(7).Text = "Wks No"
        lstTrackMNRB.ColumnHeaders(7).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(8).Text = "D/N No"
        lstTrackMNRB.ColumnHeaders(8).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(9).Text = "Remarks"
        lstTrackMNRB.ColumnHeaders(9).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(10).Text = "A/C Period"
        lstTrackMNRB.ColumnHeaders(10).Alignment = lvwColumnLeft
        
    ElseIf OptMNRB.Value = True Then
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Receipt No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Receipt Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "Rec't Amt"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Member No"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(5).Text = "Member Name"
        lstTrackMNRB.ColumnHeaders(5).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(6).Text = "Bordx No"
        lstTrackMNRB.ColumnHeaders(6).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(7).Text = "Wks No"
        lstTrackMNRB.ColumnHeaders(7).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(8).Text = "D/N No"
        lstTrackMNRB.ColumnHeaders(8).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(9).Text = "Remarks"
        lstTrackMNRB.ColumnHeaders(9).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(10).Text = "A/C Period"
        lstTrackMNRB.ColumnHeaders(10).Alignment = lvwColumnLeft
        
    End If
End Sub

Private Sub Option3_Click()
    
    TxtMBMNo.Text = ""
    TxtMBMNo.Enabled = False
    txtMBMName.Text = ""
    txtMBMName.Enabled = False
    TxtClaimNo1.Text = ""
    TxtClaimNo1.Enabled = False
    TxtClaimNo2.Text = ""
    TxtClaimNo2.Enabled = False
    txtDebitFr.Text = ""
    txtDebitFr.Enabled = False
    txtDebitTo.Text = ""
    txtDebitTo.Enabled = False
    txtRECNoFr.Text = ""
    txtRECNoFr.Enabled = True
    txtRECNoTo.Text = ""
    txtRECNoTo.Enabled = True
    txtAmtFr.Text = ""
    txtAmtFr.Enabled = False
    TxtAmtTo.Text = ""
    TxtAmtTo.Enabled = False
    cboACPeriodMthFr.Text = ""
    cboACPeriodMthFr.Enabled = False
    cboACPeriodMthTo.Text = ""
    cboACPeriodMthTo.Enabled = False
    cboACPeriodYRFr.Text = ""
    cboACPeriodYRFr.Enabled = False
    cboACPeriodYRTo.Text = ""
    cboACPeriodYRTo.Enabled = False
        
    If OptMBM.Value = True Then
        lstReceiptTracking.Visible = True
        lstTrackMNRB.ListItems.Clear
        lstTrackMNRB.Visible = False
        
        lstReceiptTracking.ColumnHeaders(1).Text = "Receipt No"
        lstReceiptTracking.ColumnHeaders(2).Text = "Receipt Date"
        lstReceiptTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstReceiptTracking.ColumnHeaders(3).Text = "No of Case"
        lstReceiptTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstReceiptTracking.ColumnHeaders(4).Text = "Amt Receipt Fr MBM"
        lstReceiptTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstReceiptTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf OptPayor.Value = True Then
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Receipt No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Receipt Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "No of Case"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Amt Receipt Fr Payor"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf OptMNRB.Value = True Then
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Receipt No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Receipt Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "No of Case"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Amt Receipt Fr MNRB"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
    End If
End Sub

Private Sub Option4_Click()
    TxtMBMNo.Text = ""
    TxtMBMNo.Enabled = False
    txtMBMName.Text = ""
    txtMBMName.Enabled = False
    TxtClaimNo1.Text = ""
    TxtClaimNo1.Enabled = False
    TxtClaimNo2.Text = ""
    TxtClaimNo2.Enabled = False
    txtDebitFr.Text = ""
    txtDebitFr.Enabled = False
    txtDebitTo.Text = ""
    txtDebitTo.Enabled = False
    txtRECNoFr.Text = ""
    txtRECNoFr.Enabled = True
    txtRECNoTo.Text = ""
    txtRECNoTo.Enabled = True
    txtAmtFr.Text = ""
    txtAmtFr.Enabled = False
    TxtAmtTo.Text = ""
    TxtAmtTo.Enabled = False
    cboACPeriodMthFr.Text = ""
    cboACPeriodMthFr.Enabled = False
    cboACPeriodMthTo.Text = ""
    cboACPeriodMthTo.Enabled = False
    cboACPeriodYRFr.Text = ""
    cboACPeriodYRFr.Enabled = False
    cboACPeriodYRTo.Text = ""
    cboACPeriodYRTo.Enabled = False
        
    If OptMBM.Value = True Then
        lstReceiptTracking.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstTrackMNRB.Visible = False
        
        lstReceiptTracking.ColumnHeaders(1).Text = "Payor"
        lstReceiptTracking.ColumnHeaders(2).Text = "No of Receipt"
        lstReceiptTracking.ColumnHeaders(2).Alignment = lvwColumnRight
        lstReceiptTracking.ColumnHeaders(3).Text = "Amt Receipt Fr MBM"
        lstReceiptTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        
        For AA = 4 To 10
            lstReceiptTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf OptPayor.Value = True Then
        lstTrackMNRB.Visible = True
        lstTrackMNRB.ListItems.Clear
        lstReceiptTracking.Visible = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Payor"
        lstTrackMNRB.ColumnHeaders(2).Text = "No of Receipt"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(3).Text = "Amt Receipt Fr Payor"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        
        For AA = 4 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
    ElseIf OptMNRB.Value = True Then
        lstTrackMNRB.Visible = True
        lstTrackMNRB.ListItems.Clear
        lstReceiptTracking.Visible = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Payor"
        lstTrackMNRB.ColumnHeaders(2).Text = "No of Receipt"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(3).Text = "Amt Receipt Fr MNRB"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        
        For AA = 4 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
    End If
End Sub

Private Sub OptPayor_Click()
    
    If Option2.Value = True Then
        
        TxtMBMNo.Enabled = True
        txtMBMName.Enabled = True
        TxtClaimNo1.Enabled = True
        TxtClaimNo2.Enabled = True
        txtDebitFr.Enabled = True
        txtDebitTo.Enabled = True
        txtRECNoFr.Enabled = True
        txtRECNoTo.Enabled = True
        txtAmtFr.Enabled = True
        TxtAmtTo.Enabled = True
        cboACPeriodMthFr.Enabled = True
        cboACPeriodMthTo.Enabled = True
        cboACPeriodYRFr.Enabled = True
        cboACPeriodYRTo.Enabled = True
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Receipt No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Receipt Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "Rec't Amt"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Member No"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(5).Text = "Member Name"
        lstTrackMNRB.ColumnHeaders(5).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(6).Text = "Bordx No"
        lstTrackMNRB.ColumnHeaders(6).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(7).Text = "Wks No"
        lstTrackMNRB.ColumnHeaders(7).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(8).Text = "D/N No"
        lstTrackMNRB.ColumnHeaders(8).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(9).Text = "Remarks"
        lstTrackMNRB.ColumnHeaders(9).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(10).Text = "A/C Period"
        lstTrackMNRB.ColumnHeaders(10).Alignment = lvwColumnLeft
        
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
    
    ElseIf Option1.Value = True Then
    
        TxtMBMNo.Text = ""
        TxtMBMNo.Enabled = False
        txtMBMName.Text = ""
        txtMBMName.Enabled = False
        TxtClaimNo1.Text = ""
        TxtClaimNo1.Enabled = False
        TxtClaimNo2.Text = ""
        TxtClaimNo2.Enabled = False
        txtDebitFr.Text = ""
        txtDebitFr.Enabled = False
        txtDebitTo.Text = ""
        txtDebitTo.Enabled = False
        txtRECNoFr.Text = ""
        txtRECNoFr.Enabled = False
        txtRECNoTo.Text = ""
        txtRECNoTo.Enabled = False
        txtAmtFr.Text = ""
        txtAmtFr.Enabled = False
        TxtAmtTo.Text = ""
        TxtAmtTo.Enabled = False
        cboACPeriodMthFr.Text = ""
        cboACPeriodMthFr.Enabled = False
        cboACPeriodMthTo.Text = ""
        cboACPeriodMthTo.Enabled = False
        cboACPeriodYRFr.Text = ""
        cboACPeriodYRFr.Enabled = False
        cboACPeriodYRTo.Text = ""
        cboACPeriodYRTo.Enabled = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Bordx No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Bordx Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "No of Case"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Amt Receipt Fr Payor"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
        
    ElseIf Option3.Value = True Then
    
        TxtMBMNo.Text = ""
        TxtMBMNo.Enabled = False
        txtMBMName.Text = ""
        txtMBMName.Enabled = False
        TxtClaimNo1.Text = ""
        TxtClaimNo1.Enabled = False
        TxtClaimNo2.Text = ""
        TxtClaimNo2.Enabled = False
        txtDebitFr.Text = ""
        txtDebitFr.Enabled = False
        txtDebitTo.Text = ""
        txtDebitTo.Enabled = False
        txtRECNoFr.Text = ""
        txtRECNoFr.Enabled = False
        txtRECNoTo.Text = ""
        txtRECNoTo.Enabled = False
        txtAmtFr.Text = ""
        txtAmtFr.Enabled = False
        TxtAmtTo.Text = ""
        TxtAmtTo.Enabled = False
        cboACPeriodMthFr.Text = ""
        cboACPeriodMthFr.Enabled = False
        cboACPeriodMthTo.Text = ""
        cboACPeriodMthTo.Enabled = False
        cboACPeriodYRFr.Text = ""
        cboACPeriodYRFr.Enabled = False
        cboACPeriodYRTo.Text = ""
        cboACPeriodYRTo.Enabled = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Receipt No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Receipt Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "No of Case"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Amt Receipt Fr Payor"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
        
    ElseIf Option4.Value = True Then
    
        TxtMBMNo.Text = ""
        TxtMBMNo.Enabled = False
        txtMBMName.Text = ""
        txtMBMName.Enabled = False
        TxtClaimNo1.Text = ""
        TxtClaimNo1.Enabled = False
        TxtClaimNo2.Text = ""
        TxtClaimNo2.Enabled = False
        txtDebitFr.Text = ""
        txtDebitFr.Enabled = False
        txtDebitTo.Text = ""
        txtDebitTo.Enabled = False
        txtRECNoFr.Text = ""
        txtRECNoFr.Enabled = True
        txtRECNoTo.Text = ""
        txtRECNoTo.Enabled = True
        txtAmtFr.Text = ""
        txtAmtFr.Enabled = False
        TxtAmtTo.Text = ""
        TxtAmtTo.Enabled = False
        cboACPeriodMthFr.Text = ""
        cboACPeriodMthFr.Enabled = False
        cboACPeriodMthTo.Text = ""
        cboACPeriodMthTo.Enabled = False
        cboACPeriodYRFr.Text = ""
        cboACPeriodYRFr.Enabled = False
        cboACPeriodYRTo.Text = ""
        cboACPeriodYRTo.Enabled = False
    
        lstTrackMNRB.ColumnHeaders(1).Text = "Payor"
        lstTrackMNRB.ColumnHeaders(2).Text = "No of Receipt"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(3).Text = "Amt Receipt Fr Payor"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        
        For AA = 4 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
    
    End If
End Sub

'**********************End Sorted ListView *******************************
Private Sub txtAmtFr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAmtTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

'******************* Start Option Button *******************************

Private Sub OptMBM_Click()
    
    If Option2.Value = True Then
        
        TxtMBMNo.Enabled = True
        txtMBMName.Enabled = True
        TxtClaimNo1.Enabled = True
        TxtClaimNo2.Enabled = True
        txtDebitFr.Enabled = True
        txtDebitTo.Enabled = True
        txtRECNoFr.Enabled = True
        txtRECNoTo.Enabled = True
        txtAmtFr.Enabled = True
        TxtAmtTo.Enabled = True
        cboACPeriodMthFr.Enabled = True
        cboACPeriodMthTo.Enabled = True
        cboACPeriodYRFr.Enabled = True
        cboACPeriodYRTo.Enabled = True
        
        lstReceiptTracking.ColumnHeaders(1).Text = "Receipt No"
        lstReceiptTracking.ColumnHeaders(2).Text = "Receipt Date"
        lstReceiptTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstReceiptTracking.ColumnHeaders(3).Text = "Amount"
        lstReceiptTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstReceiptTracking.ColumnHeaders(4).Text = "Payor"
        lstReceiptTracking.ColumnHeaders(4).Alignment = lvwColumnCenter
        lstReceiptTracking.ColumnHeaders(5).Text = "Member No"
        lstReceiptTracking.ColumnHeaders(5).Alignment = lvwColumnLeft
        lstReceiptTracking.ColumnHeaders(6).Text = "Member Name"
        lstReceiptTracking.ColumnHeaders(6).Alignment = lvwColumnLeft
        lstReceiptTracking.ColumnHeaders(7).Text = "Claim No"
        lstReceiptTracking.ColumnHeaders(7).Alignment = lvwColumnLeft
        lstReceiptTracking.ColumnHeaders(8).Text = "D/N No"
        lstReceiptTracking.ColumnHeaders(8).Alignment = lvwColumnLeft
        lstReceiptTracking.ColumnHeaders(9).Text = "Remarks"
        lstReceiptTracking.ColumnHeaders(9).Alignment = lvwColumnLeft
        lstReceiptTracking.ColumnHeaders(10).Text = "A/C Period"
        lstReceiptTracking.ColumnHeaders(10).Alignment = lvwColumnLeft
        
        lstReceiptTracking.Visible = True
        lstTrackMNRB.ListItems.Clear
        lstTrackMNRB.Visible = False
    
    ElseIf Option1.Value = True Then
        
        TxtMBMNo.Text = ""
        TxtMBMNo.Enabled = False
        txtMBMName.Text = ""
        txtMBMName.Enabled = False
        TxtClaimNo1.Text = ""
        TxtClaimNo1.Enabled = False
        TxtClaimNo2.Text = ""
        TxtClaimNo2.Enabled = False
        txtDebitFr.Text = ""
        txtDebitFr.Enabled = False
        txtDebitTo.Text = ""
        txtDebitTo.Enabled = False
        txtRECNoFr.Text = ""
        txtRECNoFr.Enabled = False
        txtRECNoTo.Text = ""
        txtRECNoTo.Enabled = False
        txtAmtFr.Text = ""
        txtAmtFr.Enabled = False
        TxtAmtTo.Text = ""
        TxtAmtTo.Enabled = False
        cboACPeriodMthFr.Text = ""
        cboACPeriodMthFr.Enabled = False
        cboACPeriodMthTo.Text = ""
        cboACPeriodMthTo.Enabled = False
        cboACPeriodYRFr.Text = ""
        cboACPeriodYRFr.Enabled = False
        cboACPeriodYRTo.Text = ""
        cboACPeriodYRTo.Enabled = False
        
        lstReceiptTracking.ColumnHeaders(1).Text = "Bordx No"
        lstReceiptTracking.ColumnHeaders(2).Text = "Bordx Date"
        lstReceiptTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstReceiptTracking.ColumnHeaders(3).Text = "No of Case"
        lstReceiptTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstReceiptTracking.ColumnHeaders(4).Text = "Amt Receipt Fr MBM"
        lstReceiptTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstReceiptTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
        lstReceiptTracking.Visible = True
        lstTrackMNRB.ListItems.Clear
        lstTrackMNRB.Visible = False
    
    ElseIf Option3.Value = True Then
        
        TxtMBMNo.Text = ""
        TxtMBMNo.Enabled = False
        txtMBMName.Text = ""
        txtMBMName.Enabled = False
        TxtClaimNo1.Text = ""
        TxtClaimNo1.Enabled = False
        TxtClaimNo2.Text = ""
        TxtClaimNo2.Enabled = False
        txtDebitFr.Text = ""
        txtDebitFr.Enabled = False
        txtDebitTo.Text = ""
        txtDebitTo.Enabled = False
        txtRECNoFr.Text = ""
        txtRECNoFr.Enabled = False
        txtRECNoTo.Text = ""
        txtRECNoTo.Enabled = False
        txtAmtFr.Text = ""
        txtAmtFr.Enabled = False
        TxtAmtTo.Text = ""
        TxtAmtTo.Enabled = False
        cboACPeriodMthFr.Text = ""
        cboACPeriodMthFr.Enabled = False
        cboACPeriodMthTo.Text = ""
        cboACPeriodMthTo.Enabled = False
        cboACPeriodYRFr.Text = ""
        cboACPeriodYRFr.Enabled = False
        cboACPeriodYRTo.Text = ""
        cboACPeriodYRTo.Enabled = False
        
        lstReceiptTracking.ColumnHeaders(1).Text = "Receipt No"
        lstReceiptTracking.ColumnHeaders(2).Text = "Receipt Date"
        lstReceiptTracking.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstReceiptTracking.ColumnHeaders(3).Text = "No of Case"
        lstReceiptTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        lstReceiptTracking.ColumnHeaders(4).Text = "Amt Receipt Fr MBM"
        lstReceiptTracking.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstReceiptTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
        lstReceiptTracking.Visible = True
        lstTrackMNRB.ListItems.Clear
        lstTrackMNRB.Visible = False
        
    ElseIf Option4.Value = True Then
        
        TxtMBMNo.Text = ""
        TxtMBMNo.Enabled = False
        txtMBMName.Text = ""
        txtMBMName.Enabled = False
        TxtClaimNo1.Text = ""
        TxtClaimNo1.Enabled = False
        TxtClaimNo2.Text = ""
        TxtClaimNo2.Enabled = False
        txtDebitFr.Text = ""
        txtDebitFr.Enabled = False
        txtDebitTo.Text = ""
        txtDebitTo.Enabled = False
        txtRECNoFr.Text = ""
        txtRECNoFr.Enabled = True
        txtRECNoTo.Text = ""
        txtRECNoTo.Enabled = True
        txtAmtFr.Text = ""
        txtAmtFr.Enabled = False
        TxtAmtTo.Text = ""
        TxtAmtTo.Enabled = False
        cboACPeriodMthFr.Text = ""
        cboACPeriodMthFr.Enabled = False
        cboACPeriodMthTo.Text = ""
        cboACPeriodMthTo.Enabled = False
        cboACPeriodYRFr.Text = ""
        cboACPeriodYRFr.Enabled = False
        cboACPeriodYRTo.Text = ""
        cboACPeriodYRTo.Enabled = False
        
        lstReceiptTracking.ColumnHeaders(1).Text = "Payor"
        lstReceiptTracking.ColumnHeaders(2).Text = "No of Receipt"
        lstReceiptTracking.ColumnHeaders(2).Alignment = lvwColumnRight
        lstReceiptTracking.ColumnHeaders(3).Text = "Amt Receipt Fr MBM"
        lstReceiptTracking.ColumnHeaders(3).Alignment = lvwColumnRight
        
        For AA = 4 To 10
            lstReceiptTracking.ColumnHeaders(AA).Text = ""
        Next AA
        
        lstReceiptTracking.Visible = True
        lstTrackMNRB.ListItems.Clear
        lstTrackMNRB.Visible = False
    
    End If
End Sub

Private Sub OptMNRB_Click()
    If Option2.Value = True Then
        
        TxtMBMNo.Enabled = True
        txtMBMName.Enabled = True
        TxtClaimNo1.Enabled = True
        TxtClaimNo2.Enabled = True
        txtDebitFr.Enabled = True
        txtDebitTo.Enabled = True
        txtRECNoFr.Enabled = True
        txtRECNoTo.Enabled = True
        txtAmtFr.Enabled = True
        TxtAmtTo.Enabled = True
        cboACPeriodMthFr.Enabled = True
        cboACPeriodMthTo.Enabled = True
        cboACPeriodYRFr.Enabled = True
        cboACPeriodYRTo.Enabled = True
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Receipt No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Receipt Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "Rec't Amt"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Member No"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(5).Text = "Member Name"
        lstTrackMNRB.ColumnHeaders(5).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(6).Text = "Bordx No"
        lstTrackMNRB.ColumnHeaders(6).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(7).Text = "Wks No"
        lstTrackMNRB.ColumnHeaders(7).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(8).Text = "D/N No"
        lstTrackMNRB.ColumnHeaders(8).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(9).Text = "Remarks"
        lstTrackMNRB.ColumnHeaders(9).Alignment = lvwColumnLeft
        lstTrackMNRB.ColumnHeaders(10).Text = "A/C Period"
        lstTrackMNRB.ColumnHeaders(10).Alignment = lvwColumnLeft
        
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
    
    ElseIf Option1.Value = True Then
        
        TxtMBMNo.Text = ""
        TxtMBMNo.Enabled = False
        txtMBMName.Text = ""
        txtMBMName.Enabled = False
        TxtClaimNo1.Text = ""
        TxtClaimNo1.Enabled = False
        TxtClaimNo2.Text = ""
        TxtClaimNo2.Enabled = False
        txtDebitFr.Text = ""
        txtDebitFr.Enabled = False
        txtDebitTo.Text = ""
        txtDebitTo.Enabled = False
        txtRECNoFr.Text = ""
        txtRECNoFr.Enabled = False
        txtRECNoTo.Text = ""
        txtRECNoTo.Enabled = False
        txtAmtFr.Text = ""
        txtAmtFr.Enabled = False
        TxtAmtTo.Text = ""
        TxtAmtTo.Enabled = False
        cboACPeriodMthFr.Text = ""
        cboACPeriodMthFr.Enabled = False
        cboACPeriodMthTo.Text = ""
        cboACPeriodMthTo.Enabled = False
        cboACPeriodYRFr.Text = ""
        cboACPeriodYRFr.Enabled = False
        cboACPeriodYRTo.Text = ""
        cboACPeriodYRTo.Enabled = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Bordx No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Bordx Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "No of Case"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Amt Receipt Fr MNRB"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
        
    ElseIf Option3.Value = True Then
        
        TxtMBMNo.Text = ""
        TxtMBMNo.Enabled = False
        txtMBMName.Text = ""
        txtMBMName.Enabled = False
        TxtClaimNo1.Text = ""
        TxtClaimNo1.Enabled = False
        TxtClaimNo2.Text = ""
        TxtClaimNo2.Enabled = False
        txtDebitFr.Text = ""
        txtDebitFr.Enabled = False
        txtDebitTo.Text = ""
        txtDebitTo.Enabled = False
        txtRECNoFr.Text = ""
        txtRECNoFr.Enabled = False
        txtRECNoTo.Text = ""
        txtRECNoTo.Enabled = False
        txtAmtFr.Text = ""
        txtAmtFr.Enabled = False
        TxtAmtTo.Text = ""
        TxtAmtTo.Enabled = False
        cboACPeriodMthFr.Text = ""
        cboACPeriodMthFr.Enabled = False
        cboACPeriodMthTo.Text = ""
        cboACPeriodMthTo.Enabled = False
        cboACPeriodYRFr.Text = ""
        cboACPeriodYRFr.Enabled = False
        cboACPeriodYRTo.Text = ""
        cboACPeriodYRTo.Enabled = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Receipt No"
        lstTrackMNRB.ColumnHeaders(2).Text = "Receipt Date"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstTrackMNRB.ColumnHeaders(3).Text = "No of Case"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(4).Text = "Amt Receipt Fr MNRB"
        lstTrackMNRB.ColumnHeaders(4).Alignment = lvwColumnRight
        
        For AA = 5 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
        
        lstTrackMNRB.Visible = True
        lstReceiptTracking.ListItems.Clear
        lstReceiptTracking.Visible = False
        
    ElseIf Option4.Value = True Then
    
        TxtMBMNo.Text = ""
        TxtMBMNo.Enabled = False
        txtMBMName.Text = ""
        txtMBMName.Enabled = False
        TxtClaimNo1.Text = ""
        TxtClaimNo1.Enabled = False
        TxtClaimNo2.Text = ""
        TxtClaimNo2.Enabled = False
        txtDebitFr.Text = ""
        txtDebitFr.Enabled = False
        txtDebitTo.Text = ""
        txtDebitTo.Enabled = False
        txtRECNoFr.Text = ""
        txtRECNoFr.Enabled = True
        txtRECNoTo.Text = ""
        txtRECNoTo.Enabled = True
        txtAmtFr.Text = ""
        txtAmtFr.Enabled = False
        TxtAmtTo.Text = ""
        TxtAmtTo.Enabled = False
        cboACPeriodMthFr.Text = ""
        cboACPeriodMthFr.Enabled = False
        cboACPeriodMthTo.Text = ""
        cboACPeriodMthTo.Enabled = False
        cboACPeriodYRFr.Text = ""
        cboACPeriodYRFr.Enabled = False
        cboACPeriodYRTo.Text = ""
        cboACPeriodYRTo.Enabled = False
        
        lstTrackMNRB.Visible = True
        lstTrackMNRB.ListItems.Clear
        lstReceiptTracking.Visible = False
        
        lstTrackMNRB.ColumnHeaders(1).Text = "Payor"
        lstTrackMNRB.ColumnHeaders(2).Text = "No of Receipt"
        lstTrackMNRB.ColumnHeaders(2).Alignment = lvwColumnRight
        lstTrackMNRB.ColumnHeaders(3).Text = "Amt Receipt Fr MNRB"
        lstTrackMNRB.ColumnHeaders(3).Alignment = lvwColumnRight
        
        For AA = 4 To 10
            lstTrackMNRB.ColumnHeaders(AA).Text = ""
        Next AA
    
    End If
End Sub
'******************* End Option Button *******************************

Private Sub CboPayor_Change()
If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayor.Text = ValidValue
                CboPayor.SelStart = 0
                CboPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub txtRecDate1_LostFocus()
    If IsDate(txtRecDate1) Then
    Else
        If txtRecDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRecDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtRecDate2_LostFocus()
    If IsDate(txtRecDate2) Then
    Else
        If txtRecDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtRecDate2.SetFocus
        End If
    End If
End Sub

Private Sub cboHLTCode_Change()
If cboHLTCode.Text = "" Then
    If InStr(cboHLTCode.Text, "null") Then
       cboHLTCode.Enabled = False
    End If
End If
End Sub

Private Sub cboHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHLTCode.Text, cboHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode.Text, cboHLTCode.SelStart + cboHLTCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboHLTCode.ListCount - 1
 
        If NewText = LCase(Left(cboHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHLTCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboHLTCode.Text = ValidValue
                cboHLTCode.SelStart = 0
                cboHLTCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub ReceiptMemberByPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstReceiptTracking.ListItems.Clear
            
    'Sql = " Select PAYCode, COUNT(ReceiptNo) AS NoofCase, SUM(ReceiptMBMAmt) AS ReceiptMBMAmt From VIEW_ReceiptfrMBMTracking2 where 0 = 0 "
    
    'sql2 = "Group By PAYCode "
         
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend + sql2
    'Else
    '    Sql = Sql + sql2
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "10"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
                
            With rst2
                If Len(!PAYCode) Then
                    Set Record = lstReceiptTracking.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = lstReceiptTracking.ListItems.Add(, , "Empty")
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(1) = Trim(!NoofCase)
                Else
                    Record.SubItems(1) = "0"
                End If
                If Len(!ReceiptMBMAmt) Then
                    Record.SubItems(2) = Format(!ReceiptMBMAmt, "#,##0.00")
                Else
                    Record.SubItems(2) = "0.00"
                End If
                                                
                Current = Current + 1
                lblcurrent = Current
                If Len(!ReceiptMBMAmt) Then
                    TotalAmt = TotalAmt + !ReceiptMBMAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub ReceiptPayorByPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstTrackMNRB.ListItems.Clear
            
    'Sql = " Select PAYCode, COUNT(ReceiptNo) AS NoofCase, SUM(ReceiptPAYAmt) AS ReceiptPAYAmt From VIEW_ReceiptfrPayorTracking2 where 0 = 0 "
    
    'sql2 = "Group By PAYCode "
     
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend + sql2
    'Else
    '    Sql = Sql + sql2
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "11"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
        
            With rst2
                If Len(!PAYCode) Then
                    Set Record = lstTrackMNRB.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = lstTrackMNRB.ListItems.Add(, , "Empty")
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(1) = Trim(!NoofCase)
                Else
                    Record.SubItems(1) = "0"
                End If
                If Len(!ReceiptPAYAmt) Then
                    Record.SubItems(2) = Format(!ReceiptPAYAmt, "#,##0.00")
                Else
                    Record.SubItems(2) = "0.00"
                End If
                                                
                Current = Current + 1
                lblcurrent = Current
                If Len(!ReceiptPAYAmt) Then
                    TotalAmt = TotalAmt + !ReceiptPAYAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub ReceiptMNRBByPayor(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    total = 0
    Current = 0
    TotalAmt = 0
    
    lstTrackMNRB.ListItems.Clear
            
    'Sql = " Select PAYCode, COUNT(ReceiptNo) AS NoofCase, SUM(ReceiptMNRBAmt) AS ReceiptMNRBAmt From VIEW_ReceiptfrMNRBTracking2 where 0 = 0 "
    
    'sql2 = " Group By PAYCode "
     
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend + sql2
    'Else
    '    Sql = Sql + sql2
    'End If
    
    'rst2.Open Sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "12"
    Set cmd.ActiveConnection = medixmasterconnPayment
        cmd.CommandText = "ReceiptTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdEXPtoExcel.Enabled = True
        cmdPrint.Enabled = True
        CmdSearch.Enabled = True
    Else
        Do Until rst2.EOF
                
            With rst2
                If Len(!PAYCode) Then
                    Set Record = lstTrackMNRB.ListItems.Add(, , !PAYCode)
                Else
                    Set Record = lstTrackMNRB.ListItems.Add(, , "Empty")
                End If
                If Len(!NoofCase) Then
                    Record.SubItems(1) = Trim(!NoofCase)
                Else
                    Record.SubItems(1) = "0"
                End If
                If Len(!ReceiptMNRBAmt) Then
                   Record.SubItems(2) = Format(!ReceiptMNRBAmt, "#,##0.00")
                Else
                    Record.SubItems(2) = "0.00"
                End If
                                                
                Current = Current + 1
                lblcurrent = Current
                If Len(!ReceiptMNRBAmt) Then
                    TotalAmt = TotalAmt + !ReceiptMNRBAmt
                End If
                lbltotalAmt = Format(TotalAmt, "#,##0.00")
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    cmdPrint.Enabled = True
    CmdSearch.Enabled = True
End Sub


