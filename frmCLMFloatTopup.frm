VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCLMFloatTopup 
   Caption         =   "Claim Float Topup"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CheckAll 
      Caption         =   "Check All"
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame FrameReceiptPV 
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   6120
      Width           =   11625
      Begin VB.TextBox txtCHQNo 
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtTURecNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   10
         Top             =   200
         Width           =   1575
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   6360
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   240
         Width           =   5175
      End
      Begin MSMask.MaskEdBox txtTURecDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtTURecAmt 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   12
         Top             =   680
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtCHQDate 
         Height          =   285
         Left            =   3840
         TabIndex        =   14
         Top             =   465
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Chq No"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Chq Date"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   48
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "TU Rec Date"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "TU Rec No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Topup Rec Amt"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   700
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1380
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   11655
      Begin VB.TextBox txtTUNoFr 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   0
         Top             =   195
         Width           =   1335
      End
      Begin VB.TextBox txtTUNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   1
         Top             =   195
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtTUDateFr 
         Height          =   315
         Left            =   7920
         TabIndex        =   6
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTUDateTo 
         Height          =   315
         Left            =   9840
         TabIndex        =   7
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTURecDateFr 
         Height          =   315
         Left            =   7920
         TabIndex        =   8
         Top             =   485
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTURecDateTo 
         Height          =   315
         Left            =   9840
         TabIndex        =   9
         Top             =   485
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   2
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   3
         Top             =   450
         Width           =   1335
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   705
         Width           =   3255
      End
      Begin VB.ComboBox CboTopupStatus 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   980
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
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
         Left            =   240
         TabIndex        =   47
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "TU Rec No from"
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
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "To"
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
         Index           =   0
         Left            =   3285
         TabIndex        =   45
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "TU Rec Date from"
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
         Index           =   3
         Left            =   6360
         TabIndex        =   44
         Top             =   525
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Index           =   3
         Left            =   9360
         TabIndex        =   43
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "TU Date from"
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
         Index           =   2
         Left            =   6360
         TabIndex        =   21
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "To"
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
         Index           =   2
         Left            =   9360
         TabIndex        =   20
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "TU No from"
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
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "To"
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
         Index           =   3
         Left            =   3285
         TabIndex        =   18
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Topup Status"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1000
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstTUListing 
      Height          =   4200
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7408
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TU Rec No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "TU Rec Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "TU Rec Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "TU No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "TU Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Req Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Chq No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Chq Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "TUCode"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   28
      Top             =   7200
      Width           =   11655
      Begin VB.TextBox lblSelectedAmt 
         Height          =   285
         Left            =   5160
         TabIndex        =   57
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox lbltotalAmt 
         Height          =   285
         Left            =   2880
         TabIndex        =   56
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox lblcurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   6720
         TabIndex        =   38
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S&earch"
         Height          =   330
         Left            =   8040
         TabIndex        =   37
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   8760
         TabIndex        =   36
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   35
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10320
         TabIndex        =   34
         Top             =   220
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Height          =   330
         Left            =   7320
         TabIndex        =   31
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "&Print to File"
         Height          =   330
         Left            =   9360
         TabIndex        =   30
         Top             =   220
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   7320
         TabIndex        =   29
         Top             =   220
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Topup  Rec Amt"
         Height          =   255
         Left            =   1605
         TabIndex        =   41
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   840
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   4065
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label lblMG 
      Caption         =   "MedicaGen Float Balance :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   53
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label lblMGAmt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   52
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "MedicaLife Float Balance :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label lblMLAmt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   50
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claim Float Topup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmCLMFloatTopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim tmpPrevRecAmt As Double
Dim tmpTUNo As String
Dim intExit As Integer
Dim tmpTURecNo As String
Dim tmpChkFloatAmt As Double
Dim tmpPayCode As String
Dim tmpFloatBal As Double
Dim tmpCOVPayCode As String
Dim tmpchkPAYCode As String
Dim m_blnDirection(1 To 11) As Boolean


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

Private Sub CboTopupStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboTopupStatus.Text, CboTopupStatus.SelStart) + Chr(KeyAscii) + Mid(CboTopupStatus.Text, CboTopupStatus.SelStart + CboTopupStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboTopupStatus.ListCount - 1
 
        If NewText = LCase(Left(CboTopupStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboTopupStatus.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboTopupStatus.Text = ValidValue
                CboTopupStatus.SelStart = 0
                CboTopupStatus.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub chkCheckItems_Click()

End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    CboTopupStatus.ListIndex = 0
    lstTUListing.ListItems.Clear
    lblSelectedAmt = 0
    txtTURecAmt = 0
    EditFlag = False
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    CmdEdit.Visible = False
    CmdUpdate.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdEXPtoExcel_Click()
    Screen.MousePointer = vbHourglass
    
    If lstTUListing.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstTUListing)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSave_Click()
Dim AA As Integer
Dim ItmChk As Boolean
Dim BB As Integer
Dim tmpSql As String
Dim tmpTUNo As String
Dim startTrans As Boolean
Dim bSucess  As Boolean
Dim DiffPay As Boolean
Dim UpdateStatus As Boolean

    ItmChk = False
    DiffPay = False
    tmpPayCode = ""
        AA = lstTUListing.ListItems.Count
        For BB = 1 To lstTUListing.ListItems.Count
            If lstTUListing.ListItems(BB).Checked = True Then
                If ItmChk = False Then
                    ItmChk = True
                End If
                    If tmpchkPAYCode <> "" Then
                        If tmpchkPAYCode <> lstTUListing.ListItems(BB).SubItems(10) Then
                            DiffPay = True
                            Exit For
                        End If
                    End If
                    tmpPayCode = Trim(IIf(lstTUListing.ListItems(BB).SubItems(10) = "TUG", "TUG", "TUL"))
                    tmpchkPAYCode = lstTUListing.ListItems(BB).SubItems(10)
            End If
        Next BB
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        tmpCOVPayCode = IIf(Trim(tmpPayCode) = "TUG", "XG", "XL")
        'IIf(IsNumeric(txtINVAmount), txtINVAmount, 0)
        Call SearchFloatAmt("And FTPayCode = '" & Trim(tmpCOVPayCode) & "'")
        'tmpFloatBal = CDbl(tmpChkFloatAmt) - CDbl(txtTURecAmt)
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    
    If AA = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf ItmChk = False Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf txtTURecDate = "__/__/____" Then
        MsgBox "Please Enter FT Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtTURecDate.SetFocus
    ElseIf Trim(txtchqNo) = "" Then
       MsgBox "Please Enter Cheque No!", vbCritical
       txtchqNo.SetFocus
    ElseIf txtchqDate = "__/__/____" Then
        MsgBox "Please Enter Cheque Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtchqDate.SetFocus
    'ElseIf CDbl(txtTURecAmt) = "" Then
    '    MsgBox "Please Enter TU Receipt Amount!", vbCritical
    '    txtTURecAmt.SetFocus
    ElseIf DiffPay = True Then
        MsgBox "Different Payors for this transaction is not allowed! ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf CDbl(lblSelectedAmt) <> CDbl(txtTURecAmt) Then
        MsgBox "Please check Float Transfer Amount! ", vbOKOnly + vbCritical, DialogBoxTitle
    'ElseIf CDbl(tmpFloatBal) <= 0 Then
    '    MsgBox "Insufficient Float! ", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Screen.MousePointer = vbHourglass
     '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
     '   medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
                        
        startTrans = True
        medixmasterconnWS.beginTrans
        medixmasterconnPayment.beginTrans
            
            tmpTURecNo = GenerateTURecNo(tmpPayCode)
            
            
            For BB = 1 To lstTUListing.ListItems.Count
                If lstTUListing.ListItems(BB).Checked = True Then
                    tmpTUNo = lstTUListing.ListItems(BB).SubItems(4)
                    UpdateStatus = True
                    Call UpdateFTFloat(tmpTUNo)
                End If
            Next BB
            If UpdateStatus = True Then
                Call SaveTUFloat
            End If

            Call UpdateFloatAmt
            medixmasterconnWS.commitTrans
            medixmasterconnPayment.commitTrans
            startTrans = False
            txtTURecNo = tmpTURecNo
            'TxtReceiptNo = tmpRecNo
            'txtPVNo = tmpFTPVNo
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            'If tmpFloatBal < 300000 Then
            '    MsgBox "Float Balance is " & tmpFloatBal & " for " & tmpCOVPayCode, vbOKOnly + vbInformation, DialogBoxTitle
            'End If
            
            tmpSql = ""
            tmpSql = " and  CLMTURecNo = '" & Trim(tmpTURecNo) & "'"
            lstTUListing.ListItems.Clear
            Call CLMTUListing(tmpSql)
        EditFlag = False
        
        
        
     '   medixmasterconnWS.Close
    '    Set medixmasterconnWS = Nothing
     '   medixmasterconnPayment.Close
     '   Set medixmasterconnPayment = Nothing
        Screen.MousePointer = vbDefault
        'Call ClearContraDetail
        'Call ClearRecDetail
        lblSelectedAmt = 0
        lbltotalAmt = 0
    End If
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If bSucess = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnPayment.RollbackTrans
    End If
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnPayment.Close
   ' Set medixmasterconnPayment = Nothing
    MsgBox Err.Description
End Sub

Private Sub CmdSearch_Click()
On Error GoTo errRun
Dim strAppend As String
    strAppend = ""
    CmdSearch.Enabled = False
    'cmdprint.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    Call ClearFrameCase
    Screen.MousePointer = vbHourglass
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Call SearchRecords
    'medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing
    Screen.MousePointer = vbDefault
    CmdSearch.Enabled = True
    CmdEdit.Visible = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    Exit Sub

errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstTUListing.ListItems.Clear
    lblcurrent = ""
    lbltotalAmt = ""
    lblSelectedAmt = ""
    'cmdprint.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    CmdEdit.Visible = True
  '  medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
End Sub

Private Sub CmdStop_Click()
    intExit = 1
End Sub

Private Sub cmdUpdate_Click()
Dim AA As Integer
Dim ItmChk As Boolean
Dim BB As Integer
Dim tmpSql As String
Dim tmpTUNo As String
Dim startTrans As Boolean
Dim bSucess  As Boolean
Dim DiffPay As Boolean
    ItmChk = False
    DiffPay = False
    tmpPayCode = ""
        
    AA = lstTUListing.ListItems.Count
    tmpPayCode = lstTUListing.SelectedItem.SubItems(10)
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        tmpCOVPayCode = IIf(Trim(tmpPayCode) = "TUG", "XG", "XL")
        Call SearchFloatAmt("And FTPayCode = '" & Trim(tmpCOVPayCode) & "'")
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    
    If AA = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf txtTURecDate = "__/__/____" Then
        MsgBox "Please Enter FT Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtTURecDate.SetFocus
    ElseIf Trim(txtchqNo) = "" Then
       MsgBox "Please Enter Cheque No!", vbCritical
       txtchqNo.SetFocus
    ElseIf txtchqDate = "__/__/____" Then
        MsgBox "Please Enter Cheque Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtchqDate.SetFocus
    ElseIf DiffPay = True Then
        MsgBox "Different Payors for this transaction is not allowed! ", vbOKOnly + vbCritical, DialogBoxTitle
    'ElseIf CDbl(lblSelectedAmt) <> CDbl(txtTURecAmt) Then
    '    MsgBox "Please check Float Transfer Amount! ", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Screen.MousePointer = vbHourglass
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
       ' medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
                        
        startTrans = True
        medixmasterconnWS.beginTrans
        medixmasterconnPayment.beginTrans
            
            tmpTURecNo = Trim(lstTUListing.SelectedItem.SubItems(1))
            Call UpdateTUFloat
            'Call UpdateFTFloat(tmpTURecNo)
            Call UpdateFloatAmt
            medixmasterconnWS.commitTrans
            medixmasterconnPayment.commitTrans
            startTrans = False
            txtTURecNo = tmpTURecNo
            'TxtReceiptNo = tmpRecNo
            'txtPVNo = tmpFTPVNo
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            'If tmpFloatBal < 300000 Then
            '    MsgBox "Float Balance is " & tmpFloatBal & " for " & tmpCOVPayCode, vbOKOnly + vbInformation, DialogBoxTitle
            'End If
            
            tmpSql = ""
            tmpSql = " and  CLMTURecNo = '" & Trim(tmpTURecNo) & "'"
            lstTUListing.ListItems.Clear
            Call CLMTUListing(tmpSql)
        EditFlag = False
        
        
        
 '       medixmasterconnWS.Close
   '     Set medixmasterconnWS = Nothing
    ''    medixmasterconnPayment.Close
    '    Set medixmasterconnPayment = Nothing
        Screen.MousePointer = vbDefault
        'Call ClearContraDetail
        'Call ClearRecDetail
        lblSelectedAmt = 0
        lbltotalAmt = 0
    End If
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If bSucess = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnPayment.RollbackTrans
    End If
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnPayment.Close
  '  Set medixmasterconnPayment = Nothing
    MsgBox Err.Description
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
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        EditFlag = False
        Call ComboListing
        txtTURecAmt = 0
        lblSelectedAmt = 0
        lbltotalAmt = 0
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset

    PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

    With CboTopupStatus
        .AddItem "0 - Not Received"
        .AddItem "1 - Received"
        .Text = "0 - Not Received"
    End With
End Sub

Private Sub ClearFrameCase()
    txtTURecNo = ""
    txtTURecDate = "__/__/____"
    txtchqNo = ""
    txtchqDate = "__/__/____"
    txtTURecAmt = 0
    TxtRemark.Text = ""
End Sub

Private Function SearchRecords()
Dim sql As String
Dim strAppend As String
Dim tmpPayCode As String
    strAppend = ""
    
    
    If Len(Trim(txtTUNoFr)) Then
        strAppend = strAppend + " AND  CLMTUNo >= '" & Trim(txtTUNoFr) & "'"
    End If
    
    If Len(Trim(txtTUNoTo)) Then
        strAppend = strAppend + " AND  CLMTUNo <= '" & Trim(txtTUNoTo) & "'"
    End If
    
    If Len(Trim(Text2)) Then
        strAppend = strAppend + " AND  CLMTURecNo >= '" & Trim(Text2) & "'"
    End If
    
    If Len(Trim(Text1)) Then
        strAppend = strAppend + " AND  CLMTURecNo <= '" & Trim(Text1) & "'"
    End If
    
    
    If Len(Trim(CboPayor)) Then
        If Mid(CboPayor, 1, 2) = "XG" Then
            tmpPayCode = "TUG"
        Else
            tmpPayCode = "TUL"
        End If
        strAppend = strAppend + " AND  Substring(CLMTUNo,1,3) = '" & Trim(tmpPayCode) & "'"
    End If
            
    'If Len(Trim(CboPayor)) Then
    '    StrAppend = StrAppend + " AND  PAYCode = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
    'End If
    'If Len(Trim(txtFTDateFr)) > 0 And IsDate(txtFTDateFr) = True Then
    '    StrAppend = StrAppend + " And CLMFTDate >= '" & Format(txtFTDateFr, "mm/dd/yyyy") & "'"
    'End If
    
    'If Len(Trim(txtFTDateTo)) > 0 And IsDate(txtFTDateTo) = True Then
    '    StrAppend = StrAppend + " And CLMFTDate <= '" & Format(txtFTDateTo, "mm/dd/yyyy") & "'"
    'End If
    
    If Len(Trim(txtTUDateFR)) > 0 And IsDate(txtTUDateFR) = True Then
        strAppend = strAppend + " And CLMTUDate >= '" & Format(txtTUDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtTUDateTo)) > 0 And IsDate(txtTUDateTo) = True Then
        strAppend = strAppend + " And CLMTUDate <= '" & Format(txtTUDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtTURecDateFr)) > 0 And IsDate(txtTURecDateFr) = True Then
        strAppend = strAppend + " And CLMTURecDate >= '" & Format(txtTURecDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtTURecDateTo)) > 0 And IsDate(txtTURecDateTo) = True Then
        strAppend = strAppend + " And CLMTURecDate <= '" & Format(txtTURecDateTo, "mm/dd/yyyy") & "'"
    End If
    
    'If Len(Trim(txtPVDateFr)) > 0 And IsDate(txtPVDateFr) = True Then
    '    StrAppend = StrAppend + " And CLMPVDate >= '" & Format(txtPVDateFr, "mm/dd/yyyy") & "'"
    'End If
    
    'If Len(Trim(txtPVDateTo)) > 0 And IsDate(txtPVDateTo) = True Then
    '    StrAppend = StrAppend + " And CLMPVDate <= '" & Format(txtPVDateTo, "mm/dd/yyyy") & "'"
    'End If

    If Len(Trim(CboTopupStatus)) Then
        If Mid(CboTopupStatus, 1, 1) = "0" Then
            strAppend = strAppend + " AND  (CLMTURecNo Is null OR CLMTURecNo = '') "
        ElseIf Mid(CboTopupStatus, 1, 1) = "1" Then
            strAppend = strAppend + " AND  (CLMTURecNo Is NOT null) "
        End If
    End If
    
    Call CLMTUListing(strAppend)
    Call SearchFloatAmt("")
End Function

Private Sub CLMTUListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim Current As String
Dim Check As Boolean
Dim strsql As String
Dim RecFound As Boolean
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Record As ListItem
Dim BordxAPPStatus As String
Dim strAppendCase As Integer
    Current = 0
    'If ChkScan.value = 0 Then
        lstTUListing.ListItems.Clear
        lblcurrent = 0
        lblSelectedAmt = 0
        lbltotalAmt = 0
    'End If
        strAppendCase = 2
    Set cmd1.ActiveConnection = medixmasterconnWS
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_ClaimFloatTransfer"
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("@strAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd1.Parameters.Append Prm2
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    rst2.CursorLocation = adUseClient
    Set rst2 = cmd1.Execute
    RecFound = False
    

    
    Do
        Do Until rst2.EOF

        RecFound = True
        With rst2
            If Len(Trim(!CLMTURecNo)) Then
                Set Record = lstTUListing.ListItems.Add(, , "Received")
            Else
                Set Record = lstTUListing.ListItems.Add(, , "")
            End If
            Record.SubItems(1) = IIf(Len(!CLMTURecNo), !CLMTURecNo, "")
            Record.SubItems(2) = IIf(Len(!CLMTURecDate), Format(!CLMTURecDate, "dd/mm/yyyy"), "")
            Record.SubItems(3) = Format(IIf(IsNumeric(!CLMTURecAmt), !CLMTURecAmt, 0), "#,##0.00")
            Record.SubItems(4) = IIf(Len(!CLMTUNo), !CLMTUNo, "")
            Record.SubItems(5) = IIf(Trim(!CLMTUDate) <> "", Format(!CLMTUDate, "dd/mm/yyyy"), "")
            Record.SubItems(6) = Format(IIf(Len(!CLMTUAmt), !CLMTUAmt, 0), "#,##0.00")
            Record.SubItems(7) = IIf(Len(!CLMCHQNo), !CLMCHQNo, "")
            Record.SubItems(8) = IIf(Trim(!CLMCHQDate) <> "", Format(!CLMCHQDate, "dd/mm/yyyy"), "")
            Record.SubItems(9) = IIf(Trim(Len(!CLMTURemarks)), !CLMTURemarks, "")
            Record.SubItems(10) = IIf(Trim(!CLMTUNo) <> "", Mid(!CLMTUNo, 1, 3), "")
            lbltotalAmt = CDbl(lbltotalAmt) + Format(CDbl(!CLMTUAmt), "#,##0.00")
            Current = Current + 1
            lblcurrent = Current
        End With
        rst2.MoveNext
            
        DoEvents
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
         
        Loop
       
    Loop Until Check = False
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    
    If RecFound = False Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        lblcurrent = 0
    End If
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub lstTUListing_Click()
    If lstTUListing.SelectedItem.Text = "Received" Then
        tmpPrevRecAmt = lstTUListing.SelectedItem.SubItems(6)
        Call LvDisplayData
    End If
End Sub

Private Sub lstTUListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstTUListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstTUListing, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstTUListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstTUListing, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstTUListing, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstTUListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstTUListing, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstTUListing, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstTUListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstTUListing, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstTUListing, ColumnHeader.Index, ldtString, m_blnDirection(11)
    End Select
End Sub

Private Sub lstTUListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Text <> "Received" Then
        If Item.Checked = True Then
            Item.Text = "Receive"
            txtTURecAmt = CDbl(txtTURecAmt) + CDbl(Item.SubItems(6))
            lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(Item.SubItems(6))
        ElseIf Item.Checked = False Then
            Item.Text = ""
            If txtTURecAmt > 0 Then
                txtTURecAmt = CDbl(txtTURecAmt) - CDbl(Item.SubItems(6))
                lblSelectedAmt = CDbl(lblSelectedAmt) - CDbl(Item.SubItems(6))
            End If
        End If
    Else
        Item.Checked = False
        MsgBox "Float has been alloacated for this item", vbOKOnly + vbCritical, DialogBoxTitle
    End If

End Sub

Private Sub LvDisplayData(Optional Item As ListItem, Optional ItemCheck As Boolean)
Dim lstRec As ListItem

    Call ClearFrameCase
    If lstTUListing.ListItems.Count Then
        If ItemCheck = True Then
            Set lstRec = Item
            lstTUListing.SelectedItem.Selected = False
            lstRec.Selected = True
        Else
            Set lstRec = lstTUListing.SelectedItem
        End If
            txtTURecNo = IIf(Len(Trim(lstRec.SubItems(1))), lstRec.SubItems(1), "")
            If Len(lstRec.SubItems(2)) Then
                txtTURecDate = Format(lstRec.SubItems(2), "DD/MM/YYYY")
            End If
            If Len(lstRec.SubItems(3)) Then
                txtTURecAmt = Format(lstRec.SubItems(3), "#,##0.00")
            End If
            If Len(lstRec.SubItems(7)) Then
                txtchqNo = lstRec.SubItems(7)
            End If
            If Len(lstRec.SubItems(8)) Then
                txtchqDate = Format(lstRec.SubItems(8), "DD/MM/YYYY")
            End If

            
            TxtRemark = IIf(Len(lstRec.SubItems(9)), lstRec.SubItems(9), "")
            CmdSave.Enabled = False
    End If
End Sub
Private Sub SearchFloatAmt(strsqlSearch As String)
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
    
    strsqlFT = " Select FTPayCode, FTFloatAmt from CLMFloatAmt where 0=0"
    strsqlFT = strsqlFT + strsqlSearch
    rst.Open strsqlFT, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        
        If rst.recordCount = 1 Then
            tmpCOVPayCode = rst!FTPayCode
            tmpChkFloatAmt = rst!FTFloatAmt
        Else
            Do While Not rst.EOF
                If rst!FTPayCode = "XG" Then
                    lblMGAmt = Format(rst!FTFloatAmt, "#,##0.00")
                Else
                    lblMLAmt = Format(rst!FTFloatAmt, "#,##0.00")
                End If
            rst.MoveNext
            Loop
        End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub UpdateFloatAmt()
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
Dim tmpFTFloatAmt As Double
Dim tmpFTFloatBLAmt As Double
    strsqlFT = ""
    'rst.Close
    'Set rst = Nothing
    'tmpCOVPayCode = "XL"
    strsqlFT = " Select FTPayCode, FTFloatAmt from CLMFloatAmt Where FTPayCode = '" & Trim(tmpCOVPayCode) & "'"
    rst.Open strsqlFT, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        tmpFTFloatAmt = rst!FTFloatAmt
        If EditFlag = False Then
            tmpFTFloatBLAmt = CDbl(tmpFTFloatAmt) + CDbl(txtTURecAmt)
        Else
            tmpFTFloatBLAmt = (CDbl(tmpFTFloatAmt) - CDbl(tmpPrevRecAmt)) + CDbl(txtTURecAmt)
        End If
        
        rst!FTFloatAmt = tmpFTFloatBLAmt
        rst.Update
        
        If tmpPayCode = "TUG" Then
            lblMGAmt = tmpFTFloatBLAmt
        Else
            lblMLAmt = tmpFTFloatBLAmt
        End If
    rst.Close
    Set rst = Nothing

End Sub

Private Sub SaveTUFloat()
Dim strsqlTU As String
Dim tmpTURecDate As String
Dim tmpChqDate As String
Dim tmpDate As String

    tmpTURecDate = Format(txtTURecDate, "yyyy/mm/dd")
    tmpChqDate = Format(txtchqDate, "yyyy/mm/dd")
    tmpDate = Format(Date, "yyyy/mm/dd")
    
    
    strsqlTU = ""
    strsqlTU = "Insert Into CLMFloatTopupAmt(CLMTURecNo, CLMTURecDate,CLMTURecAmt,CLMCHQNo, " & _
                "CLMCHQDate, CLMTURegUser, CLMTURegDate, CLMTULastUpdateUser,CLMTULastUpdateDate)  " & _
                "Values('" & tmpTURecNo & "','" & tmpTURecDate & "', " & _
                "" & txtTURecAmt & ",'" & txtchqNo & "','" & tmpChqDate & "', " & _
                "'" & Logon & "','" & tmpDate & "', '" & Logon & "','" & tmpDate & "')"
    medixmasterconnWS.Execute strsqlTU

End Sub

Private Sub UpdateTUFloat()
Dim strsqlTU As String
Dim tmpTURecDate As String
Dim tmpChqDate As String
Dim tmpDate As String
Dim rst As New ADODB.Recordset
    tmpTURecDate = Format(txtTURecDate, "yyyy/mm/dd")
    tmpChqDate = Format(txtchqDate, "yyyy/mm/dd")
    tmpDate = Format(Date, "yyyy/mm/dd")
    
    strsqlTU = ""
    strsqlTU = " Select  CLMTURecNo,CLMTURecNo, CLMTURecDate,CLMTURecAmt,CLMCHQNo, " & _
                "CLMCHQDate, CLMTURegUser, CLMTURegDate, CLMTULastUpdateUser,CLMTULastUpdateDate " & _
                "from CLMFloatTopupAmt " & _
                " Where CLMTURecNo = '" & Trim(tmpTURecNo) & "'"
    
    rst.Open strsqlTU, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        'If rst.recordCount = 0 Then
        '    rst.AddNew
        '    savedstatus = False
        'ElseIf rst.recordCount > 0 Then
        '    savedstatus = True
        'End If
        If rst.recordCount > 0 Then
            Do While Not rst.EOF
                With rst
                    !CLMTURecDate = tmpTURecDate
                    !CLMTURecAmt = txtTURecAmt
                    !CLMCHQNo = txtchqNo
                    !CLMCHQDate = txtchqDate
                    !CLMTULastUpdateUser = Logon
                    !CLMTULastUpdateDate = tmpDate
                End With
                rst.Update
            rst.MoveNext
            Loop
        End If
        rst.Close
        Set rst = Nothing
End Sub


Private Sub UpdateFTFloat(tmpTUNo As String)
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
Dim savedstatus As Boolean
    strsqlFT = ""
    'rst.Close
    'Set rst = Nothing
    strsqlFT = " Select  CLMTURecNo,CLMTUNo from CLMFloatTransfer " & _
                " Where CLMTUNo = '" & Trim(tmpTUNo) & "'"
    
    rst.Open strsqlFT, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        If rst.recordCount = 0 Then
            rst.AddNew
            savedstatus = False
        ElseIf rst.recordCount > 0 Then
            savedstatus = True
        End If
            Do While Not rst.EOF
                With rst
                    !CLMTURecNo = Trim(tmpTURecNo)
                End With
                rst.Update
            rst.MoveNext
            Loop
        rst.Close
        Set rst = Nothing
End Sub

Private Sub txtChqDate_LostFocus()
    If IsDate(txtchqDate) Then
    Else
        If txtchqDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtchqDate.SetFocus
        End If
    End If
End Sub

Private Sub txtTURecDateFr_LostFocus()
    If IsDate(txtTURecDateFr) Then
    Else
        If txtTURecDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtTURecDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtTUDateFR_LostFocus()
    If IsDate(txtTUDateFR) Then
    Else
        If txtTUDateFR <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtTUDateFR.SetFocus
        End If
    End If
End Sub


Private Sub txtTUDateTo_LostFocus()
    If IsDate(txtTUDateTo) Then
    Else
        If txtTUDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtTUDateTo.SetFocus
        End If
    End If

End Sub


Private Sub txtTURecDateTo_LostFocus()
    If IsDate(txtTURecDateTo) Then
    Else
        If txtTURecDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtTURecDateTo.SetFocus
        End If
    End If
End Sub

Private Sub CheckAll_Click()
Dim LstCnt As Integer

Dim counttransfered As Integer
counttransfered = 0

    LstCnt = lstTUListing.ListItems.Count
    
    txtTURecAmt = CDbl(0)
    lblSelectedAmt = CDbl(0)
        
        If LstCnt = 0 Then
            CheckAll.Value = 0
        Else
            If CheckAll.Value = 1 Then
                For LstCnt = 1 To LstCnt
                    If lstTUListing.ListItems(LstCnt).Checked = False And lstTUListing.ListItems(LstCnt).Text = "Received" Then
                        'MsgBox "Float has been alloacated for this item! ", vbCritical
                         counttransfered = counttransfered + 1
                    ElseIf lstTUListing.ListItems(LstCnt).Checked = False Then
                        lstTUListing.ListItems(LstCnt).Checked = True
                        lstTUListing.ListItems(LstCnt).Text = "Receive"
                        
                        txtTURecAmt = CDbl(txtTURecAmt) + CDbl(lstTUListing.ListItems(LstCnt).SubItems(6))
                        lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(lstTUListing.ListItems(LstCnt).SubItems(6))
            
                    ElseIf lstTUListing.ListItems(LstCnt).Checked = True Then
                        lstTUListing.ListItems(LstCnt).Checked = False
                        lstTUListing.ListItems(LstCnt).Text = ""
                    End If
                Next LstCnt
            Else
                For LstCnt = 1 To LstCnt
                    If lstTUListing.ListItems(LstCnt).Checked = True Then
                        lstTUListing.ListItems(LstCnt).Checked = False
                        lstTUListing.ListItems(LstCnt).Text = ""
                    End If
                Next LstCnt
                
                txtTURecAmt = CDbl(0)
                lblSelectedAmt = CDbl(0)
            End If
            
            If lstTUListing.ListItems.Count = counttransfered And counttransfered <> 0 Then
                MsgBox "All items have been transferred! ", vbCritical
            ElseIf lstTUListing.ListItems.Count <> counttransfered And counttransfered <> 0 Then
                MsgBox "Some ( " + counttransfered + " )item have been transferred! ", vbCritical
            End If
            
        End If
End Sub
