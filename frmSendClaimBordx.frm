VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSendClaimBordx 
   Caption         =   "Send Claim Bordx"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox ChkMNRBDate 
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
         Left            =   5040
         TabIndex        =   33
         Top             =   1260
         Width           =   255
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   5040
         TabIndex        =   32
         Top             =   540
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1320
         TabIndex        =   31
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   30
         Top             =   530
         Width           =   1455
      End
      Begin VB.TextBox TxtMonth2 
         Height          =   285
         Left            =   9000
         MaxLength       =   2
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtMonth1 
         Height          =   285
         Left            =   7440
         MaxLength       =   2
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtYear2 
         Height          =   285
         Left            =   9000
         MaxLength       =   4
         TabIndex        =   24
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtYear1 
         Height          =   285
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   29
         Top             =   530
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtBordxDate1 
         Height          =   285
         Left            =   1320
         TabIndex        =   34
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDate2 
         Height          =   285
         Left            =   3240
         TabIndex        =   35
         Top             =   780
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBordxAmt2 
         Height          =   285
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   18
         Top             =   1040
         Width           =   1695
      End
      Begin VB.TextBox TxtBordxAmt1 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   17
         Top             =   1020
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtMNRBDate1 
         Height          =   315
         Left            =   1320
         TabIndex        =   19
         Top             =   1245
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMNRBDate2 
         Height          =   315
         Left            =   3240
         TabIndex        =   20
         Top             =   1245
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDPOFr 
         Height          =   315
         Left            =   7440
         TabIndex        =   25
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDPOTo 
         Height          =   315
         Left            =   9000
         TabIndex        =   26
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboStatus 
         Height          =   315
         Left            =   7440
         TabIndex        =   27
         Top             =   1000
         Width           =   2775
      End
      Begin VB.ComboBox cboCashType 
         Height          =   315
         ItemData        =   "frmSendClaimBordx.frx":0000
         Left            =   7440
         List            =   "frmSendClaimBordx.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1280
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Cash Type"
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   53
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Date Pass Out"
         Height          =   255
         Index           =   4
         Left            =   6360
         TabIndex        =   52
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "To"
         Height          =   255
         Index           =   3
         Left            =   8760
         TabIndex        =   51
         Top             =   750
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "To"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   49
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Sent Date"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   48
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   47
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   44
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   42
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx Amt"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   1030
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   40
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "A/C Month"
         Height          =   255
         Left            =   6360
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   38
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "A/C Year"
         Height          =   255
         Left            =   6360
         TabIndex        =   37
         Top             =   530
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Status"
         Height          =   255
         Index           =   0
         Left            =   6360
         TabIndex        =   36
         Top             =   1030
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   11655
      Begin MSMask.MaskEdBox TxtDate 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   195
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lstBordxDelivery 
      Height          =   4335
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7646
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Bordx Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Sent to Status"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Sent Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Sent By"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "A/C Period"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "DPO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Cash Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   6960
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   375
         Left            =   3360
         TabIndex        =   55
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   54
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   220
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9240
         TabIndex        =   9
         Top             =   220
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7680
         TabIndex        =   8
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10320
         TabIndex        =   7
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10920
         TabIndex        =   6
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6960
         TabIndex        =   5
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdToPayor 
         Caption         =   "&Update"
         Height          =   375
         Left            =   8400
         TabIndex        =   4
         Top             =   220
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Send Claim Bordx"
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
      TabIndex        =   50
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmSendClaimBordx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim BordxNoArray() As String
Dim searchCriteria As String
Dim ItemChecked As Boolean
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim Check2 As Boolean
Dim Current2 As String
Dim TotalAmt2 As String
Private m_blnDirection(1 To 10) As Boolean

Private Sub CmdCheck_Click()
    SetCheckAllItems True
End Sub

Private Sub CboPayor_Click()
    If CboPayor.Text = "" Then
        If InStr(CboPayor.Text, "null") Then
            CboPayor.Enabled = False
        End If
        TxtBordx1.SetFocus
    End If
    TxtBordx1.SetFocus
End Sub

Private Sub CboPayor_Change()
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
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



Private Sub Form_KeyPress(KeyAscii As Integer)
    'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
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
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    Current2 = 0
    TotalAmt2 = 0
    Check = False
    Check2 = False
    Label14.Visible = False
    TxtBordx2.Visible = False
    ReDim BordxNoArray(0)
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    ItemChecked = False
End Sub


'**********************Start Command Button *******************************

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        lstBordxDelivery.ListItems.Clear
        Label14.Visible = False
        TxtBordx2.Visible = False
        TxtBordx1.SetFocus
    Else
        lstBordxDelivery.ListItems.Clear
        Label14.Visible = True
        TxtBordx2.Visible = True
        TxtBordx1.SetFocus
    End If
End Sub

Private Sub cmdPrint_Click()
    'Screen.MousePointer = vbHourglass
        'Call ExportToExcelBordxToMNRB(lstBordxDelivery)
    'Screen.MousePointer = vbDefault
    
    Screen.MousePointer = vbHourglass
        
    If lstBordxDelivery.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstBordxDelivery)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CmdClear_Click()
    CboPayor.Clear
    cboStatus.Clear
    Call ClearForm(Me)
    Current2 = 0
    TotalAmt2 = 0
    Check = False
    Check2 = False
    TxtBordx1.SetFocus
    lstBordxDelivery.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    ReDim BordxNoArray(0)
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub CmdSearch_Click()
On Error GoTo errRun

    ItemChecked = False
    LblTotalAmt = ""
    lblTotal = ""
    lblCurrent = ""
    ReDim BordxNoArray(0)
    CmdSearch.Enabled = False
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    If ChkScan.Value = 1 Then
    
        If CboPayor = "" Then
            MsgBox "Please select a Payor", vbOKOnly + vbCritical, DialogBoxTitle
            CboPayor.SetFocus
        ElseIf TxtBordx1 = "" Then
            MsgBox "Please Enter Bordx No", vbOKOnly + vbCritical, DialogBoxTitle
            TxtBordx1.SetFocus
            CmdClear.Enabled = True
            CmdToPayor.Enabled = True
            CmdPrint.Enabled = True
            CmdExit.Enabled = True
            CmdSearch.Enabled = True
        Else
        
            Screen.MousePointer = vbHourglass
            
            listrecord = lstBordxDelivery.ListItems.Count
            If listrecord <> 0 Then
                
                For listloop1 = 1 To listrecord
                    Call BordxNoCheck
                Next listloop1
            End If
                If Check = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                ElseIf Check = False Then
                    CmdClear.Enabled = False
                    CmdToPayor.Enabled = False
                    CmdPrint.Enabled = False
                    CmdExit.Enabled = False
                    CmdSearch.Enabled = False
                    Call SearchRecord
                End If
                Check = False
                CmdClear.Enabled = True
                CmdToPayor.Enabled = True
                CmdPrint.Enabled = True
                CmdExit.Enabled = True
                CmdSearch.Enabled = True
            Screen.MousePointer = Default
            TxtBordx1 = ""
            TxtBordx1.SetFocus
        End If
        TxtBordx1.SetFocus
    Else
        
        Screen.MousePointer = vbHourglass
            
            Call SearchRecord
                
        Screen.MousePointer = Default
        
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        CmdPrint.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    End If
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Exit Sub
    
errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstBordxDelivery.ListItems.Clear
    lblCurrent = ""
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
    
End Sub

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim CAS As New ADODB.Recordset

    strAppend = ""
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtBordx1)) Then
            strAppend = strAppend + " AND BordxRefNo = '" & Trim(TxtBordx1) & "'"
        End If
    Else
        If Len(Trim(TxtBordx1)) Then
            strAppend = strAppend + " AND BordxRefNo >= '" & Trim(TxtBordx1) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
            strAppend = strAppend + " AND BordxRefNo <= '" & Trim(TxtBordx2) & "'"
        End If
    End If
    
    If IsDate(txtBordxDate1) = True Then
        strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(txtBordxDate2) = True Then
        strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxAmt1)) Then
        strAppend = strAppend + " AND BordxClaimAmt >= '" & Trim(TxtBordxAmt1) & "'"
    End If
    
    If Len(Trim(TxtBordxAmt2)) Then
        strAppend = strAppend + " AND BordxClaimAmt <= '" & Trim(TxtBordxAmt2) & "'"
    End If
    
    If Len(Trim(TxtMonth1)) Then
        strAppend = strAppend + " AND PAYToACPeriodMth >= '" & Trim(TxtMonth1) & "'"
    End If
    
    If Len(Trim(TxtMonth2)) Then
        strAppend = strAppend + " AND PAYToACPeriodMth <= '" & Trim(TxtMonth2) & "'"
    End If
    
    If Len(Trim(TxtYear1)) Then
        strAppend = strAppend + " AND PAYToACPeriodYR >= '" & Trim(TxtYear1) & "'"
    End If
    
    If Len(Trim(TxtYear2)) Then
        strAppend = strAppend + " AND PAYToACPeriodYR <= '" & Trim(TxtYear2) & "'"
    End If
    
    If Mid(cboStatus, 1, 1) = "1" Then
        strAppend = strAppend + " AND BordxSendToACStatus = '1'"
    End If
    
    If Mid(cboStatus, 1, 1) = "0" Then
        strAppend = strAppend + " AND BordxSendToACStatus = '0'"
    End If
    
    If IsDate(txtMNRBDate1) = True Then
        strAppend = strAppend + " And BordxSendACDate >= '" & Format(txtMNRBDate1, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(txtMNRBDate2) = True Then
        strAppend = strAppend + " And BordxSendACDate <= '" & Format(txtMNRBDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkMNRBDate.Value = 1 Then
        strAppend = strAppend + " And (BordxSendACDate Is null OR BordxSendACDate = '')"
    End If
    
    If IsDate(txtDPOFr) = True Then
        strAppend = strAppend + " And CASTokenDate >= '" & Format(txtDPOFr, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(txtDPOTo) = True Then
        strAppend = strAppend + " And CASTokenDate <= '" & Format(txtDPOTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(cboCashType) Then
        strAppend = strAppend + " And INVPayTo = '" & Mid(cboCashType, 1, 1) & "'"
    End If
    
    
    
    
    searchCriteria = strAppend
    
    If ChkScan.Value = 0 Then
        Call BordxDeliveryListing(strAppend)
    Else
        Call BordxDeliveryListing2(strAppend)
    End If
   
End Function

Private Sub CmdStop_Click()
    intExit = 1
    TxtBordx1.SetFocus
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdToPayor_Click()
Dim ii As Integer
Dim Cnt As Integer
Dim RecordFound As Boolean
Dim recordPassed As Boolean
Dim Update
    
    CmdToPayor.Enabled = False
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    If ItemChecked = True Then
        
        If TxtDate = "__/__/____" Then
            MsgBox "Please Enter in date", vbCritical
            TxtDate.SetFocus
        Else
            Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            
            If Update = vbYes Then
                Cnt = 0
                RecordFound = False
                If lstBordxDelivery.ListItems.Count <> 0 Then
                    For ii = 0 To UBound(BordxNoArray)
                        
                        If Trim(BordxNoArray(ii)) <> "" Then
                            UpdateSent BordxNoArray(ii)
                            RecordFound = True
                        End If
                    
                    Next ii
                    If RecordFound = True Then
                        MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
                    
                    Erase BordxNoArray
                    Arraycnt = 0
                    ReDim BordxNoArray(0)
                    lstBordxDelivery.ListItems.Clear
                    If ChkScan.Value = 0 Then
                        Call BordxDeliveryListing(searchCriteria)
                    End If
                End If
            End If
        End If
    Else
        MsgBox "Please select at least one record to be updated! ", vbCritical
    End If
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
    CmdToPayor.Enabled = True
End Sub
'**********************End Command Button *******************************

'**********************Start BordxDelivery Listing  *******************************
Private Sub BordxDeliveryListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
    
    total = 0
    Current = 0
    
    lstBordxDelivery.ListItems.Clear
    
    'sql = " Select BordxRefNo, BordxPAYCode, BordxDate, " & _
          " BordxClaimAmt, BordxStatus, BordxMNRBSentDate, BordxMNRBSentBy, " & _
          " BordxMNRBToStatus, PAYToACPeriodYR, PAYToACPeriodMth " & _
          " From BORDX where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "Bordx_Sent"
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
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        CmdPrint.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
            
            With rst2
                Set Record = lstBordxDelivery.ListItems.Add(, , "")
                
                If Len(rst2!BordxPAYCode) Then
                    Record.SubItems(1) = rst2!BordxPAYCode
                End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(2) = rst2!BordxRefNo
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(3) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(4) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!Bordxstatus) Then
                    Record.SubItems(5) = rst2!Bordxstatus
                End If
                If Len(rst2!BordxSendToACStatus) Then
                   Record.SubItems(6) = rst2!BordxSendToACStatus
                End If
                'If Mid(rst2!BordxRefNo, 1, 1) = "S" Then
                    If Len(rst2!BordxSendACDate) Then
                        Record.SubItems(7) = Format(rst2!BordxSendACDate, "dd/mm/yyyy")
                    End If
                'Else
                '    If Len(rst2!BordxPAYSentDate) Then
                '        Record.SubItems(7) = Format(rst2!BordxPAYSentDate, "dd/mm/yyyy")
                '    End If
                'End If
                
                If Len(rst2!BordxSendACBy) Then
                   Record.SubItems(8) = rst2!BordxSendACBy
                End If
                
                If Len(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR) Then
                    Record.SubItems(9) = Trim(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR)
                End If
                
                If Len(rst2!CASTokenDate) Then
                    Record.SubItems(10) = Format(rst2!CASTokenDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!InvPayTo) Then
                    Record.SubItems(11) = rst2!InvPayTo
                End If
                
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!BordxClaimAmt) Then
                    TotalAmt = TotalAmt + rst2!BordxClaimAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
End Sub

Private Sub BordxDeliveryListing2(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sqlCount As String
Dim ClaimAdmin As Boolean
    
    'sql = " Select BordxRefNo, BordxPAYCode, BordxDate, BordxClaimAmt, BordxStatus, " & _
          " BordxMNRBSentDate, BordxMNRBSentBy, BordxMNRBToStatus, PAYToACPeriodYR, PAYToACPeriodMth " & _
          " From BORDX where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If
        
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "Bordx_Sent"
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

    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        CmdPrint.Enabled = True
        CmdExit.Enabled = True
        CmdSearch.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        
            With rst2
                Set Record = lstBordxDelivery.ListItems.Add(, , "")
                
                If Len(rst2!BordxPAYCode) Then
                    Record.SubItems(1) = rst2!BordxPAYCode
                End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(2) = rst2!BordxRefNo
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(3) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(4) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!Bordxstatus) Then
                    Record.SubItems(5) = rst2!Bordxstatus
                End If
                If Len(rst2!BordxSendToACStatus) Then
                   Record.SubItems(6) = rst2!BordxSendToACStatus
                End If
                'If Mid(rst2!BordxRefNo, 1, 1) = "S" Then
                If Len(rst2!BordxSendACDate) Then
                    Record.SubItems(7) = Format(rst2!BordxSendACDate, "dd/mm/yyyy")
                End If
                'Else
                '    If Len(rst2!BordxPAYSentDate) Then
                '        Record.SubItems(7) = Format(rst2!BordxPAYSentDate, "dd/mm/yyyy")
                '    End If
                'End If
                If Len(rst2!BordxSendACBy) Then
                   Record.SubItems(8) = rst2!BordxSendACBy
                End If
                If Len(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR) Then
                    Record.SubItems(9) = Trim(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR)
                End If
                
                If Len(rst2!CASTokenDate) Then
                    Record.SubItems(10) = Format(rst2!CASTokenDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!InvPayTo) Then
                    Record.SubItems(11) = rst2!InvPayTo
                End If
                
                Current2 = Current2 + 1
                lblTotal = Current2
                If Len(rst2!BordxClaimAmt) Then
                    TotalAmt2 = TotalAmt2 + rst2!BordxClaimAmt
                End If
                LblTotalAmt = Format(TotalAmt2, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    CmdPrint.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
End Sub
'**********************End BordxDelivery Listing *******************************

'**********************Start Check ListView Record*******************************
Private Sub lstBordxDelivery_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim temp
Static Arraycnt As Integer
    
        
    If Item.Checked = True Then
        If Item.SubItems(6) = False Then
            Arraycnt = IIf(UBound(BordxNoArray) = 0, 1, UBound(BordxNoArray) + 1)
            ReDim Preserve BordxNoArray(Arraycnt)
            BordxNoArray(Arraycnt) = Item.SubItems(2)
            Arraycnt = Arraycnt + 1
            ItemChecked = True
        Else
            MsgBox "Record has been Updated!", vbCritical
            Item.Checked = False
        End If
            
    Else
        ' Remove from the array
        newbound = 0
    
        If UBound(BordxNoArray) <> 0 Then
            For i = 0 To UBound(BordxNoArray)
                If (Item.SubItems(2)) <> BordxNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = BordxNoArray(i)
                    newbound = newbound + 1
                End If
            Next
            Erase BordxNoArray
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim BordxNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                BordxNoArray(i) = tempArray(i)
            Next
            Erase tempArray
        Else
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            Arraycnt = 0
        End If
    End If
End Sub
'**********************End Check ListView Record*******************************

'**********************Start Update all Record functions*******************************

Private Sub UpdateSent(tBordxNo As String)
Dim strsql As String
Dim SentDate As String
Dim strsqlREM As String
Dim rst As New ADODB.Recordset
Dim strsql2 As String
Dim rst2 As New ADODB.Recordset
Dim tmpRemarks1 As String
Dim tmpRemarks2 As String
Dim tmpRemarks3 As String
Dim tmpRemarks4 As String
Dim tmpRemarks5 As String
Dim strsqlWeb As String
Dim WEB As New ADODB.Recordset

    SentDate = Format(TxtDate, "YYYY/MM/DD")
    
    If Mid(tBordxNo, 1, 1) = "S" Then
        strsql = " update Bordx set " & _
                 " BordxMNRBSentDate = '" & SentDate & "', " & _
                 " BordxMNRBToStatus = '1', " & _
                 " BordxMNRBSentBy = '" & Logon & "' " & _
                 " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
        
        medixmasterconnWS.Execute strsql
        
        strsql = " update Bordx set " & _
             " BordxPAYSentDate = '" & SentDate & "', " & _
             " BordxPayToStatus = '1', " & _
             " BordxPAYSentBy = '" & Logon & "' " & _
             " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
    
        medixmasterconnWS.Execute strsql
        
    Else
        strsql = " update Bordx set " & _
             " BordxPAYSentDate = '" & SentDate & "', " & _
             " BordxPayToStatus = '1', " & _
             " BordxPAYSentBy = '" & Logon & "' " & _
             " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
    
        medixmasterconnWS.Execute strsql
    End If
    
    strsql = " update Bordx set " & _
             " BordxSendACDate = '" & SentDate & "', " & _
             " BordxSendToACStatus = '1', " & _
             " BordxSendACBy = '" & Logon & "' " & _
             " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
    medixmasterconnWS.Execute strsql
    
    strsqlREM = ""
    strsqlREM = "SELECT dbo.CASRemarks.CASNumber, " & _
                " HISWorksheet.dbo.Claim.BordxRefNo " & _
                " FROM dbo.CASRemarks  INNER JOIN " & _
                " HISWorksheet.dbo.Claim ON " & _
                " dbo.CasRemarks.CasNumber = HISWorksheet.dbo.Claim.CasNumber " & _
                " where HISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(tBordxNo) & "'" & _
                " GROUP BY dbo.CASRemarks.CASNumber, " & _
                " HISWorksheet.dbo.Claim.BordxRefNo, " & _
                " HISWorksheet.dbo.Claim.CasNumber "
    rst.Open strsqlREM, medixmasterconn, adOpenKeyset, adLockOptimistic

    Do While Not rst.EOF
        strsqlWeb = ""
        strsqlWeb = " update CASRemarks set " & _
                " CASWebStatus = 'Y' " & _
                " Where CASNumber = '" & Trim(rst!CasNumber) & "'"
        medixmasterconn.Execute strsqlWeb
    rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    strsqlREM = ""
    strsqlREM = " SELECT HISWorksheet.dbo.Claim.BordxRefNo, " & _
                " dbo.CASRemarksTwo.CASNumber " & _
                " FROM HISWorksheet.dbo.Claim INNER JOIN " & _
                " dbo.CASRemarksTwo ON HISWorksheet.dbo.Claim.CasNumber = dbo.CASRemarksTwo.CasNumber " & _
                " where HISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(tBordxNo) & "'" & _
                " GROUP BY HISWorksheet.dbo.Claim.BordxRefNo, HISWorksheet.dbo.Claim.CASNumber, " & _
                " dbo.CASRemarksTwo.CASNumber"
    rst.Open strsqlREM, medixmasterconn, adOpenKeyset, adLockOptimistic

    Do While Not rst.EOF
        strsqlWeb = ""
        strsqlWeb = " update CASRemarksTwo set " & _
                " CASWebStatus = 'Y' " & _
                " Where CASNumber = '" & Trim(rst!CasNumber) & "'"
        medixmasterconn.Execute strsqlWeb
    rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    strsqlREM = ""
    strsqlREM = " SELECT HISWorksheet.dbo.Claim.BordxRefNo, " & _
                " dbo.CASRemarkThree.CASNumber " & _
                " FROM HISWorksheet.dbo.Claim INNER JOIN dbo.CASRemarkThree ON " & _
                " HISWorksheet.dbo.Claim.CasNumber = dbo.CASRemarkThree.CasNumber " & _
                " where HISWorksheet.dbo.Claim.BordxRefNo = '" & Trim(tBordxNo) & "'" & _
                " GROUP BY HISWorksheet.dbo.Claim.BordxRefNo, dbo.CASRemarkThree.CASNumber "
    rst.Open strsqlREM, medixmasterconn, adOpenKeyset, adLockOptimistic

    Do While Not rst.EOF
        strsqlWeb = ""
        strsqlWeb = " update CASRemarkThree set " & _
                " CASWebStatus = 'Y' " & _
                " Where CASNumber = '" & Trim(rst!CasNumber) & "'"
        medixmasterconn.Execute strsqlWeb
    rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
End Sub
'**********************End Update all Record functions*******************************

'**********************Start Sorted ListView *******************************
Private Sub lstBordxDelivery_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
    Case 5
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    End Select
End Sub
'**********************End Sorted ListView *******************************

'**********************Start Combo Listing*******************************
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim MonthStart
Dim YearStart
Dim cmd As New ADODB.Command

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    
    Do Until PAY.EOF
        With PAY
            CboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing

    With cboStatus
        .AddItem "0 - Not Sent"
        .AddItem "1 - Sent"
        .AddItem "2 - Both"
        .Text = "0 - Not Sent"
    End With

    With cboCashType
        .AddItem "H - Hopital"
        .AddItem "M - Member"
    End With

End Sub
'**********************End Combo Listing*******************************

Private Sub TxtBordxAmt1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtBordxAmt2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtBordxDate1_LostFocus()
    If IsDate(txtBordxDate1) Then
    Else
        If txtBordxDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDate2_LostFocus()
    If IsDate(txtBordxDate2) Then
    Else
        If txtBordxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate2.SetFocus
        End If
    End If
End Sub


Private Sub txtDate_LostFocus()
    If IsDate(TxtDate) Then
    Else
        If TxtDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDate.SetFocus
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
    lvCount = lstBordxDelivery.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstBordxDelivery.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Public Sub SetCheck(hwnd As Long, lItemIndex As Long, bState As Boolean)
Dim LV As LVITEM
Dim r As Long
    With LV
    .mask = LVIF_STATE
    .state = IIf(bState, &H2000, &H1000)
    .stateMask = LVIS_STATEIMAGEMASK
    End With
    r = SendMessageAny(hwnd, LVM_SETITEMSTATE, lItemIndex, LV)
End Sub

Private Function BordxNoCheck()
    If Trim(ChkStr(TxtBordx1)) = ChkStr(Trim(lstBordxDelivery.ListItems(listloop1).SubItems(2))) Then
        Check = True
    End If
End Function
