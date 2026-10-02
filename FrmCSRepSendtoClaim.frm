VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCSRepSendtoClaim 
   Caption         =   "Case Send to Claims (Repudiate)"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      Caption         =   "Selection Criteria"
      Height          =   1755
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox CboCaseStatus 
         Height          =   315
         Left            =   7560
         TabIndex        =   21
         Top             =   225
         Width           =   1575
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   ">"
         Height          =   270
         Left            =   9240
         TabIndex        =   20
         Top             =   800
         Width           =   255
      End
      Begin VB.TextBox TxtCaseNo 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   19
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   285
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   18
         Top             =   240
         Width           =   1560
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   5200
         TabIndex        =   17
         Top             =   240
         Width           =   1050
      End
      Begin VB.CheckBox chkDOR 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   5200
         TabIndex        =   16
         Top             =   525
         Width           =   495
      End
      Begin VB.CheckBox ChkDOA 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   5200
         TabIndex        =   15
         Top             =   795
         Width           =   495
      End
      Begin VB.CheckBox ChkDOD 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   5200
         TabIndex        =   14
         Top             =   1050
         Width           =   495
      End
      Begin VB.CheckBox ChkDPO 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   5200
         TabIndex        =   13
         Top             =   1320
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtDORDate1 
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Top             =   500
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate1 
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Top             =   750
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate1 
         Height          =   285
         Left            =   1200
         TabIndex        =   27
         Top             =   1010
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDPODate1 
         Height          =   285
         Left            =   1200
         TabIndex        =   28
         Top             =   1260
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDORDate2 
         Height          =   285
         Left            =   3480
         TabIndex        =   29
         Top             =   500
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate2 
         Height          =   285
         Left            =   3480
         TabIndex        =   30
         Top             =   750
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate2 
         Height          =   285
         Left            =   3480
         TabIndex        =   31
         Top             =   1010
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDPODate2 
         Height          =   285
         Left            =   3480
         TabIndex        =   32
         Top             =   1260
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   7800
         TabIndex        =   22
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox CboDept 
         Height          =   315
         Left            =   7560
         TabIndex        =   23
         Top             =   515
         Width           =   1575
      End
      Begin VB.ComboBox CboPass 
         Height          =   315
         Left            =   7560
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   800
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   255
         Left            =   3000
         TabIndex        =   46
         Top             =   280
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "Case No"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   350
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Passed by"
         Height          =   255
         Left            =   6600
         TabIndex        =   44
         Top             =   820
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Case Status"
         Height          =   255
         Left            =   6600
         TabIndex        =   43
         Top             =   250
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Dept. Stage"
         Height          =   255
         Left            =   6600
         TabIndex        =   42
         Top             =   535
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   6840
         TabIndex        =   41
         Top             =   1875
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "To"
         Height          =   255
         Left            =   3000
         TabIndex        =   40
         Top             =   1350
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   3000
         TabIndex        =   39
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   3000
         TabIndex        =   38
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "DOD from"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         ToolTipText     =   "Date of Discharged"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "DOA from"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         ToolTipText     =   "Date of Admission"
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "DOR from"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         ToolTipText     =   "Date of Registered"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "DPO from"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         ToolTipText     =   "Date Pass On"
         Top             =   1350
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   11655
      Begin VB.TextBox lblCurrent 
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10800
         TabIndex        =   2
         Top             =   180
         Width           =   615
      End
      Begin VB.TextBox txtBatchNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10200
         TabIndex        =   3
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9240
         TabIndex        =   4
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdPass 
         Caption         =   "&Pass to Claim"
         Height          =   375
         Left            =   8160
         TabIndex        =   5
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   6
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   7
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5640
         TabIndex        =   9
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
         Height          =   375
         Left            =   4800
         TabIndex        =   10
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   255
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstCaseTrackList 
      Height          =   5055
      Left            =   120
      TabIndex        =   47
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8916
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
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pass-On"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Case No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mem'ship No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "DOR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "DOA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "DOD"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "DPO"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "No of Days (DPO-DOR)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Batch No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Passed By"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Reg. By"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Hospital Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Stay Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "Cash Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Inv. Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   16
         Text            =   "Claimability"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Dept Stage"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label39 
      Caption         =   "* - Search Empty Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Case Send to Claims (Repudiate)"
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
      TabIndex        =   48
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmCSRepSendtoClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intExit As Integer
Dim listloop1 As Integer
Dim Check As Boolean
Dim Current2 As String
Dim Month As String
Dim ScanCode As String
Dim PayorCode As String
Dim DPO As String
Dim CASNo As String
Dim strAppend As String
Dim m_blnDirection(1 To 18) As Boolean

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
    Current2 = 0
    Check = False
    ChkScan.Value = 1
    Call ComboListing
End Sub

'**********************Start Command Button *******************************

Private Sub CmdBack_Click()
    Frame1.Visible = True
    TxtCaseNo.SetFocus
    lstCaseTrackList.Top = 2160
    lstCaseTrackList.Width = 11655
    lstCaseTrackList.Left = 120
    lstCaseTrackList.Height = 4935
End Sub

Private Sub CmdView_Click()
    Frame1.Visible = False
    lstCaseTrackList.Top = 480
    lstCaseTrackList.Width = 11655
    lstCaseTrackList.Left = 120
    lstCaseTrackList.Height = 6615
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
    CmdPass.Enabled = True
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstCaseTrackList.ListItems.Clear
    lblCurrent = 0
    Current2 = 0
    Check = False
End Sub

Private Sub CmdRefresh_Click()
Dim User3 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

CboPass.Clear

'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

'User3.Open "Select Distinct CASUserToken from CAS", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    strAppend = ""
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Send2Claim"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append Prm2
    User3.CursorLocation = adUseClient
    User3.CursorType = adOpenForwardOnly
    User3.CacheSize = 100
    Set User3 = cmd.Execute

    Do Until User3.EOF
        If Len(User3!CASUserToken) Then
            CboPass.AddItem User3!CASUserToken
        End If
        User3.MoveNext
    Loop
    User3.Close
    Set User3 = Nothing
    
'medixmasterconn.Close
'Set medixmasterconn = Nothing

End Sub

Private Sub CmdCheck_Click()
    SetCheckAllItems True
End Sub

Public Sub cmdSearch_Click()
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    ScanCode = TxtCaseNo
    CmdSearch.Enabled = False
    If Mid(ScanCode, 1, 2) = "LO" Then
        TxtCaseNo = ""
        Call SearchForm(ScanCode)
    Else
            
        CmdClear.Enabled = False
        CmdExit.Enabled = False
        CmdPrintFile.Enabled = False
        CmdPass.Enabled = False
            
        If ChkScan.Value = 1 Then
        
            If TxtCaseNo = "" Then
                MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtCaseNo.SetFocus
                CmdClear.Enabled = True
                CmdExit.Enabled = True
                'CmdPrint.Enabled = True
                CmdPrintFile.Enabled = True
                CmdPass.Enabled = True
            Else
                Screen.MousePointer = vbHourglass
            
                listrecord = lstCaseTrackList.ListItems.Count
                If listrecord <> 0 Then
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                
                If Check = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                ElseIf Check = False Then
                    CmdClear.Enabled = True
                    CmdExit.Enabled = True
                    CmdPrintFile.Enabled = True
                    CmdPass.Enabled = True
                    Call SearchRecord
                End If
                Check = False
                Screen.MousePointer = Default
                TxtCaseNo = ""
                TxtCaseNo.SetFocus
            End If
            CmdClear.Enabled = True
            CmdExit.Enabled = True
            CmdPrintFile.Enabled = True
            CmdPass.Enabled = True
            TxtCaseNo.SetFocus
        Else
            
            Screen.MousePointer = vbHourglass
                
                Call SearchRecord
            
            Screen.MousePointer = Default
            
            CmdClear.Enabled = True
            CmdExit.Enabled = True
            CmdPrintFile.Enabled = True
            CmdPass.Enabled = True
        End If
    End If
    CmdSearch.Enabled = True
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPass_Click()
Dim i As Integer
Dim ItemChecked As Boolean
Dim CaseNo As String
Dim Update
Dim startTrans
  
    startTrans = False
    
    ItemChecked = False
    
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
    If lstCaseTrackList.ListItems.Count = 0 Then
        MsgBox "Please select a payor before proceeding!", vbOKOnly + vbCritical
    Else
        
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                
                startTrans = True
                medixmasterconnMaint.beginTrans
                medixmasterconn.beginTrans
                DPO = Date
                txtBatchNo = GenerateCasePass(DPO)
                
                For i = 1 To lstCaseTrackList.ListItems.Count
                    If lstCaseTrackList.ListItems(i).Checked = True Then
                        ItemChecked = True
                        Exit For
                    End If
                Next i
                    
                If ItemChecked = True Then
                
                    For i = 1 To lstCaseTrackList.ListItems.Count
                        If lstCaseTrackList.ListItems(i).Checked = True Then
                            
                            CaseNo = lstCaseTrackList.ListItems(i).ListSubItems(1).Text
                            Call UpdateRecord(CaseNo)
                        End If
                    Next i
                    
                    If startTrans = True Then
                        medixmasterconnMaint.commitTrans
                        medixmasterconn.commitTrans
                    End If
                    
                    MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    
                    lstCaseTrackList.ListItems.Clear
                    Call DischargeListing(strAppend)
                Else
                    medixmasterconnMaint.commitTrans
                    medixmasterconn.commitTrans
                End If
            End If
    End If
    
  '  medixmasterconn.Close
 '   Set medixmasterconn = Nothing
        
  '  medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    
    
Exit Sub
errSave:
    MsgBox Err.Description
    If startTrans = True Then
        medixmasterconnMaint.RollbackTrans
        medixmasterconnWS.RollbackTrans
    End If
End Sub

Private Sub CmdPrintFile_Click()
    
    Screen.MousePointer = vbHourglass
    
    If lstCaseTrackList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstCaseTrackList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
  
End Sub
'**********************End Command Button *******************************

'**********************Start Update Record*******************************
Private Sub UpdateRecord(Optional CaseNo As String)
Dim REC As New ADODB.Recordset
Dim strsql As String
Dim strsqlREC As String
Dim TokenDate As String

    TokenDate = Format(Date, "YYYY/MM/DD")
    strsql = "EXEC Case_Send2ClaimRep '" & Trim(CaseNo) & "',4"
    REC.Open strsql, medixmasterconn, adOpenForwardOnly, adLockOptimistic
    
    If REC.EOF Then
        strsqlREC = "EXEC Case_SendRep2InsertRec '" & Trim(CaseNo) & "', '" & TokenDate & "', '" & txtBatchNo & "', '1', '" & Logon & "'"
    Else
        strsqlREC = "EXEC Case_SendRep2UpdateRec '" & Trim(CaseNo) & "', '" & TokenDate & "', '" & txtBatchNo & "', '1', '" & Logon & "'"
    End If
    medixmasterconn.Execute strsqlREC
    REC.Close
    Set REC = Nothing
End Sub
'**********************End Update Record*******************************

'**********************End Search Record*******************************
Private Function SearchRecord()
Dim strsql As String
Dim sql As String
Dim CAS As New ADODB.Recordset
    
    strAppend = ""
        
        If ChkScan.Value = 1 Then
            If Len(Trim(TxtCaseNo)) Then
                strAppend = strAppend + " AND CASNumber = '" & Trim(TxtCaseNo) & "'"
            End If
        Else
            If Len(Trim(TxtCaseNo)) Then
                strAppend = strAppend + " AND CASNumber >= '" & Trim(TxtCaseNo) & "'"
            End If
            
            If Len(Trim(TxtCaseNo2)) Then
                strAppend = strAppend + " AND CASNumber <= '" & Trim(TxtCaseNo2) & "'"
            End If
        End If
        
        If Mid(CboCaseStatus, 1, 1) = 0 Then
            strAppend = strAppend + " And CASStatus = 'O'"
        ElseIf Mid(CboCaseStatus, 1, 1) = 1 Then
            strAppend = strAppend + " And CASStatus = 'D'"
        ElseIf Mid(CboCaseStatus, 1, 1) = 2 Then
            strAppend = strAppend + " And (CASStatus = 'O' OR CASStatus = 'C')"
        End If
        
        If Mid(CboClaimability, 1, 1) = "R" Then
            strAppend = strAppend + " AND CASClaimability = 'R'"
        ElseIf Mid(CboClaimability, 1, 1) = "A" Then
            strAppend = strAppend + " AND CASClaimability = 'A'"
        End If
            
        If Mid(CboDept, 1, 1) = 0 Or Mid(CboDept, 1, 1) = 1 Then
           sql = ""
           sql = " AND  CASClaimAdmin = '" & Trim(Mid(CboDept, 1, 1)) & "'"
           strAppend = strAppend + sql
        End If
        
        If Len(Trim(CboPass)) Then
            strAppend = strAppend + " AND CASUserToken = '" & Trim(CboPass) & "'"
        End If
                   
        If ChkDOA.Value = 1 Then
            sql = ""
            sql = "AND (CasDateAdmitted Is null OR CasDateAdmitted = '')" & ""
            strAppend = strAppend + sql
        End If
        
        If chkDOR.Value = 1 Then
            sql = ""
            sql = "AND (CaseRegDate Is null OR CaseRegDate = '')" & ""
            strAppend = strAppend + sql
        End If
        
        If ChkDOD.Value = 1 Then
            sql = ""
            sql = "AND (CasDischargeDate Is null OR CasDischargeDate = '')" & ""
            strAppend = strAppend + sql
        End If
        
        If ChkDPO.Value = 1 Then
            sql = ""
            sql = "AND (CASTokenDate Is null OR CASTokenDate = '')" & ""
            strAppend = strAppend + sql
        End If
        
        If Len(Trim(txtDORDate1)) > 0 And IsDate(txtDORDate1) = True Then
            strAppend = strAppend + " And CaseRegDate >= '" & Format(txtDORDate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDORDate2)) > 0 And IsDate(txtDORDate2) = True Then
            strAppend = strAppend + " And CaseRegDate <= '" & Format(txtDORDate2, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDOADate1)) > 0 And IsDate(txtDOADate1) = True Then
            strAppend = strAppend + " And CasDateAdmitted >= '" & Format(txtDOADate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDOADate2)) > 0 And IsDate(txtDOADate2) = True Then
            strAppend = strAppend + " And CasDateAdmitted <= '" & Format(txtDOADate2, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDODDate1)) > 0 And IsDate(txtDODDate1) = True Then
            strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDODDate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDODDate2)) > 0 And IsDate(txtDODDate2) = True Then
            strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDODDate2, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDPODate1)) > 0 And IsDate(txtDPODate1) = True Then
            strAppend = strAppend + " And CASTokenDate >= '" & Format(txtDPODate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(txtDPODate2)) > 0 And IsDate(txtDPODate2) = True Then
            strAppend = strAppend + " And CASTokenDate <= '" & Format(txtDPODate2, "mm/dd/yyyy") & "'"
        End If
        
        Call DischargeListing(strAppend)
    
End Function
'**********************End Search Record*******************************

'**********************Start List Record*******************************

Private Sub DischargeListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Current As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String

    Current = 0
    lstCaseTrackList.ListItems.Clear
    
    strAppendCase = "3"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_Send2ClaimRep"
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

    Do
        Do Until rst2.EOF
                
            With rst2
            'If Len(!CaseNo) Then
            'Else
                If !ClaimAdmin = "1" Then
                    Set Record = lstCaseTrackList.ListItems.Add(, , "Passed")
                Else
                    Set Record = lstCaseTrackList.ListItems.Add(, , "Not")
                End If
                If Len(!CasNumber) Then
                    Record.SubItems(1) = !CasNumber
                End If
                If Len(!MBMNumber) Then
                    Record.SubItems(2) = !MBMNumber
                End If
                If Len(!PatientName) Then
                    Record.SubItems(3) = Trim(!PatientName)
                End If
                If Len(!CaseRegDate) Then
                    Record.SubItems(4) = Format(!CaseRegDate, "dd/mm/yyyy")
                End If
                If Len(!CASDateAdmitted) Then
                   Record.SubItems(5) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(!CASDischargeDate) Then
                    Record.SubItems(6) = Format(!CASDischargeDate, "dd/mm/yyyy")
                End If
                If Len(!TokenDate) Then
                    Record.SubItems(7) = Format(!TokenDate, "dd/mm/yyyy")
                End If
                If !NoofDay = "0" Then
                    Record.SubItems(8) = "1"
                ElseIf Len(!NoofDay) Then
                    Record.SubItems(8) = !NoofDay
                End If
                If Len(!TokenBatch) Then
                    Record.SubItems(9) = !TokenBatch
                End If
                If Len(!UserToken) Then
                    Record.SubItems(10) = Trim(!UserToken)
                End If
                If Len(!CaseOpenBy) Then
                    Record.SubItems(11) = Trim(!CaseOpenBy)
                End If
                If Len(!CASHOSCode) Then
                    Record.SubItems(12) = !CASHOSCode
                End If
                If Len(!CASStayType) Then
                    Record.SubItems(13) = Trim(!CASStayType)
                End If
                If Len(!CASPAYType) Then
                    Record.SubItems(14) = Trim(!CASPAYType)
                End If
                If Len(!ActualAmt) Then
                    Record.SubItems(15) = !ActualAmt
                End If
                If Len(!CASClaimability) Then
                    Record.SubItems(16) = !CASClaimability
                End If
                If Len(!ClaimAdmin) Then
                    Record.SubItems(17) = !ClaimAdmin
                End If
                
                Current = Current + 1
                lblCurrent = Current
            'End If
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    
    If lstCaseTrackList.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdPrintFile.Enabled = True
        CmdPass.Enabled = False
        lblCurrent = 0
    End If
    
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
    CmdPass.Enabled = True
End Sub
'********************** End List Record*******************************

'**********************Start Sort ListView*******************************

Private Sub lstCaseTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
    Case 16
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstCaseTrackList, ColumnHeader.Index, ldtString, m_blnDirection(18)
    End Select
End Sub
'**********************Start Sort ListView*******************************


'**********************Start Check ListView Record*******************************
Private Sub lstCaseTrackList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    If Item.Checked = True Then
        If Trim(Item.SubItems(4)) = "" Then
            MsgBox "Date of Register is Empty.", vbCritical
            Item.Checked = False
        ElseIf Trim(Item.SubItems(5)) = "" Then
            MsgBox "Date of Admitted is Empty.", vbCritical
            Item.Checked = False
        ElseIf Trim(Item.SubItems(6)) = "" Then
            MsgBox "Date of Discharge cannot be empty!", vbCritical
            Item.Checked = False
        ElseIf Trim(Item.SubItems(12)) = "" Then
            MsgBox "Hospital Name is Empty.", vbCritical
            Item.Checked = False
        ElseIf Trim(Item.SubItems(17)) = "1" Then
            MsgBox "File has been send out.", vbCritical
            Item.Checked = False
        End If
    End If
End Sub
'**********************End Check ListView Record*******************************

'**********************Start Combobox Listing *******************************
Private Sub ComboListing()
Dim User3 As New ADODB.Recordset
Dim cmd As New ADODB.Command

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchCASToken"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    User3.CursorLocation = adUseClient
    User3.CursorType = adOpenForwardOnly
    User3.CacheSize = 100
    Set User3 = cmd.Execute
    
    Do Until User3.EOF
    If Len(User3!CASUserToken) Then
        CboPass.AddItem User3!CASUserToken
    End If
    User3.MoveNext
    Loop
    User3.Close
    Set User3 = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
    With CboDept
        .AddItem "0 - Admission"
        .AddItem "1 - Claim"
        .AddItem "2 - Both"
    End With
    
    With CboCaseStatus
        .Text = "2 - Open & Close"
        .AddItem "0 - Open"
        .AddItem "1 - Delete"
        .AddItem "2 - Open & Close"
    End With
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
  
End Sub
'**********************Start Combobox Listing *******************************

'**********************Start Combobox Listed *******************************
Private Sub CboCaseStatus_Change()
    If InStr(CboCaseStatus.Text, "null") Then
       CboCaseStatus.Enabled = False
    End If
End Sub

Private Sub CboCaseStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCaseStatus.Text, CboCaseStatus.SelStart) + Chr(KeyAscii) + Mid(CboCaseStatus.Text, CboCaseStatus.SelStart + CboCaseStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCaseStatus.ListCount - 1
 
        If NewText = LCase(Left(CboCaseStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCaseStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCaseStatus.Text = ValidValue
                CboCaseStatus.SelStart = 0
                CboCaseStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClaimability_Change()
    If InStr(CboClaimability.Text, "null") Then
       CboClaimability.Enabled = False
    End If
End Sub

Private Sub CboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClaimability.Text, CboClaimability.SelStart) + Chr(KeyAscii) + Mid(CboClaimability.Text, CboClaimability.SelStart + CboClaimability.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClaimability.ListCount - 1
 
        If NewText = LCase(Left(CboClaimability.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClaimability.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClaimability.Text = ValidValue
                CboClaimability.SelStart = 0
                CboClaimability.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPass_Change()
    If InStr(CboPass.Text, "null") Then
       CboPass.Enabled = False
    End If
End Sub

Private Sub CboPass_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPass.Text, CboPass.SelStart) + Chr(KeyAscii) + Mid(CboPass.Text, CboPass.SelStart + CboPass.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPass.ListCount - 1
 
        If NewText = LCase(Left(CboPass.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPass.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPass.Text = ValidValue
                CboPass.SelStart = 0
                CboPass.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub CboDept_Change()
    If InStr(CboDept.Text, "null") Then
       CboDept.Enabled = False
    End If
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboDept.Text, CboDept.SelStart) + Chr(KeyAscii) + Mid(CboDept.Text, CboDept.SelStart + CboDept.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboDept.ListCount - 1
 
        If NewText = LCase(Left(CboDept.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboDept.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboDept.Text = ValidValue
                CboDept.SelStart = 0
                CboDept.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub txtCaseNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub txtCaseNo_LostFocus()
    TxtCaseNo = ChkStr(TxtCaseNo)
    CmdSearch.Default = False
End Sub

'**********************End Combobox Listed *******************************

Private Sub txtDORDate1_LostFocus()
    If IsDate(txtDORDate1) Then
    Else
        If txtDORDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDORDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDORDate2_LostFocus()
    If IsDate(txtDORDate2) Then
    Else
        If txtDORDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDORDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDOADate1_LostFocus()
    If IsDate(txtDOADate1) Then
    Else
        If txtDOADate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOADate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDOADate2_LostFocus()
    If IsDate(txtDOADate2) Then
    Else
        If txtDOADate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOADate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDODDate1_LostFocus()
    If IsDate(txtDODDate1) Then
    Else
        If txtDODDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDODDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDODDate2_LostFocus()
    If IsDate(txtDODDate2) Then
    Else
        If txtDODDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDODDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDPODate1_LostFocus()
    If IsDate(txtDPODate1) Then
    Else
        If txtDPODate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDPODate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDPODate2_LostFocus()
    If IsDate(txtDPODate2) Then
    Else
        If txtDPODate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDPODate2.SetFocus
        End If
    End If
End Sub

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        lstCaseTrackList.ListItems.Clear
        Label18.Visible = False
        TxtCaseNo2.Visible = False
        lblCurrent = 0
        TxtCaseNo = ""
        TxtCaseNo2 = ""
    Else
        lstCaseTrackList.ListItems.Clear
        Label18.Visible = True
        TxtCaseNo2.Visible = True
        lblCurrent = 0
        TxtCaseNo = ""
        TxtCaseNo2 = ""
    End If
    'TxtCaseNo.SetFocus
End Sub

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtCaseNo)) = ChkStr(lstCaseTrackList.ListItems(listloop1).SubItems(1)) Then
        Check = True
    End If
End Function

Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstCaseTrackList.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstCaseTrackList.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
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

