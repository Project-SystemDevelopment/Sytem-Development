VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMonthlyDNRpt 
   Caption         =   "Monthly Debit Note Report"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   53
      Top             =   7200
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   4320
         TabIndex        =   65
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox lblCurrent 
         Height          =   375
         Left            =   240
         TabIndex        =   64
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   33
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9840
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   7560
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   8280
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   8880
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1680
         TabIndex        =   55
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Total Rec'd Amt"
         Height          =   255
         Left            =   2880
         TabIndex        =   54
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2625
      Left            =   120
      TabIndex        =   34
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox ChkDN 
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
         Height          =   255
         Left            =   11040
         TabIndex        =   26
         Top             =   1420
         Width           =   500
      End
      Begin VB.CheckBox ChkInv 
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
         Height          =   255
         Left            =   11040
         TabIndex        =   23
         Top             =   1150
         Width           =   500
      End
      Begin VB.CheckBox ChkDOR 
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
         Height          =   255
         Left            =   11040
         TabIndex        =   20
         Top             =   885
         Width           =   500
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
         Height          =   255
         Left            =   11040
         TabIndex        =   17
         Top             =   600
         Width           =   500
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
         Height          =   255
         Left            =   11040
         TabIndex        =   14
         Top             =   285
         Width           =   500
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   315
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   315
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtDOADate1 
         Height          =   315
         Left            =   7320
         TabIndex        =   12
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate2 
         Height          =   315
         Left            =   9480
         TabIndex        =   13
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   525
         Width           =   3615
      End
      Begin VB.TextBox TxtVersion1 
         Height          =   315
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   3
         Top             =   750
         Width           =   1575
      End
      Begin VB.TextBox TxtVersion2 
         Height          =   315
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtMBMNo1 
         Height          =   315
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1000
         Width           =   1575
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   315
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1000
         Width           =   1575
      End
      Begin VB.ComboBox CboCaseStatus 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1290
         Width           =   3615
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   1560
         Width           =   3615
      End
      Begin VB.ComboBox CboCaseRegBy 
         Height          =   315
         Left            =   3720
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1850
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtDODDate1 
         Height          =   315
         Left            =   7320
         TabIndex        =   15
         Top             =   530
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate2 
         Height          =   315
         Left            =   9480
         TabIndex        =   16
         Top             =   530
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDORDate1 
         Height          =   330
         Left            =   7320
         TabIndex        =   18
         Top             =   810
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   315
         Left            =   7320
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDORDate2 
         Height          =   330
         Left            =   9480
         TabIndex        =   19
         Top             =   810
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   315
         Left            =   9480
         TabIndex        =   22
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDNDate1 
         Height          =   315
         Left            =   7320
         TabIndex        =   24
         Top             =   1370
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDNDate2 
         Height          =   315
         Left            =   9480
         TabIndex        =   25
         Top             =   1370
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtRecAmtfr 
         Height          =   285
         Left            =   7320
         TabIndex        =   27
         Top             =   1650
         Width           =   1455
      End
      Begin VB.TextBox txtRecAmtto 
         Height          =   285
         Left            =   9480
         TabIndex        =   28
         Top             =   1650
         Width           =   1455
      End
      Begin VB.ComboBox CboCashType 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1850
         Width           =   1575
      End
      Begin VB.ComboBox CboInvRegBy 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2130
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "Cash Type"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   1890
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Case Status"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "To"
         Height          =   255
         Left            =   9000
         TabIndex        =   61
         Top             =   1400
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "DN Date from"
         Height          =   255
         Left            =   6100
         TabIndex        =   60
         ToolTipText     =   "Date of Admission"
         Top             =   1400
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         Height          =   255
         Left            =   9000
         TabIndex        =   59
         Top             =   1130
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Inv. Date from"
         Height          =   255
         Left            =   6100
         TabIndex        =   58
         ToolTipText     =   "Date of Admission"
         Top             =   1130
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   9000
         TabIndex        =   52
         Top             =   870
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   9000
         TabIndex        =   51
         Top             =   580
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   9000
         TabIndex        =   50
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "DOR from"
         Height          =   255
         Left            =   6100
         TabIndex        =   49
         ToolTipText     =   "Date Pass On"
         Top             =   870
         Width           =   885
      End
      Begin VB.Label Label9 
         Caption         =   "DOD from"
         Height          =   255
         Left            =   6100
         TabIndex        =   48
         ToolTipText     =   "Date of Discharged"
         Top             =   580
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "DOA from"
         Height          =   255
         Left            =   6100
         TabIndex        =   47
         ToolTipText     =   "Date of Admission"
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Ver. No"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   820
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   520
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Case Reg. By"
         Height          =   495
         Left            =   3120
         TabIndex        =   44
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Wks. No"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   41
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Member No"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1100
         Width           =   840
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   1095
         Width           =   360
      End
      Begin VB.Label Label25 
         Caption         =   "To"
         Height          =   255
         Left            =   9000
         TabIndex        =   38
         Top             =   1700
         Width           =   495
      End
      Begin VB.Label Label24 
         Caption         =   "Rec'd Amt from"
         Height          =   255
         Left            =   6100
         TabIndex        =   37
         ToolTipText     =   "Approved Amt"
         Top             =   1700
         Width           =   1125
      End
      Begin VB.Label Label32 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   36
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label33 
         Caption         =   "Inv. Reg. By"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   2200
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstWorksheetTrackList 
      Height          =   4215
      Left            =   120
      TabIndex        =   56
      Top             =   3000
      Width           =   11655
      _ExtentX        =   20558
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
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Claim No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Cash Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Case Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Claimability"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Bill Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Approved Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Case Reg. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Case Reg. By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Inv. Reg. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Inv. Reg. By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Debit Note No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Debit Note Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Debit No Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "DN Open By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Amt Rec'd From MBM "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Text            =   "Rec'd Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Monthly Debit Note Report"
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
      TabIndex        =   57
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmMonthlyDNRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Private m_blnDirection(1 To 17) As Boolean

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

Private Sub CmdPrintFile_Click()
    'Call ExportToExcelAnalysis1(lstWorksheetTrackList)
    Screen.MousePointer = vbHourglass
    
    If lstWorksheetTrackList.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstWorksheetTrackList)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = Default
End Sub

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    Call ComboListing
End Sub

'******************** Start Command Button ********************************
Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstWorksheetTrackList.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
End Sub

Public Sub cmdSearch_Click()
On Error GoTo errRun
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    CmdPrintFile.Enabled = False
    CmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
        'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            Call SearchRecord
       ' medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
    Screen.MousePointer = Default
    CmdSearch.Enabled = True
Exit Sub
errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstBordxDelivery.ListItems.Clear
    lblCurrent = ""
    CmdPrintFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing

End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
'******************** End Command Button ********************************

'******************** Start Search Record ********************************

Private Function SearchRecord()
Dim StrSql As String
Dim StrAppend As String
Dim Sql As String

    
    StrAppend = ""
        
    
    If Len(Trim(CboPayor)) Then
        StrAppend = StrAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtCaseNo1)) Then
        StrAppend = StrAppend + " AND CASNumber >= '" & Trim(TxtCaseNo1) & "'"
    End If
        
    If Len(Trim(TxtCaseNo2)) Then
        StrAppend = StrAppend + " AND CASNumber <= '" & Trim(TxtCaseNo2) & "'"
    End If
        
    If Len(Trim(TxtVersion1)) Then
        StrAppend = StrAppend + " And claimversion >= '" & Trim(TxtVersion1) & "'"
    End If
    
    If Len(Trim(TxtVersion2)) Then
        StrAppend = StrAppend + " And claimversion <= '" & Trim(TxtVersion2) & "'"
    End If
      
    If Mid(CboClaimability, 1, 1) = "R" Then
        StrAppend = StrAppend + " AND CASClaimability = 'R'"
    End If
        
    If Mid(CboClaimability, 1, 1) = "A" Then
        StrAppend = StrAppend + " AND CASClaimability = 'A'"
    End If
    
    If Len(Trim(CboCashType)) Then
        StrAppend = StrAppend + " And CASPAYType = '" & Mid(CboCashType, 1, 1) & "'"
    End If
    
    If Mid(CboCaseStatus, 1, 1) = 0 Then
        StrAppend = StrAppend + " And CASStatus = 'O'"
    ElseIf Mid(CboCaseStatus, 1, 1) = 1 Then
        StrAppend = StrAppend + " And CASStatus = 'D'"
    ElseIf Mid(CboCaseStatus, 1, 1) = 2 Then
        StrAppend = StrAppend + " And (CASStatus = 'O' OR CASStatus = 'C')"
    End If
    
    If Len(Trim(CboCaseRegBy)) Then
        StrAppend = StrAppend + " AND CaseOpenBy = '" & Trim(CboCaseRegBy) & "'"
    End If
    
    If Len(Trim(CboInvRegBy)) Then
        StrAppend = StrAppend + " AND INVEnteredBy = '" & Trim(CboInvRegBy) & "'"
    End If
    
    If Len(txtRecAmtfr) Then
        StrAppend = StrAppend + " AND CLMTotalApproved >= " & Trim(txtRecAmtfr) & ""
    End If
        
    If Len(txtRecAmtto) Then
        StrAppend = StrAppend + " AND CLMTotalApproved <= " & Trim(txtRecAmtto) & ""
    End If
    
    If Len(Trim(TxtMBMNo1)) Then
        StrAppend = StrAppend + " AND MBMNumber >= '" & Trim(TxtMBMNo1) & "'"
    End If
    
    If Len(Trim(TxtMBMNo2)) Then
        StrAppend = StrAppend + " AND MBMNumber <= '" & Trim(TxtMBMNo2) & "'"
    End If
    
    If Len(Trim(txtDORDate1)) > 0 And IsDate(txtDORDate1) = True Then
        StrAppend = StrAppend + " And CaseRegDate >= '" & Format(txtDORDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDORDate2)) > 0 And IsDate(txtDORDate2) = True Then
        StrAppend = StrAppend + " And CaseRegDate <= '" & Format(txtDORDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOADate1)) > 0 And IsDate(txtDOADate1) = True Then
        StrAppend = StrAppend + " And CasDateAdmitted >= '" & Format(txtDOADate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOADate2)) > 0 And IsDate(txtDOADate2) = True Then
        StrAppend = StrAppend + " And CasDateAdmitted <= '" & Format(txtDOADate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDODDate1)) > 0 And IsDate(txtDODDate1) = True Then
        StrAppend = StrAppend + " And CasDischargeDate >= '" & Format(txtDODDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDODDate2)) > 0 And IsDate(txtDODDate2) = True Then
        StrAppend = StrAppend + " And CasDischargeDate <= '" & Format(txtDODDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDNDate1)) > 0 And IsDate(txtDNDate1) = True Then
        StrAppend = StrAppend + " And ClaimOpenDate >= '" & Format(txtDNDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDNDate2)) > 0 And IsDate(txtDNDate2) = True Then
        StrAppend = StrAppend + " And ClaimOpenDate <= '" & Format(txtDNDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDate1)) > 0 And IsDate(txtInvDate1) = True Then
        StrAppend = StrAppend + " And INVEnteredDate >= '" & Format(txtInvDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtInvDate2)) > 0 And IsDate(txtInvDate2) = True Then
        StrAppend = StrAppend + " And INVEnteredDate <= '" & Format(txtInvDate2, "mm/dd/yyyy") & "'"
    End If
    
    If ChkDOA.Value = 1 Then
        Sql = ""
        Sql = "AND (CasDateAdmitted Is null OR CasDateAdmitted = '')" & ""
        StrAppend = StrAppend + Sql
    End If
    
    If ChkDOD.Value = 1 Then
        Sql = ""
        Sql = "AND (CasDischargeDate Is null OR CasDischargeDate = '')" & ""
        StrAppend = StrAppend + Sql
    End If
    
    If ChkDOR.Value = 1 Then
        Sql = ""
        Sql = "AND (CaseRegDate Is null OR CaseRegDate = '')" & ""
        StrAppend = StrAppend + Sql
    End If
    
    If ChkInv.Value = 1 Then
        Sql = ""
        Sql = "AND (INVEnteredDate Is null OR INVEnteredDate = '')" & ""
        StrAppend = StrAppend + Sql
    End If
    
    If ChkDN.Value = 1 Then
        Sql = ""
        Sql = "AND (ClaimOpenDate Is null OR ClaimOpenDate = '')" & ""
        StrAppend = StrAppend + Sql
    End If
        
    Call WorksheetTrackingListing(StrAppend)
        
End Function
'******************** End Search Record ********************************

'******************** Start WorksheetTracking Listing ********************************

Private Sub WorksheetTrackingListing(Optional StrAppend As String)
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
    
    lstWorksheetTrackList.ListItems.Clear
    'Sql = " Select * From VIEW_Monthly_DNRpt Where 0 = 0 "
    
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend
    'End If

    'rst2.Open Sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "MonthlyDNRpt"
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
        CmdPrintFile.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        
            With rst2
                Set Record = lstWorksheetTrackList.ListItems.Add(, , (rst2!CasNumber & "-" & rst2!claimversion))
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(1) = Trim(rst2!MBMNumber)
                End If
                If Len(rst2!PatientName) Then
                    Record.SubItems(2) = Trim(rst2!PatientName)
                End If
                If Len(rst2!CASPAYType) Then
                   Record.SubItems(3) = rst2!CASPAYType
                End If
                If Len(rst2!CASStatus) Then
                   Record.SubItems(4) = rst2!CASStatus
                End If
                If Len(rst2!CASClaimability) Then
                   Record.SubItems(5) = rst2!CASClaimability
                End If
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(6) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(7) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                If Len(rst2!CaseRegDate) Then
                    Record.SubItems(8) = Format(rst2!CaseRegDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CaseOpenBy) Then
                   Record.SubItems(9) = rst2!CaseOpenBy
                End If
                If Len(rst2!INVEnteredDate) Then
                    Record.SubItems(10) = Format(rst2!INVEnteredDate, "dd/mm/yyyy")
                End If
                If Len(rst2!INVEnteredBy) Then
                   Record.SubItems(11) = rst2!INVEnteredBy
                End If
                If Len(rst2!DebitCreditRefNo) Then
                   Record.SubItems(12) = rst2!DebitCreditRefNo
                End If
                If Len(rst2!DebitCreditDate) Then
                    Record.SubItems(13) = Format(rst2!DebitCreditDate, "dd/mm/yyyy")
                End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(14) = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                End If
                If Len(rst2!ClaimOpenBy) Then
                   Record.SubItems(15) = Format(rst2!ClaimOpenBy, "dd/mm/yyyy")
                End If
                If Len(rst2!ReceiptAmt) Then
                    Record.SubItems(16) = Format(rst2!ReceiptAmt, "#,##0.00")
                End If
                If Len(rst2!ReceiptDate) Then
                    Record.SubItems(17) = Format(rst2!ReceiptDate, "dd/mm/yyyy")
                End If
                                            
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!ReceiptAmt) Then
                    TotalAmt = TotalAmt + rst2!ReceiptAmt
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
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
End Sub
'******************** End WorksheetTracking Listing ********************************

'**********************Start Sorted ListView *******************************
Private Sub lstWorksheetTrackList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(14)
    Case 15
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
    Case 16
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
    Case 18
        SortListView lstWorksheetTrackList, ColumnHeader.Index, ldtDateTime, m_blnDirection(18)
    End Select
End Sub
'**********************End Sorted ListView *******************************

'******************** Start ComboListing ********************************
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim User As New ADODB.Recordset
Dim User3 As New ADODB.Recordset
Dim cmd As New ADODB.Command


'PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
'User.Open "Select Distinct CaseOpenBy from CAS", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
'User3.Open "Select Distinct INVEnteredBy from ClaimInvoice", medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
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
            CboPayor.AddItem !PAYCode & " - " & Trim(!PAYCompanyName)
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchCASUserOpen"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    User.CursorLocation = adUseClient
    User.CursorType = adOpenForwardOnly
    User.CacheSize = 100
    Set User = cmd.Execute

    Do Until User.EOF
        With User
            If Len(!CaseOpenBy) Then
                CboCaseRegBy.AddItem !CaseOpenBy
            End If
            User.MoveNext
        End With
    Loop
    User.Close
    Set User = Nothing
  
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "SearchInvUser"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    User3.CursorLocation = adUseClient
    User3.CursorType = adOpenForwardOnly
    User3.CacheSize = 100
    Set User3 = cmd.Execute
  
    Do Until User3.EOF
        With User3
            If Len(User3!INVEnteredBy) Then
                CboInvRegBy.AddItem !INVEnteredBy
            End If
            User3.MoveNext
        End With
    Loop
    User3.Close
    Set User3 = Nothing
    
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    With CboCaseStatus
        .AddItem "0 - Open"
        .AddItem "1 - Delete"
        .AddItem "2 - Open & Close"
    End With
    
    With CboCashType
        CboCashType.AddItem "C - Cashless"
        CboCashType.AddItem "R - Reimbursement"
    End With
       
End Sub

'******************** End ComboListing ********************************

'******************** Start ComboBox Change/KeyPress ********************************

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

Private Sub CboInvRegBy_Change()
    If InStr(CboInvRegBy.Text, "null") Then
       CboInvRegBy.Enabled = False
    End If
End Sub

Private Sub CboInvRegBy_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboInvRegBy.Text, CboInvRegBy.SelStart) + Chr(KeyAscii) + Mid(CboInvRegBy.Text, CboInvRegBy.SelStart + CboInvRegBy.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboInvRegBy.ListCount - 1
 
        If NewText = LCase(Left(CboInvRegBy.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboInvRegBy.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboInvRegBy.Text = ValidValue
                CboInvRegBy.SelStart = 0
                CboInvRegBy.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPayor_Change()
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End Sub
'******************** End ComboBox Change/KeyPress ********************************


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

Private Sub txtInvDate1_LostFocus()
    If IsDate(txtInvDate1) Then
    Else
        If txtInvDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtInvDate2_LostFocus()
    If IsDate(txtInvDate2) Then
    Else
        If txtInvDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtInvDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDNDate1_LostFocus()
    If IsDate(txtDNDate1) Then
    Else
        If txtDNDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDNDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDNDate2_LostFocus()
    If IsDate(txtDNDate2) Then
    Else
        If txtDNDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDNDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtRecAmtfr_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtRecAmtto_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub CboCaseRegBy_Change()
    If InStr(CboCaseRegBy.Text, "null") Then
       CboCaseRegBy.Enabled = False
    End If
End Sub


Private Sub CboCaseRegBy_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCaseRegBy.Text, CboCaseRegBy.SelStart) + Chr(KeyAscii) + Mid(CboCaseRegBy.Text, CboCaseRegBy.SelStart + CboCaseRegBy.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCaseRegBy.ListCount - 1
 
        If NewText = LCase(Left(CboCaseRegBy.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCaseRegBy.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCaseRegBy.Text = ValidValue
                CboCaseRegBy.SelStart = 0
                CboCaseRegBy.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCashType_Change()
    If InStr(CboCashType.Text, "null") Then
       CboCashType.Enabled = False
    End If
End Sub

Private Sub CboCashType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCashType.Text, CboCashType.SelStart) + Chr(KeyAscii) + Mid(CboCashType.Text, CboCashType.SelStart + CboCashType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCashType.ListCount - 1
 
        If NewText = LCase(Left(CboCashType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCashType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCashType.Text = ValidValue
                CboCashType.SelStart = 0
                CboCashType.SelLength = Len(ValidValue)
            End If
        End If
End Sub
