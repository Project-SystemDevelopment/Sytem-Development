VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAuditReviewReport 
   Caption         =   "Audit Review Tracking"
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
   Begin MSFlexGridLib.MSFlexGrid Fgrid 
      Height          =   3255
      Left            =   120
      TabIndex        =   62
      Top             =   3960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   16
      Cols            =   8
      FixedRows       =   2
      BackColorFixed  =   -2147483638
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame5 
      Caption         =   "Reports Type 1"
      Height          =   3675
      Left            =   6240
      TabIndex        =   53
      Top             =   240
      Width           =   5415
      Begin VB.OptionButton Opt2 
         Caption         =   "Monthly AR Responses"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   1995
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Monthly Productivity "
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Value           =   -1  'True
         Width           =   1995
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   44
      Top             =   7200
      Width           =   11655
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   67
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   360
         Left            =   6360
         TabIndex        =   30
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   360
         Left            =   7080
         TabIndex        =   31
         Top             =   120
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
         Height          =   360
         Left            =   7680
         TabIndex        =   32
         Top             =   120
         Width           =   615
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
         Height          =   360
         Left            =   8280
         TabIndex        =   33
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9840
         TabIndex        =   35
         Top             =   100
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   36
         Top             =   100
         Width           =   615
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   8880
         TabIndex        =   34
         Top             =   100
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   195
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selection Criteria"
      Height          =   3675
      Left            =   120
      TabIndex        =   37
      Top             =   240
      Width           =   6135
      Begin VB.CheckBox ChkDR 
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
         Left            =   5520
         TabIndex        =   18
         Top             =   2220
         Width           =   495
      End
      Begin VB.CheckBox ChkDS 
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
         Left            =   5520
         TabIndex        =   15
         Top             =   1965
         Width           =   495
      End
      Begin VB.CheckBox ChkARStatus 
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
         Left            =   3120
         TabIndex        =   61
         Top             =   3330
         Width           =   495
      End
      Begin VB.CheckBox chkDO 
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
         Left            =   5520
         TabIndex        =   9
         Top             =   1430
         Width           =   495
      End
      Begin VB.CheckBox ChkDC 
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
         Left            =   5520
         TabIndex        =   12
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox ChkFinal 
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
         Left            =   5520
         TabIndex        =   24
         Top             =   2790
         Width           =   495
      End
      Begin VB.CheckBox ChkFirst 
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
         Left            =   5520
         TabIndex        =   21
         Top             =   2535
         Width           =   375
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   3855
      End
      Begin VB.CheckBox ChkHOS 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5520
         TabIndex        =   6
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TxtCaseNo 
         Height          =   310
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   1
         Top             =   525
         Width           =   1560
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   310
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   2
         Top             =   525
         Width           =   1680
      End
      Begin VB.TextBox TxtMBMNo1 
         Height          =   310
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   3
         Top             =   810
         Width           =   1560
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   310
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   4
         Top             =   810
         Width           =   1680
      End
      Begin VB.ComboBox CboHos 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1095
         Width           =   3855
      End
      Begin MSMask.MaskEdBox txtDODate1 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1380
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODate2 
         Height          =   285
         Left            =   3720
         TabIndex        =   8
         Top             =   1380
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDCDate2 
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         Top             =   1635
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDCDate1 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   1640
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDSDate2 
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   1900
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDRDate2 
         Height          =   285
         Left            =   3720
         TabIndex        =   17
         Top             =   2175
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDSDate1 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   1900
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDRDate1 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   2175
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtFirst1 
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   19
         Top             =   2450
         Width           =   1485
      End
      Begin VB.TextBox TxtFirst2 
         Height          =   285
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   20
         Top             =   2450
         Width           =   1725
      End
      Begin VB.TextBox TxtFinal1 
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   22
         Top             =   2700
         Width           =   1485
      End
      Begin VB.TextBox TxtFinal2 
         Height          =   285
         Left            =   3735
         MaxLength       =   3
         TabIndex        =   23
         Top             =   2700
         Width           =   1725
      End
      Begin VB.TextBox TxtDiscountAmt1 
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   25
         Top             =   2960
         Width           =   1485
      End
      Begin VB.TextBox TxtDiscountAmt2 
         Height          =   285
         Left            =   3735
         MaxLength       =   3
         TabIndex        =   26
         Top             =   2960
         Width           =   1725
      End
      Begin VB.ComboBox CboARStatus 
         Height          =   315
         Left            =   1560
         TabIndex        =   27
         Top             =   3230
         Width           =   1470
      End
      Begin VB.Label Label11 
         Caption         =   "Date Sent From"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Date Reply From"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   2235
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   64
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   63
         Top             =   2235
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   60
         Top             =   3060
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Discount Amt"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   3060
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   58
         Top             =   1700
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   57
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Date Closed From"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   1700
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Date Open From"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "A/R Status"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label38 
         Caption         =   "Final Diag."
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   2805
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   51
         Top             =   2805
         Width           =   375
      End
      Begin VB.Label Label36 
         Caption         =   "First Diag."
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   2550
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   49
         Top             =   2550
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1150
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   260
         Left            =   240
         TabIndex        =   42
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Case No"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   40
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label42 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   39
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label43 
         Caption         =   "Member No"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   900
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstCaseTrackList 
      Height          =   2775
      Left            =   120
      TabIndex        =   46
      Top             =   4440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4895
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Month"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "No of files"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "RM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "No of files"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "RM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Avg Claim per file (RM)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "No of files"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "RM"
         Object.Width           =   2540
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
      TabIndex        =   48
      Top             =   7750
      Width           =   2055
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Audit Review Tracking"
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
      TabIndex        =   47
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmAuditReviewReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intExit As Integer
Dim CC As Integer
Dim HOspitals As String
Dim Month As String
Dim intloop As Integer
Dim ScanCode As String
Dim CASNo As String
Dim m_blnDirection(1 To 8) As Boolean

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
    intloop = 0
    Call ComboListing
    Call DoInitialSettings
End Sub

'**********************Start Command Button *******************************
Private Sub CmdBack_Click()
    Frame1.Visible = True
    Frame5.Visible = True
    Fgrid.Top = 3960
    Fgrid.Width = 11775
    Fgrid.Left = 120
    Fgrid.Height = 3255
End Sub

Private Sub CmdView_Click()
    Frame1.Visible = False
    Frame5.Visible = False
    Fgrid.Top = 360
    Fgrid.Width = 11655
    Fgrid.Left = 120
    Fgrid.Height = 6855
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lstCaseTrackList.ListItems.Clear
    
    'If Opt1.value = True Then
        For FCol = 1 To 7
            Fgrid.col = FCol
            For Frow = 2 To 15
                Fgrid.row = Frow
                Fgrid = "-"
            Next
        Next
    'ElseIf Opt2.value = True Then
        'For FCol = 1 To 4
            'Fgrid.Col = FCol
            'For Frow = 2 To 15
                'Fgrid.Row = Frow
                'Fgrid = "-"
            'Next
        'Next
    'End If
    
    lblCurrent = 0
    intloop = 0
End Sub

Public Sub CmdSearch_Click()

ScanCode = TxtCaseNo
If Mid(ScanCode, 1, 2) = "LO" Then
    TxtCaseNo = ""
    Call SearchForm(ScanCode)
Else
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    CmdPrintFile.Enabled = False
        
    Screen.MousePointer = vbHourglass
        CmdSearch.Enabled = False
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call ClearFgrid
            Call SearchRecord
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        CmdSearch.Enabled = True
    Screen.MousePointer = Default
        
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
        
End If

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrintFile_Click()
    If Opt1.Value = True Then
        Call ExportToExcelWorksheet
    ElseIf Opt2.Value = True Then
        Call ExportToExcelWorksheet2
    End If
End Sub
'**********************End Command Button *******************************

'**********************End Search Record*******************************
Private Function SearchRecord()
Dim StrSql As String
Dim strAppend As String
Dim sql As String
Dim CAS As New ADODB.Recordset
    
    strAppend = ""
    HOspitals = ""
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And SUBSTRING(CASNumber,1,2) = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    Else
        'StrAppend = StrAppend + " And SUBSTRING(CASNumber,1,1) <> 'Z'"
    End If
    
    If Len(Trim(TxtCaseNo)) Then
        strAppend = strAppend + " AND CASNumber >= '" & Trim(TxtCaseNo) & "'"
    End If
        
    If Len(Trim(TxtCaseNo2)) Then
        strAppend = strAppend + " AND CASNumber <= '" & Trim(TxtCaseNo2) & "'"
    End If
    
    If Len(Trim(TxtMBMNo1)) Then
        strAppend = strAppend + " AND MBMNumber >= '" & RemStr(Trim(TxtMBMNo1)) & "'"
    End If
        
    If Len(Trim(TxtMBMNo2)) Then
        strAppend = strAppend + " AND MBMNumber <= '" & RemStr(Trim(TxtMBMNo2)) & "'"
    End If
        
    If Len(Trim(CboHos)) Then
        HOspitals = Replace(CboHos, "'", "''")
        strAppend = strAppend + " And HOSName = '" & Trim(HOspitals) & "'"
    End If
        
    If ChkHOS.Value = 1 Then
        sql = ""
        sql = "AND (HOSName Is null OR HOSName = '') " & ""
        strAppend = strAppend + sql
    End If
        
    If Len(Trim(txtDODate1)) > 0 And IsDate(txtDODate1) = True Then
        strAppend = strAppend + " And CASAROpenDate >= '" & Format(txtDODate1, "mm/dd/yyyy") & "'"
    End If
        
    If Len(Trim(txtDODate2)) > 0 And IsDate(txtDODate2) = True Then
        strAppend = strAppend + " And CASAROpenDate <= '" & Format(txtDODate2, "mm/dd/yyyy") & "'"
    End If
           
    If chkDO.Value = 1 Then
        sql = ""
        sql = "AND (CASAROpenDate Is null OR CASAROpenDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDSDate1)) > 0 And IsDate(txtDSDate1) = True Then
        strAppend = strAppend + " And CASLetterSent >= '" & Format(txtDSDate1, "mm/dd/yyyy") & "'"
    End If
        
    If Len(Trim(txtDSDate2)) > 0 And IsDate(txtDSDate2) = True Then
        strAppend = strAppend + " And CASLetterSent <= '" & Format(txtDSDate2, "mm/dd/yyyy") & "'"
    End If
           
    If ChkDS.Value = 1 Then
        sql = ""
        sql = "AND (CASLetterSent Is null OR CASLetterSent = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDRDate1)) > 0 And IsDate(txtDRDate1) = True Then
        strAppend = strAppend + " And CASHospReply >= '" & Format(txtDRDate1, "mm/dd/yyyy") & "'"
    End If
        
    If Len(Trim(txtDRDate2)) > 0 And IsDate(txtDRDate2) = True Then
        strAppend = strAppend + " And CASHospReply <= '" & Format(txtDRDate2, "mm/dd/yyyy") & "'"
    End If
           
    If ChkDR.Value = 1 Then
        sql = ""
        sql = "AND (CASHospReply Is null OR CASHospReply = '')" & ""
        strAppend = strAppend + sql
    End If
        
    If Len(Trim(txtDCDate1)) > 0 And IsDate(txtDCDate1) = True Then
        strAppend = strAppend + " And CASDTC >= '" & Format(txtDCDate1, "mm/dd/yyyy") & "'"
    End If
        
    If Len(Trim(txtDCDate2)) > 0 And IsDate(txtDCDate2) = True Then
        strAppend = strAppend + " And CASDTC <= '" & Format(txtDCDate2, "mm/dd/yyyy") & "'"
    End If
        
    If ChkDC.Value = 1 Then
        sql = ""
        sql = "AND (CASDTC Is null OR CASDTC = '')" & ""
        strAppend = strAppend + sql
    End If
        
    If Len(Trim(TxtFirst1)) Then
        strAppend = strAppend + " AND FirstICDCode >= '" & Trim(TxtFirst1) & "'"
    End If
        
    If Len(Trim(TxtFirst2)) Then
        strAppend = strAppend + " AND FirstICDCode <= '" & Trim(TxtFirst2) & "'"
    End If
        
    If ChkFirst.Value = 1 Then
        sql = ""
        sql = "AND (FirstICDCode Is null OR FirstICDCode = '') " & ""
        strAppend = strAppend + sql
    End If
        
    If Len(Trim(TxtFinal1)) Then
        strAppend = strAppend + " AND FinalICDCode >= '" & Trim(TxtFinal1) & "'"
    End If
        
    If Len(Trim(TxtFinal2)) Then
        strAppend = strAppend + " AND FinalICDCode <= '" & Trim(TxtFinal2) & "'"
    End If
            
    If ChkFinal.Value = 1 Then
        sql = ""
        sql = "AND (FinalICDCode Is null OR FinalICDCode = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If Len(TxtDiscountAmt1) Then
        strAppend = strAppend + " AND CASDisRec >= " & Trim(TxtDiscountAmt1) & ""
    End If
        
    If Len(TxtDiscountAmt2) Then
        strAppend = strAppend + " AND CASDisRec <= " & Trim(TxtDiscountAmt2) & ""
    End If
    
    If Mid(CboARStatus, 1, 1) = 0 Then
        strAppend = strAppend + " And CASARStatus = 'O'"
    ElseIf Mid(CboARStatus, 1, 1) = 1 Then
        strAppend = strAppend + " And CASARStatus = 'C'"
    ElseIf Mid(CboARStatus, 1, 1) = 2 Then
        strAppend = strAppend + " And (CASARStatus = 'O' OR CASARStatus = 'C')"
    End If
    
    If ChkARStatus.Value = 1 Then
        sql = ""
        sql = "AND (CASARStatus Is null OR CASARStatus = '') " & ""
        strAppend = strAppend + sql
    End If
    
    If Opt1.Value = True Then
        Call ProductivityRpt(strAppend)
        Call ProductivityRpt2(strAppend)
        Call FileCalculation
    ElseIf Opt2.Value = True Then
        Call ResponsesRpt(strAppend)
        Call FileCalculation2
    End If
End Function
'**********************End Search Record*******************************

'**********************Start Combobox Listing *******************************
Private Sub ComboListing()
Dim HOS As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHOS"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cmd.Execute
        
    Do Until HOS.EOF
        With HOS
            CboHos.AddItem !HosName
        End With
        HOS.MoveNext
    Loop
    HOS.Close
    Set HOS = Nothing
    
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
        End With
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
        
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
 
    
    With CboARStatus
        .Text = "2 - Open & Close"
        .AddItem "0 - Open"
        .AddItem "1 - Close"
        .AddItem "2 - Open & Close"
    End With
    
End Sub
'**********************Start Combobox Listing *******************************

'**********************Start Combobox Listed *******************************
Private Sub cboHOS_Change()
    If InStr(CboHos.Text, "null") Then
       CboHos.Enabled = False
    End If
End Sub

Private Sub cboHOS_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboHos.Text, CboHos.SelStart) + Chr(KeyAscii) + Mid(CboHos.Text, CboHos.SelStart + CboHos.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboHos.ListCount - 1
 
        If NewText = LCase(Left(CboHos.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboHos.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboHos.Text = ValidValue
                CboHos.SelStart = 0
                CboHos.SelLength = Len(ValidValue)
            End If
        End If
End Sub

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
Dim StrSql As String
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
    
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
        
    'If ChkGrpCom.value = 1 Then
    
    'CboGrpCom.Clear
    'StrSql = "select MBMGroupCompany from View_PAYGrpCom where PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
    'PAYGRP.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    'strAppendCase = "1"
    'strAppend = " AND PAYCode= '" & Mid(CboPayor, 1, 2) & "'"
    'Set cmd.ActiveConnection = medixmasterconn
    'cmd.CommandText = "Case_AuditReviewRpt"
    'cmd.CommandType = adCmdStoredProc
    'cmd.CommandTimeout = 300
    'Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    'cmd.Parameters.Append prm1
    'Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    'cmd.Parameters.Append prm2
    'PAYGRP.CursorLocation = adUseClient
    'PAYGRP.CursorType = adOpenForwardOnly
    'PAYGRP.CacheSize = 100
    'Set PAYGRP = cmd.Execute
    
    'If Not PAYGRP.EOF Then
    '    Do Until PAYGRP.EOF
    '       CboGrpCom.AddItem Trim(PAYGRP!MBMGroupCompany)
    '       PAYGRP.MoveNext
    '    Loop
    'End If
    'PAYGRP.Close
    'Set PAYGRP = Nothing
    
    'End If
End Sub

Private Sub Opt1_Click()
    
    For FCol = 1 To 7
        Fgrid.col = FCol
        For Frow = 2 To 15
            Fgrid.row = Frow
            Fgrid = "-"
        Next
    Next
    
    Call DoInitialSettings
    
End Sub

Private Sub Opt2_Click()
    For FCol = 1 To 7
        Fgrid.col = FCol
        For Frow = 2 To 15
            Fgrid.row = Frow
            Fgrid = "-"
        Next
    Next
    
    Call DoInitialSettings2
End Sub

Private Sub txtCaseNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtCaseNo_LostFocus()
    TxtCaseNo = ChkStr(TxtCaseNo)
    CmdSearch.Default = False
End Sub

'**********************End Combobox Listed *******************************

Private Sub txtDCDate1_LostFocus()
    If IsDate(txtDCDate1) Then
    Else
        If txtDCDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDCDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDCDate2_LostFocus()
    If IsDate(txtDCDate2) Then
    Else
        If txtDCDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDCDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDODate1_LostFocus()
    If IsDate(txtDODate1) Then
    Else
        If txtDODate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDODate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDODate2_LostFocus()
    If IsDate(txtDODate2) Then
    Else
        If txtDODate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDODate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDSDate1_LostFocus()
    If IsDate(txtDSDate1) Then
    Else
        If txtDSDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDSDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDSDate2_LostFocus()
    If IsDate(txtDSDate2) Then
    Else
        If txtDSDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDSDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDRDate1_LostFocus()
    If IsDate(txtDRDate1) Then
    Else
        If txtDRDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDRDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDRDate2_LostFocus()
    If IsDate(txtDRDate2) Then
    Else
        If txtDRDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDRDate2.SetFocus
        End If
    End If
End Sub


Private Sub CboARStatus_Change()
    If CboARStatus.Text = "" Then
        If InStr(CboARStatus.Text, "null") Then
           CboARStatus.Enabled = False
        End If
    End If
End Sub

Private Sub CboARStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboARStatus.Text, CboARStatus.SelStart) + Chr(KeyAscii) + Mid(CboARStatus.Text, CboARStatus.SelStart + CboARStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboARStatus.ListCount - 1
 
        If NewText = LCase(Left(CboARStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboARStatus.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboARStatus.Text = ValidValue
                CboARStatus.SelStart = 0
                CboARStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub ProductivityRpt(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
        
    TotalReceivedCase = 0
    TotalReceivedAmt = 0
    TotalLetterCase = 0
    TotalLetterAmt = 0
    
    'sql = " SELECT COUNT(CASNumber) AS NoofCase, SUM(CONVERT(Money,CASInvAmt)) AS TotalAmt, " & _
          " MONTH(CASDRC) AS Month From VIEW_ARReport WHERE CASDRC IS NOT NULL "
    
    'sql2 = " GROUP BY MONTH(CASDRC) ORDER BY MONTH(CASDRC) "
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + strAppend2 + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "2"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReviewRpt"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    Do
        Do Until rst2.EOF
                        
            With rst2
                
                If rst2!Month = "1" Then
                    Fgrid.col = 1
                    Fgrid.row = 2
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 2
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "2" Then
                    Fgrid.col = 1
                    Fgrid.row = 3
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 3
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "3" Then
                    Fgrid.col = 1
                    Fgrid.row = 4
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 4
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "4" Then
                    Fgrid.col = 1
                    Fgrid.row = 5
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 5
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "5" Then
                    Fgrid.col = 1
                    Fgrid.row = 6
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 6
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "6" Then
                    Fgrid.col = 1
                    Fgrid.row = 7
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 7
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "7" Then
                    Fgrid.col = 1
                    Fgrid.row = 8
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 8
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "8" Then
                    Fgrid.col = 1
                    Fgrid.row = 9
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 9
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "9" Then
                    Fgrid.col = 1
                    Fgrid.row = 10
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 10
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "10" Then
                    Fgrid.col = 1
                    Fgrid.row = 11
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "-" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 11
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "11" Then
                    Fgrid.col = 1
                    Fgrid.row = 12
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 12
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "12" Then
                    Fgrid.col = 1
                    Fgrid.row = 13
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 13
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                
                End If
            
            TotalReceivedCase = (CDbl(TotalReceivedCase) + IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0))
            TotalReceivedAmt = (CDbl(TotalReceivedAmt) + IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0))
            
            Fgrid.col = 1
            Fgrid.row = 15
            Fgrid = Format(IIf(IsNumeric(TotalReceivedCase) = True, TotalReceivedCase, 0))
            
            Fgrid.col = 2
            Fgrid.row = 15
            Fgrid = Format(IIf(IsNumeric(TotalReceivedAmt) = True, TotalReceivedAmt, 0), "#,##0.00")
                            
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

End Sub

Private Sub ProductivityRpt2(Optional strAppend As String)
Dim rst3 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
        
    TotalLetterCase = 0
    TotalLetterAmt = 0
    
    'sql3 = " SELECT COUNT(CASNumber) AS NoofCase, SUM(CONVERT(Money,CASInvAmt)) AS TotalAmt, " & _
           " MONTH(CASDRC) AS Month From VIEW_ARReport WHERE (CASLetterSent IS NOT NULL AND CASDRC IS NOT NULL)"
    
    'sql4 = " GROUP BY MONTH(CASDRC) ORDER BY MONTH(CASDRC) "
    
    'If Len(Trim(strAppend)) Then
    '    sql3 = sql3 + strAppend + sql4
    'Else
    '    sql3 = sql3 + sql4
    'End If
                
    'rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "3"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReviewRpt"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = cmd.Execute
    
    Do
        Do Until rst3.EOF
                   
            If rst3!Month = "1" Then
                    Fgrid.col = 3
                    Fgrid.row = 2
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 2
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "2" Then
                    Fgrid.col = 3
                    Fgrid.row = 3
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 3
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "3" Then
                    Fgrid.col = 3
                    Fgrid.row = 4
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 4
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "4" Then
                    Fgrid.col = 3
                    Fgrid.row = 5
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 5
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "5" Then
                    Fgrid.col = 3
                    Fgrid.row = 6
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 6
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "6" Then
                    Fgrid.col = 3
                    Fgrid.row = 7
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 7
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "7" Then
                    Fgrid.col = 3
                    Fgrid.row = 8
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 8
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "8" Then
                    Fgrid.col = 3
                    Fgrid.row = 9
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 9
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "9" Then
                    Fgrid.col = 3
                    Fgrid.row = 10
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 10
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "10" Then
                    Fgrid.col = 3
                    Fgrid.row = 11
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 11
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "11" Then
                    Fgrid.col = 3
                    Fgrid.row = 12
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 12
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst3!Month = "12" Then
                    Fgrid.col = 3
                    Fgrid.row = 13
                    Fgrid = IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 4
                    Fgrid.row = 13
                    Fgrid = Format(IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                
                End If
                
                TotalLetterCase = (CDbl(TotalLetterCase) + IIf(IsNumeric(rst3!NoofCase) = True, rst3!NoofCase, 0))
                TotalLetterAmt = (CDbl(TotalLetterAmt) + IIf(IsNumeric(rst3!TotalAmt) = True, rst3!TotalAmt, 0))
            
                Fgrid.col = 3
                Fgrid.row = 15
                Fgrid = Format(IIf(IsNumeric(TotalLetterCase) = True, TotalLetterCase, 0))
            
                Fgrid.col = 4
                Fgrid.row = 15
                Fgrid = Format(IIf(IsNumeric(TotalLetterAmt) = True, TotalLetterAmt, 0), "#,##0.00")
    
            rst3.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    
    rst3.Close
    Set rst3 = Nothing

End Sub

Private Sub ResponsesRpt(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String

    TotalReceivedCase = 0
    TotalReceivedAmt = 0
    TotalDiscountAmt = 0
    
    'sql = " SELECT COUNT(CASNumber) AS NoofCase, SUM(CONVERT(Money,CASInvAmt)) AS TotalAmt, " & _
          " SUM(CASDisRec) AS TotalDiscount, MONTH(CASHospReply) AS Month From VIEW_ARReport WHERE CASHospReply IS NOT NULL "
    
    'sql2 = " GROUP BY MONTH(CASHospReply) ORDER BY MONTH(CASHospReply) "
           
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend + sql2
    'Else
    '    sql = sql + sql2
    'End If
    
    'rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "4"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Case_AuditReviewRpt"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    Do
        Do Until rst2.EOF
                        
            With rst2
                
                If rst2!Month = "1" Then
                    Fgrid.col = 1
                    Fgrid.row = 2
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 2
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 2
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "2" Then
                    Fgrid.col = 1
                    Fgrid.row = 3
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 3
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 3
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "3" Then
                    Fgrid.col = 1
                    Fgrid.row = 4
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 4
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 4
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "4" Then
                    Fgrid.col = 1
                    Fgrid.row = 5
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 5
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 5
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "5" Then
                    Fgrid.col = 1
                    Fgrid.row = 6
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 6
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 6
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "6" Then
                    Fgrid.col = 1
                    Fgrid.row = 7
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 7
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 7
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "7" Then
                    Fgrid.col = 1
                    Fgrid.row = 8
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 8
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 8
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "8" Then
                    Fgrid.col = 1
                    Fgrid.row = 9
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 9
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 9
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "9" Then
                    Fgrid.col = 1
                    Fgrid.row = 10
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 10
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 10
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "10" Then
                    Fgrid.col = 1
                    Fgrid.row = 11
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 11
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 11
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "11" Then
                    Fgrid.col = 1
                    Fgrid.row = 12
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 12
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 12
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                ElseIf rst2!Month = "12" Then
                    Fgrid.col = 1
                    Fgrid.row = 13
                    Fgrid = IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0)
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 2
                    Fgrid.row = 13
                    Fgrid = Format(IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                    
                    Fgrid.col = 3
                    Fgrid.row = 13
                    Fgrid = Format(IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0), "#,##0.00")
                    If Fgrid = "" Then
                        Fgrid = 0
                    End If
                
                End If
            
            TotalReceivedCase = (CDbl(TotalReceivedCase) + IIf(IsNumeric(!NoofCase) = True, !NoofCase, 0))
            TotalDiscountAmt = (CDbl(TotalDiscountAmt) + IIf(IsNumeric(!TotalDiscount) = True, !TotalDiscount, 0))
            TotalReceivedAmt = (CDbl(TotalReceivedAmt) + IIf(IsNumeric(!TotalAmt) = True, !TotalAmt, 0))
                        
            Fgrid.col = 1
            Fgrid.row = 15
            Fgrid = Format(IIf(IsNumeric(TotalReceivedCase) = True, TotalReceivedCase, 0))
            
            Fgrid.col = 2
            Fgrid.row = 15
            Fgrid = Format(IIf(IsNumeric(TotalDiscountAmt) = True, TotalDiscountAmt, 0), "#,##0.00")
            If Fgrid = "" Then
                Fgrid = 0
            End If
            
            Fgrid.col = 3
            Fgrid.row = 15
            Fgrid = Format(IIf(IsNumeric(TotalReceivedAmt) = True, TotalReceivedAmt, 0), "#,##0.00")
                            
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

End Sub

Sub DoInitialSettings()
          
    Fgrid.row = 0
    Fgrid.col = 0
    Fgrid = "Month"
    Fgrid.col = 1
    Fgrid = "No. of files"
    Fgrid.col = 2
    Fgrid = "RM"
    Fgrid.col = 3
    Fgrid = "No. of files"
    Fgrid.col = 4
    Fgrid = "RM"
    Fgrid.col = 5
    Fgrid = "Avg Claim per file (RM)"
    Fgrid.col = 6
    Fgrid = "No. of files"
    Fgrid.col = 7
    Fgrid = "RM"
    
    Fgrid.col = 0
    
    Fgrid.ColAlignment(0) = 1
    Fgrid.ColAlignment(1) = 6
    Fgrid.ColAlignment(2) = 6
    Fgrid.ColAlignment(3) = 6
    Fgrid.ColAlignment(4) = 6
    Fgrid.ColAlignment(5) = 6
    Fgrid.ColAlignment(6) = 6
    Fgrid.ColAlignment(7) = 6
    Fgrid.ColWidth(0) = 2000
    Fgrid.ColWidth(1) = 1200
    Fgrid.ColWidth(2) = 1200
    Fgrid.ColWidth(3) = 1200
    Fgrid.ColWidth(4) = 1200
    Fgrid.ColWidth(5) = 2000
    Fgrid.ColWidth(6) = 1200
    Fgrid.ColWidth(7) = 1200
                
    Fgrid.row = 2
    Fgrid = "January"
    
    Fgrid.row = 3
    Fgrid = "February"
    
    Fgrid.row = 4
    Fgrid = "March"
    
    Fgrid.row = 5
    Fgrid = "April"
    
    Fgrid.row = 6
    Fgrid = "May"
    
    Fgrid.row = 7
    Fgrid = "June"
    
    Fgrid.row = 8
    Fgrid = "July"
    
    Fgrid.row = 9
    Fgrid = "August"
    
    Fgrid.row = 10
    Fgrid = "September"
    
    Fgrid.row = 11
    Fgrid = "October"
    
    Fgrid.row = 12
    Fgrid = "November"
    
    Fgrid.row = 13
    Fgrid = "December"
    
    Fgrid.row = 14
    Fgrid = ""
            
    Fgrid.row = 15
    Fgrid.CellFontBold = True
    Fgrid.CellFontSize = 8
    Fgrid = "Total"
 
End Sub

Sub DoInitialSettings2()
          
    Fgrid.row = 0
    Fgrid.col = 0
    Fgrid = "Month"
    Fgrid.col = 1
    Fgrid = "No of responses received"
    Fgrid.col = 2
    Fgrid = "Discounts (RM)"
    Fgrid.col = 3
    Fgrid = "Bill Amt (RM)"
    Fgrid.col = 4
    Fgrid = "Discount Coverage"
    Fgrid.col = 5
    Fgrid = ""
    Fgrid.col = 6
    Fgrid = ""
    Fgrid.col = 7
    Fgrid = ""
    
    Fgrid.col = 0
    
    Fgrid.ColAlignment(0) = 1
    Fgrid.ColAlignment(1) = 6
    Fgrid.ColAlignment(2) = 6
    Fgrid.ColAlignment(3) = 6
    Fgrid.ColAlignment(4) = 6
    Fgrid.ColAlignment(5) = 6
    Fgrid.ColAlignment(6) = 6
    Fgrid.ColAlignment(7) = 6
    Fgrid.ColWidth(0) = 2000
    Fgrid.ColWidth(1) = 2500
    Fgrid.ColWidth(2) = 2500
    Fgrid.ColWidth(3) = 2500
    Fgrid.ColWidth(4) = 2500
    Fgrid.ColWidth(5) = 2000
    Fgrid.ColWidth(6) = 1200
    Fgrid.ColWidth(7) = 1200
                
    Fgrid.row = 2
    Fgrid = "January"
    
    Fgrid.row = 3
    Fgrid = "February"
    
    Fgrid.row = 4
    Fgrid = "March"
    
    Fgrid.row = 5
    Fgrid = "April"
    
    Fgrid.row = 6
    Fgrid = "May"
    
    Fgrid.row = 7
    Fgrid = "June"
    
    Fgrid.row = 8
    Fgrid = "July"
    
    Fgrid.row = 9
    Fgrid = "August"
    
    Fgrid.row = 10
    Fgrid = "September"
    
    Fgrid.row = 11
    Fgrid = "October"
    
    Fgrid.row = 12
    Fgrid = "November"
    
    Fgrid.row = 13
    Fgrid = "December"
    
    Fgrid.row = 14
    Fgrid = ""
            
    Fgrid.row = 15
    Fgrid.CellFontBold = True
    Fgrid.CellFontSize = 8
    Fgrid = "Total"
 
End Sub

Private Sub FileCalculation(Optional strAppend As String)
Dim JanNo, JanAmt, FebNo, FebAmt As String
Dim MarNo, MarAmt, AprNo, AprAmt As String
Dim MayNo, MayAmt, JunNo, JunAmt As String
Dim JulNo, JulAmt, AugNo, AugAmt As String
Dim SepNo, SepAmt, OctNo, OctAmt As String
Dim NovNo, NovAmt, DecNo, DecAmt As String
Dim TotalNo, TotalAmt As String
Dim JanNo2, JanAmt2, FebNo2, FebAmt2 As String
Dim MarNo2, MarAmt2, AprNo2, AprAmt2 As String
Dim MayNo2, MayAmt2, JunNo2, JunAmt2 As String
Dim JulNo2, JulAmt2, AugNo2, AugAmt2 As String
Dim SepNo2, SepAmt2, OctNo2, OctAmt2 As String
Dim NovNo2, NovAmt2, DecNo2, DecAmt2 As String
Dim TotalNo2, TotalAmt2 As String
Dim JanNo3, JanAmt3, FebNo3, FebAmt3 As String
Dim MarNo3, MarAmt3, AprNo3, AprAmt3 As String
Dim MayNo3, MayAmt3, JunNo3, JunAmt3 As String
Dim JulNo3, JulAmt3, AugNo3, AugAmt3 As String
Dim SepNo3, SepAmt3, OctNo3, OctAmt3 As String
Dim NovNo3, NovAmt3, DecNo3, DecAmt3 As String
Dim TotalNo3, TotalAmt3 As String

    JanNo = "0"
    JanAmt = "0"
    FebNo = "0"
    FebAmt = "0"
    MarNo = "0"
    MarAmt = "0"
    AprNo = "0"
    AprAmt = "0"
    MayNo = "0"
    MayAmt = "0"
    JunNo = "0"
    JunAmt = "0"
    JulNo = "0"
    JulAmt = "0"
    AugNo = "0"
    AugAmt = "0"
    SepNo = "0"
    SepAmt = "0"
    OctNo = "0"
    OctAmt = "0"
    NovNo = "0"
    NovAmt = "0"
    DecNo = "0"
    DecAmt = "0"
    TotalNo = "0"
    TotalAmt = "0"
    
    JanNo2 = "0"
    JanAmt2 = "0"
    FebNo2 = "0"
    FebAmt2 = "0"
    MarNo2 = "0"
    MarAmt2 = "0"
    AprNo2 = "0"
    AprAmt2 = "0"
    MayNo2 = "0"
    MayAmt2 = "0"
    JunNo2 = "0"
    JunAmt2 = "0"
    JulNo2 = "0"
    JulAmt2 = "0"
    AugNo2 = "0"
    AugAmt2 = "0"
    SepNo2 = "0"
    SepAmt2 = "0"
    OctNo2 = "0"
    OctAmt2 = "0"
    NovNo2 = "0"
    NovAmt2 = "0"
    DecNo2 = "0"
    DecAmt2 = "0"
    TotalNo2 = "0"
    TotalAmt2 = "0"
    
    JanNo3 = "0"
    JanAmt3 = "0"
    FebNo3 = "0"
    FebAmt3 = "0"
    MarNo3 = "0"
    MarAmt3 = "0"
    AprNo3 = "0"
    AprAmt3 = "0"
    MayNo3 = "0"
    MayAmt3 = "0"
    JunNo3 = "0"
    JunAmt3 = "0"
    JulNo3 = "0"
    JulAmt3 = "0"
    AugNo3 = "0"
    AugAmt3 = "0"
    SepNo3 = "0"
    SepAmt3 = "0"
    OctNo3 = "0"
    OctAmt3 = "0"
    NovNo3 = "0"
    NovAmt3 = "0"
    DecNo3 = "0"
    DecAmt3 = "0"
    TotalNo3 = "0"
    TotalAmt3 = "0"
    
    '===================== Calculation 1 ==================================
    'Jan
    Fgrid.row = 2
    Fgrid.col = 4
    JanNo = Fgrid
    Fgrid.col = 3
    JanAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(JanNo) = False Then
        JanNo = 0
    End If
    
    If IsNumeric(JanAmt) = False Then
        JanAmt = 0
    End If
    If JanNo <> "0" And JanAmt <> "0" Then
        Fgrid = (JanNo / JanAmt)
    Else
        Fgrid = 0
    End If
    
    'Feb
    Fgrid.row = 3
    Fgrid.col = 4
    FebNo = Fgrid
    Fgrid.col = 3
    FebAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(FebNo) = False Then
        FebNo = 0
    End If
    
    If IsNumeric(FebAmt) = False Then
        FebAmt = 0
    End If
    If FebNo <> "0" And FebAmt <> "0" Then
        Fgrid = (FebNo / FebAmt)
    Else
        Fgrid = 0
    End If
    
    'Mar
    Fgrid.row = 4
    Fgrid.col = 4
    MarNo = Fgrid
    Fgrid.col = 3
    MarAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(MarNo) = False Then
        MarNo = 0
    End If
    
    If IsNumeric(MarAmt) = False Then
        MarAmt = 0
    End If
    If MarNo <> "0" And MarAmt <> "0" Then
        Fgrid = (MarNo / MarAmt)
    Else
        Fgrid = 0
    End If
    
    'Apr
    Fgrid.row = 5
    Fgrid.col = 4
    AprNo = Fgrid
    Fgrid.col = 3
    AprAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(AprNo) = False Then
        AprNo = 0
    End If
    
    If IsNumeric(AprAmt) = False Then
        AprAmt = 0
    End If
    If AprNo <> "0" And AprAmt <> "0" Then
        Fgrid = (AprNo / AprAmt)
    Else
        Fgrid = 0
    End If
    
    'May
    Fgrid.row = 6
    Fgrid.col = 4
    MayNo = Fgrid
    Fgrid.col = 3
    MayAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(MayNo) = False Then
        MayNo = 0
    End If
    
    If IsNumeric(MayAmt) = False Then
        MayAmt = 0
    End If
    If MayNo <> "0" And MayAmt <> "0" Then
        Fgrid = (MayNo / MayAmt)
    Else
        Fgrid = 0
    End If
    
    'Jun
    Fgrid.row = 7
    Fgrid.col = 4
    JunNo = Fgrid
    Fgrid.col = 3
    JunAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(JunNo) = False Then
        JunNo = 0
    End If
    
    If IsNumeric(JunAmt) = False Then
        JunAmt = 0
    End If
    If JunNo <> "0" And JunAmt <> "0" Then
        Fgrid = (JunNo / JunAmt)
    Else
        Fgrid = 0
    End If
    
    'Jul
    Fgrid.row = 8
    Fgrid.col = 4
    JulNo = Fgrid
    Fgrid.col = 3
    JulAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(JulNo) = False Then
        JulNo = 0
    End If
    
    If IsNumeric(JulAmt) = False Then
        JulAmt = 0
    End If
    If JulNo <> "0" And JulAmt <> "0" Then
        Fgrid = (JulNo / JulAmt)
    Else
        Fgrid = 0
    End If
    
    'Aug
    Fgrid.row = 9
    Fgrid.col = 4
    AugNo = Fgrid
    Fgrid.col = 3
    AugAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(AugNo) = False Then
        AugNo = 0
    End If
    
    If IsNumeric(AugAmt) = False Then
        AugAmt = 0
    End If
    If AugNo <> "0" And AugAmt <> "0" Then
        Fgrid = (AugNo / AugAmt)
    Else
        Fgrid = 0
    End If
    
    'Sept
    Fgrid.row = 10
    Fgrid.col = 4
    SepNo = Fgrid
    Fgrid.col = 3
    SepAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(SepNo) = False Then
        SepNo = 0
    End If
    
    If IsNumeric(SepAmt) = False Then
        SepAmt = 0
    End If
    If SepNo <> "0" And SepAmt <> "0" Then
        Fgrid = (SepNo / SepAmt)
    Else
        Fgrid = 0
    End If
    
    'Oct
    Fgrid.row = 11
    Fgrid.col = 4
    OctNo = Fgrid
    Fgrid.col = 3
    OctAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(OctNo) = False Then
        OctNo = 0
    End If
    
    If IsNumeric(OctAmt) = False Then
        OctAmt = 0
    End If
    If OctNo <> "0" And OctAmt <> "0" Then
        Fgrid = (OctNo / OctAmt)
    Else
        Fgrid = 0
    End If
    
    'Nov
    Fgrid.row = 12
    Fgrid.col = 4
    NovNo = Fgrid
    Fgrid.col = 3
    NovAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(NovNo) = False Then
        NovNo = 0
    End If
    
    If IsNumeric(NovAmt) = False Then
        NovAmt = 0
    End If
    If NovNo <> "0" And NovAmt <> "0" Then
        Fgrid = (NovNo / NovAmt)
    Else
        Fgrid = 0
    End If
    
    'Dec
    Fgrid.row = 13
    Fgrid.col = 4
    DecNo = Fgrid
    Fgrid.col = 3
    DecAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(DecNo) = False Then
        DecNo = 0
    End If
    
    If IsNumeric(DecAmt) = False Then
        DecAmt = 0
    End If
    If DecNo <> "0" And DecAmt <> "0" Then
        Fgrid = (DecNo / DecAmt)
    Else
        Fgrid = 0
    End If
    
    'Total 1
    Fgrid.row = 15
    Fgrid.col = 4
    TotalNo = Fgrid
    Fgrid.col = 3
    TotalAmt = Fgrid

    Fgrid.col = 5
    If IsNumeric(TotalNo) = False Then
        TotalNo = 0
    End If
    
    If IsNumeric(TotalAmt) = False Then
        TotalAmt = 0
    End If
    If TotalNo <> "0" And TotalAmt <> "0" Then
        Fgrid = (TotalNo / TotalAmt)
    Else
        Fgrid = 0
    End If
    '===================== Calculation 1 ==================================
    
    '===================== Calculation 2 ==================================
    'Jan
    Fgrid.row = 2
    Fgrid.col = 3
    JanNo2 = Fgrid
    Fgrid.col = 1
    JanAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(JanNo2) = False Then
        JanNo2 = 0
    End If
    
    If IsNumeric(JanAmt2) = False Then
        JanAmt2 = 0
    End If
    If JanNo2 <> "0" And JanAmt2 <> "0" Then
        Fgrid = Round(JanNo2 / JanAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Feb
    Fgrid.row = 3
    Fgrid.col = 3
    FebNo2 = Fgrid
    Fgrid.col = 1
    FebAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(FebNo2) = False Then
        FebNo2 = 0
    End If
    
    If IsNumeric(FebAmt2) = False Then
        FebAmt2 = 0
    End If
    If FebNo2 <> "0" And FebAmt2 <> "0" Then
        Fgrid = Round(FebNo2 / FebAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Mar
    Fgrid.row = 4
    Fgrid.col = 3
    MarNo2 = Fgrid
    Fgrid.col = 1
    MarAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(MarNo2) = False Then
        MarNo2 = 0
    End If
    
    If IsNumeric(MarAmt2) = False Then
        MarAmt2 = 0
    End If
    If MarNo2 <> "0" And MarAmt2 <> "0" Then
        Fgrid = Round(MarNo2 / MarAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Apr
    Fgrid.row = 5
    Fgrid.col = 3
    AprNo2 = Fgrid
    Fgrid.col = 1
    AprAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(AprNo2) = False Then
        AprNo2 = 0
    End If
    
    If IsNumeric(AprAmt2) = False Then
        AprAmt2 = 0
    End If
    If AprNo2 <> "0" And AprAmt2 <> "0" Then
        Fgrid = Round(AprNo2 / AprAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'May
    Fgrid.row = 6
    Fgrid.col = 3
    MayNo2 = Fgrid
    Fgrid.col = 1
    MayAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(MayNo2) = False Then
        MayNo2 = 0
    End If
    
    If IsNumeric(MayAmt2) = False Then
        MayAmt2 = 0
    End If
    If MayNo2 <> "0" And MayAmt2 <> "0" Then
        Fgrid = Round(MayNo2 / MayAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Jun
    Fgrid.row = 7
    Fgrid.col = 3
    JunNo2 = Fgrid
    Fgrid.col = 1
    JunAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(JunNo2) = False Then
        JunNo2 = 0
    End If
    
    If IsNumeric(JunAmt2) = False Then
        JunAmt2 = 0
    End If
    If JunNo2 <> "0" And JunAmt2 <> "0" Then
        Fgrid = Round(JunNo2 / JunAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Jul
    Fgrid.row = 8
    Fgrid.col = 3
    JulNo2 = Fgrid
    Fgrid.col = 1
    JulAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(JulNo2) = False Then
        JulNo2 = 0
    End If
    
    If IsNumeric(JulAmt2) = False Then
        JulAmt2 = 0
    End If
    If JulNo2 <> "0" And JulAmt2 <> "0" Then
        Fgrid = Round(JulNo2 / JulAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Aug
    Fgrid.row = 9
    Fgrid.col = 3
    AugNo2 = Fgrid
    Fgrid.col = 1
    AugAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(AugNo2) = False Then
        AugNo2 = 0
    End If
    
    If IsNumeric(AugAmt2) = False Then
        AugAmt2 = 0
    End If
    If AugNo2 <> "0" And AugAmt2 <> "0" Then
        Fgrid = Round(AugNo2 / AugAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Sept
    Fgrid.row = 10
    Fgrid.col = 3
    SepNo2 = Fgrid
    Fgrid.col = 1
    SepAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(SepNo2) = False Then
        SepNo2 = 0
    End If
    
    If IsNumeric(SepAmt2) = False Then
        SepAmt2 = 0
    End If
    If SepNo2 <> "0" And SepAmt2 <> "0" Then
        Fgrid = Round(SepNo2 / SepAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Oct
    Fgrid.row = 11
    Fgrid.col = 3
    OctNo2 = Fgrid
    Fgrid.col = 1
    OctAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(OctNo2) = False Then
        OctNo2 = 0
    End If
    
    If IsNumeric(OctAmt2) = False Then
        OctAmt2 = 0
    End If
    If OctNo2 <> "0" And OctAmt2 <> "0" Then
        Fgrid = Round(OctNo2 / OctAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Nov
    Fgrid.row = 12
    Fgrid.col = 3
    NovNo2 = Fgrid
    Fgrid.col = 1
    NovAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(NovNo2) = False Then
        NovNo2 = 0
    End If
    
    If IsNumeric(NovAmt2) = False Then
        NovAmt2 = 0
    End If
    If NovNo2 <> "0" And NovAmt2 <> "0" Then
        Fgrid = Round(NovNo2 / NovAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Dec
    Fgrid.row = 13
    Fgrid.col = 3
    DecNo2 = Fgrid
    Fgrid.col = 1
    DecAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(DecNo2) = False Then
        DecNo2 = 0
    End If
    
    If IsNumeric(DecAmt2) = False Then
        DecAmt2 = 0
    End If
    If DecNo2 <> "0" And DecAmt2 <> "0" Then
        Fgrid = Round(DecNo2 / DecAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Total 1
    Fgrid.row = 15
    Fgrid.col = 3
    TotalNo2 = Fgrid
    Fgrid.col = 1
    TotalAmt2 = Fgrid

    Fgrid.col = 6
    If IsNumeric(TotalNo2) = False Then
        TotalNo2 = 0
    End If
    
    If IsNumeric(TotalAmt2) = False Then
        TotalAmt2 = 0
    End If
    If TotalNo2 <> "0" And TotalAmt2 <> "0" Then
        Fgrid = Round(TotalNo2 / TotalAmt2 * 100, 0)
    Else
        Fgrid = 0
    End If
    '===================== Calculation 2 ==================================
    
    '===================== Calculation 2 ==================================
    'Jan
    Fgrid.row = 2
    Fgrid.col = 4
    JanNo3 = Fgrid
    Fgrid.col = 2
    JanAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(JanNo3) = False Then
        JanNo3 = 0
    End If
    
    If IsNumeric(JanAmt3) = False Then
        JanAmt3 = 0
    End If
    If JanNo3 <> "0" And JanAmt3 <> "0" Then
        Fgrid = Round(JanNo3 / JanAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Feb
    Fgrid.row = 3
    Fgrid.col = 4
    FebNo3 = Fgrid
    Fgrid.col = 2
    FebAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(FebNo3) = False Then
        FebNo3 = 0
    End If
    
    If IsNumeric(FebAmt3) = False Then
        FebAmt3 = 0
    End If
    If FebNo3 <> "0" And FebAmt3 <> "0" Then
        Fgrid = Round(FebNo3 / FebAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Mar
    Fgrid.row = 4
    Fgrid.col = 4
    MarNo3 = Fgrid
    Fgrid.col = 2
    MarAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(MarNo3) = False Then
        MarNo3 = 0
    End If
    
    If IsNumeric(MarAmt3) = False Then
        MarAmt3 = 0
    End If
    If MarNo3 <> "0" And MarAmt3 <> "0" Then
        Fgrid = Round(MarNo3 / MarAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Apr
    Fgrid.row = 5
    Fgrid.col = 4
    AprNo3 = Fgrid
    Fgrid.col = 2
    AprAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(AprNo3) = False Then
        AprNo3 = 0
    End If
    
    If IsNumeric(AprAmt3) = False Then
        AprAmt3 = 0
    End If
    If AprNo3 <> "0" And AprAmt3 <> "0" Then
        Fgrid = Round(AprNo3 / AprAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'May
    Fgrid.row = 6
    Fgrid.col = 4
    MayNo3 = Fgrid
    Fgrid.col = 2
    MayAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(MayNo3) = False Then
        MayNo3 = 0
    End If
    
    If IsNumeric(MayAmt3) = False Then
        MayAmt3 = 0
    End If
    If MayNo3 <> "0" And MayAmt3 <> "0" Then
        Fgrid = Round(MayNo3 / MayAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Jun
    Fgrid.row = 7
    Fgrid.col = 4
    JunNo3 = Fgrid
    Fgrid.col = 2
    JunAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(JunNo3) = False Then
        JunNo3 = 0
    End If
    
    If IsNumeric(JunAmt3) = False Then
        JunAmt3 = 0
    End If
    If JunNo3 <> "0" And JunAmt3 <> "0" Then
        Fgrid = Round(JunNo3 / JunAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Jul
    Fgrid.row = 8
    Fgrid.col = 4
    JulNo3 = Fgrid
    Fgrid.col = 2
    JulAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(JulNo3) = False Then
        JulNo3 = 0
    End If
    
    If IsNumeric(JulAmt3) = False Then
        JulAmt3 = 0
    End If
    If JulNo3 <> "0" And JulAmt3 <> "0" Then
        Fgrid = Round(JulNo3 / JulAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Aug
    Fgrid.row = 9
    Fgrid.col = 4
    AugNo3 = Fgrid
    Fgrid.col = 2
    AugAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(AugNo3) = False Then
        AugNo3 = 0
    End If
    
    If IsNumeric(AugAmt3) = False Then
        AugAmt3 = 0
    End If
    If AugNo3 <> "0" And AugAmt3 <> "0" Then
        Fgrid = Round(AugNo3 / AugAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Sept
    Fgrid.row = 10
    Fgrid.col = 4
    SepNo3 = Fgrid
    Fgrid.col = 2
    SepAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(SepNo3) = False Then
        SepNo3 = 0
    End If
    
    If IsNumeric(SepAmt3) = False Then
        SepAmt3 = 0
    End If
    If SepNo3 <> "0" And SepAmt3 <> "0" Then
        Fgrid = Round(SepNo3 / SepAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Oct
    Fgrid.row = 11
    Fgrid.col = 4
    OctNo3 = Fgrid
    Fgrid.col = 2
    OctAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(OctNo3) = False Then
        OctNo3 = 0
    End If
    
    If IsNumeric(OctAmt3) = False Then
        OctAmt3 = 0
    End If
    If OctNo3 <> "0" And OctAmt3 <> "0" Then
        Fgrid = Round(OctNo3 / OctAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Nov
    Fgrid.row = 12
    Fgrid.col = 4
    NovNo3 = Fgrid
    Fgrid.col = 2
    NovAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(NovNo3) = False Then
        NovNo3 = 0
    End If
    
    If IsNumeric(NovAmt3) = False Then
        NovAmt3 = 0
    End If
    If NovNo3 <> "0" And NovAmt3 <> "0" Then
        Fgrid = Round(NovNo3 / NovAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Dec
    Fgrid.row = 13
    Fgrid.col = 4
    DecNo3 = Fgrid
    Fgrid.col = 2
    DecAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(DecNo3) = False Then
        DecNo3 = 0
    End If
    
    If IsNumeric(DecAmt3) = False Then
        DecAmt3 = 0
    End If
    If DecNo3 <> "0" And DecAmt3 <> "0" Then
        Fgrid = Round(DecNo3 / DecAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Total 1
    Fgrid.row = 15
    Fgrid.col = 4
    TotalNo3 = Fgrid
    Fgrid.col = 2
    TotalAmt3 = Fgrid

    Fgrid.col = 7
    If IsNumeric(TotalNo3) = False Then
        TotalNo3 = 0
    End If
    
    If IsNumeric(TotalAmt3) = False Then
        TotalAmt3 = 0
    End If
    If TotalNo3 <> "0" And TotalAmt3 <> "0" Then
        Fgrid = Round(TotalNo3 / TotalAmt3 * 100, 0)
    Else
        Fgrid = 0
    End If
    '===================== Calculation 3 ==================================
                    
End Sub

Private Sub FileCalculation2(Optional strAppend As String)
Dim JanNo, JanAmt, FebNo, FebAmt As String
Dim MarNo, MarAmt, AprNo, AprAmt As String
Dim MayNo, MayAmt, JunNo, JunAmt As String
Dim JulNo, JulAmt, AugNo, AugAmt As String
Dim SepNo, SepAmt, OctNo, OctAmt As String
Dim NovNo, NovAmt, DecNo, DecAmt As String
Dim TotalNo, TotalAmt As String

    JanNo = "0"
    JanAmt = "0"
    FebNo = "0"
    FebAmt = "0"
    MarNo = "0"
    MarAmt = "0"
    AprNo = "0"
    AprAmt = "0"
    MayNo = "0"
    MayAmt = "0"
    JunNo = "0"
    JunAmt = "0"
    JulNo = "0"
    JulAmt = "0"
    AugNo = "0"
    AugAmt = "0"
    SepNo = "0"
    SepAmt = "0"
    OctNo = "0"
    OctAmt = "0"
    NovNo = "0"
    NovAmt = "0"
    DecNo = "0"
    DecAmt = "0"
    TotalNo = "0"
    TotalAmt = "0"
            
    '===================== Calculation 1 ==================================
    'Jan
    Fgrid.row = 2
    Fgrid.col = 2
    JanNo = Fgrid
    Fgrid.col = 3
    JanAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(JanNo) = False Then
        JanNo = 0
    End If
    
    If IsNumeric(JanAmt) = False Then
        JanAmt = 0
    End If
    If JanNo <> "0" And JanAmt <> "0" Then
        Fgrid = Round(JanNo / JanAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Feb
    Fgrid.row = 3
    Fgrid.col = 2
    FebNo = Fgrid
    Fgrid.col = 3
    FebAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(FebNo) = False Then
        FebNo = 0
    End If
    
    If IsNumeric(FebAmt) = False Then
        FebAmt = 0
    End If
    If FebNo <> "0" And FebAmt <> "0" Then
        Fgrid = Round(FebNo / FebAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Mar
    Fgrid.row = 4
    Fgrid.col = 2
    MarNo = Fgrid
    Fgrid.col = 3
    MarAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(MarNo) = False Then
        MarNo = 0
    End If
    
    If IsNumeric(MarAmt) = False Then
        MarAmt = 0
    End If
    If MarNo <> "0" And MarAmt <> "0" Then
        Fgrid = Round(MarNo / MarAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Apr
    Fgrid.row = 5
    Fgrid.col = 2
    AprNo = Fgrid
    Fgrid.col = 3
    AprAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(AprNo) = False Then
        AprNo = 0
    End If
    
    If IsNumeric(AprAmt) = False Then
        AprAmt = 0
    End If
    If AprNo <> "0" And AprAmt <> "0" Then
        Fgrid = Round(AprNo / AprAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'May
    Fgrid.row = 6
    Fgrid.col = 2
    MayNo = Fgrid
    Fgrid.col = 4
    MayAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(MayNo) = False Then
        MayNo = 0
    End If
    
    If IsNumeric(MayAmt) = False Then
        MayAmt = 0
    End If
    If MayNo <> "0" And MayAmt <> "0" Then
        Fgrid = Round(MayNo / MayAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Jun
    Fgrid.row = 7
    Fgrid.col = 2
    JunNo = Fgrid
    Fgrid.col = 3
    JunAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(JunNo) = False Then
        JunNo = 0
    End If
    
    If IsNumeric(JunAmt) = False Then
        JunAmt = 0
    End If
    If JunNo <> "0" And JunAmt <> "0" Then
        Fgrid = Round(JunNo / JunAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Jul
    Fgrid.row = 8
    Fgrid.col = 2
    JulNo = Fgrid
    Fgrid.col = 3
    JulAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(JulNo) = False Then
        JulNo = 0
    End If
    
    If IsNumeric(JulAmt) = False Then
        JulAmt = 0
    End If
    If JulNo <> "0" And JulAmt <> "0" Then
        Fgrid = Round(JulNo / JulAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Aug
    Fgrid.row = 9
    Fgrid.col = 2
    AugNo = Fgrid
    Fgrid.col = 3
    AugAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(AugNo) = False Then
        AugNo = 0
    End If
    
    If IsNumeric(AugAmt) = False Then
        AugAmt = 0
    End If
    If AugNo <> "0" And AugAmt <> "0" Then
        Fgrid = Round(AugNo / AugAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Sept
    Fgrid.row = 10
    Fgrid.col = 2
    SepNo = Fgrid
    Fgrid.col = 3
    SepAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(SepNo) = False Then
        SepNo = 0
    End If
    
    If IsNumeric(SepAmt) = False Then
        SepAmt = 0
    End If
    If SepNo <> "0" And SepAmt <> "0" Then
        Fgrid = Round(SepNo / SepAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Oct
    Fgrid.row = 11
    Fgrid.col = 2
    OctNo = Fgrid
    Fgrid.col = 3
    OctAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(OctNo) = False Then
        OctNo = 0
    End If
    
    If IsNumeric(OctAmt) = False Then
        OctAmt = 0
    End If
    If OctNo <> "0" And OctAmt <> "0" Then
        Fgrid = Round(OctNo / OctAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Nov
    Fgrid.row = 12
    Fgrid.col = 2
    NovNo = Fgrid
    Fgrid.col = 3
    NovAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(NovNo) = False Then
        NovNo = 0
    End If
    
    If IsNumeric(NovAmt) = False Then
        NovAmt = 0
    End If
    If NovNo <> "0" And NovAmt <> "0" Then
        Fgrid = Round(NovNo / NovAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Dec
    Fgrid.row = 13
    Fgrid.col = 2
    DecNo = Fgrid
    Fgrid.col = 3
    DecAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(DecNo) = False Then
        DecNo = 0
    End If
    
    If IsNumeric(DecAmt) = False Then
        DecAmt = 0
    End If
    If DecNo <> "0" And DecAmt <> "0" Then
        Fgrid = Round(DecNo / DecAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    
    'Total 1
    Fgrid.row = 15
    Fgrid.col = 2
    TotalNo = Fgrid
    Fgrid.col = 3
    TotalAmt = Fgrid

    Fgrid.col = 4
    If IsNumeric(TotalNo) = False Then
        TotalNo = 0
    End If
    
    If IsNumeric(TotalAmt) = False Then
        TotalAmt = 0
    End If
    If TotalNo <> "0" And TotalAmt <> "0" Then
        Fgrid = Round(TotalNo / TotalAmt * 100, 0)
    Else
        Fgrid = 0
    End If
    '===================== Calculation 1 ==================================
                    
End Sub

Public Function ExportToExcelWorksheet()
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET

    Screen.MousePointer = vbHourglass
    
    'Set objExcel = New excel.Application
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized

    'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\CaseTracking\ProductivityRpt.xlt")
    Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Tracking\ProductivityRpt.xlt")
    Set objWorksheet = objWorkbook.Sheets(1)
    'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
        
    'RECEIVED FROM CLAIMS - NO OF FILES
    Fgrid.col = 1
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "B") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "B") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "B") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "B") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "B") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "B") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "B") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "B") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "B") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "B") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "B") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "B") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "B") = Fgrid
    End If
    
    'RECEIVED FROM CLAIM - RM
    Fgrid.col = 2
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "C") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "C") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "C") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "C") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "C") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "C") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "C") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "C") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "C") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "C") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "C") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "C") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "C") = Fgrid
    End If
    
    'AR LETTER SENT - NO OF FILES
    Fgrid.col = 3
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "D") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "D") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "D") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "D") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "D") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "D") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "D") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "D") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "D") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "D") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "D") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "D") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "D") = Fgrid
    End If
    
    'AR LETTER SENT - RM
    Fgrid.col = 4
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "E") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "E") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "E") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "E") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "E") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "E") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "E") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "E") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "E") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "E") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "E") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "E") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "E") = Fgrid
    End If
    
    'AR COVERAGE - AVG CLAIM PER FILE (RM)
    Fgrid.col = 5
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "G") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "G") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "G") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "G") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "G") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "G") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "G") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "G") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "G") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "G") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "G") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "G") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "G") = Fgrid
    End If
    
    'AR COVERAGE - NO OF FILES
    Fgrid.col = 6
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "H") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "H") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "H") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "H") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "H") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "H") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "H") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "H") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "H") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "H") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "H") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "H") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "H") = Fgrid
    End If
    
    'AR COVERAGE - RM
    Fgrid.col = 7
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "I") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "I") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "I") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "I") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "I") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "I") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "I") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "I") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "I") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "I") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "I") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "I") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "I") = Fgrid
    End If
    
    
        
    Screen.MousePointer = vbDefault
    
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    
    Exit Function
ExitFunction:
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Screen.MousePointer = vbDefault
    objExcel.Quit
    Exit Function
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    GoTo ExitFunction
    
End Function

Public Function ExportToExcelWorksheet2()
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET

    Screen.MousePointer = vbHourglass
    
    'Set objExcel = New excel.Application
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized

    'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\CaseTracking\ResponsesRpt.xlt")
    Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Tracking\ResponsesRpt.xlt")
    Set objWorksheet = objWorkbook.Sheets(1)
    'Turn off the alerts, otherwise user will have to confirm my actions.
    objExcel.DisplayAlerts = False
    
    Set objWorksheet = objWorkbook.Worksheets.Item(1)
    'No of Responses Received
    Fgrid.col = 1
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "B") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "B") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "B") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "B") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "B") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "B") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "B") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "B") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "B") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "B") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "B") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "B") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "B") = Fgrid
    End If
    
    'Discount - RM
    Fgrid.col = 2
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "C") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "C") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "C") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "C") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "C") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "C") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "C") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "C") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "C") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "C") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "C") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "C") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "C") = Fgrid
    End If
    
    'Bill Amt - RM
    Fgrid.col = 3
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "D") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "D") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "D") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "D") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "D") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "D") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "D") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "D") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "D") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "D") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "D") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "D") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "D") = Fgrid
    End If
    
    'Discount Coverage
    Fgrid.col = 4
    Fgrid.row = 2
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(7, "E") = Fgrid
    End If
    Fgrid.row = 3
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(8, "E") = Fgrid
    End If
    Fgrid.row = 4
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(9, "E") = Fgrid
    End If
    Fgrid.row = 5
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(10, "E") = Fgrid
    End If
    Fgrid.row = 6
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(11, "E") = Fgrid
    End If
    Fgrid.row = 7
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(12, "E") = Fgrid
    End If
    Fgrid.row = 8
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(13, "E") = Fgrid
    End If
    Fgrid.row = 9
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(14, "E") = Fgrid
    End If
    Fgrid.row = 10
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(15, "E") = Fgrid
    End If
    Fgrid.row = 11
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(16, "E") = Fgrid
    End If
    Fgrid.row = 12
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(17, "E") = Fgrid
    End If
    Fgrid.row = 13
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(18, "E") = Fgrid
    End If
    Fgrid.row = 15
    If Trim(Fgrid) <> "" Then
        objWorksheet.Cells(20, "E") = Fgrid
    End If
      
        
    Screen.MousePointer = vbDefault
    
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    
    Exit Function
ExitFunction:
    objExcel.Interactive = True
    objExcel.DisplayAlerts = True
    Screen.MousePointer = vbDefault
    objExcel.Quit
    Exit Function
HANDLE_ERROR:
    MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
    Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    GoTo ExitFunction
    
End Function

Private Sub ClearFgrid()
Dim col As Integer
Dim row As Integer
Dim grid As String

    For col = 1 To 7
        For row = 2 To 15
            Fgrid.col = col
            Fgrid.row = row
            If Fgrid <> "" Then
                grid = "-"
            End If
            Fgrid = grid
        Next
    Next
    
End Sub

