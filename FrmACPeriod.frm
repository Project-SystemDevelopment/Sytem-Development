VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmACPeriod 
   Caption         =   "Bordx A/C Period"
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
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   120
      TabIndex        =   25
      Top             =   6360
      Width           =   11655
      Begin VB.ComboBox CboPeriodMth 
         Height          =   315
         ItemData        =   "FrmACPeriod.frx":0000
         Left            =   1680
         List            =   "FrmACPeriod.frx":0002
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CboPeriodYr 
         Height          =   315
         ItemData        =   "FrmACPeriod.frx":0004
         Left            =   4560
         List            =   "FrmACPeriod.frx":0006
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label39 
         Caption         =   "* - Search Empty Date"
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
         Left            =   6360
         TabIndex        =   39
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "A/C Period Mth"
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
         TabIndex        =   38
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "A/C Period Yr"
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
         Left            =   3240
         TabIndex        =   37
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   4560
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
         Height          =   375
         Left            =   6240
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9120
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7800
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10080
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10680
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7080
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdToPayor 
         Caption         =   "&Update"
         Height          =   375
         Left            =   8400
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstBordxDelivery 
      Height          =   4695
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8281
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
         Text            =   "Date Sent to Payor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Date Returned Fr Payor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Date Sent to MNRB"
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
         Text            =   "AC Period Status"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   120
      TabIndex        =   27
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   5040
         TabIndex        =   4
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox ChkACPeriod 
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
         TabIndex        =   11
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtBordxAmt2 
         Height          =   285
         Left            =   9240
         MaxLength       =   8
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtBordxAmt1 
         Height          =   285
         Left            =   7320
         MaxLength       =   8
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtBordxDate1 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   930
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
         TabIndex        =   6
         Top             =   930
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtACPeriod1 
         Height          =   285
         Left            =   7320
         TabIndex        =   9
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtACPeriod2 
         Height          =   285
         Left            =   9240
         TabIndex        =   10
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   8880
         TabIndex        =   36
         Top             =   645
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   35
         Top             =   650
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   650
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   32
         Top             =   1000
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   8880
         TabIndex        =   30
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx Amt"
         Height          =   255
         Left            =   6240
         TabIndex        =   29
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "A/C Period"
         Height          =   255
         Left            =   6240
         TabIndex        =   28
         Top             =   645
         Width           =   1095
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Bordx A/C Period"
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
      TabIndex        =   40
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmACPeriod"
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
Private m_blnDirection(1 To 11) As Boolean

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
    Call ComboListing
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
    Screen.MousePointer = vbHourglass
        'Call ExportToExcelBordxACPeriod(lstBordxDelivery)
    If lstBordxDelivery.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstBordxDelivery)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    CboPayor.Clear
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
    Call ComboListing
End Sub

Private Sub cmdSearch_Click()
    Dim abc As String
    ItemChecked = False
    LblTotalAmt = ""
    lblTotal = ""
    lblCurrent = ""
    ReDim BordxNoArray(0)
    CmdSearch.Enabled = False
    CmdToPayor.Enabled = False
    If ChkScan.Value = 1 Then
        If CboPayor = "" Then
            MsgBox "Please select a Payor", vbOKOnly + vbCritical, DialogBoxTitle
            CboPayor.SetFocus
        ElseIf TxtBordx1 = "" Then
            MsgBox "Please Enter Bordx No", vbOKOnly + vbCritical, DialogBoxTitle
            TxtBordx1.SetFocus
            CmdClear.Enabled = True
            CmdToPayor.Enabled = True
            cmdPrint.Enabled = True
            CmdExit.Enabled = True
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
                    cmdPrint.Enabled = False
                    CmdExit.Enabled = False
                    Call SearchRecord
                End If
                Check = False
                CmdClear.Enabled = True
                CmdToPayor.Enabled = True
                cmdPrint.Enabled = True
                CmdExit.Enabled = True
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
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    End If
    CmdSearch.Enabled = True
    CmdToPayor.Enabled = True
End Sub

Private Function SearchRecord()
On Error Resume Next
Dim strsql As String
Dim strappend As String
Dim sql As String
Dim CAS As New ADODB.Recordset
    strappend = ""
    If Len(Trim(CboPayor)) Then
        strappend = strappend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtBordx1)) Then
            strappend = strappend + " AND BordxRefNo = '" & Trim(TxtBordx1) & "'"
        End If
    Else
        If Len(Trim(TxtBordx1)) Then
            strappend = strappend + " AND BordxRefNo >= '" & Trim(TxtBordx1) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
            strappend = strappend + " AND BordxRefNo <= '" & Trim(TxtBordx2) & "'"
        End If
    End If
    
    If IsDate(txtBordxDate1) = True Then
        strappend = strappend + " And BordxDate >= '" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(txtBordxDate2) = True Then
        strappend = strappend + " And BordxDate <= '" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxAmt1)) Then
        strappend = strappend + " AND BordxClaimAmt >= '" & Trim(TxtBordxAmt1) & "'"
    End If
    
    If Len(Trim(TxtBordxAmt2)) Then
        strappend = strappend + " AND BordxClaimAmt <= '" & Trim(TxtBordxAmt2) & "'"
    End If
    
    If Mid(TxtACPeriod1, 1, 1) = 0 Then
        strappend = strappend + " AND PAYToACPeriodMth >= '" & Mid(TxtACPeriod1, 2, 1) & "'"
    ElseIf Mid(TxtACPeriod1, 1, 1) = 1 Then
        strappend = strappend + " AND PAYToACPeriodMth >= '" & Mid(TxtACPeriod1, 1, 2) & "'"
        strappend = strappend + " AND PAYToACPeriodYR >= '" & Mid(TxtACPeriod1, 4, 4) & "'"
    End If
    
    If Mid(TxtACPeriod2, 1, 1) = 0 Then
        strappend = strappend + " AND PAYToACPeriodMth <= '" & Mid(TxtACPeriod2, 2, 1) & "'"
    ElseIf Mid(TxtACPeriod2, 1, 1) = 1 Then
        strappend = strappend + " AND PAYToACPeriodMth <= '" & Mid(TxtACPeriod2, 1, 2) & "'"
        strappend = strappend + " AND PAYToACPeriodYR <= '" & Mid(TxtACPeriod2, 4, 4) & "'"
    End If
    If ChkACPeriod.Value = 1 Then
        sql = ""
        sql = " AND (PAYToACPeriodMth Is null OR PAYToACPeriodMth = '') AND (PAYToACPeriodYR Is null OR PAYToACPeriodYR = '') " & ""
        strappend = strappend + sql
    End If
    searchCriteria = strappend
    If ChkScan.Value = 0 Then
        Call BordxDeliveryListing(strappend)
    Else
        Call BordxDeliveryListing2(strappend)
    End If
End Function

Private Sub CmdStop_Click()
    intExit = 1
    TxtBordx1.SetFocus
    CmdToPayor.Enabled = True
    cmdPrint.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub
Private Sub CmdExit_Click()
        Unload Me
End Sub

Private Sub CmdToPayor_Click()
On Error Resume Next
Dim ii As Integer
Dim cnt As Integer
Dim recordFound As Boolean
Dim recordPassed As Boolean
Dim Update
    If ItemChecked = True Then
        CmdSearch.Enabled = False
        CmdToPayor.Enabled = False
        If CboPeriodMth = "" Then
            MsgBox "Month cannot be Empty.", vbCritical
            CboPeriodMth.SetFocus
        ElseIf CboPeriodYr = "" Then
            MsgBox "Year cannot be empty!", vbCritical
            CboPeriodYr.SetFocus
        Else
            Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                cnt = 0
                recordFound = False
                If lstBordxDelivery.ListItems.Count <> 0 Then
                    For ii = 0 To UBound(BordxNoArray)
                        If Trim(BordxNoArray(ii)) <> "" Then
                            UpdateACPeriod BordxNoArray(ii)
                            recordFound = True
                        End If
                    Next ii
                    If recordFound = True Then
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
        CmdSearch.Enabled = True
        CmdToPayor.Enabled = True
    Else
        MsgBox "Please select at least one record to be updated! ", vbCritical
    End If
End Sub
'**********************End Command Button *******************************

'**********************Start BordxDelivery Listing  *******************************
Private Sub BordxDeliveryListing(Optional strappend As String)
On Error Resume Next
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
    Total = 0
    Current = 0
    lstBordxDelivery.ListItems.Clear
    sql = " Select BordxRefNo, BordxPAYCode, BordxDate, BordxClaimAmt, BordxMNRBSentDate, " & _
          " BordxPAYSentDate, BordxPAYRecDate, PAYToACPeriodYR, PAYToACPeriodMth, " & _
          " ACPeriodStatus, BordxStatus From BORDX Where 0 = 0 "
    If Len(Trim(strappend)) Then
        sql = sql + strappend
    End If
    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdToPayor.Enabled = True
        cmdPrint.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
    Else
        Do
        Do Until rst2.EOF
            With rst2
                Set Record = lstBordxDelivery.ListItems.Add(, , "")
                If Len(!BordxPAYCode) Then
                    Record.SubItems(1) = !BordxPAYCode
                End If
                If Len(!BordxRefNo) Then
                    Record.SubItems(2) = !BordxRefNo
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(3) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                If Len(!BordxClaimAmt) Then
                    Record.SubItems(4) = Format(!BordxClaimAmt, "#,##0.00")
                End If
                If Len(!Bordxstatus) Then
                    Record.SubItems(5) = !Bordxstatus
                End If
                If Len(!BordxPAYSentDate) Then
                    Record.SubItems(6) = Format(!BordxPAYSentDate, "dd/mm/yyyy")
                End If
                If Len(!BordxPAYRecDate) Then
                    Record.SubItems(7) = Format(!BordxPAYRecDate, "dd/mm/yyyy")
                End If
                If Len(!BordxMNRBSentDate) Then
                    Record.SubItems(8) = Format(!BordxMNRBSentDate, "dd/mm/yyyy")
                End If
                If Len(!PAYToACPeriodMth & "/" & !PAYToACPeriodYR) Then
                    Record.SubItems(9) = Trim(!PAYToACPeriodMth & "/" & !PAYToACPeriodYR)
                End If
                If Len(!ACPeriodStatus) Then
                    Record.SubItems(10) = !ACPeriodStatus
                End If
                Current = Current + 1
                lblCurrent = Current
                If Len(!BordxClaimAmt) Then
                    TotalAmt = TotalAmt + !BordxClaimAmt
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
    CmdToPayor.Enabled = True
    cmdPrint.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub BordxDeliveryListing2(Optional strappend As String)
On Error Resume Next
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim ClaimAdmin As Boolean
    sql = " Select BordxRefNo, BordxPAYCode, BordxDate, BordxClaimAmt, BordxMNRBSentDate, " & _
          " BordxPAYSentDate, BordxPAYRecDate, PAYToACPeriodYR, PAYToACPeriodMth, " & _
          " ACPeriodStatus, BordxStatus From BORDX Where 0 = 0 "
    If Len(Trim(strappend)) Then
        sql = sql + strappend
    End If
    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdToPayor.Enabled = True
        cmdPrint.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
    Else
        Do
        Do Until rst2.EOF
            With rst2
                Set Record = lstBordxDelivery.ListItems.Add(, , "")
                If Len(!BordxPAYCode) Then
                    Record.SubItems(1) = !BordxPAYCode
                End If
                If Len(!BordxRefNo) Then
                    Record.SubItems(2) = !BordxRefNo
                End If
                If Len(!BordxDate) Then
                    Record.SubItems(3) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                If Len(!BordxClaimAmt) Then
                    Record.SubItems(4) = Format(!BordxClaimAmt, "#,##0.00")
                End If
                If Len(!Bordxstatus) Then
                    Record.SubItems(5) = !Bordxstatus
                End If
                If Len(!BordxPAYSentDate) Then
                    Record.SubItems(6) = Format(!BordxPAYSentDate, "dd/mm/yyyy")
                End If
                If Len(!BordxPAYRecDate) Then
                    Record.SubItems(7) = Format(!BordxPAYRecDate, "dd/mm/yyyy")
                End If
                If Len(!BordxMNRBSentDate) Then
                    Record.SubItems(8) = Format(!BordxMNRBSentDate, "dd/mm/yyyy")
                End If
                If Len(!PAYToACPeriodMth & "/" & !PAYToACPeriodYR) Then
                    Record.SubItems(9) = Trim(!PAYToACPeriodMth & "/" & !PAYToACPeriodYR)
                End If
                If Len(!ACPeriodStatus) Then
                    Record.SubItems(10) = !ACPeriodStatus
                End If
                Current2 = Current2 + 1
                lblTotal = Current2
                If Len(!BordxClaimAmt) Then
                    TotalAmt2 = TotalAmt2 + !BordxClaimAmt
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
    CmdToPayor.Enabled = True
    cmdPrint.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub
'**********************End BordxDelivery Listing *******************************

'**********************Start Check ListView Record*******************************
Private Sub lstBordxDelivery_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim temp
Static Arraycnt As Integer
    If Item.Checked = True Then
        If Item.SubItems(10) = False Then
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
'**********************End Check ListView Record***************************************

'**********************Start Update all Record functions*******************************
Private Sub UpdateACPeriod(tBordxNo As String)
On Error Resume Next
Dim strsql As String
Dim ACPeriod As New ADODB.Recordset
    strsql = " update Bordx set " & _
             " PAYToACPeriodMth = '" & Trim(CboPeriodMth) & "', " & _
             " ACPeriodStatus = '1', " & _
             " PAYToACPeriodYR = '" & Trim(CboPeriodYr) & "' " & _
             " Where BordxRefNo = '" & Trim(tBordxNo) & "'"
    medixmasterconnWS.Execute strsql
End Sub
'**********************End Update all Record functions*******************************

'**********************Start Sorted ListView *******************************
Private Sub lstBordxDelivery_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtstring, m_blnDirection(1)
    Case 2
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtstring, m_blnDirection(2)
    Case 3
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtstring, m_blnDirection(3)
    Case 4
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtstring, m_blnDirection(4)
    Case 5
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lstBordxDelivery, ColumnHeader.Index, ldtstring, m_blnDirection(11)
    End Select
End Sub
'**********************End Sorted ListView *******************************

'**********************Start Combo Listing*******************************
Private Sub ComboListing()
On Error Resume Next
Dim PAY As New ADODB.Recordset
Dim MonthStart
Dim YearStart
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
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
                CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
            End With
            PAY.MoveNext
        Loop
        PAY.Close
        Set PAY = Nothing
        For MonthStart = 1 To 12
            With CboPeriodMth
                .AddItem MonthStart
            End With
        Next MonthStart
        For YearStart = 1999 To 3000
            With CboPeriodYr
                .AddItem YearStart
            End With
        Next YearStart
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

