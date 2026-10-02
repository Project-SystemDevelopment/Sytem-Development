VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBordxTracking2 
   Caption         =   "Bordx Tracking"
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox cboReportList 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   3500
      End
      Begin VB.Label Label20 
         Caption         =   "Report Type"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   280
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   34
      Top             =   6960
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   375
         Left            =   4200
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   7440
         TabIndex        =   21
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10560
         TabIndex        =   25
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9960
         TabIndex        =   24
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   8160
         TabIndex        =   22
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   8880
         TabIndex        =   23
         Top             =   220
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3240
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1680
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1440
         TabIndex        =   36
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1560
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lstBordxTracking 
      Height          =   3855
      Left            =   120
      TabIndex        =   39
      Top             =   3120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6800
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Bordx Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "CLM Status"
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
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Date Sent to A/C"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "A/C Period"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   120
      TabIndex        =   40
      Top             =   960
      Width           =   11655
      Begin VB.ComboBox CboHLTCode 
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.CheckBox ChkBordxNo 
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
         Left            =   5040
         TabIndex        =   26
         Top             =   585
         Width           =   495
      End
      Begin VB.CheckBox ChkBordxAmt 
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
         Left            =   5040
         TabIndex        =   28
         Top             =   1395
         Width           =   495
      End
      Begin VB.CheckBox ChkBordxDate 
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
         Left            =   5040
         TabIndex        =   27
         Top             =   1140
         Width           =   495
      End
      Begin VB.CheckBox ChkDTAC 
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
         Left            =   10560
         TabIndex        =   33
         Top             =   1270
         Width           =   495
      End
      Begin VB.CheckBox ChkDTP 
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
         Left            =   10560
         TabIndex        =   30
         Top             =   500
         Width           =   495
      End
      Begin VB.CheckBox ChkDTMNRB 
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
         Left            =   10560
         TabIndex        =   32
         Top             =   1010
         Width           =   495
      End
      Begin VB.CheckBox ChkDFP 
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
         Left            =   10560
         TabIndex        =   31
         Top             =   750
         Width           =   495
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   4
         Top             =   530
         Width           =   1695
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
         Left            =   10560
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin MSMask.MaskEdBox TxtACPeriod1 
         Height          =   285
         Left            =   6840
         TabIndex        =   11
         Top             =   240
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
         Left            =   8760
         TabIndex        =   12
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDTP2 
         Height          =   285
         Left            =   8760
         TabIndex        =   14
         Top             =   495
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDFP2 
         Height          =   285
         Left            =   8760
         TabIndex        =   16
         Top             =   750
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDTMNRB2 
         Height          =   285
         Left            =   8760
         TabIndex        =   18
         Top             =   1005
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDTAC2 
         Height          =   285
         Left            =   8760
         TabIndex        =   20
         Top             =   1245
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   3
         Top             =   530
         Width           =   1455
      End
      Begin MSMask.MaskEdBox TxtDTP1 
         Height          =   285
         Left            =   6840
         TabIndex        =   13
         Top             =   495
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDFP1 
         Height          =   285
         Left            =   6840
         TabIndex        =   15
         Top             =   750
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDTMNRB1 
         Height          =   285
         Left            =   6840
         TabIndex        =   17
         Top             =   1005
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDTAC1 
         Height          =   285
         Left            =   6840
         TabIndex        =   19
         Top             =   1245
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   780
         Width           =   3615
      End
      Begin MSMask.MaskEdBox txtBordxDate2 
         Height          =   300
         Left            =   3240
         TabIndex        =   7
         Top             =   1070
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDate1 
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   1070
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBordxAmt1 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   8
         Top             =   1340
         Width           =   1455
      End
      Begin VB.TextBox TxtBordxAmt2 
         Height          =   285
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1340
         Width           =   1695
      End
      Begin VB.ComboBox CboHospital 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1590
         Width           =   3615
      End
      Begin VB.Label Label19 
         Caption         =   "Health Type"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   1740
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Date to A/C"
         Height          =   255
         Left            =   5640
         TabIndex        =   58
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   8400
         TabIndex        =   57
         Top             =   1350
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Date to MNRB"
         Height          =   255
         Left            =   5640
         TabIndex        =   56
         Top             =   1100
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "To"
         Height          =   255
         Left            =   8400
         TabIndex        =   55
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Date Fr Payor"
         Height          =   255
         Left            =   5640
         TabIndex        =   54
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   8400
         TabIndex        =   53
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Date to Payor"
         Height          =   255
         Left            =   5640
         TabIndex        =   52
         Top             =   550
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   8400
         TabIndex        =   51
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "A/C Period"
         Height          =   255
         Left            =   5640
         TabIndex        =   49
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx Amt"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1485
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   47
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   45
         Top             =   1100
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   42
         Top             =   615
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   8400
         TabIndex        =   41
         Top             =   300
         Width           =   255
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Bordx Tracking"
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
      TabIndex        =   59
      Top             =   0
      Width           =   11805
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
      Left            =   120
      TabIndex        =   50
      Top             =   7700
      Width           =   2055
   End
End
Attribute VB_Name = "FrmBordxTracking2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim searchCriteria As String
Dim ScanCode As String
Dim CC As Integer
Dim Payor As String
Dim Hospital As String
Dim TotalActAmt As Variant
Dim TotalAppAmt As Variant
Dim TotalAllAppAmt As Variant
Dim NoInv As Variant
Dim ActAmt As Variant
Dim AppAmt As Variant
Dim NoInv2 As Variant
Dim ActAmt2 As Variant
Dim AppAmt2 As Variant
Dim NoInv3 As Variant
Dim ActAmt3 As Variant
Dim AppAmt3 As Variant
Dim NoInv4 As Variant
Dim ActAmt4 As Variant
Dim AppAmt4 As Variant
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

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

'**********************Start Command Button *******************************

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        'Call ExportToExcelBordxTracking(lstBordxTracking)
    If lstBordxTracking.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstBordxTracking)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdClear_Click()
    CboPayor.Clear
    CboHLTCode.Clear
    CboHospital.Clear
    Call ClearForm(Me)
    lstBordxTracking.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
'    medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

Private Sub cmdSearch_Click()
On Error GoTo errRun
    ScanCode = TxtBordx1
    If Len(cboReportList) = False Then
        MsgBox "Please select a Report Type.", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    
    If Mid(ScanCode, 1, 2) = "LO" Then
        TxtBordx1 = ""
        Call SearchForm(ScanCode)
    Else
        LblTotalAmt = ""
        lblTotal = ""
        lblCurrent = ""
        
        CmdClear.Enabled = False
        cmdPrint.Enabled = False
        CmdExit.Enabled = False
        CmdSearch.Enabled = False
        Screen.MousePointer = vbHourglass
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

            Call SearchRecord
     '   medixmasterconnWS.Close
     '   Set medixmasterconnWS = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
        
        Screen.MousePointer = Default
        CmdSearch.Enabled = True
    End If
    Exit Sub

errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstBordxTracking.ListItems.Clear
    lblCurrent = ""
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnWS.Close
    'Set medixmasterconnWS = Nothing

End Sub

Private Sub cboReportList_Click()
    If cboReportList.ListIndex = 0 Then
        Call Option1
    ElseIf cboReportList.ListIndex = 1 Then
        Call Option2
    ElseIf cboReportList.ListIndex = 2 Then
        Call Option3
    ElseIf cboReportList.ListIndex = 3 Then
        Call Option4
    ElseIf cboReportList.ListIndex = 4 Then
        Call Option5
    ElseIf cboReportList.ListIndex = 5 Then
        Call Option6
    ElseIf cboReportList.ListIndex = 6 Then
        Call Option7
    End If

End Sub

Private Function SearchRecord()
Dim strsql As String
Dim strappend As String
Dim sql As String
Dim CAS As New ADODB.Recordset
Dim HOspitals As String

    strappend = ""
    HOspitals = ""
    
    If cboReportList.ListIndex = 0 Or cboReportList.ListIndex = 1 Or cboReportList.ListIndex = 6 Then
    
        If Len(Trim(CboHLTCode)) Then
            strappend = strappend + " And HLTCode = '" & Trim(Mid(CboHLTCode, 1, IIf(InStr(1, CboHLTCode, "-") = 0, 2, InStr(1, CboHLTCode, "-") - 1))) & "'"
        End If
        
        If Len(Trim(CboPayor)) Then
            strappend = strappend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
        End If
        
        If Len(Trim(TxtBordx1)) Then
            strappend = strappend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
            strappend = strappend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
        
        If ChkBordxNo.Value = 1 Then
            sql = ""
            sql = " AND (BordxRefNo Is null OR BordxRefNo = '')"
            strappend = strappend + sql
        End If
        
        If IsDate(txtBordxDate1) = True Then
            strappend = strappend + " And BordxDate >= '" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(txtBordxDate2) = True Then
            strappend = strappend + " And BordxDate <= '" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
        End If
        
        If ChkBordxDate.Value = 1 Then
            sql = ""
            sql = " AND (BordxDate Is null OR BordxDate = '')"
            strappend = strappend + sql
        End If
        
        If Len(Trim(TxtBordxAmt1)) Then
            strappend = strappend + " AND BordxClaimAmt >= '" & Trim(TxtBordxAmt1) & "'"
        End If
        
        If Len(Trim(TxtBordxAmt2)) Then
            strappend = strappend + " AND BordxClaimAmt <= '" & Trim(TxtBordxAmt2) & "'"
        End If
        
        If ChkBordxAmt.Value = 1 Then
            sql = ""
            sql = " AND BordxClaimAmt Is null"
            strappend = strappend + sql
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
        
        If IsDate(TxtDTP1) = True Then
            strappend = strappend + " And BordxPAYSentDate >= '" & Format(TxtDTP1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(TxtDTP2) = True Then
            strappend = strappend + " And BordxPAYSentDate <= '" & Format(TxtDTP2, "mm/dd/yyyy") & "'"
        End If
        
        If IsDate(TxtDFP1) = True Then
            strappend = strappend + " And BordxPAYRecDate >= '" & Format(TxtDFP1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(TxtDFP2) = True Then
            strappend = strappend + " And BordxPAYRecDate <= '" & Format(TxtDFP2, "mm/dd/yyyy") & "'"
        End If
        
        If IsDate(TxtDTMNRB1) = True Then
            strappend = strappend + " And BordxMNRBSentDate >= '" & Format(TxtDTMNRB1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(TxtDTMNRB2) = True Then
            strappend = strappend + " And BordxMNRBSentDate <= '" & Format(TxtDTMNRB2, "mm/dd/yyyy") & "'"
        End If
        
        If IsDate(TxtDTAC1) = True Then
            strappend = strappend + " And BordxSendACDate >= '" & Format(TxtDTAC1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(TxtDTAC2) = True Then
            strappend = strappend + " And BordxSendACDate <= '" & Format(TxtDTAC2, "mm/dd/yyyy") & "'"
        End If
        
        If ChkACPeriod.Value = 1 Then
            sql = ""
            sql = " AND (PAYToACPeriodMth Is null OR PAYToACPeriodMth = '') AND (PAYToACPeriodYR Is null OR PAYToACPeriodYR = '') " & ""
            strappend = strappend + sql
        End If
        
        If ChkDTP.Value = 1 Then
            sql = ""
            sql = " AND (BordxPAYSentDate Is null OR BordxPAYSentDate = '') AND (PAYToACPeriodYR Is null OR PAYToACPeriodYR = '') " & ""
            strappend = strappend + sql
        End If
        
        If ChkDFP.Value = 1 Then
            sql = ""
            sql = " AND (BordxPAYRecDate Is null OR BordxPAYRecDate = '') AND (PAYToACPeriodYR Is null OR PAYToACPeriodYR = '') " & ""
            strappend = strappend + sql
        End If
        
        If ChkDTMNRB.Value = 1 Then
            sql = ""
            sql = " AND (BordxMNRBSentDate Is null OR BordxMNRBSentDate = '') AND (PAYToACPeriodYR Is null OR PAYToACPeriodYR = '') " & ""
            strappend = strappend + sql
        End If
        
        If ChkDTAC.Value = 1 Then
            sql = ""
            sql = " AND (BordxSendACDate Is null OR BordxSendACDate = '') AND (PAYToACPeriodYR Is null OR PAYToACPeriodYR = '') " & ""
            strappend = strappend + sql
        End If
        
    Else
        If Len(Trim(CboPayor)) And cboReportList.ListIndex = 5 Then
            strappend = strappend + " And Payor = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
        End If
        
        If Len(Trim(CboPayor)) And cboReportList.ListIndex = 3 Then
            strappend = strappend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
        End If
        
        If Len(Trim(CboHospital)) Then
            If cboReportList.ListIndex = 2 Or cboReportList.ListIndex = 4 Then
                HOspitals = Replace(CboHospital, "'", "''")
                strappend = strappend + " And HOSName = '" & Trim(HOspitals) & "'"
            End If
        End If
        
        If IsDate(TxtDTAC1) = True Then
            strappend = strappend + " And ClaimOpenDate >= '" & Format(TxtDTAC1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(TxtDTAC2) = True Then
            strappend = strappend + " And ClaimOpenDate <= '" & Format(TxtDTAC2, "mm/dd/yyyy") & "'"
        End If
        
        If ChkDTAC.Value = 1 Then
            sql = ""
            sql = " AND (ClaimOpenDate Is null OR ClaimOpenDate = '') AND (PAYToACPeriodYR Is null OR PAYToACPeriodYR = '') " & ""
            strappend = strappend + sql
        End If
    
    End If
        
    searchCriteria = strappend
    
    If cboReportList.ListIndex = 1 Then
        Call BordxTracking(strappend)
    ElseIf cboReportList.ListIndex = 0 Then
        Call BordxTracking2(strappend)
    ElseIf cboReportList.ListIndex = 2 Then
        Call HospInvNotBordx(strappend)
    ElseIf cboReportList.ListIndex = 3 Then
        Call MBMInvNotBordx(strappend)
    ElseIf cboReportList.ListIndex = 4 Then
        Call HospInvRegMonth(strappend)
    ElseIf cboReportList.ListIndex = 5 Then
        Call MBMInvRegMonth(strappend)
    ElseIf cboReportList.ListIndex = 6 Then
        Call BordxTrackingByPayor(strappend)
    End If
    
End Function

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub


Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub
'**********************End Command Button *******************************

'**********************Start BordxDelivery Listing  *******************************
Private Sub BordxTracking(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
    
    Total = 0
    Current = 0
    
    lstBordxTracking.ListItems.Clear
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
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
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        
            With rst2
                Set Record = lstBordxTracking.ListItems.Add(, , rst2!BordxPAYCode)
                
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(1) = rst2!BordxRefNo
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(2) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(3) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!Bordxstatus) Then
                    Record.SubItems(4) = rst2!Bordxstatus
                End If
                If Len(rst2!ClaimStatus) Then
                    Record.SubItems(5) = rst2!ClaimStatus
                End If
                If Len(rst2!BordxPAYSentDate) Then
                    Record.SubItems(6) = Format(rst2!BordxPAYSentDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxPAYRecDate) Then
                    Record.SubItems(7) = Format(rst2!BordxPAYRecDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxMNRBSentDate) Then
                    Record.SubItems(8) = Format(rst2!BordxMNRBSentDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxSendACDate) Then
                    Record.SubItems(9) = Format(rst2!BordxSendACDate, "dd/mm/yyyy")
                End If
                If Len(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR) Then
                    Record.SubItems(10) = Trim(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR)
                End If
                If Len(rst2!ReceiptNo) Then
                    Record.SubItems(11) = rst2!ReceiptNo
                End If
                If Len(rst2!ReceiptDate) Then
                    Record.SubItems(12) = Format(rst2!ReceiptDate, "dd/mm/yyyy")
                End If
                If Len(rst2!RECAccPeriodMth & "/" & rst2!RECAccPeriodYR) Then
                    Record.SubItems(13) = Trim(rst2!RECAccPeriodMth & "/" & rst2!RECAccPeriodYR)
                End If
                If Len(rst2!ReceiptNoPAY) Then
                    Record.SubItems(14) = rst2!ReceiptNoPAY
                End If
                If Len(rst2!ReceiptDatePay) Then
                    Record.SubItems(15) = Format(rst2!ReceiptDatePay, "dd/mm/yyyy")
                End If
                If Len(rst2!RECAccPeriodMthPay & "/" & rst2!RECAccPeriodYRPay) Then
                    Record.SubItems(16) = Trim(rst2!RECAccPeriodMthPay & "/" & rst2!RECAccPeriodYRPay)
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
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub BordxTracking2(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
    
    Total = 0
    Current = 0
    
    lstBordxTracking.ListItems.Clear
    strAppendCase = "2"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
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
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
            
            With rst2
                Set Record = lstBordxTracking.ListItems.Add(, , rst2!BordxPAYCode)
                
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(1) = rst2!BordxRefNo
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(2) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(3) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!Bordxstatus) Then
                    Record.SubItems(4) = rst2!Bordxstatus
                End If
                If Len(rst2!ClaimStatus) Then
                    Record.SubItems(5) = rst2!ClaimStatus
                End If
                If Len(rst2!BordxPAYSentDate) Then
                    Record.SubItems(6) = Format(rst2!BordxPAYSentDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxPAYRecDate) Then
                    Record.SubItems(7) = Format(rst2!BordxPAYRecDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxMNRBSentDate) Then
                    Record.SubItems(8) = Format(rst2!BordxMNRBSentDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxSendACDate) Then
                    Record.SubItems(9) = Format(rst2!BordxSendACDate, "dd/mm/yyyy")
                End If
                If Len(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR) Then
                    Record.SubItems(10) = Trim(rst2!PAYToACPeriodMth & "/" & rst2!PAYToACPeriodYR)
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
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub HospInvNotBordx(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
    
    Hospital = ""
    Total = 0
    Current = 0
    TotalActAmt = 0
    TotalAppAmt = 0
    TotalAllAppAmt = 0
    NoInv = 0
    ActAmt = 0
    AppAmt = 0
    NoInv2 = 0
    ActAmt2 = 0
    AppAmt2 = 0
    NoInv3 = 0
    ActAmt3 = 0
    AppAmt3 = 0
    NoInv4 = 0
    ActAmt4 = 0
    AppAmt4 = 0
    
    lstBordxTracking.ListItems.Clear
    strAppendCase = "3"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
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
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        
            With rst2
                Set Record = lstBordxTracking.ListItems.Add(, , rst2!HosName)
                Hospital = Replace(!HosName, "'", "''")
                
                '========= < 31 Day ===============================
                Call CountHOSInvNotBordx31
                
                If Len(NoInv) Then
                   Record.SubItems(1) = NoInv
                End If
                If Len(ActAmt) Then
                   Record.SubItems(2) = Format(ActAmt, "#,##0.00")
                End If
                If Len(AppAmt) Then
                   Record.SubItems(3) = Format(AppAmt, "#,##0.00")
                End If
                '========= < 31 Day ===============================
                
                '========= < 61 Day ===============================
                Call CountHOSInvNotBordx61
                
                If Len(NoInv2) Then
                   Record.SubItems(4) = NoInv2
                End If
                If Len(ActAmt2) Then
                   Record.SubItems(5) = Format(ActAmt2, "#,##0.00")
                End If
                If Len(AppAmt2) Then
                   Record.SubItems(6) = Format(AppAmt2, "#,##0.00")
                End If
                '========= < 61 Day ===============================
                
                '========= < 91 Day ===============================
                Call CountHOSInvNotBordx91
                
                If Len(NoInv3) Then
                   Record.SubItems(7) = NoInv3
                End If
                If Len(ActAmt3) Then
                   Record.SubItems(8) = Format(ActAmt3, "#,##0.00")
                End If
                If Len(AppAmt3) Then
                   Record.SubItems(9) = Format(AppAmt3, "#,##0.00")
                End If
                '========= < 91 Day ===============================
                
                '========= > 91 Day ===============================
                Call CountHOSInvNotBordx912
                
                If Len(NoInv4) Then
                   Record.SubItems(10) = NoInv4
                End If
                If Len(ActAmt4) Then
                   Record.SubItems(11) = Format(ActAmt4, "#,##0.00")
                End If
                If Len(AppAmt4) Then
                   Record.SubItems(12) = Format(AppAmt4, "#,##0.00")
                End If
                '========= > 91 Day ===============================
                
                Record.SubItems(13) = (CDbl(NoInv) + CDbl(NoInv2) + CDbl(NoInv3) + CDbl(NoInv4))
                TotalActAmt = (CDbl(ActAmt) + CDbl(ActAmt2) + CDbl(ActAmt3) + CDbl(ActAmt4))
                Record.SubItems(14) = Format(TotalActAmt, "#,##0.00")
                TotalAppAmt = (CDbl(AppAmt) + CDbl(AppAmt2) + CDbl(AppAmt3) + CDbl(AppAmt4))
                Record.SubItems(15) = Format(TotalAppAmt, "#,##0.00")
                                
                Current = Current + 1
                lblCurrent = Current
                'If Len(rst2!BordxClaimAmt) Then
                TotalAllAppAmt = TotalAllAppAmt + TotalAppAmt
                'End If
                LblTotalAmt = Format(TotalAllAppAmt, "#,##0.00")
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
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub HospInvRegMonth(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
Dim Month As String
   
    Total = 0
    Current = 0
    Month = ""
        
    lstBordxTracking.ListItems.Clear
    strAppendCase = "5"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
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
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!HosName) Then
                    Set Record = lstBordxTracking.ListItems.Add(, , rst2!HosName)
                Else
                    Set Record = lstBordxTracking.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = !Year & "/" & Month
                    Else
                        Record.SubItems(1) = !Year & "/" & !Month
                    End If
                End If
                
                If Len(!NoInv) Then
                   Record.SubItems(2) = !NoInv
                End If
                If Len(!ActAmt) Then
                   Record.SubItems(3) = Format(!ActAmt, "#,##0.00")
                End If
                If Len(!AppAmt) Then
                   Record.SubItems(4) = Format(!AppAmt, "#,##0.00")
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!AppAmt) Then
                    TotalAmt = TotalAmt + rst2!AppAmt
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
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub MBMInvNotBordx(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
    
    Total = 0
    Current = 0
    TotalActAmt = 0
    TotalAppAmt = 0
    TotalAllAppAmt = 0
    NoInv = 0
    ActAmt = 0
    AppAmt = 0
    NoInv2 = 0
    ActAmt2 = 0
    AppAmt2 = 0
    NoInv3 = 0
    ActAmt3 = 0
    AppAmt3 = 0
    NoInv4 = 0
    ActAmt4 = 0
    AppAmt4 = 0
    Payor = ""
    
    lstBordxTracking.ListItems.Clear
    strAppendCase = "4"
    Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
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
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        
            With rst2
                Set Record = lstBordxTracking.ListItems.Add(, , rst2!PAYCompanyName)
                Payor = !PAYCode
                
                '========= < 31 Day ===============================
                Call CountMBMInvNotBordx31
                
                If Len(NoInv) Then
                   Record.SubItems(1) = NoInv
                End If
                If Len(ActAmt) Then
                   Record.SubItems(2) = Format(ActAmt, "#,##0.00")
                End If
                If Len(AppAmt) Then
                   Record.SubItems(3) = Format(AppAmt, "#,##0.00")
                End If
                '========= < 31 Day ===============================
                
                '========= < 61 Day ===============================
                Call CountMBMInvNotBordx61
                
                If Len(NoInv2) Then
                   Record.SubItems(4) = NoInv2
                End If
                If Len(ActAmt2) Then
                   Record.SubItems(5) = Format(ActAmt2, "#,##0.00")
                End If
                If Len(AppAmt2) Then
                   Record.SubItems(6) = Format(AppAmt2, "#,##0.00")
                End If
                '========= < 61 Day ===============================
                
                '========= < 91 Day ===============================
                Call CountMBMInvNotBordx91
                
                If Len(NoInv3) Then
                   Record.SubItems(7) = NoInv3
                End If
                If Len(ActAmt3) Then
                   Record.SubItems(8) = Format(ActAmt3, "#,##0.00")
                End If
                If Len(AppAmt3) Then
                   Record.SubItems(9) = Format(AppAmt3, "#,##0.00")
                End If
                '========= < 91 Day ===============================
                
                '========= > 91 Day ===============================
                Call CountMBMInvNotBordx912
                
                If Len(NoInv4) Then
                   Record.SubItems(10) = NoInv4
                End If
                If Len(ActAmt4) Then
                   Record.SubItems(11) = Format(ActAmt4, "#,##0.00")
                End If
                If Len(AppAmt4) Then
                   Record.SubItems(12) = Format(AppAmt4, "#,##0.00")
                End If
                '========= > 91 Day ===============================
                
                Record.SubItems(13) = (CDbl(NoInv) + CDbl(NoInv2) + CDbl(NoInv3) + CDbl(NoInv4))
                TotalActAmt = (CDbl(ActAmt) + CDbl(ActAmt2) + CDbl(ActAmt3) + CDbl(ActAmt4))
                Record.SubItems(14) = Format(TotalActAmt, "#,##0.00")
                TotalAppAmt = (CDbl(AppAmt) + CDbl(AppAmt2) + CDbl(AppAmt3) + CDbl(AppAmt4))
                Record.SubItems(15) = Format(TotalAppAmt, "#,##0.00")
                                
                Current = Current + 1
                lblCurrent = Current
                TotalAllAppAmt = TotalAllAppAmt + TotalAppAmt
                LblTotalAmt = Format(TotalAllAppAmt, "#,##0.00")
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
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub MBMInvRegMonth(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
Dim Month As String
   
    Total = 0
    Current = 0
    Month = ""
        
    lstBordxTracking.ListItems.Clear
    strAppendCase = "6"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
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
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!Payor) Then
                    Set Record = lstBordxTracking.ListItems.Add(, , rst2!Payor)
                Else
                    Set Record = lstBordxTracking.ListItems.Add(, , "Empty")
                End If
                
                If Len(!Month & "/" & !Year) Then
                    If Len(!Month) = 1 Then
                        Month = "0" & !Month
                        Record.SubItems(1) = !Year & "/" & Month
                    Else
                        Record.SubItems(1) = !Year & "/" & !Month
                    End If
                End If
                
                If Len(!NoInv) Then
                   Record.SubItems(2) = !NoInv
                End If
                If Len(!ActAmt) Then
                   Record.SubItems(3) = Format(!ActAmt, "#,##0.00")
                End If
                If Len(!AppAmt) Then
                   Record.SubItems(4) = Format(!AppAmt, "#,##0.00")
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!AppAmt) Then
                    TotalAmt = TotalAmt + rst2!AppAmt
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
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub
'**********************End BordxDelivery Listing *******************************

'**********************Start Sorted ListView *******************************
Private Sub lstBordxTracking_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    If cboReportList.ListIndex = 0 Then
    
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstBordxTracking, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstBordxTracking, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
        Case 4
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
        Case 8
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
        Case 9
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
        Case 10
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
        Case 11
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        End Select
    
    ElseIf cboReportList.ListIndex = 1 Then
    
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstBordxTracking, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstBordxTracking, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
        Case 4
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
        Case 8
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
        Case 9
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
        Case 10
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
        Case 11
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView lstBordxTracking, ColumnHeader.Index, ldtString, m_blnDirection(12)
        Case 13
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView lstBordxTracking, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView lstBordxTracking, ColumnHeader.Index, ldtDateTime, m_blnDirection(16)
        Case 17
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
        End Select
        
    ElseIf cboReportList.ListIndex = 2 Or cboReportList.ListIndex = 3 Then
    
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstBordxTracking, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        Case 16
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        End Select
        
    ElseIf cboReportList.ListIndex = 4 Or cboReportList.ListIndex = 5 Then
    
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstBordxTracking, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
        
    ElseIf cboReportList.ListIndex = 6 Then
    
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstBordxTracking, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstBordxTracking, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
    
    End If
End Sub
'**********************End Sorted ListView *******************************

'**********************Start Combo Listing*******************************
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim cmd As New ADODB.Command

    cboReportList.Clear
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
            CboHospital.AddItem !HosName
            HOS.MoveNext
        End With
    Loop
    HOS.Close
    Set HOS = Nothing
    
    CboHLTCode.AddItem "N" & " - " & "Non-Sihat"
    CboHLTCode.AddItem "S" & " - " & "Sihat"
    
    With cboReportList
        .AddItem "Without Receipt No."
        .AddItem "With Receipt No."
        .AddItem "Hospital Invoice Not Bordx"
        .AddItem "MBM Invoice Not Bordx"
        .AddItem "Hospital Invoice Registration Month"
        .AddItem "MBM Invoice Registration Month"
        .AddItem "By Payor"
    End With

End Sub
'**********************End Combo Listing*******************************


Private Sub Option1()
        
    CboHospital.Enabled = False
    CboHLTCode.Enabled = True
    CboPayor.Enabled = True
    TxtBordx1.Enabled = True
    TxtBordx2.Enabled = True
    ChkBordxNo.Enabled = True
    txtBordxDate1.Enabled = True
    txtBordxDate2.Enabled = True
    ChkBordxDate.Enabled = True
    TxtBordxAmt1.Enabled = True
    TxtBordxAmt2.Enabled = True
    ChkBordxAmt.Enabled = True
    TxtACPeriod1.Enabled = True
    TxtACPeriod2.Enabled = True
    TxtDTP1.Enabled = True
    TxtDTP2.Enabled = True
    TxtDFP1.Enabled = True
    TxtDFP2.Enabled = True
    TxtDTMNRB1.Enabled = True
    TxtDTMNRB2.Enabled = True
    ChkACPeriod.Enabled = True
    ChkDTP.Enabled = True
    ChkDFP.Enabled = True
    ChkDTMNRB.Enabled = True
    Label17.Caption = "Date to A/C"
            
    lstBordxTracking.ListItems.Clear
    
    lstBordxTracking.ColumnHeaders(1).Text = "Payor"
    lstBordxTracking.ColumnHeaders(2).Text = "Bordx No"
    lstBordxTracking.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstBordxTracking.ColumnHeaders(3).Text = "Bordx Date"
    lstBordxTracking.ColumnHeaders(3).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(4).Text = "Bordx Amount"
    lstBordxTracking.ColumnHeaders(4).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(5).Text = "Bordx Status"
    lstBordxTracking.ColumnHeaders(5).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(6).Text = "CLM Status"
    lstBordxTracking.ColumnHeaders(6).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(7).Text = "Date Sent to Payor"
    lstBordxTracking.ColumnHeaders(7).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(8).Text = "Date Returned Fr Payor"
    lstBordxTracking.ColumnHeaders(8).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(9).Text = "Date Sent to MNRB"
    lstBordxTracking.ColumnHeaders(9).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(10).Text = "Date Sent to A/C"
    lstBordxTracking.ColumnHeaders(10).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(11).Text = "A/C Period"
    lstBordxTracking.ColumnHeaders(11).Alignment = lvwColumnRight
    
    For CC = 12 To 17
        lstBordxTracking.ColumnHeaders(CC).Text = ""
    Next CC
    
End Sub

Private Sub Option2()
    
    CboHospital.Enabled = False
    CboHLTCode.Enabled = True
    CboPayor.Enabled = True
    TxtBordx1.Enabled = True
    TxtBordx2.Enabled = True
    ChkBordxNo.Enabled = True
    txtBordxDate1.Enabled = True
    txtBordxDate2.Enabled = True
    ChkBordxDate.Enabled = True
    TxtBordxAmt1.Enabled = True
    TxtBordxAmt2.Enabled = True
    ChkBordxAmt.Enabled = True
    TxtACPeriod1.Enabled = True
    TxtACPeriod2.Enabled = True
    TxtDTP1.Enabled = True
    TxtDTP2.Enabled = True
    TxtDFP1.Enabled = True
    TxtDFP2.Enabled = True
    TxtDTMNRB1.Enabled = True
    TxtDTMNRB2.Enabled = True
    ChkACPeriod.Enabled = True
    ChkDTP.Enabled = True
    ChkDFP.Enabled = True
    ChkDTMNRB.Enabled = True
    Label17.Caption = "Date to A/C"
    lstBordxTracking.ListItems.Clear
    
    lstBordxTracking.ColumnHeaders(1).Text = "Payor"
    lstBordxTracking.ColumnHeaders(2).Text = "Bordx No"
    lstBordxTracking.ColumnHeaders(2).Alignment = lvwColumnLeft
    lstBordxTracking.ColumnHeaders(3).Text = "Bordx Date"
    lstBordxTracking.ColumnHeaders(3).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(4).Text = "Bordx Amount"
    lstBordxTracking.ColumnHeaders(4).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(5).Text = "Bordx Status"
    lstBordxTracking.ColumnHeaders(5).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(6).Text = "CLM Status"
    lstBordxTracking.ColumnHeaders(6).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(7).Text = "Date Sent to Payor"
    lstBordxTracking.ColumnHeaders(7).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(8).Text = "Date Returned Fr Payor"
    lstBordxTracking.ColumnHeaders(8).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(9).Text = "Date Sent to MNRB"
    lstBordxTracking.ColumnHeaders(9).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(10).Text = "Date Sent to A/C"
    lstBordxTracking.ColumnHeaders(10).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(11).Text = "A/C Period"
    lstBordxTracking.ColumnHeaders(11).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(12).Text = "Receipt No (MNRB)"
    lstBordxTracking.ColumnHeaders(12).Alignment = lvwColumnLeft
    lstBordxTracking.ColumnHeaders(13).Text = "Receipt Date (MNRB)"
    lstBordxTracking.ColumnHeaders(13).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(14).Text = "Receipt A/C Period (MNRB)"
    lstBordxTracking.ColumnHeaders(14).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(15).Text = "Receipt No (Payor)"
    lstBordxTracking.ColumnHeaders(15).Alignment = lvwColumnLeft
    lstBordxTracking.ColumnHeaders(16).Text = "Receipt Date (Payor)"
    lstBordxTracking.ColumnHeaders(16).Alignment = lvwColumnCenter
    lstBordxTracking.ColumnHeaders(17).Text = "Receipt A/C Period (Payor)"
    lstBordxTracking.ColumnHeaders(17).Alignment = lvwColumnRight
    
End Sub

Private Sub Option3()
    CboHospital.Enabled = True
    CboHLTCode.Enabled = False
    CboPayor.Enabled = False
    TxtBordx1.Enabled = False
    TxtBordx2.Enabled = False
    ChkBordxNo.Enabled = False
    txtBordxDate1.Enabled = False
    txtBordxDate2.Enabled = False
    ChkBordxDate.Enabled = False
    TxtBordxAmt1.Enabled = False
    TxtBordxAmt2.Enabled = False
    ChkBordxAmt.Enabled = False
    TxtACPeriod1.Enabled = False
    TxtACPeriod2.Enabled = False
    TxtDTP1.Enabled = False
    TxtDTP2.Enabled = False
    TxtDFP1.Enabled = False
    TxtDFP2.Enabled = False
    TxtDTMNRB1.Enabled = False
    TxtDTMNRB2.Enabled = False
    ChkACPeriod.Enabled = False
    ChkDTP.Enabled = False
    ChkDFP.Enabled = False
    ChkDTMNRB.Enabled = False
    ChkDTAC.Enabled = False
    Label17.Caption = "Clm Open Date"
    lstBordxTracking.ListItems.Clear
    
    lstBordxTracking.ColumnHeaders(1).Text = "Hospital"
    lstBordxTracking.ColumnHeaders(2).Text = "No.Inv <31"
    lstBordxTracking.ColumnHeaders(2).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(3).Text = "Act Amt <31"
    lstBordxTracking.ColumnHeaders(3).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(4).Text = "App Amt <31"
    lstBordxTracking.ColumnHeaders(4).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(5).Text = "No.Inv <61"
    lstBordxTracking.ColumnHeaders(5).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(6).Text = "Act Amt <61"
    lstBordxTracking.ColumnHeaders(6).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(7).Text = "App Amt <61"
    lstBordxTracking.ColumnHeaders(7).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(8).Text = "No.Inv <91"
    lstBordxTracking.ColumnHeaders(8).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(9).Text = "Act Amt <91"
    lstBordxTracking.ColumnHeaders(9).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(10).Text = "App Amt <91"
    lstBordxTracking.ColumnHeaders(10).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(11).Text = "No.Inv >91"
    lstBordxTracking.ColumnHeaders(11).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(12).Text = "Act Amt >91"
    lstBordxTracking.ColumnHeaders(12).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(13).Text = "App Amt >91"
    lstBordxTracking.ColumnHeaders(13).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(14).Text = "Total No.Inv"
    lstBordxTracking.ColumnHeaders(14).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(15).Text = "Total Act Amt"
    lstBordxTracking.ColumnHeaders(15).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(16).Text = "Total App Amt"
    lstBordxTracking.ColumnHeaders(16).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(17).Text = ""
    lstBordxTracking.ColumnHeaders(17).Alignment = lvwColumnRight
End Sub

Private Sub Option4()
    CboHospital.Enabled = False
    CboHLTCode.Enabled = False
    CboPayor.Enabled = True
    TxtBordx1.Enabled = False
    TxtBordx2.Enabled = False
    ChkBordxNo.Enabled = False
    txtBordxDate1.Enabled = False
    txtBordxDate2.Enabled = False
    ChkBordxDate.Enabled = False
    TxtBordxAmt1.Enabled = False
    TxtBordxAmt2.Enabled = False
    ChkBordxAmt.Enabled = False
    TxtACPeriod1.Enabled = False
    TxtACPeriod2.Enabled = False
    TxtDTP1.Enabled = False
    TxtDTP2.Enabled = False
    TxtDFP1.Enabled = False
    TxtDFP2.Enabled = False
    TxtDTMNRB1.Enabled = False
    TxtDTMNRB2.Enabled = False
    ChkACPeriod.Enabled = False
    ChkDTP.Enabled = False
    ChkDFP.Enabled = False
    ChkDTMNRB.Enabled = False
    ChkDTAC.Enabled = False
    Label17.Caption = "Clm Open Date"
    lstBordxTracking.ListItems.Clear
    
    lstBordxTracking.ColumnHeaders(1).Text = "Payor"
    lstBordxTracking.ColumnHeaders(2).Text = "No.Inv <31"
    lstBordxTracking.ColumnHeaders(2).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(3).Text = "Act Amt <31"
    lstBordxTracking.ColumnHeaders(3).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(4).Text = "App Amt <31"
    lstBordxTracking.ColumnHeaders(4).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(5).Text = "No.Inv <61"
    lstBordxTracking.ColumnHeaders(5).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(6).Text = "Act Amt <61"
    lstBordxTracking.ColumnHeaders(6).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(7).Text = "App Amt <61"
    lstBordxTracking.ColumnHeaders(7).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(8).Text = "No.Inv <91"
    lstBordxTracking.ColumnHeaders(8).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(9).Text = "Act Amt <91"
    lstBordxTracking.ColumnHeaders(9).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(10).Text = "App Amt <91"
    lstBordxTracking.ColumnHeaders(10).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(11).Text = "No.Inv >91"
    lstBordxTracking.ColumnHeaders(11).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(12).Text = "Act Amt >91"
    lstBordxTracking.ColumnHeaders(12).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(13).Text = "App Amt >91"
    lstBordxTracking.ColumnHeaders(13).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(14).Text = "Total No.Inv"
    lstBordxTracking.ColumnHeaders(14).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(15).Text = "Total Act Amt"
    lstBordxTracking.ColumnHeaders(15).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(16).Text = "Total App Amt"
    lstBordxTracking.ColumnHeaders(16).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(17).Text = ""
    lstBordxTracking.ColumnHeaders(17).Alignment = lvwColumnRight
End Sub

Private Sub Option5()
    CboHospital.Enabled = True
    CboHLTCode.Enabled = False
    CboPayor.Enabled = False
    TxtBordx1.Enabled = False
    TxtBordx2.Enabled = False
    ChkBordxNo.Enabled = False
    txtBordxDate1.Enabled = False
    txtBordxDate2.Enabled = False
    ChkBordxDate.Enabled = False
    TxtBordxAmt1.Enabled = False
    TxtBordxAmt2.Enabled = False
    ChkBordxAmt.Enabled = False
    TxtACPeriod1.Enabled = False
    TxtACPeriod2.Enabled = False
    TxtDTP1.Enabled = False
    TxtDTP2.Enabled = False
    TxtDFP1.Enabled = False
    TxtDFP2.Enabled = False
    TxtDTMNRB1.Enabled = False
    TxtDTMNRB2.Enabled = False
    ChkACPeriod.Enabled = False
    ChkDTP.Enabled = False
    ChkDFP.Enabled = False
    ChkDTMNRB.Enabled = False
    ChkDTAC.Enabled = False
    Label17.Caption = "Clm Open Date"
    lstBordxTracking.ListItems.Clear
    
    lstBordxTracking.ColumnHeaders(1).Text = "Hospital"
    lstBordxTracking.ColumnHeaders(2).Text = "Inv DOR Month"
    lstBordxTracking.ColumnHeaders(2).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(3).Text = "No of Inv"
    lstBordxTracking.ColumnHeaders(3).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(4).Text = "Bill Amt"
    lstBordxTracking.ColumnHeaders(4).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(5).Text = "App Amt"
    lstBordxTracking.ColumnHeaders(5).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(6).Text = ""
    lstBordxTracking.ColumnHeaders(6).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(7).Text = ""
    lstBordxTracking.ColumnHeaders(7).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(8).Text = ""
    lstBordxTracking.ColumnHeaders(8).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(9).Text = ""
    lstBordxTracking.ColumnHeaders(9).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(10).Text = ""
    lstBordxTracking.ColumnHeaders(10).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(11).Text = ""
    lstBordxTracking.ColumnHeaders(11).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(12).Text = ""
    lstBordxTracking.ColumnHeaders(12).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(13).Text = ""
    lstBordxTracking.ColumnHeaders(13).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(14).Text = ""
    lstBordxTracking.ColumnHeaders(14).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(15).Text = ""
    lstBordxTracking.ColumnHeaders(15).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(16).Text = ""
    lstBordxTracking.ColumnHeaders(16).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(17).Text = ""
    lstBordxTracking.ColumnHeaders(17).Alignment = lvwColumnRight
End Sub

Private Sub Option6()
    CboHospital.Enabled = False
    CboHLTCode.Enabled = False
    CboPayor.Enabled = True
    TxtBordx1.Enabled = False
    TxtBordx2.Enabled = False
    ChkBordxNo.Enabled = False
    txtBordxDate1.Enabled = False
    txtBordxDate2.Enabled = False
    ChkBordxDate.Enabled = False
    TxtBordxAmt1.Enabled = False
    TxtBordxAmt2.Enabled = False
    ChkBordxAmt.Enabled = False
    TxtACPeriod1.Enabled = False
    TxtACPeriod2.Enabled = False
    TxtDTP1.Enabled = False
    TxtDTP2.Enabled = False
    TxtDFP1.Enabled = False
    TxtDFP2.Enabled = False
    TxtDTMNRB1.Enabled = False
    TxtDTMNRB2.Enabled = False
    ChkACPeriod.Enabled = False
    ChkDTP.Enabled = False
    ChkDFP.Enabled = False
    ChkDTMNRB.Enabled = False
    ChkDTAC.Enabled = False
    Label17.Caption = "Clm Open Date"
    lstBordxTracking.ListItems.Clear
    
    lstBordxTracking.ColumnHeaders(1).Text = "Payor"
    lstBordxTracking.ColumnHeaders(2).Text = "Inv DOR Month"
    lstBordxTracking.ColumnHeaders(2).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(3).Text = "No of Inv"
    lstBordxTracking.ColumnHeaders(3).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(4).Text = "Bill Amt"
    lstBordxTracking.ColumnHeaders(4).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(5).Text = "App Amt"
    lstBordxTracking.ColumnHeaders(5).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(6).Text = ""
    lstBordxTracking.ColumnHeaders(6).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(7).Text = ""
    lstBordxTracking.ColumnHeaders(7).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(8).Text = ""
    lstBordxTracking.ColumnHeaders(8).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(9).Text = ""
    lstBordxTracking.ColumnHeaders(9).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(10).Text = ""
    lstBordxTracking.ColumnHeaders(10).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(11).Text = ""
    lstBordxTracking.ColumnHeaders(11).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(12).Text = ""
    lstBordxTracking.ColumnHeaders(12).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(13).Text = ""
    lstBordxTracking.ColumnHeaders(13).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(14).Text = ""
    lstBordxTracking.ColumnHeaders(14).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(15).Text = ""
    lstBordxTracking.ColumnHeaders(15).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(16).Text = ""
    lstBordxTracking.ColumnHeaders(16).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(17).Text = ""
    lstBordxTracking.ColumnHeaders(17).Alignment = lvwColumnRight
End Sub

Private Sub Option7()
    CboHospital.Enabled = False
    CboHLTCode.Enabled = True
    CboPayor.Enabled = True
    TxtBordx1.Enabled = True
    TxtBordx2.Enabled = True
    ChkBordxNo.Enabled = True
    txtBordxDate1.Enabled = True
    txtBordxDate2.Enabled = True
    ChkBordxDate.Enabled = True
    TxtBordxAmt1.Enabled = True
    TxtBordxAmt2.Enabled = True
    ChkBordxAmt.Enabled = True
    TxtACPeriod1.Enabled = True
    TxtACPeriod2.Enabled = True
    TxtDTP1.Enabled = True
    TxtDTP2.Enabled = True
    TxtDFP1.Enabled = True
    TxtDFP2.Enabled = True
    TxtDTMNRB1.Enabled = True
    TxtDTMNRB2.Enabled = True
    ChkACPeriod.Enabled = True
    ChkDTP.Enabled = True
    ChkDFP.Enabled = True
    ChkDTMNRB.Enabled = True
    Label17.Caption = "Date to A/C"
            
    lstBordxTracking.ListItems.Clear
    
    lstBordxTracking.ColumnHeaders(1).Text = "Payor"
    lstBordxTracking.ColumnHeaders(2).Text = "No of Bordx"
    lstBordxTracking.ColumnHeaders(2).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(3).Text = "Approved Amt"
    lstBordxTracking.ColumnHeaders(3).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(4).Text = "Case Mgmt Amt"
    lstBordxTracking.ColumnHeaders(4).Alignment = lvwColumnRight
    lstBordxTracking.ColumnHeaders(5).Text = "Total Bordx Amt"
    lstBordxTracking.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For CC = 6 To 17
        lstBordxTracking.ColumnHeaders(CC).Text = ""
    Next CC
End Sub

Private Sub TxtBordx1_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtBordx1_LostFocus()
    TxtBordx1 = ChkStr(TxtBordx1)
    CmdSearch.Default = False
End Sub

Private Sub TxtBordxAmt1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtBordxAmt2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtBordxDate1_LostFocus()
    If IsDate(txtBordxDate1) Then
    Else
        If txtBordxDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtBordxDate2_LostFocus()
    If IsDate(txtBordxDate2) Then
    Else
        If txtBordxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDTP1_LostFocus()
    If IsDate(TxtDTP1) Then
    Else
        If TxtDTP1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDTP1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDTP2_LostFocus()
    If IsDate(TxtDTP2) Then
    Else
        If TxtDTP2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDTP2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDFP1_LostFocus()
    If IsDate(TxtDFP1) Then
    Else
        If TxtDFP1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDFP1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDFP2_LostFocus()
    If IsDate(TxtDFP2) Then
    Else
        If TxtDFP2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDFP2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDTMNRB1_LostFocus()
    If IsDate(TxtDTMNRB1) Then
    Else
        If TxtDTMNRB1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDTMNRB1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDTMNRB2_LostFocus()
    If IsDate(TxtDTMNRB2) Then
    Else
        If TxtDTMNRB2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDTMNRB2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDTAC1_LostFocus()
    If IsDate(TxtDTAC1) Then
    Else
        If TxtDTAC1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDTAC1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDTAC2_LostFocus()
    If IsDate(TxtDTAC2) Then
    Else
        If TxtDTAC2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDTAC2.SetFocus
        End If
    End If
End Sub

Private Sub cboPayor_Change()
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

Private Sub cboReportList_Change()
If cboReportList = "" Then
    If InStr(cboReportList.Text, "null") Then
       cboReportList.Enabled = False
    End If
End If
End Sub

Private Sub cboReportList_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboReportList.Text, cboReportList.SelStart) + Chr(KeyAscii) + Mid(cboReportList.Text, cboReportList.SelStart + cboReportList.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboReportList.ListCount - 1
 
        If NewText = LCase(Left(cboReportList.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboReportList.List(i)
        End If
    Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboReportList.Text = ValidValue
            cboReportList.SelStart = 0
            cboReportList.SelLength = Len(ValidValue)
        End If
    End If
End Sub
Private Sub cboHLTCode_Change()
If CboHLTCode.Text = "" Then
    If InStr(CboHLTCode.Text, "null") Then
       CboHLTCode.Enabled = False
    End If
End If
End Sub

Private Sub cboHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboHLTCode.Text, CboHLTCode.SelStart) + Chr(KeyAscii) + Mid(CboHLTCode.Text, CboHLTCode.SelStart + CboHLTCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboHLTCode.ListCount - 1
 
        If NewText = LCase(Left(CboHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboHLTCode.List(i)
        End If
    Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboHLTCode.Text = ValidValue
            CboHLTCode.SelStart = 0
            CboHLTCode.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CountHOSInvNotBordx31(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend1 As String
Dim Record As ListItem
    
    NoInv = 0
    ActAmt = 0
    AppAmt = 0
           
    strAppendCase = "8"
    strAppend1 = " AND HOSName = '" & Trim(Hospital) & "' AND NoofDay < 31 "
    strAppend1 = strAppend1 + strappend
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend1)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    NoInv = IIf(IsNumeric(rst2!NoInv) = True, rst2!NoInv, 0)
    ActAmt = IIf(IsNumeric(rst2!ActAmt) = True, rst2!ActAmt, 0)
    AppAmt = IIf(IsNumeric(rst2!AppAmt) = True, rst2!AppAmt, 0)
                
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountHOSInvNotBordx61(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend1 As String
Dim Record As ListItem
    
    NoInv2 = 0
    ActAmt2 = 0
    AppAmt2 = 0
          
    strAppendCase = "9"
    strAppend1 = " AND HOSName = '" & Trim(Hospital) & "' AND NoofDay between 31 AND 61"
    strAppend1 = strAppend1 + strappend
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend1)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    NoInv2 = IIf(IsNumeric(rst2!NoInv) = True, rst2!NoInv, 0)
    ActAmt2 = IIf(IsNumeric(rst2!ActAmt) = True, rst2!ActAmt, 0)
    AppAmt2 = IIf(IsNumeric(rst2!AppAmt) = True, rst2!AppAmt, 0)
                
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountHOSInvNotBordx91(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend1 As String
Dim Record As ListItem
    
    NoInv3 = 0
    ActAmt3 = 0
    AppAmt3 = 0
           
    strAppendCase = "10"
    strAppend1 = " AND HOSName = '" & Trim(Hospital) & "' AND NoofDay between 61 AND 91"
    strAppend1 = strAppend1 + strappend
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    NoInv3 = IIf(IsNumeric(rst2!NoInv) = True, rst2!NoInv, 0)
    ActAmt3 = IIf(IsNumeric(rst2!ActAmt) = True, rst2!ActAmt, 0)
    AppAmt3 = IIf(IsNumeric(rst2!AppAmt) = True, rst2!AppAmt, 0)
                
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountHOSInvNotBordx912(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend1 As String
Dim Record As ListItem
    
    NoInv4 = 0
    ActAmt4 = 0
    AppAmt4 = 0
           
    strAppendCase = "11"
    strAppend1 = " AND HOSName = '" & Trim(Hospital) & "' AND NoofDay > 91"
    strAppend1 = strAppend1 + strappend
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    NoInv4 = IIf(IsNumeric(rst2!NoInv) = True, rst2!NoInv, 0)
    ActAmt4 = IIf(IsNumeric(rst2!ActAmt) = True, rst2!ActAmt, 0)
    AppAmt4 = IIf(IsNumeric(rst2!AppAmt) = True, rst2!AppAmt, 0)
                
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountMBMInvNotBordx31(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend1 As String
Dim Record As ListItem
    
    NoInv = 0
    ActAmt = 0
    AppAmt = 0
    
    strAppendCase = "12"
    strAppend1 = " AND NoofDay < 31 AND Payor = '" & Trim(Payor) & "'"
    strAppend1 = strAppend1 + stappend
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    NoInv = IIf(IsNumeric(rst2!NoInv) = True, rst2!NoInv, 0)
    ActAmt = IIf(IsNumeric(rst2!ActAmt) = True, rst2!ActAmt, 0)
    AppAmt = IIf(IsNumeric(rst2!AppAmt) = True, rst2!AppAmt, 0)
                
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountMBMInvNotBordx61(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend1 As String
Dim Record As ListItem
    
    NoInv2 = 0
    ActAmt2 = 0
    AppAmt2 = 0
           
    strAppendCase = "13"
    strAppend1 = " AND NoofDay Between 31 AND 61 AND Payor = '" & Trim(Payor) & "'"
    strAppend1 = strAppend1 + strappend
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    NoInv2 = IIf(IsNumeric(rst2!NoInv) = True, rst2!NoInv, 0)
    ActAmt2 = IIf(IsNumeric(rst2!ActAmt) = True, rst2!ActAmt, 0)
    AppAmt2 = IIf(IsNumeric(rst2!AppAmt) = True, rst2!AppAmt, 0)
                
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountMBMInvNotBordx91(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend1 As String
Dim Record As ListItem
    
    NoInv3 = 0
    ActAmt3 = 0
    AppAmt3 = 0
       
    strAppendCase = "14"
    strAppend1 = " AND NoofDay Between 61 AND 91 AND Payor = '" & Trim(Payor) & "'"
    strAppend1 = strAppend1 + strappend
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    NoInv3 = IIf(IsNumeric(rst2!NoInv) = True, rst2!NoInv, 0)
    ActAmt3 = IIf(IsNumeric(rst2!ActAmt) = True, rst2!ActAmt, 0)
    AppAmt3 = IIf(IsNumeric(rst2!AppAmt) = True, rst2!AppAmt, 0)
                
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub CountMBMInvNotBordx912(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend1 As String
Dim Record As ListItem
    
    NoInv4 = 0
    ActAmt4 = 0
    AppAmt4 = 0
           
    strAppendCase = "15"
    strAppend1 = " AND NoofDay > 91 AND Payor = '" & Trim(Payor) & "'"
    strAppend1 = strAppend1 + strappend
    
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
    NoInv4 = IIf(IsNumeric(rst2!NoInv) = True, rst2!NoInv, 0)
    ActAmt4 = IIf(IsNumeric(rst2!ActAmt) = True, rst2!ActAmt, 0)
    AppAmt4 = IIf(IsNumeric(rst2!AppAmt) = True, rst2!AppAmt, 0)
                
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub BordxTrackingByPayor(Optional strappend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim CASMgmtFees As Variant
    
    Total = 0
    Current = 0
    CASMgmtFees = 0
    lstBordxTracking.ListItems.Clear
    strAppendCase = "7"
    
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxTracking"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strappend)
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
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
                
            With rst2
                
                If Len(rst2!BordxPAYCode) Then
                    Set Record = lstBordxTracking.ListItems.Add(, , rst2!BordxPAYCode)
                Else
                    Set Record = lstBordxTracking.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!TotalCase) Then
                    Record.SubItems(1) = rst2!TotalCase
                Else
                    Record.SubItems(1) = "0"
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(2) = Format(rst2!BordxClaimAmt, "#,##0.00")
                Else
                    Record.SubItems(2) = "0.00"
                End If
                
                If Len(rst2!CLMCaseMgmtFees) Then
                    Record.SubItems(3) = Format(rst2!CLMCaseMgmtFees, "#,##0.00")
                Else
                    Record.SubItems(3) = "0.00"
                End If
                
                'Record.SubItems(4) = Format(CDbl(Record.SubItems(2)) + CDbl(Record.SubItems(3)), "#,##0.00")
                '26/04/2005 - TAKE OUT THE MgntFee with Acc Shirley advice
                Record.SubItems(4) = Format(CDbl(Record.SubItems(2)), "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!CLMCaseMgmtFees) Then
                    CASMgmtFees = CASMgmtFees + rst2!CLMCaseMgmtFees
                End If
                
                If Len(rst2!BordxClaimAmt) Then
                    TotalAmt = TotalAmt + rst2!BordxClaimAmt '+ rst2!CLMCaseMgmtFees
                End If
                LblTotalAmt = Format(CDbl(TotalAmt) + CDbl(CASMgmtFees), "#,##0.00")
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
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub


