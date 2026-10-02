VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBordxControl 
   Caption         =   "Bordx Control"
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
      TabIndex        =   35
      Top             =   7200
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3960
         TabIndex        =   58
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   57
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9600
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6960
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7800
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtTotalDay 
         Height          =   285
         Left            =   5520
         TabIndex        =   36
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   6840
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   8640
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtNoofDay2 
         Height          =   255
         Left            =   5880
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNoofDay 
         Height          =   255
         Left            =   4080
         TabIndex        =   38
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1560
         TabIndex        =   42
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   960
         TabIndex        =   41
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1320
         TabIndex        =   40
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1530
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   11895
      Begin VB.CheckBox Check2 
         Caption         =   "All Non-Sihat"
         Height          =   255
         Left            =   3000
         TabIndex        =   55
         Top             =   120
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "All Sihat"
         Height          =   255
         Left            =   1320
         TabIndex        =   54
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   390
         Width           =   3255
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Caption         =   "To MNRB"
         Height          =   1275
         Left            =   8280
         TabIndex        =   31
         Top             =   120
         Width           =   1575
         Begin VB.OptionButton optBoth2 
            Caption         =   "Both"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton ChkNotMNRB 
            Caption         =   "Not Yet Sent"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton ChkMNRB 
            Caption         =   "Sent"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "To Payor"
         Height          =   1275
         Left            =   4680
         TabIndex        =   30
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton optBoth 
            Caption         =   "Both"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton ChkNotSent 
            Caption         =   "Not Yet Sent"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton ChkSent 
            Caption         =   "Sent"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "From Payor (Doc)"
         Height          =   1275
         Left            =   6480
         TabIndex        =   28
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton optBoth1 
            Caption         =   "Both"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Frame Frame5 
            Caption         =   "Frame5"
            Height          =   615
            Left            =   1800
            TabIndex        =   29
            Top             =   120
            Width           =   15
         End
         Begin VB.OptionButton ChkPartial1 
            Caption         =   "Not Yet Received"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton ChkPaid1 
            Caption         =   "Received"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "From MNRB (Money)"
         Height          =   1275
         Left            =   9960
         TabIndex        =   26
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton ChkPaid2 
            Caption         =   "Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton ChkPartial2 
            Caption         =   "Partial Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton ChkUnpaid2 
            Caption         =   "Unpaid"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1455
         End
         Begin VB.Frame Frame7 
            Caption         =   "Frame5"
            Height          =   615
            Left            =   1800
            TabIndex        =   27
            Top             =   120
            Width           =   15
         End
         Begin VB.OptionButton optBoth3 
            Caption         =   "Both"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin MSMask.MaskEdBox txtBordxDate2 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Top             =   1020
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDate1 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1020
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame10 
         Caption         =   "From Payor (Money)"
         Height          =   1215
         Left            =   10080
         TabIndex        =   48
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton optBoth4 
            Caption         =   "Both"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   960
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame5"
            Height          =   615
            Left            =   1800
            TabIndex        =   52
            Top             =   120
            Width           =   15
         End
         Begin VB.OptionButton ChkUnpaid3 
            Caption         =   "Unpaid"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton ChkPartial3 
            Caption         =   "Partial Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton ChkPaid3 
            Caption         =   "Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   47
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   795
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No from"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   495
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lstBordxsihat 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9128
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bordx No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "No of Claims"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Bordx Amt"
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
         Text            =   "DTP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "DRP (Doc)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "No of Days (DRP - DTP)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "DTMNRB"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "DFMNRB (Money)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "No of Days (Date Fr MNRB - Date to MNRB)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Total Days"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lstBordxNonsihat 
      Height          =   5175
      Left            =   0
      TabIndex        =   43
      Top             =   2040
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9128
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
         Text            =   "Bordx No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "No of Claims"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Bordx Amt"
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
         Text            =   "DTP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "DRP (Doc)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "No of Days (DRP - DTP)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "DFPayor (Money)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "No of Days (DTP - DPayment)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Bordx Control"
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
      Height          =   195
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label2 
      Caption         =   "Sihat"
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
      Left            =   120
      TabIndex        =   44
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Non-Sihat"
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
      Left            =   120
      TabIndex        =   45
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "FrmBordxControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Private NoofDay As String
Private NoofDay2 As String
Private TotalDay As String
Private plancode As String
Private GroupCompany As String
Public searchCriteria As String
Private m_blnDirection(1 To 15) As Boolean

Private Sub Check1_Click()
Dim AA As Integer

    Label2.Visible = True
    lstBordxNonsihat.Visible = False
    Frame3.Enabled = True
    Frame3.Visible = True
    Frame6.Enabled = True
    Frame4.Left = 4680
    Frame9.Left = 6600
    Frame6.Visible = True
    Label3.Visible = False
    lstBordxsihat.Visible = True
    lstBordxsihat.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    
    If Check1.Value = 1 Then
        Check2.Value = 0
        Check2.Enabled = False
        CboPayor.Clear
        CboPayor.Enabled = False
        
        lstBordxsihat.ColumnHeaders(1).Text = "Payor"
        lstBordxsihat.ColumnHeaders(2).Text = "Bordx No"
        lstBordxsihat.ColumnHeaders(2).Alignment = lvwColumnLeft
        lstBordxsihat.ColumnHeaders(3).Text = "Bordx Date"
        lstBordxsihat.ColumnHeaders(3).Alignment = lvwColumnCenter
        lstBordxsihat.ColumnHeaders(4).Text = "No of Claims"
        lstBordxsihat.ColumnHeaders(4).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(5).Text = "Bordx Amt"
        lstBordxsihat.ColumnHeaders(5).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(6).Text = "Bordx Status"
        lstBordxsihat.ColumnHeaders(6).Alignment = lvwColumnCenter
        lstBordxsihat.ColumnHeaders(7).Text = "DTP"
        lstBordxsihat.ColumnHeaders(7).Alignment = lvwColumnCenter
        lstBordxsihat.ColumnHeaders(8).Text = "DRP (Doc)"
        lstBordxsihat.ColumnHeaders(8).Alignment = lvwColumnCenter
        lstBordxsihat.ColumnHeaders(9).Text = "No of Days (DRP - DTP)"
        lstBordxsihat.ColumnHeaders(9).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(10).Text = "DTMNRB"
        lstBordxsihat.ColumnHeaders(10).Alignment = lvwColumnCenter
        lstBordxsihat.ColumnHeaders(11).Text = "DFMNRB (Money)"
        lstBordxsihat.ColumnHeaders(11).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(12).Text = "No of Days (Date Fr MNRB - Date to MNRB)"
        lstBordxsihat.ColumnHeaders(12).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(13).Text = "Total Days"
        lstBordxsihat.ColumnHeaders(13).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(14).Text = "Receipt No"
        lstBordxsihat.ColumnHeaders(14).Alignment = lvwColumnLeft
        lstBordxsihat.ColumnHeaders(15).Text = "Receipt Date"
        lstBordxsihat.ColumnHeaders(15).Alignment = lvwColumnCenter
        
    Else
        Check2.Value = 0
        Check2.Enabled = True
        CboPayor.Enabled = True
        Call PayorComboListing
        
        lstBordxsihat.ColumnHeaders(1).Text = "Bordx No"
        lstBordxsihat.ColumnHeaders(2).Text = "Bordx Date"
        lstBordxsihat.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstBordxsihat.ColumnHeaders(3).Text = "No of Claims"
        lstBordxsihat.ColumnHeaders(3).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(4).Text = "Bordx Amt"
        lstBordxsihat.ColumnHeaders(4).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(5).Text = "Bordx Status"
        lstBordxsihat.ColumnHeaders(5).Alignment = lvwColumnCenter
        lstBordxsihat.ColumnHeaders(6).Text = "DTP"
        lstBordxsihat.ColumnHeaders(6).Alignment = lvwColumnCenter
        lstBordxsihat.ColumnHeaders(7).Text = "DRP (Doc)"
        lstBordxsihat.ColumnHeaders(7).Alignment = lvwColumnCenter
        lstBordxsihat.ColumnHeaders(8).Text = "No of Days (DRP - DTP)"
        lstBordxsihat.ColumnHeaders(8).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(9).Text = "DTMNRB"
        lstBordxsihat.ColumnHeaders(9).Alignment = lvwColumnCenter
        lstBordxsihat.ColumnHeaders(10).Text = "DFMNRB (Money)"
        lstBordxsihat.ColumnHeaders(10).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(11).Text = "No of Days (Date Fr MNRB - Date to MNRB)"
        lstBordxsihat.ColumnHeaders(11).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(12).Text = "Total Days"
        lstBordxsihat.ColumnHeaders(12).Alignment = lvwColumnRight
        lstBordxsihat.ColumnHeaders(13).Text = "Receipt No"
        lstBordxsihat.ColumnHeaders(13).Alignment = lvwColumnLeft
        lstBordxsihat.ColumnHeaders(14).Text = "Receipt Date"
        lstBordxsihat.ColumnHeaders(14).Alignment = lvwColumnCenter
        
        For AA = 15 To 15
           lstBordxsihat.ColumnHeaders(AA).Text = ""
        Next AA
        
    End If
End Sub

Private Sub Check2_Click()
Dim AA As Integer

    Label3.Visible = True
    Label2.Visible = False
    Frame3.Enabled = False
    Frame3.Visible = False
    Frame6.Visible = False
    Frame9.Enabled = True
    Frame4.Left = 6500
    Frame9.Left = 8400
    lstBordxsihat.Visible = False
    lstBordxNonsihat.Visible = True
    lstBordxNonsihat.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    
    If Check2.Value = 1 Then
        Check1.Value = 0
        Check1.Enabled = False
        CboPayor.Clear
        CboPayor.Enabled = False
        
        lstBordxNonsihat.ColumnHeaders(1).Text = "Payor"
        lstBordxNonsihat.ColumnHeaders(2).Text = "Bordx No"
        lstBordxNonsihat.ColumnHeaders(2).Alignment = lvwColumnLeft
        lstBordxNonsihat.ColumnHeaders(3).Text = "Bordx Date"
        lstBordxNonsihat.ColumnHeaders(3).Alignment = lvwColumnCenter
        lstBordxNonsihat.ColumnHeaders(4).Text = "No of Claims"
        lstBordxNonsihat.ColumnHeaders(4).Alignment = lvwColumnRight
        lstBordxNonsihat.ColumnHeaders(5).Text = "Bordx Amt"
        lstBordxNonsihat.ColumnHeaders(5).Alignment = lvwColumnRight
        lstBordxNonsihat.ColumnHeaders(6).Text = "Bordx Status"
        lstBordxNonsihat.ColumnHeaders(6).Alignment = lvwColumnCenter
        lstBordxNonsihat.ColumnHeaders(7).Text = "DTP"
        lstBordxNonsihat.ColumnHeaders(7).Alignment = lvwColumnCenter
        lstBordxNonsihat.ColumnHeaders(8).Text = "DRP (Doc)"
        lstBordxNonsihat.ColumnHeaders(8).Alignment = lvwColumnCenter
        lstBordxNonsihat.ColumnHeaders(9).Text = "No of Days (DRP - DTP)"
        lstBordxNonsihat.ColumnHeaders(9).Alignment = lvwColumnRight
        lstBordxNonsihat.ColumnHeaders(10).Text = "DFPayor (Money)"
        lstBordxNonsihat.ColumnHeaders(10).Alignment = lvwColumnCenter
        lstBordxNonsihat.ColumnHeaders(11).Text = "No of Days (DTP - DPayment)"
        lstBordxNonsihat.ColumnHeaders(11).Alignment = lvwColumnRight
        lstBordxNonsihat.ColumnHeaders(12).Text = "Receipt No"
        lstBordxNonsihat.ColumnHeaders(12).Alignment = lvwColumnLeft
        lstBordxNonsihat.ColumnHeaders(13).Text = "Receipt Date"
        lstBordxNonsihat.ColumnHeaders(13).Alignment = lvwColumnCenter
        
    Else
        Check1.Value = 0
        Check1.Enabled = True
        CboPayor.Enabled = True
        Call PayorComboListing
        
        lstBordxNonsihat.ColumnHeaders(1).Text = "Bordx No"
        lstBordxNonsihat.ColumnHeaders(2).Text = "Bordx Date"
        lstBordxNonsihat.ColumnHeaders(2).Alignment = lvwColumnCenter
        lstBordxNonsihat.ColumnHeaders(3).Text = "No of Claims"
        lstBordxNonsihat.ColumnHeaders(3).Alignment = lvwColumnRight
        lstBordxNonsihat.ColumnHeaders(4).Text = "Bordx Amt"
        lstBordxNonsihat.ColumnHeaders(4).Alignment = lvwColumnRight
        lstBordxNonsihat.ColumnHeaders(5).Text = "Bordx Status"
        lstBordxNonsihat.ColumnHeaders(5).Alignment = lvwColumnCenter
        lstBordxNonsihat.ColumnHeaders(6).Text = "DTP"
        lstBordxNonsihat.ColumnHeaders(6).Alignment = lvwColumnCenter
        lstBordxNonsihat.ColumnHeaders(7).Text = "DRP (Doc)"
        lstBordxNonsihat.ColumnHeaders(7).Alignment = lvwColumnCenter
        lstBordxNonsihat.ColumnHeaders(8).Text = "No of Days (DRP - DTP)"
        lstBordxNonsihat.ColumnHeaders(8).Alignment = lvwColumnRight
        lstBordxNonsihat.ColumnHeaders(9).Text = "DFPayor (Money)"
        lstBordxNonsihat.ColumnHeaders(9).Alignment = lvwColumnCenter
        lstBordxNonsihat.ColumnHeaders(10).Text = "No of Days (DTP - DPayment)"
        lstBordxNonsihat.ColumnHeaders(10).Alignment = lvwColumnRight
        lstBordxNonsihat.ColumnHeaders(11).Text = "Receipt No"
        lstBordxNonsihat.ColumnHeaders(11).Alignment = lvwColumnLeft
        lstBordxNonsihat.ColumnHeaders(12).Text = "Receipt Date"
        lstBordxNonsihat.ColumnHeaders(12).Alignment = lvwColumnCenter
        
        For AA = 13 To 13
           lstBordxNonsihat.ColumnHeaders(AA).Text = ""
        Next AA
        
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

'Private Sub cmdPrint_Click()
'Dim strsqlPAY As String
'Dim PAY As New ADODB.Recordset
'Dim cmd As New ADODB.Command
'Dim prm1 As ADODB.Parameter
'Dim prm2 As ADODB.Parameter
'Dim strAppendCase As String
'Dim strAppend As String

    'strAppendCase = "1"
    'strAppend = " AND PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
    'Set cmd.ActiveConnection = medixmasterconn
    'cmd.CommandText = "BordxControl"
    'cmd.CommandType = adCmdStoredProc
    'cmd.CommandTimeout = 300
    'Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    'cmd.Parameters.Append prm1
    'Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    'cmd.Parameters.Append prm2
    'PAY.CursorLocation = adUseClient
    'PAY.CursorType = adOpenForwardOnly
    'PAY.CacheSize = 100
    'Set PAY = cmd.Execute
    
    'Screen.MousePointer = vbHourglass
        
    'If lstBordxsihat.listitems.count <> 0 Then
        'If PAY!HLTCode = "N" Then
            'RptBordxWks.show
        'Else
            'RptBordxWks.show
        'End If
    'Else
        'MsgBox "Report cannot be printed without any record.", vbInformation
    'End If
              
    'PAY.Close
    'Set PAY = Nothing
    'Screen.MousePointer = Default

'End Sub

Private Sub CmdPrintFile_Click()
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim conn As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    Screen.MousePointer = vbHourglass
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If Check1.Value = 1 Then
    
        'Call ExportToExcelBordxControlSihat(lstBordxsihat)
        If lstBordxsihat.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(lstBordxsihat)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
    
    ElseIf Check2.Value = 1 Then
    
        'Call ExportToExcelBordxControlNonSihat(lstBordxNonsihat)
        If lstBordxNonsihat.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(lstBordxNonsihat)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
    
    Else
    
        If CboPayor.Text <> "" Then
            'strsqlPAY = "Select PAYCode, HLTCode from PAY where PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
            'PAY.Open strsqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
            strAppendCase = "1"
            strAppend = " AND PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "BordxControl"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
            cmd.Parameters.Append prm1
            Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
            cmd.Parameters.Append prm2
            PAY.CursorLocation = adUseClient
            PAY.CursorType = adOpenForwardOnly
            PAY.CacheSize = 100
            Set PAY = cmd.Execute
            
                If PAY!HLTCode = "N" Then
                    'Call ExportToExcelBordxControlNonSihat(lstBordxNonsihat)
                    If lstBordxNonsihat.ListItems.Count <> 0 Then
                        Call ExportListViewtoExcel(lstBordxNonsihat)
                    Else
                        MsgBox "Report cannot be printed without any record.", vbInformation
                    End If
                Else
                    'Call ExportToExcelBordxControlSihat(lstBordxsihat)
                    If lstBordxsihat.ListItems.Count <> 0 Then
                        Call ExportListViewtoExcel(lstBordxsihat)
                    Else
                        MsgBox "Report cannot be printed without any record.", vbInformation
                    End If
                End If
            
                PAY.Close
                Set PAY = Nothing
        End If
    End If
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    Screen.MousePointer = Default
End Sub

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    'Call ComboListing
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call PayorComboListing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing

End Sub

'************************** Start Command Button **************************
Private Sub CmdSearch_Click()
On Error GoTo errRun

    ItemChecked = False
    LblTotalAmt = ""
    lblCurrent = ""
    lblTotal = ""
    CmdSearch.Enabled = False
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If Check1.Value = 1 And Check2.Enabled = False Then
    
        'If CboPayor = "" Then
            'MsgBox "Please Enter Payor Code.", vbCritical
            'CboPayor.SetFocus
        'Else
        CmdClear.Enabled = False
        CmdPrintFile.Enabled = False
        CmdExit.Enabled = False
            Screen.MousePointer = vbHourglass
                Call SearchRecord
            Screen.MousePointer = Default
        'End If
        
    ElseIf Check2.Value = 1 And Check1.Enabled = False Then
    
        'If CboPayor = "" Then
        '    MsgBox "Please Enter Payor Code.", vbCritical
        '    'CboPayor.SetFocus
        'Else
        CmdClear.Enabled = False
        CmdPrintFile.Enabled = False
        CmdExit.Enabled = False
            Screen.MousePointer = vbHourglass
                Call SearchRecord
            Screen.MousePointer = Default
        'End If
        
    Else
    
        If CboPayor = "" Then
            MsgBox "Please Enter Payor Code.", vbCritical
            CboPayor.SetFocus
        
        Else
            CmdClear.Enabled = False
            CmdPrintFile.Enabled = False
            CmdExit.Enabled = False
            Screen.MousePointer = vbHourglass
                Call SearchRecord
            Screen.MousePointer = Default
        End If
    End If
 '   medixmasterconnWS.Close
 '   Set medixmasterconnWS = Nothing
 '   medixmasterconn.Close
 '   Set medixmasterconn = Nothing
    CmdSearch.Enabled = True
    Exit Sub
    
errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstBordxsihat.ListItems.Clear
    lblCurrent = ""
    CmdPrintFile.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    'cmdPrint.Enabled = True
    CmdPrintFile.Enabled = True
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

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    CboPayor.Clear
    Call PayorComboListing
    CboPayor.SetFocus
    TotalDay = 0
    NoofDay = 0
    NoofDay2 = 0
    Frame9.Enabled = True
    Frame3.Enabled = True
    Frame6.Enabled = True
    'optBoth.value = True
    'optBoth1.value = True
    'optBoth2.value = True
    'optBoth3.value = True
    lstBordxsihat.ListItems.Clear
    lstBordxNonsihat.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
End Sub
'************************** End Command Button **************************

'************************** Start Search Record **************************
Private Function SearchRecord()
Dim StrSql As String
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim sql As String
Dim CAS As New ADODB.Recordset
Dim getSel As Boolean
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset

    If Check1.Value = 1 Then
    
        If Len(Trim(TxtBordx1)) Then
            strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
            strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
        
        If IsDate(txtBordxDate1) = True Then
            strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(txtBordxDate2) = True Then
            strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
        End If
        
        If ChkSent.Value = True Then
            strAppend = strAppend + " AND BordxPAYToStatus = '1'"
        End If
        
        If ChkNotSent.Value = True Then
            strAppend = strAppend + " AND BordxPAYToStatus = '0'"
        End If
        
        If ChkPaid1.Value = True Then
            strAppend = strAppend + " AND BordxPAYDOCFrStatus = '1'"
        End If
    
        If ChkPartial1.Value = True Then
            strAppend = strAppend + " AND BordxPAYDOCFrStatus = '0'"
        End If
        
        If ChkMNRB.Value = True Then
            strAppend = strAppend + " AND BordxMNRBToStatus = '1'"
        End If
    
        If ChkNotMNRB.Value = True Then
            strAppend = strAppend + " AND BordxMNRBToStatus = '0'"
        End If
        
        If optBoth2.Value = True Then
            strAppend = strAppend + " AND (BordxMNRBToStatus = '1' or BordxMNRBToStatus = '0')"
        End If
    
        If ChkPaid2.Value = True Then
            strAppend = strAppend + " AND BordxMNRBFrStatus = 'F'"
        End If
    
        If ChkPartial2.Value = True Then
            strAppend = strAppend + " AND BordxMNRBFrStatus = 'P'"
        End If
    
        If ChkUnpaid2.Value = True Then
            strAppend = strAppend + " AND (BordxMNRBFrStatus = '' or BordxMNRBFrStatus is null)"
        End If
        
        searchCriteria = strAppend
        
        Call SihatBordxControlListing2(strAppend)
    
    ElseIf Check2.Value = 1 Then
    
        If Len(Trim(TxtBordx1)) Then
            strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
            strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
        
        If IsDate(txtBordxDate1) = True Then
            strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(txtBordxDate2) = True Then
            strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
        End If
        
        If ChkSent.Value = True Then
            strAppend = strAppend + " AND BordxPAYToStatus = '1'"
        End If
        
        If ChkNotSent.Value = True Then
            strAppend = strAppend + " AND BordxPAYToStatus = '0'"
        End If
        
        If ChkPaid1.Value = True Then
            strAppend = strAppend + " AND BordxPAYDOCFrStatus = '1'"
        End If
    
        If ChkPartial1.Value = True Then
            strAppend = strAppend + " AND BordxPAYDOCFrStatus = '0'"
        End If
    
        If ChkPaid3.Value = True Then
            strAppend = strAppend + " AND BordxPAYFrStatus = 'F'"
        End If
    
        If ChkPartial3.Value = True Then
            strAppend = strAppend + " AND BordxPAYFrStatus = 'P'"
        End If
    
        If ChkUnpaid3.Value = True Then
            strAppend = strAppend + " AND (BordxPAYFrStatus = '' or BordxPAYFrStatus is null)"
        End If
            
        searchCriteria = strAppend
            
        Call NonSihatBordxControlListing2(strAppend)
    
    Else
    
        'strsqlPAY = "Select PAYCode, HLTCode from PAY where PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
        'PAY.Open strsqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "1"
        strAppend = " AND PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "BordxControl"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        PAY.CursorLocation = adUseClient
        PAY.CursorType = adOpenForwardOnly
        PAY.CacheSize = 100
        Set PAY = cmd.Execute
            
        strAppend = ""
        
        If Len(Trim(CboPayor)) Then
            strAppend = strAppend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
        End If
        
        If Len(Trim(TxtBordx1)) Then
            strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordx1)) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
            strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordx2)) & "'"
        End If
        
        If IsDate(txtBordxDate1) = True Then
            strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(txtBordxDate2) = True Then
            strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
        End If
        
        If ChkSent.Value = True Then
            strAppend = strAppend + " AND BordxPAYToStatus = '1'"
        End If
        
        If ChkNotSent.Value = True Then
            strAppend = strAppend + " AND BordxPAYToStatus = '0'"
        End If
        
        If PAY!HLTCode = "N" Then
            If ChkPaid1.Value = True Then
                strAppend = strAppend + " AND BordxPAYDOCFrStatus = '1'"
            End If
        
            If ChkPartial1.Value = True Then
                strAppend = strAppend + " AND BordxPAYDOCFrStatus = '0'"
            End If
        
            If ChkPaid3.Value = True Then
                strAppend = strAppend + " AND BordxPAYFrStatus = 'F'"
            End If
        
            If ChkPartial3.Value = True Then
                strAppend = strAppend + " AND BordxPAYFrStatus = 'P'"
            End If
        
            If ChkUnpaid3.Value = True Then
                strAppend = strAppend + " AND (BordxPAYFrStatus = '' or BordxPAYFrStatus is null)"
            End If
            
            searchCriteria = strAppend
            
            Call NonSihatBordxControlListing(strAppend)
        Else
            
            If ChkPaid1.Value = True Then
                strAppend = strAppend + " AND BordxPAYDOCFrStatus = '1'"
            End If
        
            If ChkPartial1.Value = True Then
                strAppend = strAppend + " AND BordxPAYDOCFrStatus = '0'"
            End If
            
            If ChkMNRB.Value = True Then
                strAppend = strAppend + " AND BordxMNRBToStatus = '1'"
            End If
        
            If ChkNotMNRB.Value = True Then
                strAppend = strAppend + " AND BordxMNRBToStatus = '0'"
            End If
            
            If optBoth2.Value = True Then
                strAppend = strAppend + " AND (BordxMNRBToStatus = '1' or BordxMNRBToStatus = '0')"
            End If
        
            If ChkPaid2.Value = True Then
                strAppend = strAppend + " AND BordxMNRBFrStatus = 'F'"
            End If
        
            If ChkPartial2.Value = True Then
                strAppend = strAppend + " AND BordxMNRBFrStatus = 'P'"
            End If
        
            If ChkUnpaid2.Value = True Then
                strAppend = strAppend + " AND (BordxMNRBFrStatus = '' or BordxMNRBFrStatus is null)"
            End If
            
            searchCriteria = strAppend
            
            Call SihatBordxControlListing(strAppend)
        End If
        PAY.Close
        Set PAY = Nothing
    End If
End Function
'************************** End Search Record **************************

'************************** Start Sihat Bordx Control Listing **************************
Private Sub SihatBordxControlListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    Total = 0
    Current = 0
    NoofDay = 0
    NoofDay2 = 0
    TotalDay = 0
    
    lstBordxsihat.ListItems.Clear
            
    'sql = " Select BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt, " & _
          " BordxPAYSentDate, BordxPAYRecDocDate, " & _
          " datediff(d, BordxPAYSentDate, BordxPAYRecDocDate) AS NoofDay, " & _
          " BordxMNRBSentDate, BordxMNRBRecDate, BordxStatus, " & _
          " datediff(d, BordxMNRBSentDate, BordxMNRBRecDate) AS NoofDay2 " & _
          " From VIEW_Bordx_ControlS2 where 0 = 0 "
        
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxControl"
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

    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
            
            With rst2
                Set Record = lstBordxsihat.ListItems.Add(, , rst2!BordxRefNo)
                If Len(rst2!BordxDate) Then
                    Record.SubItems(1) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxNoClaims) Then
                   Record.SubItems(2) = rst2!BordxNoClaims
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(3) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(!Bordxstatus) Then
                   Record.SubItems(4) = !Bordxstatus
                End If
                If Len(rst2!BordxPAYSentDate) Then
                    Record.SubItems(5) = Format(rst2!BordxPAYSentDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxPAYRecDocDate) Then
                    Record.SubItems(6) = Format(rst2!BordxPAYRecDocDate, "dd/mm/yyyy")
                End If
                
                NoofDay = 0
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(7) = "1"
                    NoofDay = Record.SubItems(7)
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(7) = rst2!NoofDay
                    NoofDay = Record.SubItems(7)
                End If
                If Len(rst2!BordxMNRBSentDate) Then
                    Record.SubItems(8) = Format(rst2!BordxMNRBSentDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxMNRBRecDate) Then
                    Record.SubItems(9) = Format(rst2!BordxMNRBRecDate, "dd/mm/yyyy")
                End If
                
                NoofDay2 = 0
                
                If rst2!NoofDay2 = "0" Then
                    Record.SubItems(10) = "1"
                    NoofDay2 = Record.SubItems(10)
                ElseIf Len(rst2!NoofDay2) Then
                    Record.SubItems(10) = rst2!NoofDay2
                    NoofDay2 = Record.SubItems(10)
                End If
                
                TotalDay = CDbl(NoofDay) + CDbl(NoofDay2)
                
                If Len(TotalDay) Then
                    Record.SubItems(11) = TotalDay
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
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    
End Sub

Private Sub SihatBordxControlListing2(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim sql As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount

    Total = 0
    Current = 0
    NoofDay = 0
    NoofDay2 = 0
    TotalDay = 0
    
    lstBordxsihat.ListItems.Clear
            
    'sql = " Select BordxPAYCode, BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt, " & _
          " BordxPAYSentDate, BordxPAYRecDocDate, " & _
          " datediff(d, BordxPAYSentDate, BordxPAYRecDocDate) AS NoofDay, " & _
          " BordxMNRBSentDate, BordxMNRBRecDate, BordxStatus, " & _
          " datediff(d, BordxMNRBSentDate, BordxMNRBRecDate) AS NoofDay2 " & _
          " From VIEW_Bordx_ControlS2 Where HLTCode = 'S' "
    
    
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If

    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "2"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxControl"
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

    If rst2.EOF = True Then
    'If rstCount!totalrecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        'Total = rstCount!totalrecord
        'lblTotal = Total
        
            With rst2
                Set Record = lstBordxsihat.ListItems.Add(, , rst2!BordxPAYCode)
                If Len(rst2!BordxRefNo) Then
                   Record.SubItems(1) = Trim(rst2!BordxRefNo)
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(2) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxNoClaims) Then
                   Record.SubItems(3) = rst2!BordxNoClaims
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(4) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(!Bordxstatus) Then
                   Record.SubItems(5) = !Bordxstatus
                End If
                If Len(rst2!BordxPAYSentDate) Then
                    Record.SubItems(6) = Format(rst2!BordxPAYSentDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxPAYRecDocDate) Then
                    Record.SubItems(7) = Format(rst2!BordxPAYRecDocDate, "dd/mm/yyyy")
                End If
                
                NoofDay = 0
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(8) = "1"
                    NoofDay = Record.SubItems(8)
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(8) = rst2!NoofDay
                    NoofDay = Record.SubItems(8)
                End If
                If Len(rst2!BordxMNRBSentDate) Then
                    Record.SubItems(9) = Format(rst2!BordxMNRBSentDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxMNRBRecDate) Then
                    Record.SubItems(10) = Format(rst2!BordxMNRBRecDate, "dd/mm/yyyy")
                End If
                
                NoofDay2 = 0
                
                If rst2!NoofDay2 = "0" Then
                    Record.SubItems(11) = "1"
                    NoofDay2 = Record.SubItems(11)
                ElseIf Len(rst2!NoofDay2) Then
                    Record.SubItems(11) = rst2!NoofDay2
                    NoofDay2 = Record.SubItems(11)
                End If
                
                TotalDay = CDbl(NoofDay) + CDbl(NoofDay2)
                
                If Len(TotalDay) Then
                    Record.SubItems(12) = TotalDay
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
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
    
End Sub
'************************** End Sihat Bordx Control Listing **************************

'************************** Start Non Sihat Bordx Control Listing **************************
Private Sub NonSihatBordxControlListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount As String
   
    Total = 0
    Current = 0
    TotalDay = 0
    NoofDay = 0
    NoofDay2 = 0
    plancode = ""
    GroupCompany = ""
    
    lstBordxNonsihat.ListItems.Clear
            
    'sql = " Select BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt, " & _
          " BordxPAYSentDate, BordxPAYRecDocDate, BordxPAYRecDate, BordxStatus, " & _
          " DateDiff(d, BordxPAYSentDate, BordxPAYRecDocDate) AS NoofDay, " & _
          " DateDiff(d, BordxPAYSentDate, BordxPAYRecDate) AS NoofDay2 " & _
          " From VIEW_Bordx_ControlN2 where 0 = 0 "
           
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "3"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxControl"
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
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    Else
    
    Do
        Do Until rst2.EOF
        
            With rst2
                Set Record = lstBordxNonsihat.ListItems.Add(, , !BordxRefNo)
                If Len(!BordxDate) Then
                    Record.SubItems(1) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                
                If Len(!BordxNoClaims) Then
                   Record.SubItems(2) = !BordxNoClaims
                End If
                
                If Len(!BordxClaimAmt) Then
                    Record.SubItems(3) = Format(!BordxClaimAmt, "#,##0.00")
                End If
                
                If Len(!Bordxstatus) Then
                   Record.SubItems(4) = !Bordxstatus
                End If
                
                If Len(!BordxPAYSentDate) Then
                    Record.SubItems(5) = Format(!BordxPAYSentDate, "dd/mm/yyyy")
                End If
                
                If Len(!BordxPAYRecDocDate) Then
                    Record.SubItems(6) = Format(!BordxPAYRecDocDate, "dd/mm/yyyy")
                End If
                
                NoofDay = 0
                
                If !NoofDay = "0" Then
                    Record.SubItems(7) = "1"
                    NoofDay = Record.SubItems(7)
                ElseIf Len(!NoofDay) Then
                    Record.SubItems(7) = !NoofDay
                    NoofDay = Record.SubItems(7)
                End If
                
                If Len(!BordxPAYRecDate) Then
                    Record.SubItems(8) = Format(!BordxPAYRecDate, "dd/mm/yyyy")
                End If
                
                NoofDay2 = 0
                
                If !NoofDay2 = "0" Then
                    Record.SubItems(9) = "1"
                    NoofDay2 = Record.SubItems(9)
                ElseIf Len(!NoofDay2) Then
                    Record.SubItems(9) = !NoofDay2
                    NoofDay2 = Record.SubItems(9)
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
    CmdClear.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub NonSihatBordxControlListing2(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim strCount As String
   
    Total = 0
    Current = 0
    TotalDay = 0
    NoofDay = 0
    NoofDay2 = 0
    plancode = ""
    GroupCompany = ""
    
    lstBordxNonsihat.ListItems.Clear
            
    'sql = " Select BordxPAYCode, BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt, " & _
          " BordxPAYSentDate, BordxPAYRecDocDate, BordxPAYRecDate, BordxStatus, " & _
          " DateDiff(d, BordxPAYSentDate, BordxPAYRecDocDate) AS NoofDay, " & _
          " DateDiff(d, BordxPAYSentDate, BordxPAYRecDate) AS NoofDay2 " & _
          " From VIEW_Bordx_ControlN2 Where 0 = 0 "
           
    'If Len(Trim(strAppend)) Then
    '    sql = sql + strAppend
    'End If
    
    'rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    strAppendCase = "4"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "BordxControl"
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
        
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdPrintFile.Enabled = True
        CmdExit.Enabled = True
    Else
    
    Do
        Do Until rst2.EOF
        
            With rst2
                Set Record = lstBordxNonsihat.ListItems.Add(, , !BordxPAYCode)
                                
                If Len(!BordxRefNo) Then
                   Record.SubItems(1) = Trim(!BordxRefNo)
                End If
                
                If Len(!BordxDate) Then
                    Record.SubItems(2) = Format(!BordxDate, "dd/mm/yyyy")
                End If
                
                If Len(!BordxNoClaims) Then
                   Record.SubItems(3) = !BordxNoClaims
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
                
                If Len(!BordxPAYRecDocDate) Then
                    Record.SubItems(7) = Format(!BordxPAYRecDocDate, "dd/mm/yyyy")
                End If
                
                NoofDay = 0
                
                If !NoofDay = "0" Then
                    Record.SubItems(8) = "1"
                    NoofDay = Record.SubItems(8)
                ElseIf Len(!NoofDay) Then
                    Record.SubItems(8) = !NoofDay
                    NoofDay = Record.SubItems(8)
                End If
                
                If Len(!BordxPAYRecDate) Then
                    Record.SubItems(9) = Format(!BordxPAYRecDate, "dd/mm/yyyy")
                End If
                
                NoofDay2 = 0
                
                If !NoofDay2 = "0" Then
                    Record.SubItems(10) = "1"
                    NoofDay2 = Record.SubItems(10)
                ElseIf Len(!NoofDay2) Then
                    Record.SubItems(10) = !NoofDay2
                    NoofDay2 = Record.SubItems(10)
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
    CmdClear.Enabled = True
    CmdPrintFile.Enabled = True
    CmdExit.Enabled = True
End Sub
'************************** End Non sihat Bordx Control Listing **************************

'************************** Start ComboListing **************************

Private Sub PayorComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
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
    
 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End Sub

'************************** End ComboListing **************************

Private Sub cboPayor_Click()
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    'strsqlPAY = "Select PAYCode, HLTCode from PAY where PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
    'PAY.Open strsqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    strAppendCase = "1"
    strAppend = " AND PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "BordxControl"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
       
        If PAY!HLTCode = "N" Then
            Label3.Visible = True
            Label2.Visible = False
            Frame3.Enabled = False
            Frame3.Visible = False
            Frame6.Visible = False
            Frame6.Enabled = False
            Frame4.Left = 6500
            Frame9.Left = 8400
            lstBordxsihat.Visible = False
            lstBordxNonsihat.Visible = True
            lstBordxNonsihat.ListItems.Clear
            lblCurrent = 0
            lblTotal = 0
            LblTotalAmt = 0
        Else
            Label2.Visible = True
            lstBordxNonsihat.Visible = False
            Label3.Visible = False
            Frame3.Visible = True
            Frame3.Enabled = True
            Frame4.Left = 4680
            Frame9.Left = 6600
            Frame6.Visible = True
            Frame6.Enabled = True
            lstBordxsihat.Visible = True
            lstBordxsihat.ListItems.Clear
            lblCurrent = 0
            lblTotal = 0
            LblTotalAmt = 0
        End If
        PAY.Close
        Set PAY = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
End Sub

Private Sub CboPayor_Change()
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If Len(Mid(CboPayor, 1, 2)) Then
        'strsqlPAY = "Select PAYCode, HLTCode from PAY where PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
        'PAY.Open strsqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
        strAppendCase = "1"
        strAppend = " AND PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "BordxControl"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        PAY.CursorLocation = adUseClient
        PAY.CursorType = adOpenForwardOnly
        PAY.CacheSize = 100
        Set PAY = cmd.Execute
        
        If InStr(CboPayor.Text, "null") Then
           CboPayor.Enabled = False
        End If
        
        If PAY!HLTCode = "N" Then
            Label3.Visible = True
            Label2.Visible = False
            Frame3.Enabled = False
            Frame3.Visible = False
            Frame6.Visible = False
            Frame9.Enabled = True
            Frame4.Left = 6500
            Frame9.Left = 8400
            lstBordxsihat.Visible = False
            lstBordxNonsihat.Visible = True
            lstBordxNonsihat.ListItems.Clear
            lblCurrent = 0
            lblTotal = 0
            LblTotalAmt = 0
        Else
            Label2.Visible = True
            lstBordxNonsihat.Visible = False
            Frame3.Enabled = True
            Frame3.Visible = True
            Frame6.Enabled = True
            Frame4.Left = 4680
            Frame9.Left = 6600
            Frame6.Visible = True
            Label3.Visible = False
            lstBordxsihat.Visible = True
            lstBordxsihat.ListItems.Clear
            lblCurrent = 0
            lblTotal = 0
            LblTotalAmt = 0
        End If
        PAY.Close
        Set PAY = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
    End If
End Sub

Private Sub lstBordxsihat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    If Check1.Value = 0 Then
    
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstBordxsihat, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
        Case 3
            SortListView lstBordxsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstBordxsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
        Case 6
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView lstBordxsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
        Case 9
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
        Case 10
            SortListView lstBordxsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView lstBordxsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView lstBordxsihat, ColumnHeader.Index, ldtString, m_blnDirection(12)
        Case 13
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        End Select
    Else
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstBordxsihat, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstBordxsihat, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
        Case 4
            SortListView lstBordxsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstBordxsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
        Case 8
            SortListView lstBordxsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
        Case 10
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
        Case 11
            SortListView lstBordxsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView lstBordxsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView lstBordxsihat, ColumnHeader.Index, ldtString, m_blnDirection(13)
        Case 14
            SortListView lstBordxsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(14)
        End Select
    End If
End Sub

Private Sub lstBordxNonsihat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    If Check2.Value = 0 Then
    
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
        Case 3
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
        Case 6
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
        Case 9
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtString, m_blnDirection(10)
        Case 11
            SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
        End Select
        
    Else
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstBordxNonsihat, ColumnHeader.Index, ldtDateTime, m_blnDirection(12)
    End Select
    
    End If
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

