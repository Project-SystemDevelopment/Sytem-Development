VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintBordxBSN 
   Caption         =   "Print Bordx - BSN"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   7440
      Width           =   11655
      Begin VB.TextBox lblBordxNo 
         Height          =   285
         Left            =   4800
         TabIndex        =   35
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3000
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   720
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdPrintSummary 
         Caption         =   "&Print Report"
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
         Left            =   9600
         TabIndex        =   25
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
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
         Left            =   10680
         TabIndex        =   24
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
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
         Left            =   8880
         TabIndex        =   23
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdCCRpt 
         Caption         =   "&CC Report"
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
         Left            =   9720
         TabIndex        =   26
         Top             =   180
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblDate 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6240
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Bordx No"
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
         Index           =   1
         Left            =   4080
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
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
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
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
         Left            =   2040
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   11655
      Begin VB.TextBox TxtBordxNo 
         Height          =   285
         Left            =   4800
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   10320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox TxtBordxNoFr 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   195
         Width           =   2055
      End
      Begin VB.TextBox TxtBordxNoTo 
         Height          =   285
         Left            =   3840
         TabIndex        =   1
         Top             =   195
         Width           =   2055
      End
      Begin VB.ComboBox cboHOS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmPrintBordxBSN.frx":0000
         Left            =   1320
         List            =   "frmPrintBordxBSN.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   4575
      End
      Begin MSMask.MaskEdBox txtDOADate1 
         Height          =   255
         Left            =   7320
         TabIndex        =   3
         Top             =   150
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate1 
         Height          =   255
         Left            =   7320
         TabIndex        =   5
         Top             =   435
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate2 
         Height          =   255
         Left            =   9120
         TabIndex        =   4
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate2 
         Height          =   255
         Left            =   9120
         TabIndex        =   6
         Top             =   440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   255
         Left            =   7320
         TabIndex        =   7
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   255
         Left            =   9120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
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
         Index           =   1
         Left            =   8760
         TabIndex        =   21
         Top             =   735
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Bordx Date"
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
         Index           =   1
         Left            =   6240
         TabIndex        =   20
         ToolTipText     =   "Date of Discharged"
         Top             =   735
         Width           =   1215
      End
      Begin VB.Label Label15 
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
         Index           =   0
         Left            =   8760
         TabIndex        =   19
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label14 
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
         Left            =   8760
         TabIndex        =   18
         Top             =   200
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "DOD"
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
         Index           =   0
         Left            =   6240
         TabIndex        =   17
         ToolTipText     =   "Date of Discharged"
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "DOA"
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
         Left            =   6240
         TabIndex        =   16
         ToolTipText     =   "Date of Admission"
         Top             =   200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
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
         Index           =   1
         Left            =   10680
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Hospital"
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
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No"
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
         Left            =   240
         TabIndex        =   12
         Top             =   210
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstPrintBordx 
      Height          =   6135
      Left            =   120
      TabIndex        =   31
      Top             =   1320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10821
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Plan Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Claim No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Patient Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Member No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Hospital"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Bordx Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Bill Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Approve Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Recovery Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Case Mgmt Fee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Debit Credit Ref No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Print Status"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "HLT Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Payor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Pln"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "BTBG"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "BordxDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Print Bordx - BSN"
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
      Height          =   220
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmPrintBordxBSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SearchType As String
Dim TmpBordxDate As String
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim tmpTotalCases As Double

Private Sub cboHOS_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboHOS, KeyAscii)
End Sub

Private Sub cmdCCRpt_Click()
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

        Call SearchCC
 '   medixmasterconn.Close
 '   medixmasterconnWS.Close
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Function SearchCC()
Dim strsql As String
Dim strAppend As String
Dim sql As String

    strAppend = ""
    
    If Len(Trim(TxtBordxNoFr)) Then
        strAppend = strAppend + " And BordxRefNo >= '" & ChkStr(TxtBordxNoFr) & "'"
    End If

    If Len(Trim(TxtBordxNoTo)) Then
        strAppend = strAppend + " And BordxRefNo <= '" & ChkStr(TxtBordxNoTo) & "'"
    End If

    If Len(Trim(cboHOS)) Then
        strAppend = strAppend + " And HOSName = '" & ChkStr(cboHOS) & "'"
    End If
    
    If Len(Trim(txtDOADate1)) > 0 And IsDate(txtDOADate1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(txtDOADate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOADate2)) > 0 And IsDate(txtDOADate2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(txtDOADate2, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(txtDODDate1)) > 0 And IsDate(txtDODDate1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(txtDODDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDODDate2)) > 0 And IsDate(txtDODDate2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(txtDODDate2, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(txtBordxDateFr)) > 0 And IsDate(txtBordxDateFr) = True Then
        strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtBordxDateTo)) > 0 And IsDate(txtBordxDateTo) = True Then
        strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(cboHOS) Then
        strAppend = strAppend + " And HOSName = '" & cboHOS & "'"
    End If
    
    Call GenerateCCReports(strAppend)
End Function

Private Sub CmdPrintSummary_Click()
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
    SearchType = "1"
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    If Trim(TxtBordxNoFr) = "" Or Trim(TxtBordxNoTo) = "" Then
        MsgBox "Please enter Bordx Reference No", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        'CmdClear.Enabled = False
        Call SearchRecord
    End If
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default

End Sub
Private Sub ComboKeyPress(tmpComboBox As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(tmpComboBox.Text, tmpComboBox.SelStart) + Chr(KeyAscii) + Mid(tmpComboBox.Text, tmpComboBox.SelStart + tmpComboBox.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To tmpComboBox.ListCount - 1
        If NewText = LCase(Left(tmpComboBox.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = tmpComboBox.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           tmpComboBox.Text = ValidValue
           tmpComboBox.SelStart = 0
           tmpComboBox.SelLength = Len(ValidValue)
         End If
    End If
End Sub

Private Sub cmdSearch_Click()
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
    SearchType = "2"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    CmdPrintSummary.Enabled = True
    If Trim(TxtBordxNoFr) = "" Or Trim(TxtBordxNoTo) = "" Then
        MsgBox "Please enter Bordx Reference No", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        'CmdClear.Enabled = False
        Call SearchRecord
    End If
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub combolist()
Dim tmpSelField As String
Dim tmpTable As String
Dim tmpSelCriteria As String
    cboHOS.Clear
        tmpSelField = "SELECT HOSCode as tmpCode, HOSName as tmpName "
        tmpTable = "HOS"
        tmpSelCriteria = "0=0"
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, cboHOS)
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing

End Sub

Private Sub LoadComboListing(tmpSelField As String, tmpSelCriteria As String, tmpTableName As String, tmpComboBox As ComboBox)
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter

    tmpComboBox.Clear

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, tmpSelField)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, tmpTableName)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 255, tmpSelCriteria)
    cmd.Parameters.Append Prm3
    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    
    Do Until TmpRec.EOF
        With TmpRec
        If tmpTableName = "HOS" Then
            If Len(!tmpName) Then
                tmpComboBox.AddItem !tmpName
            End If
        End If
            TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
End Sub
Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String

    strAppend = ""
    
    If Len(Trim(TxtBordxNoFr)) Then
        strAppend = strAppend + " And BordxRefNo >= '" & ChkStr(TxtBordxNoFr) & "'"
    End If

    If Len(Trim(TxtBordxNoTo)) Then
        strAppend = strAppend + " And BordxRefNo <= '" & ChkStr(TxtBordxNoTo) & "'"
    End If

    If Len(Trim(cboHOS)) Then
        strAppend = strAppend + " And HOSName = '" & ChkStr(cboHOS) & "'"
    End If
    
    If Len(Trim(txtDOADate1)) > 0 And IsDate(txtDOADate1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(txtDOADate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOADate2)) > 0 And IsDate(txtDOADate2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(txtDOADate2, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(txtDODDate1)) > 0 And IsDate(txtDODDate1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(txtDODDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDODDate2)) > 0 And IsDate(txtDODDate2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(txtDODDate2, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(txtBordxDateFr)) > 0 And IsDate(txtBordxDateFr) = True Then
        strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtBordxDateTo)) > 0 And IsDate(txtBordxDateTo) = True Then
        strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(cboHOS) Then
        strAppend = strAppend + " And HOSName = '" & cboHOS & "'"
    End If
    
    If SearchType = "2" Then
        Call PrintBordxListing(strAppend)
    Else
        Call GeneratePaymentSum
    End If
End Function

Private Sub PrintBordxListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim PlanCheck As Boolean
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset

    total = 0
    Current = 0
    PlanCheck = False
    lstPrintBordx.ListItems.Clear
               
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "Print_Bordx"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 1)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
     If rst2.EOF = True Then
        MsgBox "No record found ", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrintSummary.Enabled = False
    Else
        
    'Do
        Do Until rst2.EOF
        total = rst2.recordCount

             With rst2
                Set Record = lstPrintBordx.ListItems.Add(, , rst2!PLNCode)
                If Len(rst2!CasNumber & "-" & rst2!claimversion) Then
                    Record.SubItems(1) = (rst2!CasNumber & "-" & rst2!claimversion)
                End If
                If Len(rst2!PatientName) Then
                    Record.SubItems(2) = rst2!PatientName
                End If
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(3) = rst2!MBMNumber
                End If
                If Len(rst2!HosName) Then
                   Record.SubItems(4) = rst2!HosName
                End If
                                
                If rst2!MBMBTBG = "Y" Then
                    If Len(rst2!CLMPayableByMedix2H) Then
                        Record.SubItems(5) = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                    End If
                Else
                    If Len(rst2!CLMTotalApproved) Then
                        Record.SubItems(5) = Format(rst2!CLMTotalApproved, "#,##0.00")
                    End If
                End If
                
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(6) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(7) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(8) = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                End If
                'If Len(rst2!CLMCaseMgmtFees) Then
                '    Record.SubItems(9) = Format(rst2!CLMCaseMgmtFees, "#,##0.00")
                'End If
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(10) = rst2!DebitCreditRefNo
                End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(11) = rst2!BordxRefNo
                End If
                  
                
                If Len(rst2!HLTCode) Then
                    Record.SubItems(13) = rst2!HLTCode
                End If
                If Len(rst2!PAYCode) Then
                    Record.SubItems(14) = rst2!PAYCode
                End If
                If Len(rst2!PLNCode) Then
                    Record.SubItems(15) = rst2!PLNCode
                End If
                If Trim(rst2!MBMBTBG) = "Y" Then
                    Record.SubItems(15) = rst2!MBMBTBG
                Else
                    Record.SubItems(15) = "N"
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(17) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                
                Current = Current + 1
                lblCurrent = Current
                
                
                If rst2!MBMBTBG = "Y" And rst2!InvPayTo = "H" Then
                    TotalAmt = TotalAmt + rst2!CLMPayableByMedix2H
                ElseIf rst2!MBMBTBG = "Y" And rst2!InvPayTo = "M" Then
                    TotalAmt = TotalAmt + rst2!CLMPayableByMedix2M
                Else
                'If Len(rst2!MBMBTBG) Then
                    TotalAmt = TotalAmt + rst2!TotalMgmtFee
                End If
                
                'TotalAmt = TotalAmt + rst2!CLMTotalApproved
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
            End With
            rst2.MoveNext
            'DoEvents
                'If intExit = 1 Then
                '   Check = False
                '   Exit Do
                'End If
        Loop
    'Loop Until Check = False
    End If
    'intExit = 0
    rst2.Close
    Set rst2 = Nothing
    'CmdClear.Enabled = True
End Sub

Private Sub Form_Load()
    Call combolist
End Sub

Private Sub GeneratePaymentSum()

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim strsql As String
Dim rst2 As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim tmpReceiptNo As String
Dim PVCnt As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHosName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpSubMBMPaid As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As String
Dim tmpPatientName As String
Dim strsqlCOV As String
Dim RstCov As New ADODB.Recordset
Dim totalMBMPaid As Double

    Screen.MousePointer = vbHourglass

        
        
        
        
        Call GenerateMASWingRpt
        'Call GenerateAWUSRpt
        'Call GenerateSarawakRpt
        'Call GenerateMemberRpt
        'Call GenerateRecoveryRpt
        'Call GenerateStaffClaimRpt
        'Call GenerateDailyAllwRpt
        'medixmasterconnPayment.Close
        'Set medixmasterconnPayment = Nothing
        'objExcel.DisplayAlerts = True
        'objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing

        Exit Sub
'HANDLE_ERROR:
'        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'        Set objWorksheet = Nothing
'        Set objWorkbook = Nothing
'        GoTo ExitFunction
End Sub

Private Sub GenerateMASWingRpt()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim strsql As String
Dim rst2 As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim tmpReceiptNo As String
Dim PVCnt As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim cmd3 As New ADODB.Command
Dim Prm5 As New ADODB.Parameter
Dim prm6 As New ADODB.Parameter
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim prm8 As New ADODB.Parameter
Dim cmd5 As New ADODB.Command
Dim Prm9 As New ADODB.Parameter
Dim prm10 As New ADODB.Parameter

Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHosName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double
Dim strsqlCOV As String
Dim RstCov As New ADODB.Recordset
Dim tmpPatientName As String
Dim tmpSubMBMPaid As Double
Dim totalMBMPaid As Double
Dim PVReimCnt As Double
Dim tmpPatientRel As String
Dim PLN As New ADODB.Recordset
Dim strsqlPLN As String
Dim PLNDesc As String
    Screen.MousePointer = vbHourglass

    If lblBordxNo = "" Then
    
        MsgBox "Please Select Bordx", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\PayAdviceSummaryBSN.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        'On Error GoTo HANDLE_ERROR
        RecFound = False
        strsql = ""
        strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "CLM_PayAdviceHOSBSN"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 2)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        PVCnt = 0
        tmpTotalCases = 0
        tmpTotalActual = 0
        tmpTotalHos = 0
        tmpHOSPayee = ""
        tmpSubMBMPaid = 0
        RecFound = False
        tmpSeqNo = 0
        totalMBMPaid = 0
        PVReimCnt = 0
        tmpPatientRel = ""
        
        Do While Not rst2.EOF
            RecFound = True
            PVCnt = PVCnt + 1
            If PVCnt = 1 Then
                objWorksheet.Cells(2, "J") = lblBordxNo
                objWorksheet.Cells(3, "J") = lblDate
            End If
                tmpSeqNo = CDbl(tmpSeqNo) + 1
                objWorksheet.Cells(16 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                objWorksheet.Cells(16 + CDbl(PVCnt), "B") = rst2!HosName
                objWorksheet.Cells(16 + CDbl(PVCnt), "E") = rst2!TotalCases
                tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                objWorksheet.Cells(16 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                tmpTotalActual = tmpTotalActual + Format(CDbl(rst2!TotalActual), "#,##0.00")
                totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                objWorksheet.Cells(16 + CDbl(PVCnt), "G") = Format(rst2!PayabletoHOS, "#,##0.00")
                objWorksheet.Cells(16 + CDbl(PVCnt), "H") = CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS)
                
                'objWorksheet.Cells(16 + CDbl(PVCnt), "J").Borders.LineStyle = True
                'objWorksheet.Cells(16 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
            rst2.MoveNext
        Loop
          '  objWorksheet.Cells(15, "E") = Format(tmpTotalCases, "#,##0.00")
          '  objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
          '  objWorksheet.Cells(15, "G") = Format(totalMBMPaid, "#,##0.00")
          '  objWorksheet.Cells(15, "H") = Format(tmpTotalHos, "#,##0.00")
          '  'objWorksheet.Cells(8, "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
          '  objWorksheet.Cells(15, "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
          '  objWorksheet.Cells(15, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
          '  objWorksheet.Cells(15, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
          '  objWorksheet.Cells(15, "H").Borders(xlEdgeBottom).LineStyle = xlContinuous
            
            objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "E").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
            objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "F").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "H") = Format(totalMBMPaid, "#,##0.00")
            objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "G").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(tmpTotalHos, "#,##0.00")
            
        rst2.Close
        Set rst2 = Nothing
        
        'By Hosp State
        
        If RecFound = True Then

                strsql = ""
                strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
                Set cmd3.ActiveConnection = medixmasterconnWS
                cmd3.CommandText = "CLM_PayAdviceHOSBSN"
                cmd3.CommandType = adCmdStoredProc
                cmd3.CommandTimeout = 300
                Set Prm5 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd3.Parameters.Append Prm5
                Set prm6 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 3)
                cmd3.Parameters.Append prm6
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd3.Execute
                
                PVCnt = 0
                tmpTotalCases = 0
                tmpTotalActual = 0
                tmpTotalHos = 0
                tmpHOSPayee = ""
                tmpSubMBMPaid = 0
                RecFound = False
                tmpSeqNo = 0
                totalMBMPaid = 0
                PVReimCnt = 0
                tmpPatientRel = ""
                
                Do While Not rst2.EOF
                    RecFound = True
                    PVCnt = PVCnt + 1
                    Set objWorksheet = objWorkbook.Sheets(2)
                    objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(2)
                    
                    If PVCnt = 1 Then
                        objWorksheet.Cells(2, "J") = lblBordxNo
                        objWorksheet.Cells(3, "J") = lblDate
                    End If
                        tmpSeqNo = CDbl(tmpSeqNo) + 1
                        objWorksheet.Cells(16 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                        objWorksheet.Cells(16 + CDbl(PVCnt), "B") = rst2!HOSState
                        objWorksheet.Cells(16 + CDbl(PVCnt), "E") = rst2!TotalCases
                        tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                        objWorksheet.Cells(16 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                        tmpTotalActual = tmpTotalActual + Format(CDbl(rst2!TotalActual), "#,##0.00")
                        totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                        objWorksheet.Cells(16 + CDbl(PVCnt), "G") = Format(rst2!PayabletoHOS, "#,##0.00")
                        objWorksheet.Cells(16 + CDbl(PVCnt), "H") = CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS)
                        
                        'objWorksheet.Cells(16 + CDbl(PVCnt), "J").Borders.LineStyle = True
                        'objWorksheet.Cells(16 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                        tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                        tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
                    rst2.MoveNext
                Loop
                  '  objWorksheet.Cells(15, "E") = Format(tmpTotalCases, "#,##0.00")
                  '  objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
                  '  objWorksheet.Cells(15, "G") = Format(totalMBMPaid, "#,##0.00")
                  '  objWorksheet.Cells(15, "H") = Format(tmpTotalHos, "#,##0.00")
                  '  'objWorksheet.Cells(8, "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "H").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H") = Format(totalMBMPaid, "#,##0.00")
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(tmpTotalHos, "#,##0.00")
                    
                rst2.Close
                Set rst2 = Nothing
            End If
        
        'By Department
        
        If RecFound = True Then

                strsql = ""
                strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
                Set cmd4.ActiveConnection = medixmasterconnWS
                cmd4.CommandText = "CLM_PayAdviceHOSBSN"
                cmd4.CommandType = adCmdStoredProc
                cmd4.CommandTimeout = 300
                Set Prm7 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd4.Parameters.Append Prm7
                Set prm8 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 4)
                cmd4.Parameters.Append prm8
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd4.Execute
                
                PVCnt = 0
                tmpTotalCases = 0
                tmpTotalActual = 0
                tmpTotalHos = 0
                tmpHOSPayee = ""
                tmpSubMBMPaid = 0
                RecFound = False
                tmpSeqNo = 0
                totalMBMPaid = 0
                PVReimCnt = 0
                tmpPatientRel = ""
                
                Do While Not rst2.EOF
                    RecFound = True
                    PVCnt = PVCnt + 1
                    Set objWorksheet = objWorkbook.Sheets(3)
                    objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(3)
                    
                    If PVCnt = 1 Then
                        objWorksheet.Cells(2, "J") = lblBordxNo
                        objWorksheet.Cells(3, "J") = lblDate
                    End If
                        tmpSeqNo = CDbl(tmpSeqNo) + 1
                        objWorksheet.Cells(16 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                        objWorksheet.Cells(16 + CDbl(PVCnt), "B") = rst2!MBMDepartment
                        objWorksheet.Cells(16 + CDbl(PVCnt), "E") = rst2!TotalCases
                        tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                        objWorksheet.Cells(16 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                        tmpTotalActual = tmpTotalActual + Format(CDbl(rst2!TotalActual), "#,##0.00")
                        totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                        objWorksheet.Cells(16 + CDbl(PVCnt), "G") = Format(rst2!PayabletoHOS, "#,##0.00")
                        objWorksheet.Cells(16 + CDbl(PVCnt), "H") = CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS)
                        
                        'objWorksheet.Cells(16 + CDbl(PVCnt), "J").Borders.LineStyle = True
                        'objWorksheet.Cells(16 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                        tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                        tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
                    rst2.MoveNext
                Loop
                  '  objWorksheet.Cells(15, "E") = Format(tmpTotalCases, "#,##0.00")
                  '  objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
                  '  objWorksheet.Cells(15, "G") = Format(totalMBMPaid, "#,##0.00")
                  '  objWorksheet.Cells(15, "H") = Format(tmpTotalHos, "#,##0.00")
                  '  'objWorksheet.Cells(8, "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "H").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H") = Format(totalMBMPaid, "#,##0.00")
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(tmpTotalHos, "#,##0.00")
                    
                rst2.Close
                Set rst2 = Nothing
            End If
        
            
        'By DOD Month
        
        If RecFound = True Then
                strsql = ""
                strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
                Set cmd5.ActiveConnection = medixmasterconnWS
                cmd5.CommandText = "CLM_PayAdviceHOSBSN"
                cmd5.CommandType = adCmdStoredProc
                cmd5.CommandTimeout = 300
                Set Prm9 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd5.Parameters.Append Prm9
                Set prm10 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 5)
                cmd5.Parameters.Append prm10
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd5.Execute
                
                PVCnt = 0
                tmpTotalCases = 0
                tmpTotalActual = 0
                tmpTotalHos = 0
                tmpHOSPayee = ""
                tmpSubMBMPaid = 0
                RecFound = False
                tmpSeqNo = 0
                totalMBMPaid = 0
                PVReimCnt = 0
                tmpPatientRel = ""
                
                Do While Not rst2.EOF
                    RecFound = True
                    'PLNCode
                    PVCnt = PVCnt + 1
                    Set objWorksheet = objWorkbook.Sheets(4)
                    objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(4)
                    
                    If PVCnt = 1 Then
                        objWorksheet.Cells(2, "J") = lblBordxNo
                        objWorksheet.Cells(3, "J") = lblDate
                    End If
                        tmpSeqNo = CDbl(tmpSeqNo) + 1
                        objWorksheet.Cells(16 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                        objWorksheet.Cells(16 + CDbl(PVCnt), "B") = rst2!DODMonth & "/" & rst2!DODYear
                        objWorksheet.Cells(16 + CDbl(PVCnt), "E") = rst2!TotalCases
                        tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                        objWorksheet.Cells(16 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                        tmpTotalActual = tmpTotalActual + Format(CDbl(rst2!TotalActual), "#,##0.00")
                        totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                        objWorksheet.Cells(16 + CDbl(PVCnt), "G") = Format(rst2!PayabletoHOS, "#,##0.00")
                        objWorksheet.Cells(16 + CDbl(PVCnt), "H") = CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS)
                        
                        'objWorksheet.Cells(16 + CDbl(PVCnt), "J").Borders.LineStyle = True
                        'objWorksheet.Cells(16 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                        tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                        tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
                    rst2.MoveNext
                Loop
                  '  objWorksheet.Cells(15, "E") = Format(tmpTotalCases, "#,##0.00")
                  '  objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
                  '  objWorksheet.Cells(15, "G") = Format(totalMBMPaid, "#,##0.00")
                  '  objWorksheet.Cells(15, "H") = Format(tmpTotalHos, "#,##0.00")
                  '  'objWorksheet.Cells(8, "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  '  objWorksheet.Cells(15, "H").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "H") = Format(totalMBMPaid, "#,##0.00")
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G").Font.Bold = True
                    objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(tmpTotalHos, "#,##0.00")
                    
                rst2.Close
                Set rst2 = Nothing
            End If
            'General Listing
            If RecFound = True Then
                strsql = ""
                strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
                Set cmd2.ActiveConnection = medixmasterconnWS
                cmd2.CommandText = "CLM_PayAdviceHOSBSN"
                cmd2.CommandType = adCmdStoredProc
                cmd2.CommandTimeout = 300
                Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 3000, strsql)
                cmd2.Parameters.Append Prm3
                Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 3000, 1)
                cmd2.Parameters.Append prm4
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd2.Execute
            
            
                PVCnt = 0
                'tmpTotalCases = 0
                tmpTotalActual = 0
                tmpTotalHos = 0
                TMPWS = 1
                Do While Not rst2.EOF
                    PVCnt = PVCnt + 1
                    PLNDesc = ""
                    strsqlPLN = "Select PLNCode ,PLNDescription from PLN where PLNCode = '" & rst2!PLNCode & "'"
                    PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                    
                    If PLN.recordCount > 0 Then
                        'PLNDesc = Mid(PLN!PLNDescription, 1, IIf(InStr(1, PLN!PLNDescription, "-") = 0, 4, InStr(4, PLN!PLNDescription, "-") - 1))
                        PLNDesc = Mid(PLN!PLNDescription, InStr(1, PLN!PLNDescription, "-") + 1)
                    End If
                    PLN.Close
                    Set PLN = Nothing
                    
                    Set objWorksheet = objWorkbook.Sheets(5)
                    objExcel.DisplayAlerts = False
                    Set objWorksheet = objWorkbook.Worksheets.Item(5)
                    
                    If PVCnt = 1 Then
                        objWorksheet.Cells(2, "L") = lblBordxNo
                        objWorksheet.Cells(3, "L") = lblDate
                    End If
                    

                        If Trim(rst2!MBMPatientCovID) <> "00" Then
                            strsqlCOV = "Select MBMCName,MBMCRELCode from MBMCoveredPersons where MBMCNumber = '" & Trim(rst2!MBMNumber) & "' and MBMCCoverID = '" & Trim(rst2!MBMPatientCovID) & "'"
                            RstCov.Open strsqlCOV, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                                Do While Not RstCov.EOF
                                    tmpPatientName = IIf(Len(RstCov!MBMCName), RstCov!MBMCName, "")
                                    tmpPatientRel = IIf(Len(RstCov!MBMCRELCode), RstCov!MBMCRELCode, "")
                                    RstCov.MoveNext
                                Loop
                                RstCov.Close
                                Set RstCov = Nothing
                                
                            Else
                                tmpPatientName = rst2!MBMName
                                tmpPatientRel = "E"
                        End If
                        
                        'Format(rst2!CLMTotalActual, "#,##0.00")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "A") = PVCnt
                        objWorksheet.Cells(14 + CDbl(PVCnt), "B") = rst2!CasNumber & "-" & rst2!claimversion
                        objWorksheet.Cells(14 + CDbl(PVCnt), "C") = Right(PLNDesc, 2)
                        If Trim(tmpPatientRel) = "E" Then
                            objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "E"
                        ElseIf Trim(tmpPatientRel) = "W" Then
                            objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "S"
                        ElseIf Trim(tmpPatientRel) = "H" Then
                            objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "S"
                        ElseIf Trim(tmpPatientRel) = "D" Then
                            objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "C"
                        ElseIf Trim(tmpPatientRel) = "S" Then
                            objWorksheet.Cells(14 + CDbl(PVCnt), "D") = "C"
                        Else
                            objWorksheet.Cells(14 + CDbl(PVCnt), "D") = ""
                        End If
                        objWorksheet.Cells(14 + CDbl(PVCnt), "E") = rst2!MBMName
                        objWorksheet.Cells(14 + CDbl(PVCnt), "G") = tmpPatientName
                        objWorksheet.Cells(14 + CDbl(PVCnt), "I") = rst2!MBMPolicyNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "K") = rst2!MBMNumber
                        objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(rst2!CASDateAdmitted, "YYYY/MM/DD")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "O") = Format(rst2!CASDischargeDate, "YYYY/MM/DD")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "P") = IIf(Left(rst2!CLMWSRemarks, 4) = "POST", rst2!CLMWSRemarks, "")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "Q") = rst2!DIAFinalDiag1 & "," & rst2!DIAFinalDiag2 & "," & rst2!DIAFinalDiag3
                        objWorksheet.Cells(14 + CDbl(PVCnt), "R") = rst2!HosName
                        
                        If rst2!ModalityCode = "N" Then
                            objWorksheet.Cells(14 + CDbl(PVCnt), "S") = "NS"
                        ElseIf rst2!ModalityCode = "S" Then
                            objWorksheet.Cells(14 + CDbl(PVCnt), "S") = "S"
                        Else
                            objWorksheet.Cells(14 + CDbl(PVCnt), "S") = rst2!ModalityCode
                        End If
                        objWorksheet.Cells(14 + CDbl(PVCnt), "T") = Format(rst2!CLMTotalActual, "#,##0.00")
                        tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                        tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMPayableByMedix2H, "#,##0.00"))
                        objWorksheet.Cells(14 + CDbl(PVCnt), "V") = Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "U") = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                        tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMPayableByMedix2H)
                    
                    rst2.MoveNext
                Loop
                rst2.Close
                Set rst2 = Nothing
                RecFound = False
                

                PVCnt = PVCnt + 2
                objWorksheet.Cells(14 + CDbl(PVCnt), "T").Borders(xlEdgeTop).LineStyle = xlContinuous
                objWorksheet.Cells(14 + CDbl(PVCnt), "T").Borders(xlEdgeBottom).LineStyle = xlDouble
                objWorksheet.Cells(14 + CDbl(PVCnt), "T").Font.Bold = True
                objWorksheet.Cells(14 + CDbl(PVCnt), "T") = Format(tmpsubAct, "#,##0.00")
                objWorksheet.Cells(14 + CDbl(PVCnt), "V").Borders(xlEdgeTop).LineStyle = xlContinuous
                objWorksheet.Cells(14 + CDbl(PVCnt), "V").Borders(xlEdgeBottom).LineStyle = xlDouble
                objWorksheet.Cells(14 + CDbl(PVCnt), "V").Font.Bold = True
                objWorksheet.Cells(14 + CDbl(PVCnt), "V") = Format(tmpSubMBMPaid, "#,##0.00")
                objWorksheet.Cells(14 + CDbl(PVCnt), "U").Borders(xlEdgeTop).LineStyle = xlContinuous
                objWorksheet.Cells(14 + CDbl(PVCnt), "U").Borders(xlEdgeBottom).LineStyle = xlDouble
                objWorksheet.Cells(14 + CDbl(PVCnt), "U").Font.Bold = True
                objWorksheet.Cells(14 + CDbl(PVCnt), "U") = Format(tmpSubApp, "#,##0.00")
                
                objWorksheet.Cells(14 + CDbl(PVCnt), "T") = Format(tmpsubAct, "#,##0.00")
                objWorksheet.Cells(14 + CDbl(PVCnt), "V") = Format(tmpSubMBMPaid, "#,##0.00")
                objWorksheet.Cells(14 + CDbl(PVCnt), "U") = Format(tmpSubApp, "#,##0.00")
        End If
        
        
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault

        Exit Sub

End Sub

Private Sub GenerateMemberRpt()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim strsql As String
Dim rst2 As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim tmpReceiptNo As String
Dim PVCnt As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHosName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double
Dim tmptax As Double
Dim tmpTotalMBM As Double
Dim tmpTotalTax As Double
Dim tmpTotalDueMBM As Double

    Screen.MousePointer = vbHourglass

    If lblBordxNo = "" Then
    
        MsgBox "Please Select Bordx", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\PayAdviceMBMSummary.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        'On Error GoTo HANDLE_ERROR
        RecFound = False
        strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "CLM_PayAdviceHOS"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 5)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        tmptax = 0
        PVCnt = 0
        tmpTotalCases = 0
        tmpTotalActual = 0
        tmpTotalHos = 0
        tmpHOSPayee = ""
        tmpSeqNo = 0
        tmpTotalMBM = 0
        tmpTotalTax = 0
        tmpTotalDueMBM = 0
        tmptax = 0
        RecFound = False
        Do While Not rst2.EOF
            RecFound = True
            PVCnt = PVCnt + 1
            If PVCnt = 1 Then
                objWorksheet.Cells(9, "B") = lblBordxNo
                objWorksheet.Cells(10, "B") = lblDate
            End If
                tmpSeqNo = CDbl(tmpSeqNo) + 1
                objWorksheet.Cells(13 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                objWorksheet.Cells(13 + CDbl(PVCnt), "B") = rst2!MBMEmployeeNo
                objWorksheet.Cells(13 + CDbl(PVCnt), "C") = rst2!MBMEmail
                objWorksheet.Cells(13 + CDbl(PVCnt), "D") = rst2!MBMName
                
                objWorksheet.Cells(13 + CDbl(PVCnt), "I") = rst2!MBMName
                
                objWorksheet.Cells(13 + CDbl(PVCnt), "N") = rst2!CasNumber & "-" & rst2!claimversion
                If Format(rst2!CLMTax, "#,##0.00") <> Format(rst2!CLMPayableByMedix2M, "#,##0.00") Then
                    objWorksheet.Cells(13 + CDbl(PVCnt), "O") = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                Else
                    objWorksheet.Cells(13 + CDbl(PVCnt), "O") = Format(0, "#,##0.00")
                End If
                objWorksheet.Cells(13 + CDbl(PVCnt), "P") = IIf(IsNumeric(rst2!CLMTax), Format(rst2!CLMTax, "#,##0.00"), Format(0, "#,##0.00"))
                tmptax = IIf(IsNumeric(rst2!CLMTax), Format(rst2!CLMTax, "#,##0.00"), Format(0, "#,##0.00"))
                
                objWorksheet.Cells(13 + CDbl(PVCnt), "Q") = rst2!BordxDate
                
                objWorksheet.Cells(13 + CDbl(PVCnt), "R") = Format(0, "#,##0.00")
                
                If Format(rst2!CLMTax, "#,##0.00") <> Format(rst2!CLMPayableByMedix2M, "#,##0.00") Then
                    objWorksheet.Cells(13 + CDbl(PVCnt), "S") = Format(CDbl(rst2!CLMPayableByMedix2M) + CDbl(tmptax), "#,##0.00")
                Else
                    objWorksheet.Cells(13 + CDbl(PVCnt), "S") = Format(CDbl(tmptax), "#,##0.00")
                End If
                'objWorksheet.Cells(14 + CDbl(PVCnt), "S").Borders.LineStyle = True
                'objWorksheet.Cells(14 + CDbl(PVCnt), "U").Borders(xlEdgeBottom).LineStyle = xlContinuous
                tmpTotalCases = tmpTotalCases + 1
                
                If Format(rst2!CLMTax, "#,##0.00") <> Format(rst2!CLMPayableByMedix2M, "#,##0.00") Then
                    tmpTotalMBM = CDbl(tmpTotalMBM) + CDbl(rst2!CLMPayableByMedix2M)
                    tmpTotalMBM = Format(tmpTotalMBM, "#,##0.00")
                End If
                
                tmpTotalTax = CDbl(tmpTotalTax) + CDbl(tmptax)
                tmpTotalTax = Format(tmpTotalTax, "#,##0.00")
                
                If Format(rst2!CLMTax, "#,##0.00") <> Format(rst2!CLMPayableByMedix2M, "#,##0.00") Then
                    tmpTotalDueMBM = tmpTotalDueMBM + (CDbl(rst2!CLMPayableByMedix2M) + CDbl(tmptax))
                    tmpTotalDueMBM = Format(tmpTotalDueMBM, "#,##0.00")
                Else
                    tmpTotalDueMBM = tmpTotalDueMBM + CDbl(tmptax)
                    tmpTotalDueMBM = Format(tmpTotalDueMBM, "#,##0.00")
                End If
            rst2.MoveNext
        Loop
            'objWorksheet.Cells(15, "J") = Format(tmpTotalMBM, "#,##0.00")
            'objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
            'objWorksheet.Cells(15, "H") = Format(tmpTotalHos, "#,##0.00")
            'objWorksheet.Cells(15, "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
            'objWorksheet.Cells(15, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
            'objWorksheet.Cells(15, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
            'objWorksheet.Cells(15, "H").Borders(xlEdgeBottom).LineStyle = xlContinuous
            
            objWorksheet.Cells(14 + CDbl(PVCnt), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(14 + CDbl(PVCnt), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(14 + CDbl(PVCnt), "O").Font.Bold = True
            objWorksheet.Cells(14 + CDbl(PVCnt), "O") = Format(tmpTotalMBM, "#,##0.00")
            objWorksheet.Cells(14 + CDbl(PVCnt), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(14 + CDbl(PVCnt), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(14 + CDbl(PVCnt), "P").Font.Bold = True
            objWorksheet.Cells(14 + CDbl(PVCnt), "P") = Format(tmpTotalTax, "#,##0.00")
            objWorksheet.Cells(14 + CDbl(PVCnt), "R").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(14 + CDbl(PVCnt), "R").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(14 + CDbl(PVCnt), "R").Font.Bold = True
            objWorksheet.Cells(14 + CDbl(PVCnt), "R") = Format(0, "#,##0.00")
            objWorksheet.Cells(14 + CDbl(PVCnt), "S").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(14 + CDbl(PVCnt), "S").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(14 + CDbl(PVCnt), "S").Font.Bold = True
            objWorksheet.Cells(14 + CDbl(PVCnt), "S") = Format(tmpTotalDueMBM, "#,##0.00")

        rst2.Close
        Set rst2 = Nothing
        
        
            If RecFound = True Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Claim\" & "\CoverStaffPayment.DOT")
                End If
                
                    'bookmark = False
                    objdoc.Bookmarks("TotalInv").Range.Text = tmpTotalCases
                    objdoc.Bookmarks("BordxDate").Range.Text = lblDate
                    objdoc.Bookmarks("BordxNo").Range.Text = lblBordxNo
                
        End If
        
        
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing

        Exit Sub
'HANDLE_ERROR:
'        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'        Set objWorksheet = Nothing
'        Set objWorkbook = Nothing
'        GoTo ExitFunction
End Sub

Private Sub GenerateRecoveryRpt()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim strsql As String
Dim rst2 As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim tmpReceiptNo As String
Dim PVCnt As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHosName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double
Dim tmpTotalRecovery As Double
Dim cov As New ADODB.Recordset
    Screen.MousePointer = vbHourglass

    If lblBordxNo = "" Then
    
        MsgBox "Please Select Bordx", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\PayAdviceRecovery.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        'On Error GoTo HANDLE_ERROR
        RecFound = False
        strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "CLM_PayAdviceHOS"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 6)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        PVCnt = 0
        tmpTotalCases = 0
        tmpTotalActual = 0
        tmpTotalHos = 0
        tmpHOSPayee = ""
        tmpSeqNo = 0
        tmpTotalRecovery = 0
        RecFound = False
        Do While Not rst2.EOF
            RecFound = True
            PVCnt = PVCnt + 1
            If PVCnt = 1 Then
                objWorksheet.Cells(1, "B") = lblBordxNo
                objWorksheet.Cells(2, "B") = lblDate
            End If
                tmpSeqNo = CDbl(tmpSeqNo) + 1
                objWorksheet.Cells(5 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                objWorksheet.Cells(5 + CDbl(PVCnt), "B") = rst2!MBMEmployeeNo
                objWorksheet.Cells(5 + CDbl(PVCnt), "C") = rst2!MBMName
                If Trim(rst2!MBMPatientCovID) = "00" Then
                    objWorksheet.Cells(5 + CDbl(PVCnt), "F") = rst2!MBMName
                Else
                    strsql = "Select MBMCNumber, MBMCName,MBMCCoverID from MBMCoveredPersons where MBMCNumber = '" & rst2!MBMNumber & "' And MBMCCoverID = '" & rst2!MBMPatientCovID & "'"
                    cov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    'strsqlHOS = "Select HOSPayee from HOS where HOSName = '" & objWorksheet.Cells(7, "A") & "'"
                    'HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic

                    Do While Not cov.EOF
                        objWorksheet.Cells(5 + CDbl(PVCnt), "F") = cov!MBMCName
                        cov.MoveNext
                    Loop
                    cov.Close
                    Set cov = Nothing
                End If
                objWorksheet.Cells(5 + CDbl(PVCnt), "J") = rst2!CasNumber & "-" & rst2!claimversion
                objWorksheet.Cells(5 + CDbl(PVCnt), "K") = rst2!HosName
                objWorksheet.Cells(5 + CDbl(PVCnt), "N") = rst2!CASDischargeDate
                objWorksheet.Cells(5 + CDbl(PVCnt), "O") = Format(rst2!CLMBTBGAmt * -1, "#,##0.00")
                tmpTotalCases = CDbl(tmpTotalCases) + 1
                tmpTotalRecovery = CDbl(tmpTotalRecovery) + CDbl(rst2!CLMBTBGAmt * -1)
                
            rst2.MoveNext
        Loop
            objWorksheet.Cells(7 + CDbl(PVCnt), "J").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(7 + CDbl(PVCnt), "J").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(7 + CDbl(PVCnt), "J").Font.Bold = True
            objWorksheet.Cells(7 + CDbl(PVCnt), "J") = "Total Claims"

            objWorksheet.Cells(7 + CDbl(PVCnt), "K").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(7 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(7 + CDbl(PVCnt), "K").Font.Bold = True
            objWorksheet.Cells(7 + CDbl(PVCnt), "K") = tmpTotalCases
            
            objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(7 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(7 + CDbl(PVCnt), "L").Font.Bold = True
            
            objWorksheet.Cells(7 + CDbl(PVCnt), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(7 + CDbl(PVCnt), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(7 + CDbl(PVCnt), "M").Font.Bold = True
            
            objWorksheet.Cells(7 + CDbl(PVCnt), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(7 + CDbl(PVCnt), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(7 + CDbl(PVCnt), "N").Font.Bold = True
            objWorksheet.Cells(7 + CDbl(PVCnt), "N") = "Total Amt"

            objWorksheet.Cells(7 + CDbl(PVCnt), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(7 + CDbl(PVCnt), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(7 + CDbl(PVCnt), "O").Font.Bold = True
            objWorksheet.Cells(7 + CDbl(PVCnt), "O") = Format(tmpTotalRecovery, "#,##0.00")

        rst2.Close
        Set rst2 = Nothing
        
        
            If RecFound = True Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Claim\" & "\CoverRecovery.DOT")
                End If
                
                    'bookmark = False
                    objdoc.Bookmarks("TotalInv").Range.Text = tmpTotalCases
                    objdoc.Bookmarks("BordxDate").Range.Text = lblDate
                    objdoc.Bookmarks("BordxNo").Range.Text = lblBordxNo

                
        End If
        
        
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing

        Exit Sub
'HANDLE_ERROR:
'        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'        Set objWorksheet = Nothing
'        Set objWorkbook = Nothing
'        GoTo ExitFunction
End Sub

Private Sub GenerateStaffClaimRpt()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim strsql As String
Dim rst2 As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim tmpReceiptNo As String
Dim PVCnt As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHosName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double
Dim tmptax As Double
Dim tmpTotalMBM As Double
Dim tmpTotalTax As Double
Dim tmpTotalDueMBM As Double

    Screen.MousePointer = vbHourglass

    If lblBordxNo = "" Then
    
        MsgBox "Please Select Bordx", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\PayAdviceStaffclaim.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        'On Error GoTo HANDLE_ERROR
        RecFound = False
        strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "CLM_PayAdviceHOS"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 5)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        tmptax = 0
        PVCnt = 0
        tmpTotalCases = 0
        tmpTotalActual = 0
        tmpTotalHos = 0
        tmpHOSPayee = ""
        tmpSeqNo = 0
        tmpTotalMBM = 0
        tmpTotalTax = 0
        tmpTotalDueMBM = 0
        tmptax = 0
        RecFound = False
        Do While Not rst2.EOF
            RecFound = True
            
            If PVCnt = 1 Then
                objWorksheet.Cells(2, "B") = lblBordxNo
                objWorksheet.Cells(3, "B") = lblDate
            End If
                If Format(rst2!CLMTax, "#,##0.00") <> Format(rst2!CLMPayableByMedix2M, "#,##0.00") Then
                    PVCnt = PVCnt + 1
                    objWorksheet.Cells(6 + CDbl(PVCnt), "A") = rst2!MBMEmployeeNo
                    objWorksheet.Cells(6 + CDbl(PVCnt), "B") = rst2!MBMName
                    objWorksheet.Cells(6 + CDbl(PVCnt), "C") = rst2!MBMEmail
                    objWorksheet.Cells(6 + CDbl(PVCnt), "D") = Format(CDbl(rst2!CLMPayableByMedix2M) + CDbl(tmptax), "#,##0.00")
                    objWorksheet.Cells(6 + CDbl(PVCnt), "E") = rst2!CasNumber & "-" & rst2!claimversion
                    objWorksheet.Cells(6 + CDbl(PVCnt), "F") = lblDate
                End If
            rst2.MoveNext
        Loop

        rst2.Close
        Set rst2 = Nothing
        
        
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing

        Exit Sub
'HANDLE_ERROR:
'        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'        Set objWorksheet = Nothing
'        Set objWorkbook = Nothing
'        GoTo ExitFunction
End Sub

Private Sub GenerateDailyAllwRpt()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim strsql As String
Dim rst2 As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim tmpReceiptNo As String
Dim PVCnt As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHosName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double
Dim tmptax As Double
Dim tmpTotalMBM As Double
Dim tmpTotalTax As Double
Dim tmpTotalDueMBM As Double

    Screen.MousePointer = vbHourglass

    If lblBordxNo = "" Then
    
        MsgBox "Please Select Bordx", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\PayAdviceDailyAllw.XLT")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        'On Error GoTo HANDLE_ERROR
        RecFound = False
        strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "CLM_PayAdviceHOS"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 5)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        tmptax = 0
        PVCnt = 0
        tmpTotalCases = 0
        tmpTotalActual = 0
        tmpTotalHos = 0
        tmpHOSPayee = ""
        tmpSeqNo = 0
        tmpTotalMBM = 0
        tmpTotalTax = 0
        tmpTotalDueMBM = 0
        tmptax = 0
        RecFound = False
        Do While Not rst2.EOF
            RecFound = True
           ' PVCnt = PVCnt + 1
            If PVCnt = 1 Then
                objWorksheet.Cells(2, "B") = lblBordxNo
                objWorksheet.Cells(3, "B") = lblDate
            End If
            
                If Format(rst2!CLMTax, "#,##0.00") <> Format(rst2!CLMPayableByMedix2M, "#,##0.00") Then
                Else
                    PVCnt = PVCnt + 1
                    objWorksheet.Cells(6 + CDbl(PVCnt), "A") = rst2!MBMEmployeeNo
                    objWorksheet.Cells(6 + CDbl(PVCnt), "B") = rst2!MBMName
                    objWorksheet.Cells(6 + CDbl(PVCnt), "C") = rst2!MBMEmail
                    objWorksheet.Cells(6 + CDbl(PVCnt), "D") = Format(rst2!CLMTax, "###0.00")
                    objWorksheet.Cells(6 + CDbl(PVCnt), "E") = rst2!CasNumber & rst2!claimversion
                    objWorksheet.Cells(6 + CDbl(PVCnt), "F") = Format(lblDate, "dd-mmm-yyyy")
                End If
            
            rst2.MoveNext
        Loop
            'objWorksheet.Cells(15, "J") = Format(tmpTotalMBM, "#,##0.00")
            'objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
            'objWorksheet.Cells(15, "H") = Format(tmpTotalHos, "#,##0.00")
            'objWorksheet.Cells(15, "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
            'objWorksheet.Cells(15, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
            'objWorksheet.Cells(15, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
            'objWorksheet.Cells(15, "H").Borders(xlEdgeBottom).LineStyle = xlContinuous
            
            objWorksheet.Cells(14 + CDbl(PVCnt), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(14 + CDbl(PVCnt), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(14 + CDbl(PVCnt), "N").Font.Bold = True
            objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(tmpTotalMBM, "#,##0.00")
            objWorksheet.Cells(14 + CDbl(PVCnt), "O").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(14 + CDbl(PVCnt), "O").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(14 + CDbl(PVCnt), "O").Font.Bold = True
            objWorksheet.Cells(14 + CDbl(PVCnt), "O") = Format(tmpTotalTax, "#,##0.00")
            objWorksheet.Cells(14 + CDbl(PVCnt), "P").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(14 + CDbl(PVCnt), "P").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(14 + CDbl(PVCnt), "P").Font.Bold = True
            objWorksheet.Cells(14 + CDbl(PVCnt), "P") = Format(0, "#,##0.00")
            objWorksheet.Cells(14 + CDbl(PVCnt), "Q").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(14 + CDbl(PVCnt), "Q").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(14 + CDbl(PVCnt), "Q").Font.Bold = True
            objWorksheet.Cells(14 + CDbl(PVCnt), "Q") = Format(tmpTotalDueMBM, "#,##0.00")

        rst2.Close
        Set rst2 = Nothing
        
        
            If RecFound = True Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Claim\" & "\CoverStaffPayment.DOT")
                End If
                
                    'bookmark = False
                    objdoc.Bookmarks("TotalInv").Range.Text = tmpTotalCases
                    objdoc.Bookmarks("BordxDate").Range.Text = lblDate
                    objdoc.Bookmarks("BordxNo").Range.Text = lblBordxNo
                
        End If
        
        
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing

        Exit Sub
'HANDLE_ERROR:
'        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'        Set objWorksheet = Nothing
'        Set objWorkbook = Nothing
'        GoTo ExitFunction
End Sub

Private Sub GenerateAWUSRpt()
 Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim strsql As String
Dim rst2 As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim tmpReceiptNo As String
Dim PVCnt As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHosName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double
Dim strsqlCOV As String
Dim RstCov As New ADODB.Recordset
Dim tmpPatientName As String
Dim tmpSubMBMPaid As Double
Dim totalMBMPaid As Double

    Screen.MousePointer = vbHourglass

    If lblBordxNo = "" Then
    
        MsgBox "Please Select Bordx", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\PayAdviceSummary.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        'On Error GoTo HANDLE_ERROR
        RecFound = False
        strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "' And (MBMGroupCategory <> 'MASWING SDN BHD') And (HOSState = 'Sabah') and INVPayTo = 'H'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "CLM_PayAdviceHOS"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 2)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        PVCnt = 0
        tmpTotalCases = 0
        tmpTotalActual = 0
        tmpTotalHos = 0
        tmpHOSPayee = ""
        tmpSeqNo = 0
        tmpSubMBMPaid = 0
        totalMBMPaid = 0
        RecFound = False
        Do While Not rst2.EOF
            RecFound = True
            PVCnt = PVCnt + 1
            If PVCnt = 1 Then
                objWorksheet.Cells(7, "J") = lblBordxNo
                objWorksheet.Cells(8, "J") = lblDate
            End If
                tmpSeqNo = CDbl(tmpSeqNo) + 1
                objWorksheet.Cells(16 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                objWorksheet.Cells(16 + CDbl(PVCnt), "B") = rst2!HosName
                objWorksheet.Cells(16 + CDbl(PVCnt), "E") = rst2!TotalCases
                tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                objWorksheet.Cells(16 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                tmpTotalActual = tmpTotalActual + Format(CDbl(rst2!TotalActual), "#,##0.00")
                objWorksheet.Cells(16 + CDbl(PVCnt), "G") = CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS)
                totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                objWorksheet.Cells(16 + CDbl(PVCnt), "H") = Format(rst2!PayabletoHOS, "#,##0.00")
                objWorksheet.Cells(16 + CDbl(PVCnt), "J").Borders.LineStyle = True
                objWorksheet.Cells(16 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
            rst2.MoveNext
        Loop
            objWorksheet.Cells(15, "E") = Format(tmpTotalCases, "#,##0.00")
            objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
            objWorksheet.Cells(15, "G") = Format(totalMBMPaid, "#,##0.00")
            
            objWorksheet.Cells(15, "H") = Format(tmpTotalHos, "#,##0.00")
            'objWorksheet.Cells(8, "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
            objWorksheet.Cells(15, "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
            objWorksheet.Cells(15, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
            objWorksheet.Cells(15, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
            objWorksheet.Cells(15, "H").Borders(xlEdgeBottom).LineStyle = xlContinuous
            
            objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "E").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
            objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "F").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
            objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "G").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(totalMBMPaid, "#,##0.00")
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "H") = Format(tmpTotalHos, "#,##0.00")

        rst2.Close
        Set rst2 = Nothing
        
        
            If RecFound = True Then
                strsql = ""
                strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "' And (MBMGroupCategory <> 'MASWING SDN BHD') And (HOSState = 'Sabah') and INVPayTo = 'H'"
                Set cmd2.ActiveConnection = medixmasterconnWS
                cmd2.CommandText = "CLM_PayAdviceHOS"
                cmd2.CommandType = adCmdStoredProc
                cmd2.CommandTimeout = 300
                Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd2.Parameters.Append Prm3
                Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 1)
                cmd2.Parameters.Append prm4
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd2.Execute
            
            
                PVCnt = 0
                'tmpTotalCases = 0
                tmpTotalActual = 0
                tmpTotalHos = 0
                TMPWS = 1
                tmpSeqNo = 0
                Do While Not rst2.EOF
                    PVCnt = PVCnt + 1
                    If PVCnt = 1 Then
                        If TMPWS <> 1 Then
                            TMPWS = TMPWS - 1
                            tmpSeqNo = CDbl(tmpSeqNo) + 1
                            Set objWorksheet = objWorkbook.Sheets(2)
                            objExcel.DisplayAlerts = False
                            Set objWorksheet = objWorkbook.Worksheets.Item(2)
                            objWorksheet.Cells(2, "L") = rst2!HOSCode & "-" & lblBordxNo & "-" & Format(tmpSeqNo, "000") & "-" & Format(lblDate, "yy")
                            objWorksheet.Cells(3, "L") = lblDate

                            objWorksheet.Cells(7, "A") = rst2!HosName
                            objWorksheet.Cells(8, "A") = rst2!HOSAdd1
                            objWorksheet.Cells(9, "A") = rst2!HOSAdd2
                            objWorksheet.Cells(10, "A") = rst2!HOSAdd3
                            objWorksheet.Cells(11, "A") = "Tel : " & rst2!HOSTelNo & "   " & "Fax : " & rst2!HOSFaxNo
                            
                            tmpsubAct = 0
                            tmpSubApp = 0
                            tmpSubMBMPaid = 0
                        End If
                    End If
                        'If objWorksheet.Cells(7, "A") = rst2!HOSName Then
                        '    tmpHOSPayee = rst2!HOSPayee
                        'End If
                        If objWorksheet.Cells(7, "A") <> rst2!HosName Then
                            If PVCnt <> 1 Then
                                PVCnt = PVCnt + 1
                                strsqlHOS = "Select HOSPayee from HOS where HOSName = '" & objWorksheet.Cells(7, "A") & "'"
                                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
       
                                If HOS.recordCount Then
                                    tmpHOSPayee = HOS!HOSPayee
                                Else
                                    tmpHOSPayee = ""
                                End If
                                HOS.Close
                                Set HOS = Nothing
                                
                                objWorksheet.Cells(14 + CDbl(PVCnt), "L").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "L").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "L") = Format(tmpsubAct, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M") = Format(tmpSubMBMPaid, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "N").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(tmpSubApp, "#,##0.00")
                                
                                objWorksheet.Cells(14 + CDbl(PVCnt), "L") = Format(tmpsubAct, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M") = Format(tmpSubMBMPaid, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(tmpSubApp, "#,##0.00")
                                PVCnt = PVCnt + 1
                                objWorksheet.Cells(16 + CDbl(PVCnt), "B") = "Bank Simpanan Nasional"
                                
                                objWorksheet.Cells(20 + CDbl(PVCnt), "B") = "Verified By"
                                objWorksheet.Cells(20 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(20 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(20 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(20 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                
                                objWorksheet.Cells(24 + CDbl(PVCnt), "B") = "Approved By"
                                objWorksheet.Cells(24 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(24 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(24 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(24 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                
                                
                                objWorksheet.Cells(28 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(28 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(28 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(28 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(28 + CDbl(PVCnt), "B") = "Bank"
                                objWorksheet.Cells(32 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(32 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(32 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(32 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(32 + CDbl(PVCnt), "B") = "Date"
                                objWorksheet.Cells(26 + CDbl(PVCnt), "B") = "Payable to  "
                                objWorksheet.Cells(26 + CDbl(PVCnt), "C") = tmpHOSPayee
                            End If
                                TMPWS = TMPWS + 1
                                
                                Set objWorksheet = objWorkbook.Sheets(TMPWS)
                                objExcel.DisplayAlerts = False
                                Set objWorksheet = objWorkbook.Worksheets.Item(TMPWS)
                                tmpSeqNo = CDbl(tmpSeqNo) + 1
                                objWorksheet.Cells(2, "L") = rst2!HOSCode & "-" & lblBordxNo & "-" & Format(tmpSeqNo, "000") & "-" & Format(lblDate, "yy")
                                objWorksheet.Cells(3, "L") = lblDate
                                
                                objWorksheet.Cells(7, "A") = rst2!HosName
                                objWorksheet.Cells(8, "A") = rst2!HOSAdd1
                                objWorksheet.Cells(9, "A") = rst2!HOSAdd2
                                objWorksheet.Cells(10, "A") = rst2!HOSAdd3
                                objWorksheet.Cells(11, "A") = "Tel : " & rst2!HOSTelNo & "   " & "Fax : " & rst2!HOSFaxNo
                                tmpsubAct = 0
                                tmpSubApp = 0
                                tmpSubMBMPaid = 0
                                PVCnt = 0
                                PVCnt = PVCnt + 1
                        End If
                        
                        If Trim(rst2!MBMPatientCovID) <> "00" Then
                            strsqlCOV = "Select MBMCName from MBMCoveredPersons where MBMCNumber = '" & Trim(rst2!MBMNumber) & "' and MBMCCoverID = '" & Trim(rst2!MBMPatientCovID) & "'"
                            RstCov.Open strsqlCOV, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                                Do While Not RstCov.EOF
                                    tmpPatientName = IIf(Len(RstCov!MBMCName), RstCov!MBMCName, "")
                                    RstCov.MoveNext
                                Loop
                                RstCov.Close
                                Set RstCov = Nothing
                                
                            Else
                                tmpPatientName = rst2!MBMName
                        End If
                        
                        objWorksheet.Cells(14 + CDbl(PVCnt), "A") = PVCnt
                        objWorksheet.Cells(14 + CDbl(PVCnt), "B") = rst2!CasNumber & "-" & rst2!claimversion
                        objWorksheet.Cells(14 + CDbl(PVCnt), "C") = Format(rst2!CASDischargeDate, "YYYY/MM/DD")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "D") = rst2!MBMEmployeeNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "E") = rst2!MBMName
                        objWorksheet.Cells(14 + CDbl(PVCnt), "H") = tmpPatientName
                        objWorksheet.Cells(14 + CDbl(PVCnt), "K") = rst2!INVNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "L") = Format(rst2!CLMTotalActual, "#,##0.00")
                        tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                        tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMPayableByMedix2H, "#,##0.00"))
                        objWorksheet.Cells(14 + CDbl(PVCnt), "M") = Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                        tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMPayableByMedix2H)
                        
                    rst2.MoveNext
                Loop
                rst2.Close
                Set rst2 = Nothing
                RecFound = False
                If RecFound = False Then
                    strsqlHOS = "Select HOSPayee from HOS where HOSName = '" & objWorksheet.Cells(7, "A") & "'"
                    HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic

                    If HOS.recordCount Then
                        tmpHOSPayee = HOS!HOSPayee
                    Else
                        tmpHOSPayee = ""
                    End If
                    HOS.Close
                    Set HOS = Nothing

                    PVCnt = PVCnt + 2
                    objWorksheet.Cells(14 + CDbl(PVCnt), "L").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "L").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "L") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "M").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "M") = Format(tmpSubMBMPaid, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "N").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(tmpSubApp, "#,##0.00")
                    
                    objWorksheet.Cells(14 + CDbl(PVCnt), "L") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "M") = Format(tmpSubMBMPaid, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(tmpSubApp, "#,##0.00")
                    PVCnt = PVCnt + 1
                    objWorksheet.Cells(16 + CDbl(PVCnt), "B") = "Bank Simpanan Nasioanl"
                    objWorksheet.Cells(18 + CDbl(PVCnt), "B") = "Verified By"
                    objWorksheet.Cells(18 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    objWorksheet.Cells(20 + CDbl(PVCnt), "B") = "Verified By"
                    objWorksheet.Cells(20 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(20 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(20 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(20 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    objWorksheet.Cells(22 + CDbl(PVCnt), "B") = "Approved By"
                    objWorksheet.Cells(22 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(22 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(22 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(22 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
                 
                    objWorksheet.Cells(28 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(28 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(28 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(28 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(28 + CDbl(PVCnt), "B") = "Bank"
                    objWorksheet.Cells(32 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(32 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(32 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(32 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(32 + CDbl(PVCnt), "B") = "Date"
                    objWorksheet.Cells(22 + CDbl(PVCnt), "B") = "Payable to  "
                    objWorksheet.Cells(22 + CDbl(PVCnt), "C") = tmpHOSPayee
                End If
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Claim\" & "\CoverPanelPayment.DOT")
                End If
                
                    'bookmark = False
                    objdoc.Bookmarks("TotalInv").Range.Text = tmpTotalCases
                    objdoc.Bookmarks("BordxDate").Range.Text = lblDate
                    objdoc.Bookmarks("BordxNo").Range.Text = lblBordxNo
                    objdoc.Bookmarks("Remarks").Range.Text = " - SABAH"
                    objdoc.Bookmarks("Type").Range.Text = "B"
                
        End If
        
        
        'Call GenerateMASWingRpt
        'Call GenerateCCReports
        'medixmasterconnPayment.Close
        'Set medixmasterconnPayment = Nothing
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing

        Exit Sub
'HANDLE_ERROR:
'        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'        Set objWorksheet = Nothing
'        Set objWorkbook = Nothing
'        GoTo ExitFunction
End Sub

Private Sub GenerateSarawakRpt()
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim strsql As String
Dim rst2 As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim tmpReceiptNo As String
Dim PVCnt As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHosName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double
Dim strsqlCOV As String
Dim RstCov As New ADODB.Recordset
Dim tmpPatientName As String
Dim tmpSubMBMPaid As Double
Dim totalMBMPaid As Double

    Screen.MousePointer = vbHourglass

    If lblBordxNo = "" Then
    
        MsgBox "Please Select Bordx", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\PayAdviceSummary.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        'On Error GoTo HANDLE_ERROR
        RecFound = False
        strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "' And (MBMGroupCategory <> 'MASWING SDN BHD') And (HOSState = 'Sarawak') and INVPayTo = 'H'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "CLM_PayAdviceHOS"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 2)
        cmd.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        PVCnt = 0
        tmpTotalCases = 0
        tmpTotalActual = 0
        tmpTotalHos = 0
        tmpHOSPayee = ""
        tmpSeqNo = 0
        tmpSubMBMPaid = 0
        RecFound = False
        Do While Not rst2.EOF
            RecFound = True
            PVCnt = PVCnt + 1
            If PVCnt = 1 Then
                objWorksheet.Cells(7, "J") = lblBordxNo
                objWorksheet.Cells(8, "J") = lblDate
            End If
                tmpSeqNo = CDbl(tmpSeqNo) + 1
                objWorksheet.Cells(16 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "000")
                objWorksheet.Cells(16 + CDbl(PVCnt), "B") = rst2!HosName
                objWorksheet.Cells(16 + CDbl(PVCnt), "E") = rst2!TotalCases
                tmpTotalCases = tmpTotalCases + CDbl(rst2!TotalCases)
                objWorksheet.Cells(16 + CDbl(PVCnt), "F") = Format(rst2!TotalActual, "#,##0.00")
                tmpTotalActual = tmpTotalActual + Format(CDbl(rst2!TotalActual), "#,##0.00")
                objWorksheet.Cells(16 + CDbl(PVCnt), "G") = CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS)
                totalMBMPaid = CDbl(totalMBMPaid) + (CDbl(rst2!TotalActual) - CDbl(rst2!PayabletoHOS))
                objWorksheet.Cells(16 + CDbl(PVCnt), "H") = Format(rst2!PayabletoHOS, "#,##0.00")
                objWorksheet.Cells(16 + CDbl(PVCnt), "J").Borders.LineStyle = True
                objWorksheet.Cells(16 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
            rst2.MoveNext
        Loop
            objWorksheet.Cells(15, "E") = Format(tmpTotalCases, "#,##0.00")
            objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
            objWorksheet.Cells(15, "G") = Format(totalMBMPaid, "#,##0.00")
            
            objWorksheet.Cells(15, "H") = Format(tmpTotalHos, "#,##0.00")
            'objWorksheet.Cells(8, "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
            objWorksheet.Cells(15, "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
            objWorksheet.Cells(15, "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
            objWorksheet.Cells(15, "G").Borders(xlEdgeBottom).LineStyle = xlContinuous
            objWorksheet.Cells(15, "H").Borders(xlEdgeBottom).LineStyle = xlContinuous
            
            objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "E").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "E") = Format(tmpTotalCases, "#,##0.00")
            objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "F").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
            objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "G").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "G").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(totalMBMPaid, "#,##0.00")
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "H") = Format(tmpTotalHos, "#,##0.00")

        rst2.Close
        Set rst2 = Nothing
        
        
            If RecFound = True Then
                strsql = ""
                strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "' And (MBMGroupCategory <> 'MASWING SDN BHD') And (HOSState = 'Sarawak') and INVPayTo = 'H'"
                Set cmd2.ActiveConnection = medixmasterconnWS
                cmd2.CommandText = "CLM_PayAdviceHOS"
                cmd2.CommandType = adCmdStoredProc
                cmd2.CommandTimeout = 300
                Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd2.Parameters.Append Prm3
                Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 1)
                cmd2.Parameters.Append prm4
                rst2.CursorLocation = adUseClient
                rst2.CursorType = adOpenForwardOnly
                rst2.CacheSize = 100
                Set rst2 = cmd2.Execute
            
            
                PVCnt = 0
                'tmpTotalCases = 0
                tmpTotalActual = 0
                tmpTotalHos = 0
                TMPWS = 1
                tmpSeqNo = 0
                Do While Not rst2.EOF
                    PVCnt = PVCnt + 1
                    If PVCnt = 1 Then
                        If TMPWS <> 1 Then
                        TMPWS = TMPWS - 1
                            tmpSeqNo = CDbl(tmpSeqNo) + 1
                            Set objWorksheet = objWorkbook.Sheets(2)
                            objExcel.DisplayAlerts = False
                            Set objWorksheet = objWorkbook.Worksheets.Item(2)
                            objWorksheet.Cells(2, "L") = rst2!HOSCode & "-" & lblBordxNo & "-" & Format(tmpSeqNo, "000") & "-" & Format(lblDate, "yy")
                            objWorksheet.Cells(3, "L") = lblDate

                            objWorksheet.Cells(7, "A") = rst2!HosName
                            objWorksheet.Cells(8, "A") = rst2!HOSAdd1
                            objWorksheet.Cells(9, "A") = rst2!HOSAdd2
                            objWorksheet.Cells(10, "A") = rst2!HOSAdd3
                            objWorksheet.Cells(11, "A") = "Tel : " & rst2!HOSTelNo & "   " & "Fax : " & rst2!HOSFaxNo
                            
                            tmpsubAct = 0
                            tmpSubApp = 0
                            tmpSubMBMPaid = 0
                        End If
                    End If
                        'If objWorksheet.Cells(7, "A") = rst2!HOSName Then
                        '    tmpHOSPayee = rst2!HOSPayee
                        'End If
                        If objWorksheet.Cells(7, "A") <> rst2!HosName Then
                            If PVCnt <> 1 Then
                                PVCnt = PVCnt + 1
                                strsqlHOS = "Select HOSPayee from HOS where HOSName = '" & objWorksheet.Cells(7, "A") & "'"
                                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
       
                                If HOS.recordCount Then
                                    tmpHOSPayee = HOS!HOSPayee
                                Else
                                    tmpHOSPayee = ""
                                End If
                                HOS.Close
                                Set HOS = Nothing
                                
                                objWorksheet.Cells(14 + CDbl(PVCnt), "L").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "L").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "L") = Format(tmpsubAct, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M") = Format(tmpSubMBMPaid, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "N").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(tmpSubApp, "#,##0.00")
                                
                                objWorksheet.Cells(14 + CDbl(PVCnt), "L") = Format(tmpsubAct, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "M") = Format(tmpSubMBMPaid, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(tmpSubApp, "#,##0.00")
                                PVCnt = PVCnt + 1
                                objWorksheet.Cells(16 + CDbl(PVCnt), "B") = "Bank Simpanan Nasioanl"
                                objWorksheet.Cells(20 + CDbl(PVCnt), "B") = "Verified By"
                                objWorksheet.Cells(20 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(20 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(20 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(20 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                
                                objWorksheet.Cells(24 + CDbl(PVCnt), "B") = "Approved By"
                                objWorksheet.Cells(24 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(24 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(24 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(24 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                
                                objWorksheet.Cells(30 + CDbl(PVCnt), "B") = "Cheque No"
                                objWorksheet.Cells(30 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(30 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(30 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(30 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                
                                objWorksheet.Cells(28 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(28 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(28 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(28 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(28 + CDbl(PVCnt), "B") = "Bank"
                                objWorksheet.Cells(32 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(32 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(32 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(32 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                objWorksheet.Cells(32 + CDbl(PVCnt), "B") = "Date"
                                objWorksheet.Cells(26 + CDbl(PVCnt), "B") = "Payable to  "
                                objWorksheet.Cells(26 + CDbl(PVCnt), "C") = tmpHOSPayee
                            End If
                                TMPWS = TMPWS + 1
                                
                                Set objWorksheet = objWorkbook.Sheets(TMPWS)
                                objExcel.DisplayAlerts = False
                                Set objWorksheet = objWorkbook.Worksheets.Item(TMPWS)
                                tmpSeqNo = CDbl(tmpSeqNo) + 1
                                objWorksheet.Cells(2, "L") = rst2!HOSCode & "-" & lblBordxNo & "-" & Format(tmpSeqNo, "000") & "-" & Format(lblDate, "yy")
                                objWorksheet.Cells(3, "L") = lblDate
                                
                                objWorksheet.Cells(7, "A") = rst2!HosName
                                objWorksheet.Cells(8, "A") = rst2!HOSAdd1
                                objWorksheet.Cells(9, "A") = rst2!HOSAdd2
                                objWorksheet.Cells(10, "A") = rst2!HOSAdd3
                                objWorksheet.Cells(11, "A") = "Tel : " & rst2!HOSTelNo & "   " & "Fax : " & rst2!HOSFaxNo
                                tmpsubAct = 0
                                tmpSubApp = 0
                                tmpSubMBMPaid = 0
                                PVCnt = 0
                                PVCnt = PVCnt + 1
                        End If
                        
                        If Trim(rst2!MBMPatientCovID) <> "00" Then
                            strsqlCOV = ""
                            strsqlCOV = "Select MBMCName from MBMCoveredPersons where MBMCNumber = '" & Trim(rst2!MBMNumber) & "' and MBMCCoverID = '" & Trim(rst2!MBMPatientCovID) & "'"
                            RstCov.Open strsqlCOV, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                               ' RSTCOV.Close
                               ' Set RSTCOV = Nothing
                                Do While Not RstCov.EOF
                                    tmpPatientName = IIf(Len(RstCov!MBMCName), RstCov!MBMCName, "")
                                    RstCov.MoveNext
                                Loop
                                    RstCov.Close
                                    Set RstCov = Nothing

                            Else
                                tmpPatientName = rst2!MBMName
                        End If
                        
                        objWorksheet.Cells(14 + CDbl(PVCnt), "A") = PVCnt
                        objWorksheet.Cells(14 + CDbl(PVCnt), "B") = rst2!CasNumber & "-" & rst2!claimversion
                        objWorksheet.Cells(14 + CDbl(PVCnt), "C") = Format(rst2!CASDischargeDate, "YYYY/MM/DD")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "D") = rst2!MBMEmployeeNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "E") = rst2!MBMName
                        objWorksheet.Cells(14 + CDbl(PVCnt), "H") = tmpPatientName
                        objWorksheet.Cells(14 + CDbl(PVCnt), "K") = rst2!INVNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "L") = Format(rst2!CLMTotalActual, "#,##0.00")
                        tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                        tmpSubMBMPaid = CDbl(tmpSubMBMPaid) + (Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMPayableByMedix2H, "#,##0.00"))
                        objWorksheet.Cells(14 + CDbl(PVCnt), "M") = Format(rst2!CLMTotalActual, "#,##0.00") - Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                        tmpSubApp = CDbl(tmpSubApp) + CDbl(rst2!CLMPayableByMedix2H)
                        
                    rst2.MoveNext
                Loop
                rst2.Close
                Set rst2 = Nothing
                RecFound = False
                If RecFound = False Then
                    strsqlHOS = "Select HOSPayee from HOS where HOSName = '" & objWorksheet.Cells(7, "A") & "'"
                    HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic

                    If HOS.recordCount Then
                        tmpHOSPayee = HOS!HOSPayee
                    Else
                        tmpHOSPayee = ""
                    End If
                    HOS.Close
                    Set HOS = Nothing

                    PVCnt = PVCnt + 2
                    objWorksheet.Cells(14 + CDbl(PVCnt), "L").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "L").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "L") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "M").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "M") = Format(tmpSubMBMPaid, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "N").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "N").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "N").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(tmpSubApp, "#,##0.00")
                    
                    objWorksheet.Cells(14 + CDbl(PVCnt), "L") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "M") = Format(tmpSubMBMPaid, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "N") = Format(tmpSubApp, "#,##0.00")
                    PVCnt = PVCnt + 1
                    objWorksheet.Cells(16 + CDbl(PVCnt), "B") = "Bank Simpanan Nasioanl"
                    objWorksheet.Cells(18 + CDbl(PVCnt), "B") = "Verified By"
                    objWorksheet.Cells(18 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(18 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    objWorksheet.Cells(20 + CDbl(PVCnt), "B") = "Verified By"
                    objWorksheet.Cells(20 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(20 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(20 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(20 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    objWorksheet.Cells(22 + CDbl(PVCnt), "B") = "Approved By"
                    objWorksheet.Cells(22 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(22 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(22 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(22 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    objWorksheet.Cells(30 + CDbl(PVCnt), "B") = "Cheque No"
                    objWorksheet.Cells(30 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(30 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(30 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(30 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    
                    objWorksheet.Cells(28 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(28 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(28 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(28 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(28 + CDbl(PVCnt), "B") = "Bank"
                    objWorksheet.Cells(32 + CDbl(PVCnt), "C").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(32 + CDbl(PVCnt), "D").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(32 + CDbl(PVCnt), "E").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(32 + CDbl(PVCnt), "F").Borders(xlEdgeBottom).LineStyle = xlContinuous
                    objWorksheet.Cells(32 + CDbl(PVCnt), "B") = "Date"
                    objWorksheet.Cells(22 + CDbl(PVCnt), "B") = "Payable to  "
                    objWorksheet.Cells(22 + CDbl(PVCnt), "C") = tmpHOSPayee
                End If
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Claim\" & "\CoverPanelPayment.DOT")
                End If
                
                    objdoc.Bookmarks("TotalInv").Range.Text = tmpTotalCases
                    objdoc.Bookmarks("BordxDate").Range.Text = lblDate
                    objdoc.Bookmarks("BordxNo").Range.Text = lblBordxNo
                    objdoc.Bookmarks("Remarks").Range.Text = " - SAWARAK"
                    objdoc.Bookmarks("Type").Range.Text = "C"
                
        End If
        
        
        'Call GenerateMASWingRpt
        'Call GenerateCCReports
        'medixmasterconnPayment.Close
        'Set medixmasterconnPayment = Nothing
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing

        Exit Sub
'HANDLE_ERROR:
'        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'        Set objWorksheet = Nothing
'        Set objWorkbook = Nothing
'        GoTo ExitFunction
End Sub


Private Sub GenerateCCReports(Optional strAppend As String)
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim strsql As String
Dim rst2 As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim tmpReceiptNo As String
Dim PVCnt As Integer
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim cmd2 As New ADODB.Command
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHosName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpGrandTotal As Double
Dim cmd3 As New ADODB.Command
Dim Prm5 As New ADODB.Parameter
Dim prm6 As New ADODB.Parameter
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim prm8 As New ADODB.Parameter

    Screen.MousePointer = vbHourglass

        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\MASReport.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
            'On Error GoTo HANDLE_ERROR
            
            'Summary of Other Cost Center
            'strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
            strsql = strAppend
            Set cmd.ActiveConnection = medixmasterconnWS
            cmd.CommandText = "CLM_PayAdviceHOS"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
            cmd.Parameters.Append Prm1
            Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 3)
            cmd.Parameters.Append Prm2
            rst2.CursorLocation = adUseClient
            rst2.CursorType = adOpenForwardOnly
            rst2.CacheSize = 100
            Set rst2 = cmd.Execute
            
            PVCnt = 0
            tmpTotalCases = 0
            tmpTotalActual = 0
            tmpTotalHos = 0
            tmpHOSPayee = ""
            RecFound = False
            Do While Not rst2.EOF
                RecFound = True
                PVCnt = PVCnt + 1
                    
                    objWorksheet.Cells(5 + CDbl(PVCnt), "A") = "10HEQ2510"
                    objWorksheet.Cells(5 + CDbl(PVCnt), "B") = "95010000"
                    objWorksheet.Cells(5 + CDbl(PVCnt), "C") = Format(rst2!TotalAmt, "#,##0.00")
                    objWorksheet.Cells(5 + CDbl(PVCnt), "D") = "MYR"
                    objWorksheet.Cells(5 + CDbl(PVCnt), "E") = rst2!MBMCostCenter
                    objWorksheet.Cells(5 + CDbl(PVCnt), "F") = rst2!TotalCase
                    objWorksheet.Cells(5 + CDbl(PVCnt), "G") = "EA"
                    objWorksheet.Cells(5 + CDbl(PVCnt), "H") = "Cost Allocation-GHSS " & Month(Date) & Year(Date)
                    'Cost Allocation-GHSS  January 2012
                    'objWorksheet.Cells(5 + CDbl(PVCnt), "L") = "JXXXXXX INS " & Month(Date) & Year(Date)
                   ' objWorksheet.Cells(5 + CDbl(PVCnt), "M") = "51702" & "00" & rst2!MBMCostCenter
                   ' objWorksheet.Cells(5 + CDbl(PVCnt), "O") = "JXXXXXX"
                   ' objWorksheet.Cells(5 + CDbl(PVCnt), "P") = "INS " & Month(Date) & Year(Date)
                rst2.MoveNext
            Loop
    
            rst2.Close
            Set rst2 = Nothing
            
            
            'Summary MASWing
            'strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
            strsql = strAppend
            Set cmd2.ActiveConnection = medixmasterconnWS
            cmd2.CommandText = "CLM_PayAdviceHOS"
            cmd2.CommandType = adCmdStoredProc
            cmd2.CommandTimeout = 300
            Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
            cmd2.Parameters.Append Prm3
            Set prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 4)
            cmd2.Parameters.Append prm4
            rst2.CursorLocation = adUseClient
            rst2.CursorType = adOpenForwardOnly
            rst2.CacheSize = 100
            Set rst2 = cmd2.Execute
            
            tmpTotalCases = 0
            tmpTotalActual = 0
            tmpTotalHos = 0
            tmpHOSPayee = ""
            RecFound = False
            Do While Not rst2.EOF
                RecFound = True
                PVCnt = PVCnt + 1
                    
                    objWorksheet.Cells(2 + CDbl(PVCnt), "A") = "51702"
                    objWorksheet.Cells(2 + CDbl(PVCnt), "B") = "00"
                    objWorksheet.Cells(2 + CDbl(PVCnt), "C") = Mid(rst2!MBMGroupCategory, 1, 7)
                    objWorksheet.Cells(2 + CDbl(PVCnt), "H") = Format(rst2!TotalAmt, "#,##0.00")
                    objWorksheet.Cells(2 + CDbl(PVCnt), "L") = "JXXXXXX INS " & Month(Date) & Year(Date)
                    objWorksheet.Cells(2 + CDbl(PVCnt), "M") = "51702" & "00" & Mid(rst2!MBMGroupCategory, 1, 7)
                    objWorksheet.Cells(2 + CDbl(PVCnt), "O") = "JXXXXXX"
                    objWorksheet.Cells(2 + CDbl(PVCnt), "P") = "INS " & Month(Date) & Year(Date)
                rst2.MoveNext
            Loop
    
            rst2.Close
            Set rst2 = Nothing
            
            Set objWorksheet = objWorkbook.Sheets(2)
            Set objWorksheet = objWorkbook.Worksheets.Item(2)

            'Detail of other Cost Center
            'strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "'"
            strsql = strAppend
            Set cmd3.ActiveConnection = medixmasterconnWS
            cmd3.CommandText = "CLM_PayAdviceHOS"
            cmd3.CommandType = adCmdStoredProc
            cmd3.CommandTimeout = 300
            Set Prm5 = cmd3.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
            cmd3.Parameters.Append Prm5
            Set prm6 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 7)
            cmd3.Parameters.Append prm6
            rst2.CursorLocation = adUseClient
            rst2.CursorType = adOpenForwardOnly
            rst2.CacheSize = 100
            Set rst2 = cmd3.Execute
            tmpGrandTotal = 0
            PVCnt = 0
            tmpTotalCases = 0
            tmpTotalActual = 0
            tmpTotalHos = 0
            tmpHOSPayee = ""
            RecFound = False
            Do While Not rst2.EOF
                RecFound = True
                PVCnt = PVCnt + 1
                
                    If PVCnt = "1" Then
                        objWorksheet.Cells(CDbl(PVCnt), "A") = "MAS Monthly Expenses for " & Month(Date) & "/" & Year(Date)
                        PVCnt = CDbl(PVCnt) + 1
                    End If
                    If PVCnt > "2" Then
                        If objWorksheet.Cells(3 + (CDbl(PVCnt) - 1), "C") <> "" And (rst2!MBMEmployeeNo <> objWorksheet.Cells(3 + (CDbl(PVCnt) - 1), "C")) Then
                            objWorksheet.Cells(3 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                            objWorksheet.Cells(3 + CDbl(PVCnt), "F").Font.Bold = True
                            tmpTotalActual = 0
                            PVCnt = CDbl(PVCnt) + 2
                        End If
                    End If
                    
                    If PVCnt = "2" And tmpTotalActual <> 0 Then
                        If objWorksheet.Cells(3 + (CDbl(PVCnt) - 1), "C") <> rst2!MBMEmployeeNo Then
                            objWorksheet.Cells(3 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                            tmpTotalActual = 0
                        End If
                    End If

                    objWorksheet.Cells(3 + CDbl(PVCnt), "B") = rst2!MBMCostCenter
                    objWorksheet.Cells(3 + CDbl(PVCnt), "C") = rst2!MBMEmployeeNo
                    objWorksheet.Cells(3 + CDbl(PVCnt), "D") = rst2!MBMName
                    objWorksheet.Cells(3 + CDbl(PVCnt), "E") = rst2!MBMIcBcPp
                    objWorksheet.Cells(3 + CDbl(PVCnt), "F") = Format(rst2!TotalAmt, "#,##0.00")
                    tmpGrandTotal = CDbl(tmpGrandTotal) + CDbl(rst2!TotalAmt)
                    
                    If PVCnt = "2" Then
                        tmpTotalActual = CDbl(tmpTotalActual) + CDbl(rst2!TotalAmt)
                    Else
                        tmpTotalActual = CDbl(tmpTotalActual) + CDbl(rst2!TotalAmt)
                    End If
                rst2.MoveNext
            Loop
    
            rst2.Close
            Set rst2 = Nothing
            PVCnt = CDbl(PVCnt) + 1
            objWorksheet.Cells(3 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
            objWorksheet.Cells(3 + CDbl(PVCnt), "F").Font.Bold = True
            tmpTotalActual = 0
            
            
            'Detail of MASWing
            strsql = strAppend
            Set cmd4.ActiveConnection = medixmasterconnWS
            cmd4.CommandText = "CLM_PayAdviceHOS"
            cmd4.CommandType = adCmdStoredProc
            cmd4.CommandTimeout = 300
            Set Prm7 = cmd4.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
            cmd4.Parameters.Append Prm7
            Set prm8 = cmd4.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 8)
            cmd4.Parameters.Append prm8
            rst2.CursorLocation = adUseClient
            rst2.CursorType = adOpenForwardOnly
            rst2.CacheSize = 100
            Set rst2 = cmd4.Execute
            tmpTotalActual = 0
            tmpTotalHos = 0
            tmpHOSPayee = ""
            'RecFound = False
            Do While Not rst2.EOF
                RecFound = True
                PVCnt = PVCnt + 2
                
                    'If PVCnt = "1" Then
                    '    objWorksheet.Cells(CDbl(PVCnt), "A") = "MAS Monthly Expenses for " & rst2!BordxMonth & "/" & rst2!BordxYear
                    '    PVCnt = CDbl(PVCnt) + 1
                    'End If
                    'If PVCnt > "2" Then
                        'If objWorksheet.Cells(3 + (CDbl(PVCnt) - 1), "B") <> "" And (rst2!MBMCostCenter <> objWorksheet.Cells(3 + (CDbl(PVCnt) - 1), "B")) Then
                        '    objWorksheet.Cells(3 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                        '    objWorksheet.Cells(3 + CDbl(PVCnt), "F").Font.Bold = True
                        '    tmpTotalActual = 0
                        '    PVCnt = CDbl(PVCnt) + 2
                        'End If
                    'End If
                    
                    'If PVCnt = "2" And tmpTotalActual <> 0 Then
                    '    If objWorksheet.Cells(3 + (CDbl(PVCnt) - 1), "B") <> rst2!MBMCostCenter Then
                    '        objWorksheet.Cells(3 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
                    '        tmpTotalActual = 0
                    '    End If
                    'End If

                    objWorksheet.Cells(3 + CDbl(PVCnt), "B") = Mid(rst2!MBMGroupCategory, 1, 7)
                    objWorksheet.Cells(3 + CDbl(PVCnt), "C") = rst2!MBMEmployeeNo
                    objWorksheet.Cells(3 + CDbl(PVCnt), "D") = rst2!MBMName
                    objWorksheet.Cells(3 + CDbl(PVCnt), "E") = rst2!MBMIcBcPp
                    objWorksheet.Cells(3 + CDbl(PVCnt), "F") = Format(rst2!TotalAmt, "#,##0.00")
                    tmpGrandTotal = CDbl(tmpGrandTotal) + CDbl(rst2!TotalAmt)
                    tmpTotalActual = CDbl(tmpTotalActual) + CDbl(rst2!TotalAmt)
                rst2.MoveNext
            Loop
    
            rst2.Close
            Set rst2 = Nothing
            PVCnt = CDbl(PVCnt) + 1
            objWorksheet.Cells(3 + CDbl(PVCnt), "F") = Format(tmpTotalActual, "#,##0.00")
            objWorksheet.Cells(3 + CDbl(PVCnt), "F").Font.Bold = True
            tmpTotalActual = 0
            
            PVCnt = CDbl(PVCnt) + 2
            objWorksheet.Cells(3 + CDbl(PVCnt), "E") = "Grand Total"
            objWorksheet.Cells(3 + CDbl(PVCnt), "E").Font.Bold = True
            objWorksheet.Cells(3 + CDbl(PVCnt), "F") = Format(tmpGrandTotal, "#,##0.00")
            objWorksheet.Cells(3 + CDbl(PVCnt), "F").Font.Bold = True
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        Exit Sub
'HANDLE_ERROR:
'        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'        Set objWorksheet = Nothing
'        Set objWorkbook = Nothing
'        GoTo ExitFunction
End Sub

Private Sub lstPrintBordx_Click()
    Call LvDisplayData
End Sub

Private Sub LvDisplayData(Optional Item As ListItem, Optional ItemCheck As Boolean)
Dim lstRec As ListItem

    If lstPrintBordx.ListItems.Count Then
        If Len(lstPrintBordx.SelectedItem.SubItems(11)) Then
            lblBordxNo = lstPrintBordx.SelectedItem.SubItems(11)
        End If

        If Len(lstPrintBordx.SelectedItem.SubItems(17)) Then
            lblDate = lstPrintBordx.SelectedItem.SubItems(17)
        End If
    End If
End Sub

Private Sub lstPrintBordx_KeyDown(KeyCode As Integer, Shift As Integer)
    Call LvDisplayData
End Sub

Private Sub lstPrintBordx_KeyUp(KeyCode As Integer, Shift As Integer)
    Call LvDisplayData
End Sub


