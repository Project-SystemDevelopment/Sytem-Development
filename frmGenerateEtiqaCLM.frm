VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGenerateEtiqaCLM 
   Caption         =   "Claims Data"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   7440
      Width           =   11655
      Begin VB.TextBox lblBordxNo 
         Height          =   285
         Left            =   5040
         TabIndex        =   37
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   2880
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   960
         TabIndex        =   35
         Top             =   240
         Width           =   855
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
         Height          =   135
         Left            =   7080
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   615
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblDate 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6240
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   11655
      Begin VB.TextBox TxtBordxNo 
         Height          =   285
         Left            =   4800
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   10320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox TxtBordxNoFr 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   150
         Width           =   2055
      End
      Begin VB.TextBox TxtBordxNoTo 
         Height          =   285
         Left            =   3840
         TabIndex        =   1
         Top             =   150
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
         ItemData        =   "frmGenerateEtiqaCLM.frx":0000
         Left            =   1320
         List            =   "frmGenerateEtiqaCLM.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   400
         Width           =   4575
      End
      Begin MSMask.MaskEdBox txtDOADate1 
         Height          =   255
         Left            =   7320
         TabIndex        =   4
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboCashType 
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
         ItemData        =   "frmGenerateEtiqaCLM.frx":0004
         Left            =   1320
         List            =   "frmGenerateEtiqaCLM.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   680
         Width           =   4575
      End
      Begin VB.Label Label10 
         Caption         =   "Cash Type"
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
         Left            =   240
         TabIndex        =   34
         Top             =   730
         Width           =   855
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
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
         Index           =   0
         Left            =   240
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   210
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstPrintBordx 
      Height          =   6135
      Left            =   120
      TabIndex        =   32
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
      NumItems        =   19
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
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Cash Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claims Data"
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
      TabIndex        =   33
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmGenerateEtiqaCLM"
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
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

        Call SearchCC
   ' medixmasterconn.Close
   ' medixmasterconnWS.Close
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
    
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    'If Trim(TxtBordxNoFr) = "" Or Trim(TxtBordxNoTo) = "" Then
    '    MsgBox "Please enter Bordx Reference No", vbOKOnly + vbCritical, DialogBoxTitle
    'Else
        'CmdClear.Enabled = False
        Call SearchRecord
    'End If
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

Private Sub CmdSearch_Click()
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
    SearchType = "2"
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    'If Trim(TxtBordxNoFr) = "" Or Trim(TxtBordxNoTo) = "" Then
    '    MsgBox "Please enter Bordx Reference No", vbOKOnly + vbCritical, DialogBoxTitle
    'Else
        'CmdClear.Enabled = False
        Call SearchRecord
    'End If
'    medixmasterconn.Close
'    Set medixmasterconn = Nothing
 '   medixmasterconnWS.Close
 '   Set medixmasterconnWS = Nothing
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
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call LoadComboListing(tmpSelField, tmpSelCriteria, tmpTable, cboHOS)
        'medixmasterconn.Close
       ' Set medixmasterconn = Nothing

    cboCashType.Clear
    cboCashType.AddItem "H - Hospital"
    cboCashType.AddItem "M - Member"
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
    
    strAppend = strAppend + " And (PAYCode = 'WM' OR PAYCode = 'WK') "
    
    If Len(Trim(cboCashType)) Then
        strAppend = strAppend + " And INVPAYTo = '" & Mid(cboCashType, 1, 1) & "'"
    End If
    
    
    If SearchType = "2" Then
        Call PrintBordxListing(strAppend)
    Else
        Call GeneratePaymentSum(strAppend)
    End If
End Function

Private Sub PrintBordxListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim PlanCheck As Boolean
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset

    Total = 0
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
        Total = rst2.recordCount

             With rst2
                Set Record = lstPrintBordx.ListItems.Add(, , rst2!PLNCode)
                If Len(rst2!CASNumber & "-" & rst2!ClaimVersion) Then
                    Record.SubItems(1) = (rst2!CASNumber & "-" & rst2!ClaimVersion)
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
                
                If Len(rst2!InvPayTo) Then
                    Record.SubItems(18) = rst2!InvPayTo
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

Private Sub GeneratePaymentSum(Optional strAppend As String)

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
Dim Prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHOSName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As String
Dim tmpAge As String
Dim tmpDocname As String
Dim DocPath As String
Dim DocSavePath As String

    Screen.MousePointer = vbHourglass

    'If lblBordxNo = "" Then
   '
   '     MsgBox "Please Select Bordx", vbOKOnly + vbCritical
   '     Screen.MousePointer = vbDefault
   '     Exit Sub
   ' End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\Etiqa Claims.xls")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        'On Error GoTo HANDLE_ERROR
        RecFound = False
        'strsql = ""
        'strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "' And (MBMGroupCategory <> 'MASWING SDN BHD') "
        
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "CLM_EtiqaClaims"
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
        
        PVCnt = 0
        tmpTotalCases = 0
        tmpTotalActual = 0
        tmpTotalHos = 0
        tmpHOSPayee = ""
        tmpSeqNo = 0
        RecFound = False
        Do While Not rst2.EOF
            RecFound = True
            PVCnt = PVCnt + 1
                tmpSeqNo = CDbl(tmpSeqNo) + 1
                objWorksheet.Cells(4 + CDbl(PVCnt), "A") = Format(tmpSeqNo, "0")
                objWorksheet.Cells(4 + CDbl(PVCnt), "D") = rst2!CLMLastDocDate
                'objWorksheet.Cells(4 + CDbl(PVCnt), "F") = rst2!CLMLastDocDate
                'objWorksheet.Cells(4 + CDbl(PVCnt), "G") = rst2!CasNumber & "-" & rst2!claimversion
                objWorksheet.Cells(4 + CDbl(PVCnt), "H") = "MEDIEXPRESS"
                objWorksheet.Cells(4 + CDbl(PVCnt), "I") = Mid(rst2!MBMPolicyNo, 1, 2)
                objWorksheet.Cells(4 + CDbl(PVCnt), "J") = Mid(rst2!MBMPolicyNo, 11, 8)
                objWorksheet.Cells(4 + CDbl(PVCnt), "K") = Mid(rst2!MBMPolicyNo, 7, 3)
                objWorksheet.Cells(4 + CDbl(PVCnt), "L") = rst2!MBMGroupCompany
                objWorksheet.Cells(4 + CDbl(PVCnt), "M") = "GROUP"
                objWorksheet.Cells(4 + CDbl(PVCnt), "O") = rst2!MBMName
                objWorksheet.Cells(4 + CDbl(PVCnt), "P") = rst2!MBMIcBcPp
                If rst2!MBMPatientCovID = "00" Then
                    objWorksheet.Cells(4 + CDbl(PVCnt), "R") = rst2!MBMIcBcPp
                    tmpAge = Year(rst2!MBMPayorEffDate) - Year(rst2!MBMDOB)
                Else
                    objWorksheet.Cells(4 + CDbl(PVCnt), "R") = rst2!MBMIcBcPp
                    tmpAge = Year(rst2!MBMPayorEffDate) - Year(rst2!MBMCDOB)
                End If
                objWorksheet.Cells(4 + CDbl(PVCnt), "Q") = rst2!PatientName
                objWorksheet.Cells(4 + CDbl(PVCnt), "W") = tmpAge
                objWorksheet.Cells(4 + CDbl(PVCnt), "X") = rst2!MBMPayorEffDate
                objWorksheet.Cells(4 + CDbl(PVCnt), "Y") = rst2!MBMPayorEXPDate
                objWorksheet.Cells(4 + CDbl(PVCnt), "Z") = rst2!CASDateAdmitted
                objWorksheet.Cells(4 + CDbl(PVCnt), "AA") = rst2!CASDischargeDate
                objWorksheet.Cells(4 + CDbl(PVCnt), "AB") = "HOSPITAL BENEFIT"
                objWorksheet.Cells(4 + CDbl(PVCnt), "AC") = rst2!ClaimNatureCode
                objWorksheet.Cells(4 + CDbl(PVCnt), "AD") = rst2!CASAdmissionReason
                objWorksheet.Cells(4 + CDbl(PVCnt), "AO") = Format(rst2!CLMTotalActual, "#,##0.00")
                'objWorksheet.Cells(4 + CDbl(PVCnt), "AT") = Format(rst2!MBMBTBG, "#,##0.00")
                objWorksheet.Cells(4 + CDbl(PVCnt), "AT") = Format(rst2!CLMTotalApproved, "#,##0.00")
                
                objWorksheet.Cells(4 + CDbl(PVCnt), "BS") = rst2!BordxRefNo
                
                objWorksheet.Cells(4 + CDbl(PVCnt), "BT") = rst2!CASNumber & "-" & rst2!ClaimVersion
                objWorksheet.Cells(4 + CDbl(PVCnt), "BU") = rst2!PLNCode
                objWorksheet.Cells(4 + CDbl(PVCnt), "BZ") = rst2!HosName
                
                
            rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
    'DocSavePath = LocSVRIntDoc & Trim(txtCASNumber) & "\"
    'objWorkbook.Save
    'objWorkbook.SaveAs (DocSavePath & Trim(txtCASNumber) & "-" & CboVersion & ".xls")
    
    
        If Len(Dir("C:\EtiqaData\", vbDirectory)) = 0 Then
           MkDir "C:\EtiqaData\"
        End If
        
        
        DocPath = "C:\EtiqaData\"
        tmpDocname = Mid(cboCashType, 1, 1) & Format(Date, "ddmmyyyy")
        DocSavePath = DocPath & "InternalDocuments\"
        'objWorkbook.SaveAs (DocSavePath & tmpDocname & ".xls")
        objWorkbook.SaveAs (DocPath & tmpDocname & ".xls")
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
Dim Prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHOSName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double

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
        strsql = ""
        strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "' And MBMGroupCategory = 'MASWING SDN BHD' and INVPayTo = 'H'"
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
        RecFound = False
        tmpSeqNo = 0
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
                objWorksheet.Cells(16 + CDbl(PVCnt), "H") = Format(rst2!PayabletoHOS, "#,##0.00")
                objWorksheet.Cells(16 + CDbl(PVCnt), "J").Borders.LineStyle = True
                objWorksheet.Cells(16 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
            rst2.MoveNext
        Loop
            objWorksheet.Cells(15, "E") = Format(tmpTotalCases, "#,##0.00")
            objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
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
            objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(0, "#,##0.00")
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(18 + CDbl(PVCnt), "H").Font.Bold = True
            objWorksheet.Cells(18 + CDbl(PVCnt), "H") = Format(tmpTotalHos, "#,##0.00")

        rst2.Close
        Set rst2 = Nothing
        
        
            If RecFound = True Then
                strsql = ""
                strsql = " And BordxRefNo = '" & Trim(lblBordxNo) & "' And MBMGroupCategory = 'MASWING SDN BHD' and INVPayTo = 'H'"
                Set cmd2.ActiveConnection = medixmasterconnWS
                cmd2.CommandText = "CLM_PayAdviceHOS"
                cmd2.CommandType = adCmdStoredProc
                cmd2.CommandTimeout = 300
                Set Prm3 = cmd2.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
                cmd2.Parameters.Append Prm3
                Set Prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 1)
                cmd2.Parameters.Append Prm4
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
                    If PVCnt = 1 Then
                        If TMPWS <> 1 Then
                            TMPWS = TMPWS - 1
                            Set objWorksheet = objWorkbook.Sheets(2)
                            objExcel.DisplayAlerts = False
                            Set objWorksheet = objWorkbook.Worksheets.Item(2)
                            objWorksheet.Cells(2, "I") = lblBordxNo
                            objWorksheet.Cells(3, "I") = lblDate

                            objWorksheet.Cells(7, "A") = rst2!HosName
                            objWorksheet.Cells(8, "A") = rst2!HOSAdd1
                            objWorksheet.Cells(9, "A") = rst2!HOSAdd2
                            objWorksheet.Cells(10, "A") = rst2!HOSAdd3
                            objWorksheet.Cells(11, "A") = "Tel : " & rst2!HOSTelNo & "   " & "Fax : " & rst2!HOSFaxNo
                            
                            tmpsubAct = 0
                            tmpSubApp = 0
                        End If
                    End If
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
                                
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                                
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                                PVCnt = PVCnt + 1
                                objWorksheet.Cells(16 + CDbl(PVCnt), "B") = "Malaysia Airlines Berhad"
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
                                objWorksheet.Cells(2, "I") = lblBordxNo
                                objWorksheet.Cells(3, "I") = lblDate
                                
                                objWorksheet.Cells(7, "A") = rst2!HosName
                                objWorksheet.Cells(8, "A") = rst2!HOSAdd1
                                objWorksheet.Cells(9, "A") = rst2!HOSAdd2
                                objWorksheet.Cells(10, "A") = rst2!HOSAdd3
                                objWorksheet.Cells(11, "A") = "Tel : " & rst2!HOSTelNo & "   " & "Fax : " & rst2!HOSFaxNo
                                tmpsubAct = 0
                                tmpSubApp = 0
                                PVCnt = 0
                                PVCnt = PVCnt + 1
                        End If
                        
                        objWorksheet.Cells(14 + CDbl(PVCnt), "A") = PVCnt
                        objWorksheet.Cells(14 + CDbl(PVCnt), "B") = rst2!CASNumber & "-" & rst2!ClaimVersion
                        objWorksheet.Cells(14 + CDbl(PVCnt), "C") = Format(rst2!CASDischargeDate, "YYYY/MM/DD")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "D") = rst2!MBMEmployeeNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "E") = rst2!MBMName
                        objWorksheet.Cells(14 + CDbl(PVCnt), "H") = rst2!INVNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(rst2!CLMTotalActual, "#,##0.00")
                        tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                        objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                        objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
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
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                    
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                    PVCnt = PVCnt + 1
                    objWorksheet.Cells(16 + CDbl(PVCnt), "B") = "Malaysia Airlines Berhad"
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
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Claim\" & "\CoverPanelPayment.DOT\")
                End If
                
                    'bookmark = False
                    objdoc.Bookmarks("TotalInv").Range.Text = tmpTotalCases
                    objdoc.Bookmarks("BordxDate").Range.Text = lblDate
                    objdoc.Bookmarks("BordxNo").Range.Text = lblBordxNo
                    objdoc.Bookmarks("Remarks").Range.Text = "- MASWING SDN BHD"
                    objdoc.Bookmarks("Type").Range.Text = "D"
                
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
Dim Prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHOSName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double
Dim tmpTax As Double
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
        tmpTax = 0
        PVCnt = 0
        tmpTotalCases = 0
        tmpTotalActual = 0
        tmpTotalHos = 0
        tmpHOSPayee = ""
        tmpSeqNo = 0
        tmpTotalMBM = 0
        tmpTotalTax = 0
        tmpTotalDueMBM = 0
        tmpTax = 0
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
                objWorksheet.Cells(5 + CDbl(PVCnt), "B") = rst2!MBMName
                objWorksheet.Cells(5 + CDbl(PVCnt), "H") = rst2!MBMEmployeeNo
                objWorksheet.Cells(5 + CDbl(PVCnt), "I") = rst2!CASNumber & "-" & rst2!ClaimVersion
                If Format(rst2!CLMTax, "#,##0.00") <> Format(rst2!CLMPayableByMedix2M, "#,##0.00") Then
                    objWorksheet.Cells(5 + CDbl(PVCnt), "J") = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                Else
                    objWorksheet.Cells(5 + CDbl(PVCnt), "J") = Format(0, "#,##0.00")
                End If
                objWorksheet.Cells(5 + CDbl(PVCnt), "K") = IIf(IsNumeric(rst2!CLMTax), Format(rst2!CLMTax, "#,##0.00"), Format(0, "#,##0.00"))
                tmpTax = IIf(IsNumeric(rst2!CLMTax), Format(rst2!CLMTax, "#,##0.00"), Format(0, "#,##0.00"))
                objWorksheet.Cells(5 + CDbl(PVCnt), "L") = Format(0, "#,##0.00")
                If Format(rst2!CLMTax, "#,##0.00") <> Format(rst2!CLMPayableByMedix2M, "#,##0.00") Then
                    objWorksheet.Cells(5 + CDbl(PVCnt), "M") = Format(CDbl(rst2!CLMPayableByMedix2M) + CDbl(tmpTax), "#,##0.00")
                Else
                    objWorksheet.Cells(5 + CDbl(PVCnt), "M") = Format(CDbl(tmpTax), "#,##0.00")
                End If
                objWorksheet.Cells(5 + CDbl(PVCnt), "O").Borders.LineStyle = True
                objWorksheet.Cells(5 + CDbl(PVCnt), "Q").Borders(xlEdgeBottom).LineStyle = xlContinuous
                tmpTotalCases = tmpTotalCases + 1
                
                If Format(rst2!CLMTax, "#,##0.00") <> Format(rst2!CLMPayableByMedix2M, "#,##0.00") Then
                    tmpTotalMBM = CDbl(tmpTotalMBM) + CDbl(rst2!CLMPayableByMedix2M)
                    tmpTotalMBM = Format(tmpTotalMBM, "#,##0.00")
                End If
                
                tmpTotalTax = CDbl(tmpTotalTax) + CDbl(tmpTax)
                tmpTotalTax = Format(tmpTotalTax, "#,##0.00")
                
                If Format(rst2!CLMTax, "#,##0.00") <> Format(rst2!CLMPayableByMedix2M, "#,##0.00") Then
                    tmpTotalDueMBM = tmpTotalDueMBM + (CDbl(rst2!CLMPayableByMedix2M) + CDbl(tmpTax))
                    tmpTotalDueMBM = Format(tmpTotalDueMBM, "#,##0.00")
                Else
                    tmpTotalDueMBM = tmpTotalDueMBM + CDbl(tmpTax)
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
            
            objWorksheet.Cells(6 + CDbl(PVCnt), "J").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(6 + CDbl(PVCnt), "J").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(6 + CDbl(PVCnt), "J").Font.Bold = True
            objWorksheet.Cells(6 + CDbl(PVCnt), "J") = Format(tmpTotalMBM, "#,##0.00")
            objWorksheet.Cells(6 + CDbl(PVCnt), "K").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(6 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(6 + CDbl(PVCnt), "K").Font.Bold = True
            objWorksheet.Cells(6 + CDbl(PVCnt), "K") = Format(tmpTotalTax, "#,##0.00")
            objWorksheet.Cells(6 + CDbl(PVCnt), "L").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(6 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(6 + CDbl(PVCnt), "L").Font.Bold = True
            objWorksheet.Cells(6 + CDbl(PVCnt), "L") = Format(0, "#,##0.00")
            objWorksheet.Cells(6 + CDbl(PVCnt), "M").Borders(xlEdgeTop).LineStyle = xlContinuous
            objWorksheet.Cells(6 + CDbl(PVCnt), "M").Borders(xlEdgeBottom).LineStyle = xlDouble
            objWorksheet.Cells(6 + CDbl(PVCnt), "M").Font.Bold = True
            objWorksheet.Cells(6 + CDbl(PVCnt), "M") = Format(tmpTotalDueMBM, "#,##0.00")

        rst2.Close
        Set rst2 = Nothing
        
        
            If RecFound = True Then
                Set objword = Nothing
                If objword Is Nothing Then
                    Set objword = New Word.Application
                    objword.Visible = True
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Claim\" & "\CoverStaffPayment.DOT\")
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
Dim Prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHOSName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double
Dim tmpTotalRecovery As Double
Dim COV As New ADODB.Recordset
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
                    COV.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    'strsqlHOS = "Select HOSPayee from HOS where HOSName = '" & objWorksheet.Cells(7, "A") & "'"
                    'HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic

                    Do While Not COV.EOF
                        objWorksheet.Cells(5 + CDbl(PVCnt), "F") = COV!MBMCName
                        COV.MoveNext
                    Loop
                    COV.Close
                    Set COV = Nothing
                End If
                objWorksheet.Cells(5 + CDbl(PVCnt), "J") = rst2!CASNumber & "-" & rst2!ClaimVersion
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
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Claim\" & "\CoverRecovery.DOT\")
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
Dim Prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHOSName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double

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
                objWorksheet.Cells(16 + CDbl(PVCnt), "H") = Format(rst2!PayabletoHOS, "#,##0.00")
                objWorksheet.Cells(16 + CDbl(PVCnt), "J").Borders.LineStyle = True
                objWorksheet.Cells(16 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
            rst2.MoveNext
        Loop
            objWorksheet.Cells(15, "E") = Format(tmpTotalCases, "#,##0.00")
            objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
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
            objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(0, "#,##0.00")
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
                Set Prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 1)
                cmd2.Parameters.Append Prm4
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
                            objWorksheet.Cells(2, "I") = lblBordxNo & "-" & Format(tmpSeqNo, "000")
                            objWorksheet.Cells(3, "I") = lblDate

                            objWorksheet.Cells(7, "A") = rst2!HosName
                            objWorksheet.Cells(8, "A") = rst2!HOSAdd1
                            objWorksheet.Cells(9, "A") = rst2!HOSAdd2
                            objWorksheet.Cells(10, "A") = rst2!HOSAdd3
                            objWorksheet.Cells(11, "A") = "Tel : " & rst2!HOSTelNo & "   " & "Fax : " & rst2!HOSFaxNo
                            
                            tmpsubAct = 0
                            tmpSubApp = 0
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
                                
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                                
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                                PVCnt = PVCnt + 1
                                objWorksheet.Cells(16 + CDbl(PVCnt), "B") = "Malaysia Airlines Berhad"
                                
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
                                objWorksheet.Cells(2, "I") = lblBordxNo & "-" & Format(tmpSeqNo, "000")
                                objWorksheet.Cells(3, "I") = lblDate
                                
                                objWorksheet.Cells(7, "A") = rst2!HosName
                                objWorksheet.Cells(8, "A") = rst2!HOSAdd1
                                objWorksheet.Cells(9, "A") = rst2!HOSAdd2
                                objWorksheet.Cells(10, "A") = rst2!HOSAdd3
                                objWorksheet.Cells(11, "A") = "Tel : " & rst2!HOSTelNo & "   " & "Fax : " & rst2!HOSFaxNo
                                tmpsubAct = 0
                                tmpSubApp = 0
                                PVCnt = 0
                                PVCnt = PVCnt + 1
                        End If
                        
                        objWorksheet.Cells(14 + CDbl(PVCnt), "A") = PVCnt
                        objWorksheet.Cells(14 + CDbl(PVCnt), "B") = rst2!CASNumber & "-" & rst2!ClaimVersion
                        objWorksheet.Cells(14 + CDbl(PVCnt), "C") = Format(rst2!CASDischargeDate, "YYYY/MM/DD")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "D") = rst2!MBMEmployeeNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "E") = rst2!MBMName
                        objWorksheet.Cells(14 + CDbl(PVCnt), "H") = rst2!INVNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(rst2!CLMTotalActual, "#,##0.00")
                        tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                        objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                        objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
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
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                    
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                    PVCnt = PVCnt + 1
                    objWorksheet.Cells(16 + CDbl(PVCnt), "B") = "Malaysia Airlines Berhad"
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
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Claim\" & "\CoverPanelPayment.DOT\")
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
Dim Prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHOSName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpSeqNo As Double

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
                objWorksheet.Cells(16 + CDbl(PVCnt), "H") = Format(rst2!PayabletoHOS, "#,##0.00")
                objWorksheet.Cells(16 + CDbl(PVCnt), "J").Borders.LineStyle = True
                objWorksheet.Cells(16 + CDbl(PVCnt), "L").Borders(xlEdgeBottom).LineStyle = xlContinuous
                tmpTotalHos = CDbl(tmpTotalHos) + CDbl(rst2!PayabletoHOS)
                tmpTotalHos = Format(tmpTotalHos, "#,##0.00")
            rst2.MoveNext
        Loop
            objWorksheet.Cells(15, "E") = Format(tmpTotalCases, "#,##0.00")
            objWorksheet.Cells(15, "F") = Format(tmpTotalActual, "#,##0.00")
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
            objWorksheet.Cells(18 + CDbl(PVCnt), "G") = Format(0, "#,##0.00")
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
                Set Prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 1)
                cmd2.Parameters.Append Prm4
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
                            objWorksheet.Cells(2, "I") = lblBordxNo & "-" & Format(tmpSeqNo, "000")
                            objWorksheet.Cells(3, "I") = lblDate

                            objWorksheet.Cells(7, "A") = rst2!HosName
                            objWorksheet.Cells(8, "A") = rst2!HOSAdd1
                            objWorksheet.Cells(9, "A") = rst2!HOSAdd2
                            objWorksheet.Cells(10, "A") = rst2!HOSAdd3
                            objWorksheet.Cells(11, "A") = "Tel : " & rst2!HOSTelNo & "   " & "Fax : " & rst2!HOSFaxNo
                            
                            tmpsubAct = 0
                            tmpSubApp = 0
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
                                
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeTop).LineStyle = xlContinuous
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlDouble
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K").Font.Bold = True
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                                
                                objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                                objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                                objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                                PVCnt = PVCnt + 1
                                objWorksheet.Cells(16 + CDbl(PVCnt), "B") = "Malaysia Airlines Berhad"
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
                                objWorksheet.Cells(2, "I") = lblBordxNo & "-" & Format(tmpSeqNo, "000")
                                objWorksheet.Cells(3, "I") = lblDate
                                
                                objWorksheet.Cells(7, "A") = rst2!HosName
                                objWorksheet.Cells(8, "A") = rst2!HOSAdd1
                                objWorksheet.Cells(9, "A") = rst2!HOSAdd2
                                objWorksheet.Cells(10, "A") = rst2!HOSAdd3
                                objWorksheet.Cells(11, "A") = "Tel : " & rst2!HOSTelNo & "   " & "Fax : " & rst2!HOSFaxNo
                                tmpsubAct = 0
                                tmpSubApp = 0
                                PVCnt = 0
                                PVCnt = PVCnt + 1
                        End If
                        
                        objWorksheet.Cells(14 + CDbl(PVCnt), "A") = PVCnt
                        objWorksheet.Cells(14 + CDbl(PVCnt), "B") = rst2!CASNumber & "-" & rst2!ClaimVersion
                        objWorksheet.Cells(14 + CDbl(PVCnt), "C") = Format(rst2!CASDischargeDate, "YYYY/MM/DD")
                        objWorksheet.Cells(14 + CDbl(PVCnt), "D") = rst2!MBMEmployeeNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "E") = rst2!MBMName
                        objWorksheet.Cells(14 + CDbl(PVCnt), "H") = rst2!INVNo
                        objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(rst2!CLMTotalActual, "#,##0.00")
                        tmpsubAct = CDbl(tmpsubAct) + CDbl(rst2!CLMTotalActual)
                        objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                        objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
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
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeTop).LineStyle = xlContinuous
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlDouble
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K").Font.Bold = True
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                    
                    objWorksheet.Cells(14 + CDbl(PVCnt), "I") = Format(tmpsubAct, "#,##0.00")
                    objWorksheet.Cells(14 + CDbl(PVCnt), "J") = "0"
                    objWorksheet.Cells(14 + CDbl(PVCnt), "K") = Format(tmpSubApp, "#,##0.00")
                    PVCnt = PVCnt + 1
                    objWorksheet.Cells(16 + CDbl(PVCnt), "B") = "Malaysia Airlines Berhad"
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
                    Set objdoc = objword.Documents.Add(LocCardPath & "\Claim\" & "\CoverPanelPayment.DOT\")
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
Dim Prm4 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpTotalActual As Double
Dim tmpTotalMBMDue As Double
Dim tmpTotalHos As Double
Dim tmpHOSName As String
Dim TMPWS As Integer
Dim tmpsubAct As Double
Dim tmpSubApp As Double
Dim tmpHOSPayee As String
Dim strsqlHOS As String
Dim HOS As New ADODB.Recordset
Dim tmpGrandTotal As Double
Dim cmd3 As New ADODB.Command
Dim Prm5 As New ADODB.Parameter
Dim Prm6 As New ADODB.Parameter
Dim cmd4 As New ADODB.Command
Dim Prm7 As New ADODB.Parameter
Dim Prm8 As New ADODB.Parameter

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
                    
                    objWorksheet.Cells(2 + CDbl(PVCnt), "A") = "51702"
                    objWorksheet.Cells(2 + CDbl(PVCnt), "B") = "00"
                    objWorksheet.Cells(2 + CDbl(PVCnt), "C") = Mid(rst2!MBMCostCenter, 1, 2)
                    objWorksheet.Cells(2 + CDbl(PVCnt), "D") = Mid(rst2!MBMCostCenter, 3, 3)
                    objWorksheet.Cells(2 + CDbl(PVCnt), "E") = Mid(rst2!MBMCostCenter, 6, 3)
                    objWorksheet.Cells(2 + CDbl(PVCnt), "H") = Format(rst2!TotalAmt, "#,##0.00")
                    objWorksheet.Cells(2 + CDbl(PVCnt), "L") = "JXXXXXX INS " & Month(Date) & Year(Date)
                    objWorksheet.Cells(2 + CDbl(PVCnt), "M") = "51702" & "00" & rst2!MBMCostCenter
                    objWorksheet.Cells(2 + CDbl(PVCnt), "O") = "JXXXXXX"
                    objWorksheet.Cells(2 + CDbl(PVCnt), "P") = "INS " & Month(Date) & Year(Date)
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
            Set Prm4 = cmd2.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 4)
            cmd2.Parameters.Append Prm4
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
            Set Prm6 = cmd3.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 7)
            cmd3.Parameters.Append Prm6
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
            Set Prm8 = cmd4.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 8)
            cmd4.Parameters.Append Prm8
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


