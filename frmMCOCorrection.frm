VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMCOCorrection 
   Caption         =   "MCO Fee Correction"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   60
      TabIndex        =   14
      Top             =   240
      Width           =   11775
      Begin VB.ComboBox CboPayor 
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
         Left            =   1200
         TabIndex        =   0
         Top             =   200
         Width           =   4335
      End
      Begin VB.TextBox TxtMBMNoTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   25
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox TxtMBMNoFr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtBordxDate1 
         Height          =   315
         Left            =   6960
         TabIndex        =   3
         Top             =   200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDate2 
         Height          =   315
         Left            =   8760
         TabIndex        =   4
         Top             =   200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBatchNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         MaxLength       =   3
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtBatchNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8760
         MaxLength       =   3
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label8 
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
         Left            =   8400
         TabIndex        =   21
         Top             =   520
         Width           =   495
      End
      Begin VB.Label Label7 
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
         Left            =   3240
         TabIndex        =   20
         Top             =   520
         Width           =   495
      End
      Begin VB.Label Label6 
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
         Left            =   8400
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Batch No."
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
         Left            =   5760
         TabIndex        =   18
         Top             =   520
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Member No."
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
         Left            =   120
         TabIndex        =   17
         Top             =   520
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Bordx. Date"
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
         Left            =   5760
         TabIndex        =   16
         Top             =   240
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lstMBM 
      Height          =   5295
      Left            =   60
      TabIndex        =   13
      Top             =   1200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9340
      View            =   3
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Payor Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Bordx.Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Batch"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Invoice No."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Inv. Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "MBMNo"
         Text            =   "Member No."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Cover ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Member Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "MCO"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Effective Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Expiry Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   60
      TabIndex        =   22
      Top             =   6480
      Width           =   11775
      Begin VB.TextBox txtMCOFee 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   20
         TabIndex        =   24
         Top             =   200
         Width           =   1575
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
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
         Left            =   8400
         TabIndex        =   9
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
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
         Left            =   9240
         TabIndex        =   10
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "St&op"
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
         Left            =   7560
         TabIndex        =   8
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
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
         Left            =   10080
         TabIndex        =   11
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
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
         Height          =   315
         Left            =   10920
         TabIndex        =   12
         Top             =   200
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
         Height          =   315
         Left            =   6720
         TabIndex        =   7
         Top             =   200
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "MCO Fee"
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
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   " MCO Fee Correction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   220
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmMCOCorrection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemChecked As Boolean
Dim intExit As Integer

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
   
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute

    Do While Not PAY.EOF
        With PAY
            CboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboPayor, KeyAscii)
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
Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstMBM.listitems.Clear
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    If lstMBM.listitems.Count > 0 Then
        Call ExportListViewtoExcel(lstMBM)
    Else
        MsgBox "Please select record!", vbCritical + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Sub CmdSearch_Click()
    CmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    CmdPrint.Enabled = False
    CmdClear.Enabled = False
    CmdUpdate.Enabled = False
    CmdExit.Enabled = False
    
    Call MBMInquiryList
    
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    CmdUpdate.Enabled = True
    CmdExit.Enabled = True

    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    Screen.MousePointer = Default
    CmdSearch.Enabled = True

End Sub

Private Sub MBMInquiryList()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sqlstr As String
Dim strCriteria As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim strsqlCOV As String
Dim RstCov As New ADODB.Recordset
Dim Check As Boolean

    strCriteria = ""
    If CboPayor <> "" Then
        strCriteria = strCriteria & " AND PAYCode =  '" & Mid(CboPayor, 1, 2) & "'"
    End If
    If txtBordxDate1 <> "__/__/____" Then
        strCriteria = strCriteria & " AND MBMBordxDate >= '" & Format(txtBordxDate1, "MM/DD/YYYY") & "'"
    End If
    If txtBordxDate2 <> "__/__/____" Then
        strCriteria = strCriteria & " AND MBMBordxDate <= '" & Format(txtBordxDate2, "MM/DD/YYYY") & "'"
    End If
    If TxtMBMNoFr <> "" Then
        strCriteria = strCriteria & " AND MBMNumber >=  '" & Trim(TxtMBMNoFr) & "'"
    End If
    If TxtMBMNoTo <> "" Then
        strCriteria = strCriteria & " AND MBMNumber <=  '" & (TxtMBMNoTo) & "'"
    End If
    If txtBatchNo1 <> "" Then
        strCriteria = strCriteria & " AND MBMBatchNo >=  '" & (txtBatchNo1) & "'"
    End If
    
    If txtBatchNo2 <> "" Then
        strCriteria = strCriteria & " AND MBMBatchNo <=  '" & (txtBatchNo2) & "'"
    End If
    
    lstMBM.listitems.Clear
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Search_MBM"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
    cmd.Parameters.Append Prm1
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    Do
        Do Until rst2.EOF
                    
            With rst2
                Set Record = lstMBM.listitems.Add(, , "")
                Record.SubItems(1) = !PAYCode
                Record.SubItems(2) = IIf(Len(!MBMBordxDate), Trim(!MBMBordxDate), "")
                Record.SubItems(3) = IIf(Len(!MBMBatchNo), Trim(!MBMBatchNo), "")
                Record.SubItems(4) = IIf(Len(!MBMInvoiceNo), Trim(!MBMInvoiceNo), "")
                Record.SubItems(5) = IIf(Len(!MBMInvoiceDate), Trim(!MBMInvoiceDate), "")
                Record.SubItems(6) = IIf(Len(!MBMNumber), Trim(!MBMNumber), "")
                Record.SubItems(7) = IIf(Len(!MBMCOVID), Trim(!MBMCOVID), "")
                Record.SubItems(8) = IIf(Len(!MBMName), Trim(!MBMName), "")
                Record.SubItems(9) = IIf(Len(!MBMMCOFee), !MBMMCOFee, "")
                If Len(!MBMPayorEffDate) Then
                    Record.SubItems(10) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                End If
                If Len(!MBMPayorEXPDate) Then
                   Record.SubItems(11) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                End If
                                                
                If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                    strsqlCOV = ""
                    strsqlCOV = " Select MBMCNumber,MBMCCoverID,MBMInvoiceNo,MBMInvoiceDate,MBMCMCO,MBMCName,MBMCBatchNo,MBMCBordxDate From MBMCoveredPersons Where MBMCNumber = '" & Trim(!MBMNumber) & "'"
                    RstCov.Open strsqlCOV, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                    Do While Not RstCov.EOF
                        Set Record = lstMBM.listitems.Add(, , "")
                        Record.SubItems(1) = IIf(Len(!PAYCode), Trim(!PAYCode), "")
                        Record.SubItems(2) = IIf(Len(RstCov!MBMCBordxDate), Trim(RstCov!MBMCBordxDate), "")
                        Record.SubItems(3) = IIf(Len(RstCov!MBMCBatchNo), Trim(RstCov!MBMCBatchNo), "")
                        Record.SubItems(4) = IIf(Len(RstCov!MBMInvoiceNo), Trim(RstCov!MBMInvoiceNo), "")
                        Record.SubItems(5) = IIf(Len(RstCov!MBMInvoiceDate), Trim(RstCov!MBMInvoiceDate), "")
                        Record.SubItems(6) = IIf(Len(RstCov!MBMCNumber), Trim(RstCov!MBMCNumber), "")
                        Record.SubItems(7) = IIf(Len(RstCov!MBMCCoverID), Trim(RstCov!MBMCCoverID), "")
                        Record.SubItems(8) = IIf(Len(RstCov!MBMCName), Trim(RstCov!MBMCName), "")
                        Record.SubItems(9) = IIf(Len(RstCov!MBMCMCO), RstCov!MBMCMCO, "")
                        If Len(!MBMPayorEffDate) Then
                            Record.SubItems(10) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                        End If
                        If Len(!MBMPayorEXPDate) Then
                           Record.SubItems(11) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                        End If
                        RstCov.MoveNext
                    Loop
                    RstCov.Close
                    Set RstCov = Nothing
                End If
                                                            
                                                            
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
    If lstMBM.listitems.Count = 0 Then
        MsgBox "Record Not Found!", vbCritical + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Sub CmdStop_Click()
    intExit = 1
End Sub

Private Sub CmdUpdate_Click()
On Error GoTo errSave
Dim RecordSaved
Dim RstMBM As New ADODB.Recordset
Dim strsql As String
Dim i As Integer
Dim startTrans As Boolean
        
        startTrans = False
        ItemChecked = False
           
        If Trim(txtMCOFee) = "" Then
            MsgBox "Please Enter MCO Fee!", vbCritical + vbOKOnly, DialogBoxTitle
            txtMCOFee.SetFocus
        Else
        
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                CmdPrint.Enabled = False
                CmdClear.Enabled = False
                CmdUpdate.Enabled = False
                CmdExit.Enabled = False
                CmdSearch.Enabled = False
                
               ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
               
                Screen.MousePointer = vbHourglass
                startTrans = True
                medixmasterconn.beginTrans
                
                For i = 1 To lstMBM.listitems.Count
                    If lstMBM.listitems(i).Checked = True Then
                        If lstMBM.listitems(i).SubItems(7) = "00" Then
                            strsql = "Update MBMOthers set MBMMCOFee = " & Trim(txtMCOFee) & " Where MBMNumber = '" & lstMBM.listitems(i).SubItems(6) & "'"
                        Else
                            strsql = "Update MBMCoveredPersons set MBMCMCO = " & Trim(txtMCOFee) & " Where MBMCNumber = '" & lstMBM.listitems(i).SubItems(6) & "' And MBMCCoverID = '" & lstMBM.listitems(i).SubItems(7) & "'"
                        End If
                        medixmasterconn.Execute strsql
                    End If
                Next i

                medixmasterconn.commitTrans
                MsgBox "Record updated !", vbOKOnly
            
                lstMBM.listitems.Clear
                Call MBMInquiryList
               ' medixmasterconn.Close
               ' Set medixmasterconn = Nothing
                CmdPrint.Enabled = True
                CmdClear.Enabled = True
                CmdUpdate.Enabled = True
                CmdExit.Enabled = True
                CmdSearch.Enabled = True
            Screen.MousePointer = Default
            
        End If
    End If
    Exit Sub
errSave:
    MsgBox Err.Description
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    If startTrans = True Then
        medixmasterconn.RollbackTrans
    End If
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    CmdUpdate.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
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
Dim rsSearch As New ADODB.Recordset
Dim sql As String
    intExit = 0
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub lstMBM_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If lstMBM.listitems.Count > 0 Then
        ItemChecked = True
        Call ProcessCheckedItem(Item)
    End If
End Sub

Public Function ProcessCheckedItem(Item As ListItem)
    If Item.SubItems(4) = "" Then
        If Item.Checked = True Then
            Item.Text = "Checked"
        Else
            Item.Text = ""
        End If
    Else
        MsgBox "Invoice has been generated! ", vbCritical
        Item.Checked = False
        Item.Text = ""
    End If
End Function

Private Sub txtBatchNo1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub txtBatchNo2_KeyPress(KeyAscii As Integer)
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

Private Sub txtMCOFee_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
