VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTrustAC 
   Caption         =   "Reconciliation"
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
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   16
      Top             =   240
      Width           =   11655
      Begin VB.CommandButton cmdRepeatClaim 
         Caption         =   "Repeat"
         Height          =   285
         Left            =   4920
         TabIndex        =   29
         Top             =   510
         Width           =   735
      End
      Begin VB.CommandButton cmdRepeatBordx 
         Caption         =   "Repeat"
         Height          =   285
         Left            =   4920
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cboHLTType 
         Height          =   315
         Left            =   6600
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   6600
         TabIndex        =   5
         Top             =   520
         Width           =   3495
      End
      Begin VB.TextBox txtBordxNoTo 
         Height          =   285
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   1
         Top             =   270
         Width           =   1335
      End
      Begin VB.TextBox txtBordxNoFr 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   0
         Top             =   270
         Width           =   1335
      End
      Begin VB.TextBox txtCASNo1 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   6
         Top             =   510
         Width           =   1335
      End
      Begin VB.TextBox txtCASNo2 
         Height          =   285
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   7
         Top             =   510
         Width           =   1335
      End
      Begin VB.ComboBox cboReconBORStatus 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   770
         Width           =   3135
      End
      Begin VB.ComboBox cboReconCASStatus 
         Height          =   315
         Left            =   6600
         TabIndex        =   3
         Top             =   810
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   555
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   345
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Recon Status Bordx"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   825
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Status"
         Height          =   255
         Left            =   5760
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "HLT Type"
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   555
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Payor"
         Height          =   255
         Left            =   5760
         TabIndex        =   17
         Top             =   525
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   120
      TabIndex        =   9
      Top             =   7320
      Width           =   11655
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1440
         TabIndex        =   31
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox lblCurRecord 
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmpUpdate 
         Caption         =   "&Update Recon"
         Height          =   330
         Left            =   7920
         TabIndex        =   27
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   7200
         TabIndex        =   10
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   330
         Left            =   9120
         TabIndex        =   11
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   330
         Left            =   6480
         TabIndex        =   8
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10800
         TabIndex        =   13
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10080
         TabIndex        =   12
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstTrustAC 
      Height          =   5775
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10186
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
      NumItems        =   27
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Case Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Claim No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Rec MNRB"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Rec MBM"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Rec Medix"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Rec Oth"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "RecPayor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Total Rec"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Paid2Hos"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Paid2MBM"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Paid2Medix"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Paid2Oth"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Paid2Payor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Bank Charges"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Contra"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Total Pay"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Cash Status/(Deficit)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "O/S Hosp To/(Fr)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "O/S MBM To/(Fr)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Gain/(Loss)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "O/S Check Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "App Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "Pay2Hos"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Pay/(Rec)MBM"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Check Total"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Reconciliation"
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
      TabIndex        =   21
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmTrustAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Private m_blnDirection(1 To 28) As Boolean

Private Sub cboHLTType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHLTType.Text, cboHLTType.SelStart) + Chr(KeyAscii) + Mid(cboHLTType.Text, cboHLTType.SelStart + cboHLTType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboHLTType.ListCount - 1
 
        If NewText = LCase(Left(cboHLTType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHLTType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboHLTType.Text = ValidValue
                cboHLTType.SelStart = 0
                cboHLTType.SelLength = Len(ValidValue)
            End If
        End If

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

Private Sub cboReconBORStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboReconBORStatus.Text, cboReconBORStatus.SelStart) + Chr(KeyAscii) + Mid(cboReconBORStatus.Text, cboReconBORStatus.SelStart + cboReconBORStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboReconBORStatus.ListCount - 1
 
        If NewText = LCase(Left(cboReconBORStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboReconBORStatus.Text = ValidValue
                cboReconBORStatus.SelStart = 0
                cboReconBORStatus.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboReconCASStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboReconCASStatus.Text, cboReconCASStatus.SelStart) + Chr(KeyAscii) + Mid(cboReconCASStatus.Text, cboReconCASStatus.SelStart + cboReconCASStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboReconCASStatus.ListCount - 1
 
        If NewText = LCase(Left(cboReconCASStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboReconCASStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboReconCASStatus.Text = ValidValue
                cboReconCASStatus.SelStart = 0
                cboReconCASStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cmdRepeatBordx_Click()
    If Len(txtBordxNoFr) Then
        txtBordxNoTo = txtBordxNoFr
    End If
End Sub

Private Sub cmdRepeatClaim_Click()
    If Len(txtCASNo1) Then
        txtCASNo2 = txtCASNo1
    End If
End Sub

Private Sub cmpUpdate_Click()
Dim CASNo As String
Dim CASVer As String
Dim status As String
Dim Update
Dim ItemChecked As Boolean
Dim i As Integer
    
    ItemChecked = False
    CASNo = ""
    CASVer = ""
    status = ""
    CmdSearch.Enabled = False
    cmpUpdate.Enabled = False
    
    For i = 1 To lstTrustAC.ListItems.Count
        If lstTrustAC.ListItems(i).Checked = True Then
            ItemChecked = True
        'Exit For
        End If
    Next i
    
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

    If lstTrustAC.ListItems.Count = 0 Or ItemChecked = False Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    Else
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                i = 0
                ItemChecked = False
                For i = 1 To lstTrustAC.ListItems.Count
                    If lstTrustAC.ListItems(i).Checked = True Then
                        ItemChecked = True
                    Exit For
                    End If
                Next i
                
                If ItemChecked = True Then
                    For i = 1 To lstTrustAC.ListItems.Count
                        If lstTrustAC.ListItems(i).Checked = True Then
                            CASNo = Trim(Mid(lstTrustAC.ListItems(i).ListSubItems(3).Text, 1, IIf(InStr(1, lstTrustAC.ListItems(i).ListSubItems(3).Text, "-") = 0, 10, InStr(1, lstTrustAC.ListItems(i).ListSubItems(3).Text, "-") - 1)))
                            CASVer = Trim(Mid(lstTrustAC.ListItems(i).ListSubItems(3).Text, InStr(1, lstTrustAC.ListItems(i).ListSubItems(3).Text, "-") + 1, Len(lstTrustAC.ListItems(i).ListSubItems(3).Text) - InStr(1, lstTrustAC.ListItems(i).ListSubItems(3).Text, "-")))
                
                            Call UpdateRecord(CASNo, CASVer)
                            
                            CASNo = ""
                            CASVer = ""
                        End If
                    Next i
                    
                    MsgBox "Record Updated", vbOKOnly + vbInformation, DialogBoxTitle

                    lstTrustAC.ListItems.Clear
                End If
            End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    CmdSearch.Enabled = True
    cmpUpdate.Enabled = True

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
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
    medixmasterconn.Close
    Set medixmasterconn = Nothing

    intExit = 0
    lblCurRecord = 0
End Sub

'=============== Click Command Button ============================
Private Sub CmdPrint_Click()
    'Screen.MousePointer = vbHourglass
        'Call ExportToExcelAnalysis1(lstTrustAC)
    'Screen.MousePointer = vbDefault
    
    Screen.MousePointer = vbHourglass
        
    If lstTrustAC.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstTrustAC)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub CmdSearch_Click()
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
    cmpUpdate.Enabled = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        Call SearchRecord
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    CmdSearch.Enabled = True
    cmpUpdate.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lstTrustAC.ListItems.Clear
    lblCurRecord = 0
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
End Sub

Private Sub cmdExit_Click()
        Unload Me
End Sub
'=============== Click Command Button ============================

'************************Start ComboListing**************************

Private Sub ComboListing()
Dim Pay As New ADODB.Recordset
Dim HLT As New ADODB.Recordset
Dim cmd As New ADODB.Command

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Pay.CursorLocation = adUseClient
    Pay.CursorType = adOpenForwardOnly
    Pay.CacheSize = 100
    Set Pay = cmd.Execute
    
    Do Until Pay.EOF
        With Pay
            CboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            Pay.MoveNext
        End With
    Loop
    Pay.Close
    Set Pay = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHLTType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HLT.CursorLocation = adUseClient
    HLT.CursorType = adOpenForwardOnly
    HLT.CacheSize = 100
    Set HLT = cmd.Execute
    
    Do Until HLT.EOF
        With HLT
            cboHLTType.AddItem !HLTCode & " - " & !HLTDescription
            HLT.MoveNext
        End With
    Loop
    HLT.Close
    Set HLT = Nothing
    
    cboReconCASStatus.AddItem "Y" & " - " & "Yes"
    cboReconCASStatus.AddItem "N" & " - " & "No"

    cboReconBORStatus.AddItem "Y" & " - " & "Yes"
    cboReconBORStatus.AddItem "N" & " - " & "No"

End Sub
'************************Start ComboListing**************************

Private Function SearchRecord()
Dim strAppend As String
Dim HOspitals As String
    
    strAppend = ""
    HOspitals = ""
    
    If Len(cboHLTType) Then
        strAppend = strAppend + " And HLTCode = '" & Mid(cboHLTType, 1, 1) & "'"
    End If

    If Len(Trim(txtCASNo1)) Then
        strAppend = strAppend + " And CASNumber >= '" & Trim(txtCASNo1) & "'"
    End If
    
    If Len(Trim(txtCASNo2)) Then
        strAppend = strAppend + " And CASNumber <= '" & Trim(txtCASNo2) & "'"
    End If
    
    If Len(Trim(txtBordxNoFr)) Then
        strAppend = strAppend + " And BordxRefNo >= '" & Trim(txtBordxNoFr) & "'"
    End If
    
    If Len(Trim(txtBordxNoTo)) Then
        strAppend = strAppend + " And BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
    End If
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    'If Len(Trim(txtPVDateFr)) > 0 And IsDate(txtPVDateFr) = True Then
    '    StrAppend = StrAppend + " And (PVHOSDate >= '" & Format(txtPVDateFr, "mm/dd/yyyy") & "' or PVDate >= '" & Format(txtPVDateFr, "mm/dd/yyyy") & "')"
    'End If
    
    'If Len(Trim(txtPVDateTo)) > 0 And IsDate(txtPVDateTo) = True Then
    '    StrAppend = StrAppend + " And (PVHOSDate <= '" & Format(txtPVDateTo, "mm/dd/yyyy") & "' or PVDate <= '" & Format(txtPVDateTo, "mm/dd/yyyy") & "')"
    'End If
    
    'If Len(Trim(txtRECDateFr)) > 0 And IsDate(txtRECDateFr) = True Then
    '    StrAppend = StrAppend + " And (MBMReceiptDate >= '" & Format(txtRECDateFr, "mm/dd/yyyy") & "' or MNRBReceiptDate >= '" & Format(txtRECDateFr, "mm/dd/yyyy") & "')"
    'End If
    
    'If Len(Trim(txtRECDateTo)) > 0 And IsDate(txtRECDateTo) = True Then
    '    StrAppend = StrAppend + " And (MNRBReceiptDate <= '" & Format(txtRECDateTo, "mm/dd/yyyy") & "' or MNRBReceiptDate <= '" & Format(txtRECDateTo, "mm/dd/yyyy") & "')"
    'End If
    
    If Len(cboReconCASStatus) Then
        If Mid(cboReconCASStatus, 1, 1) = "Y" Then
            strAppend = strAppend + " And ReconFlag = '" & Mid(cboReconCASStatus, 1, 1) & "'"
        Else
            strAppend = strAppend + " And ReconFlag is null "
        End If
    End If
    
    Call TrustACListing(strAppend)
   
End Function

Private Sub TrustACListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim Total As String
Dim Current As String
Dim MedixPaidMBM As Double
Dim MedixPaidHOS As Double
Dim Paid2Medix  As Double
Dim Paid2Others As Double
Dim TotalPaid As Double
Dim Pay2MBM As Double
Dim Pay2HOS As Double
Dim OSMBM As Double
Dim OSHos As Double
Dim Pay2Medix As Double
Dim Pay2Others As Double
Dim Pay2PayAmt As Double
Dim BankCharges As Double
Dim RecMBM As Double
Dim RecMNRB As Double
Dim RecOthers As Double
Dim RecPayor As Double
Dim totalREc As Double
Dim RecMedix As Double
Dim TotalApp As Double
Dim CashPos As Double
Dim MedixPay As Double
Dim GainLoss As Double
Dim Contra As Double
Dim CheckTotal As Double
Dim MBMRecovery As Double
    
    Total = 0
    Current = 0
    lstTrustAC.ListItems.Clear
    sql = " Select BordxRefNo,CASNumber, ClaimVersion, MBMAmt, MNRBAmt, CLMTotalApproved," & _
            " HOSPaidAmt, MBMPaidAmt, CLMPayableByMedix2H, CLMPayableByMedix2M, " & _
            "ContraCaseAmt, PaymentBankCharge, ReconFlag, " & _
            "HLTCode,MedixRecAmt,GainLoss,PAYPaidAmt,RecPayorAmt " & _
            "From VIEW_TrustAC where 0 = 0 "

    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
   
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
    Do
        Do Until rst2.EOF
                            
            With rst2
                
                Set Record = lstTrustAC.ListItems.Add(, , "")
                
                If Trim(Len(!ReconFlag)) Then
                    Record.SubItems(1) = "Yes"
                'Else
                '    Record.SubItems(1) = "No"
                End If
                
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(2) = Trim(rst2!BordxRefNo)
                End If

                
                If Len(rst2!CasNumber & "-" & rst2!claimversion) Then
                    Record.SubItems(3) = Trim(rst2!CasNumber & "-" & rst2!claimversion)
                End If
                
     
                If Len(rst2!MNRBAmt) Then
                    Record.SubItems(4) = Format(rst2!MNRBAmt, "#,##0.00;(#,##0.00)")
                    RecMNRB = rst2!MNRBAmt
                Else
                    Record.SubItems(4) = "0.00"
                    RecMNRB = 0
                End If

                If Len(rst2!MBMAmt) Then
                    Record.SubItems(5) = Format(rst2!MBMAmt, "#,##0.00;(#,##0.00)")
                    RecMBM = rst2!MBMAmt
                Else
                    Record.SubItems(5) = "0.00"
                    RecMBM = 0
                End If
                
                If Len(rst2!MedixRecAmt) Then
                    Record.SubItems(6) = Format(rst2!MedixRecAmt, "#,#.00;(#,#.00)")
                    RecMedix = Format(rst2!MedixRecAmt, "#,#.00;(#,#.00)")
                Else
                    Record.SubItems(6) = "0.00"
                    RecMedix = 0
                End If
                
                'If Len(rst2!OthersReceipt) Then
                '    Record.SubItems(16) = Format(rst2!OthersReceipt, "#,#.00;(#,#.00)")
                    'RecMedix = Format(rst2!OthersReceipt, "#,#.00;(#,#.00)")
                'Else
                    Record.SubItems(7) = "0.00"
                    RecOthers = 0
                'End If
                
                If Len(rst2!RecPayorAmt) Then
                     Record.SubItems(8) = Format(rst2!RecPayorAmt, "#,#.00;(#,#.00)")
                     RecPayor = Format(rst2!RecPayorAmt, "#,#.00;(#,#.00)")
                Else
                     Record.SubItems(8) = "0.00"
                    RecPayor = 0
                End If
                
                Record.SubItems(9) = Format(CDbl(RecMNRB) + CDbl(RecMBM) + CDbl(RecMedix) + CDbl(RecOthers) + CDbl(RecPayor), "#,##0.00;(#,##0.00)")
                totalREc = CDbl(RecMNRB) + CDbl(RecMBM) + CDbl(RecMedix) + CDbl(RecOthers) + CDbl(RecPayor)
                
                If Len(rst2!HOSPaidAmt) Then
                    Record.SubItems(10) = Format(rst2!HOSPaidAmt, "#,##0.00;(#,##0.00)")
                    MedixPaidHOS = Format(rst2!HOSPaidAmt, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(10) = "0.00"
                    MedixPaidHOS = 0
                End If
                
                If Len(rst2!MBMPaidAmt) Then
                    Record.SubItems(11) = Format(rst2!MBMPaidAmt, "#,##0.00;(#,##0.00)")
                    MedixPaidMBM = Format(rst2!MBMPaidAmt, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(11) = "0.00"
                    MedixPaidMBM = 0
                End If
                
                'If Len(rst2!MedixPaidAmt) Then
                 '   Record.SubItems(5) = Format(rst2!MedixPaidAmt, "#,##0.00;(#,##0.00)")
                 '   MedixPaidMBM = Format(rst2!MedixPaidAmt, "#,##0.00;(#,##0.00)")
                'Else
                    Record.SubItems(12) = "0.00"
                    Paid2Medix = 0
                'End If
                
                'If Len(rst2!OthersPaidAmt) Then
                 '   Record.SubItems(5) = Format(rst2!OthersPaidAmt, "#,##0.00;(#,##0.00)")
                 '   MedixPaidMBM = Format(rst2!OthersPaidAmt, "#,##0.00;(#,##0.00)")
                'Else
                    Record.SubItems(13) = "0.00"
                    Paid2Others = 0
                'End If
                
                If Len(rst2!PAYPaidAmt) Then
                    Record.SubItems(14) = Format(rst2!PaymentBankCharge, "#,##0.00;(#,##0.00)")
                    Pay2PayAmt = rst2!PAYPaidAmt
                Else
                    Record.SubItems(14) = "0.00"
                    Pay2PayAmt = 0
                End If
                
                If Len(rst2!PaymentBankCharge) Then
                    Record.SubItems(15) = Format(rst2!PaymentBankCharge, "#,##0.00;(#,##0.00)")
                    BankCharges = rst2!PaymentBankCharge
                Else
                    Record.SubItems(15) = "0.00"
                    BankCharges = 0
                End If
                
                If Len(rst2!ContraCaseAmt) Then
                    Record.SubItems(16) = Format(rst2!ContraCaseAmt, "#,##0.00;(#,##0.00)")
                    Contra = Format(rst2!ContraCaseAmt, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(16) = "0.00"
                    Contra = 0
                End If
                
                TotalPaid = Format(CDbl(MedixPaidHOS) + CDbl(MedixPaidMBM) + CDbl(Paid2Medix) + CDbl(Paid2Others) + CDbl(BankCharges) + CDbl(Contra) + CDbl(Pay2PayAmt), "#,##0.00;(#,##0.00)")
                Record.SubItems(17) = Format(CDbl(MedixPaidHOS) + CDbl(MedixPaidMBM) + CDbl(Paid2Medix) + CDbl(Paid2Others) + CDbl(BankCharges) + CDbl(Contra), "#,##0.00;(#,##0.00)")
                
                'CashPosition
                'Record.SubItems(18) = Format(CDbl(totalREc) - CDbl(IIf(TotalPaid < 0, (TotalPaid * -1), TotalPaid)), "#,##0.00;(#,##0.00)")
                Record.SubItems(18) = Format(CDbl(totalREc) - TotalPaid, "#,##0.00;(#,##0.00)")
                'CashPos = Format(CDbl(totalREc) - CDbl(IIf(TotalPaid < 0, (TotalPaid * -1), TotalPaid)), "#,##0.00;(#,##0.00)")
                CashPos = Format(CDbl(totalREc) - TotalPaid, "#,##0.00;(#,##0.00)")
                If Len(rst2!CLMPayableByMedix2H) Then
                    Record.SubItems(24) = Format(rst2!CLMPayableByMedix2H, "#,##0.00;(#,##0.00)")
                    Pay2HOS = Format(rst2!CLMPayableByMedix2H, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(24) = "0.00"
                    Pay2HOS = 0
                End If
                
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(25) = Format(rst2!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                    Pay2MBM = Format(rst2!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                Else
                    Record.SubItems(25) = "0.00"
                    Pay2MBM = 0
                End If
                
                Record.SubItems(19) = Format(CDbl(Pay2HOS) - CDbl(MedixPaidHOS), "#,##0.00;(#,##0.00)")
                OSHos = Format(CDbl(Pay2HOS) - CDbl(MedixPaidHOS), "#,##0.00;(#,##0.00)")
                
                Record.SubItems(20) = Format(CDbl(RecMBM) + CDbl(Pay2MBM) - CDbl(MedixPaidMBM), "#,##0.00;(#,##0.00)")
                OSMBM = Format(CDbl(RecMBM) + CDbl(Pay2MBM) - CDbl(MedixPaidMBM), "#,##0.00;(#,##0.00)")
                
                Record.SubItems(21) = Format(GainLoss, "#,##0.00;(#,##0.00)")
                GainLoss = Format(GainLoss, "#,##0.00;(#,##0.00)")
                
                Record.SubItems(22) = Format((CDbl(OSMBM) + CDbl(OSHos)) - CDbl(GainLoss), "#,##0.00;(#,##0.00)")
                'GainLoss = Format(GainLoss, "#,##0.00;(#,##0.00)")
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(23) = Format(rst2!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    TotalApp = rst2!CLMTotalApproved
                Else
                    Record.SubItems(23) = "0.00"
                    TotalApp = 0
                End If
                
                'Record.SubItems(23) = Format(CDbl(RecMedix) - CDbl(Paid2Medix), "#,##0.00;(#,##0.00)")
                'Pay2Medix = Format(CDbl(RecMedix) - CDbl(Paid2Medix), "#,##0.00;(#,##0.00)")
                
                'Record.SubItems(24) = Format(CDbl(RecOthers) - CDbl(Paid2Others), "#,##0.00;(#,##0.00)")
                'Pay2Others = Format(CDbl(RecOthers) - CDbl(Paid2Others), "#,##0.00;(#,##0.00)")
                
                
                'Record.SubItems(25) = Format(CDbl((RecMNRB) + CDbl(RecMBM) + CDbl(RecMedix) + CDbl(RecOthers)) - CDbl(TotalApp), "#,##0.00;(#,##0.00)")
                'GainLoss = Format(CDbl((RecMNRB) + CDbl(RecMBM) + CDbl(RecMedix) + CDbl(RecOthers)) - CDbl(TotalApp), "#,##0.00;(#,##0.00)")
                
                
                'Record.SubItems(24) = Format(CDbl(GainLoss) - CDbl(Contra), "#,##0.00;(#,##0.00)")
                'CheckTotal = Format(CDbl(GainLoss) - CDbl(Contra), "#,##0.00;(#,##0.00)")
                
                Record.SubItems(26) = Format(CDbl(TotalApp) - (CDbl(Pay2HOS) + CDbl(Pay2MBM)), "#,##0.00;(#,##0.00)")
                CheckTotal = Format(CDbl(TotalApp) - (CDbl(Pay2HOS) + CDbl(Pay2MBM)), "#,##0.00;(#,##0.00)")
                
                'If Len(CaseStatus) Then
                '    Record.SubItems(25) = CaseStatus
                'End If
                
                Current = Current + 1
                lblCurRecord = Current
                
                
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
End Sub

Private Sub lstTrustAC_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstTrustAC, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstTrustAC, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstTrustAC, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstTrustAC, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstTrustAC, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstTrustAC, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstTrustAC, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstTrustAC, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstTrustAC, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstTrustAC, ColumnHeader.Index, ldtDateTime, m_blnDirection(12)
    Case 13
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    Case 14
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    Case 15
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
    Case 16
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
    Case 17
        SortListView lstTrustAC, ColumnHeader.Index, ldtString, m_blnDirection(17)
    Case 18
        SortListView lstTrustAC, ColumnHeader.Index, ldtDateTime, m_blnDirection(18)
    Case 19
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
    Case 20
        SortListView lstTrustAC, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lstTrustAC, ColumnHeader.Index, ldtDateTime, m_blnDirection(21)
    Case 22
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(22)
    Case 23
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
    Case 24
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
    Case 25
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
    Case 26
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
    Case 27
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
    Case 28
        SortListView lstTrustAC, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
    End Select
End Sub

Private Sub lstTrustAC_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ProcessCheckedItem(Item)
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
                
    If Item.Checked = True Then
        If Trim(Item.SubItems(1)) = "Yes" Then
            MsgBox "Record has been flaged!", vbCritical
            Item.Checked = False
        Else
            Item.SubItems(1) = "Yes"
        End If
    Else
        Item.SubItems(1) = ""
    End If
End Function

Private Sub UpdateRecord(CASNo As String, CASVer As String)
Dim Update
Dim strsql As String
Dim REC As New ADODB.Recordset
    
    strsql = " Select CASNumber,ClmVersion, ReconFlag,ReconBy, ReconDate from ReconFlag " & _
                " Where CASNumber = '" & Trim(CASNo) & "' AND CLMVersion = '" & Trim(CASVer) & "'"
    REC.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
    If REC.recordCount = 0 Then
        REC.AddNew
    End If
        With REC
            !CasNumber = CASNo
            !clmVersion = CASVer
            !ReconFlag = "Y"
            !ReconBy = Logon
            !ReconDate = Date
        End With
        REC.Update
    REC.Close
    Set REC = Nothing

    strsql = " Select CASNumber,ClmVersion, ReconFlag,ReconBy, ReconDate from ReconFlag " & _
                " Where CASNumber = '" & Trim(CASNo) & "' AND CLMVersion = '" & Trim(CASVer) & "'"
    REC.Open strsql, medixneptuneconnPayment, adOpenKeyset, adLockOptimistic
    
    If REC.recordCount = 0 Then
        REC.AddNew
    End If
        With REC
            !CasNumber = CASNo
            !clmVersion = CASVer
            !ReconFlag = "Y"
            !ReconBy = Logon
            !ReconDate = Date
        End With
        REC.Update
    REC.Close
    Set REC = Nothing

End Sub
