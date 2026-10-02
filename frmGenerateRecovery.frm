VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGenerateRecovery 
   Caption         =   "Generate Recovery Listing"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   11865
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   7200
      Width           =   11775
      Begin VB.TextBox LblTotalAmt 
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   9480
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7920
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10200
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10920
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdPrintDetails 
         Caption         =   "Print"
         Height          =   375
         Left            =   8640
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   50
      TabIndex        =   7
      Top             =   240
      Width           =   11775
      Begin MSMask.MaskEdBox TxtBordxDate2 
         Height          =   315
         Left            =   7560
         TabIndex        =   4
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordxDate1 
         Height          =   315
         Left            =   5760
         TabIndex        =   3
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   3240
      End
      Begin VB.TextBox TxtBordxNo 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   500
         Width           =   1455
      End
      Begin VB.TextBox TxtBordxNoTo 
         Height          =   285
         Left            =   2985
         TabIndex        =   2
         Top             =   500
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtWSDateTo 
         Height          =   315
         Left            =   7560
         TabIndex        =   6
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtWSDateFr 
         Height          =   315
         Left            =   5760
         TabIndex        =   5
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   26
         Top             =   510
         Width           =   375
      End
      Begin VB.Label Label70 
         Caption         =   "WS Reg Date"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   25
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         Height          =   255
         Index           =   1
         Left            =   2720
         TabIndex        =   23
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label70 
         Caption         =   "Bordx Date"
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   22
         Top             =   250
         Width           =   975
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   21
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstPrintBordx 
      Height          =   6015
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10610
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
         Text            =   "Claim No"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Patient Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Hospital Name"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Bill Amt"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Recovery Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "BordxAmt"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Recovery Listing"
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
      TabIndex        =   20
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmGenerateRecovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intExit As Integer
Dim ScanCode As String
Dim RePrintStatus As Boolean
Dim CancelPrinting As Boolean
Dim bPrintPass As Boolean
Dim tmpPrintStatus As Boolean
Dim tmpRptByDept As Boolean
Dim GrpCompany As String
Dim tmpRptFormat As String
Dim strAppend As String
Dim tmpRptType As String
Private m_blnDirection(1 To 13) As Boolean

Private Sub CboGrpCom_Change()
    Call SearchGrpCompany
End Sub

Private Sub CboGrpCom_LostFocus()
    Call SearchGrpCompany
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

Private Sub cmdPrintDetails_Click()
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

        Call PrintTemplate
    
   ' medixmasterconn.Close
   ' medixmasterconnWS.Close
End Sub

Private Sub Form_Load()
Dim AA As Integer
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    cmdPrintDetails.Enabled = False
    Call combolist
End Sub

'******************** Start Command Button ********************************
Private Sub CmdClear_Click()
    lstPrintBordx.ListItems.Clear
    TxtBordxNo.Text = 0
    TxtBordxNo.Text = ""
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    cmdPrintDetails.Enabled = False
    TxtBordxNo.SetFocus
End Sub

Private Sub CmdSearch_Click()

    Screen.MousePointer = vbHourglass
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"

    If Trim(CboGrpCom) = "" Then
        MsgBox "Please select Group Company", vbOKOnly + vbCritical, DialogBoxTitle
        TxtBordxNo.SetFocus
    Else
            CmdClear.Enabled = False
            Call SearchRecords
            'StrAppend = " And BordxRefNo = '" & ChkStr(Trim(TxtBordxNo)) & "'"
            cmdPrintDetails.Enabled = True
            Call PrintBordxListing(strAppend)
    End If
  '  medixmasterconn.Close
  '  medixmasterconnWS.Close
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub PrintBordxListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim TotRec As String
Dim CurrRec As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Check As Boolean

    TotRec = 0
    CurrRec = 0
    Check = False
    lstPrintBordx.ListItems.Clear
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "ClaimReport1"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, "234")
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "No record found ", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Do
            Do Until rst2.EOF
                TotRec = rst2.recordCount
                lblTotal = TotRec
                 With rst2
                    Set Record = lstPrintBordx.ListItems.Add(, , !CASNumber & "-" & !ClaimVersion)
                    Record.SubItems(1) = IIf(Len(Trim(!PatientName)), Trim(!PatientName), "")
                    Record.SubItems(2) = IIf(Len(Trim(!MBMNumber)), Trim(!MBMNumber), "")
                    Record.SubItems(3) = IIf(Len(Trim(!HosName)), Trim(!HosName), "")
                    Record.SubItems(4) = IIf(Len(rst2!CLMTotalActual), Format(rst2!CLMTotalActual, "#,##0.00"), 0)
                    Record.SubItems(5) = IIf(Len(rst2!RecoveryAmt), Format(rst2!RecoveryAmt * -1, "#,##0.00"), 0)
                    Record.SubItems(6) = IIf(Len(rst2!CLMBordxAmt), Format(rst2!CLMBordxAmt, "#,##0.00"), 0)
                    Record.SubItems(7) = IIf(Len(Trim(!BordxRefNo)), Trim(!BordxRefNo), "")
                    CurrRec = CurrRec + 1
                    lblCurrent = CurrRec
                    TotalAmt = TotalAmt + rst2!CLMBordxAmt
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
End Sub

Private Sub txtBordxDate1_LostFocus()
    If TxtBordxDate1 <> "__/__/____" Then
        If IsDate(TxtBordxDate1) = False Then
            MsgBox "Invalid date format!", vbCritical
        End If
    End If

End Sub

Private Sub txtBordxDate2_LostFocus()
    If TxtBordxDate2 <> "__/__/____" Then
        If IsDate(TxtBordxDate2) = False Then
            MsgBox "Invalid date format!", vbCritical
        End If
    End If

End Sub

Private Sub TxtBordxNo_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub ComboKeyPress(tmpCombo As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(tmpCombo.Text, tmpCombo.SelStart) + Chr(KeyAscii) + Mid(tmpCombo.Text, tmpCombo.SelStart + tmpCombo.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To tmpCombo.ListCount - 1
       
        If NewText = LCase(Left(tmpCombo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = tmpCombo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               tmpCombo.Text = ValidValue
               tmpCombo.SelStart = 0
               tmpCombo.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub PrintTemplate()
Dim StrSql1 As String
Dim strsqlPlutoBor As String
Dim PlutoBor As New ADODB.Recordset
Dim sqlinsert As String
Dim sqlbor As String
Dim TmpBordxDate As String
Dim tmpDateVisit As String
Dim TempMBMName As String
Dim TempMBMDept As String
Dim tempPath As String
Dim PLUTOStatus As Boolean
Dim BOR As New ADODB.Recordset
Dim TempRemarks As String
Dim TempGRPCompany As String
Dim TempCLNName As String
Dim TempMBMCName As String
Dim TempPatientName As String
Dim TempCASDiagnosis5 As String
Dim PVCnt As Integer
Dim tmpSeqNo As Integer
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim tmpTotalBordxAmt As String
Dim tmpTotalActual As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpTotalRecovery As String
        tmpTotalRecovery = 0
        tmpTotalBordxAmt = 0
        PVCnt = 0
        tmpSeqNo = 0
        tmpTotalActual = 0
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            '(LocCardPath & "Claim" & "\EXCLetterMAS.DOT")
                If tmpRptType = "2" Then
                    Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\" & "Recovery Listing MDEC.xls")
                ElseIf tmpRptType = "1" Then
                    Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\" & "Recovery Listing Template.xls")
                End If
            'Set objWorkbook = objExcel.Workbooks.Add("C:\Template\Airod OP Template.xls")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
    
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "ClaimReport1"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, "234")
        cmd.Parameters.Append Prm2
        BOR.CursorLocation = adUseClient
        BOR.CursorType = adOpenForwardOnly
        BOR.CacheSize = 100
        Set BOR = cmd.Execute

        Do While Not BOR.EOF
            
            'RecFound = True
            PVCnt = PVCnt + 1
                If PVCnt = "1" Then
                    objWorksheet.Cells(7, "C") = BOR!PAYCompanyName
                    objWorksheet.Cells(8, "C") = BOR!MBMPolicyNo
                    objWorksheet.Cells(9, "C") = Format(BOR!MBMPayorEffDate, "dd/mm/yyyy") & " To " & Format(BOR!MBMPayorEXPDate, "dd/mm/yyyy")
                    'objWorksheet.Cells(10, "C") = TxtBordxNo
                    objWorksheet.Cells(10, "C") = Format(Date, "dd/mm/yyyy")
                    
                End If
                If tmpRptType = "1" Then
                    tmpSeqNo = CDbl(tmpSeqNo) + 1
                    objWorksheet.Cells(12 + CDbl(PVCnt), "A") = tmpSeqNo
                    objWorksheet.Cells(12 + CDbl(PVCnt), "B") = BOR!CASNumber & "-" & BOR!ClaimVersion
                    objWorksheet.Cells(12 + CDbl(PVCnt), "C") = BOR!PatientName
                    objWorksheet.Cells(12 + CDbl(PVCnt), "D") = BOR!ICNo
                    objWorksheet.Cells(12 + CDbl(PVCnt), "E") = BOR!MBMName
                    objWorksheet.Cells(12 + CDbl(PVCnt), "F") = BOR!MBMIcBcPp
                    objWorksheet.Cells(12 + CDbl(PVCnt), "G") = BOR!MBMEmployeeNo
                    objWorksheet.Cells(12 + CDbl(PVCnt), "H") = BOR!MBMDeptCode
                    objWorksheet.Cells(12 + CDbl(PVCnt), "I") = BOR!MBMDepartment
                    objWorksheet.Cells(12 + CDbl(PVCnt), "J") = BOR!PLNCode
                    objWorksheet.Cells(12 + CDbl(PVCnt), "K") = BOR!CASDateAdmitted
                    objWorksheet.Cells(12 + CDbl(PVCnt), "L") = BOR!CASDischargeDate
                    objWorksheet.Cells(12 + CDbl(PVCnt), "M") = BOR!ClaimOpenDate
                    objWorksheet.Cells(12 + CDbl(PVCnt), "N") = BOR!HosName
                    objWorksheet.Cells(12 + CDbl(PVCnt), "O") = BOR!DIAFinalDiag1 & "/" & BOR!DIAFinalDiag2 & "/" & BOR!DIAFinalDiag3
                    objWorksheet.Cells(12 + CDbl(PVCnt), "P") = BOR!CLMBordxAmt
                    objWorksheet.Cells(12 + CDbl(PVCnt), "Q") = BOR!RecoveryAmt * -1
                    objWorksheet.Cells(12 + CDbl(PVCnt), "R") = BOR!CLMTotalActual
                    tmpTotalActual = CDbl(tmpTotalActual) + CDbl(BOR!CLMTotalActual)
                    tmpTotalBordxAmt = CDbl(tmpTotalBordxAmt) + CDbl(BOR!CLMBordxAmt)
                    tmpTotalRecovery = CDbl(tmpTotalRecovery) + CDbl(BOR!RecoveryAmt)
                ElseIf tmpRptType = "2" Then
                    tmpSeqNo = CDbl(tmpSeqNo) + 1
                    objWorksheet.Cells(12 + CDbl(PVCnt), "A") = tmpSeqNo
                    objWorksheet.Cells(12 + CDbl(PVCnt), "B") = BOR!CASNumber & "-" & BOR!ClaimVersion
                    objWorksheet.Cells(12 + CDbl(PVCnt), "C") = BOR!PatientName
                    objWorksheet.Cells(12 + CDbl(PVCnt), "D") = BOR!ICNo
                    objWorksheet.Cells(12 + CDbl(PVCnt), "E") = BOR!MBMName
                    objWorksheet.Cells(12 + CDbl(PVCnt), "F") = BOR!MBMEmployeeNo
                    objWorksheet.Cells(12 + CDbl(PVCnt), "G") = BOR!MBMCostCenter
                    objWorksheet.Cells(12 + CDbl(PVCnt), "H") = BOR!InvDate
                    objWorksheet.Cells(12 + CDbl(PVCnt), "I") = BOR!CASDateAdmitted
                    objWorksheet.Cells(12 + CDbl(PVCnt), "J") = BOR!CASDischargeDate
                    objWorksheet.Cells(12 + CDbl(PVCnt), "K") = BOR!HosName
                    objWorksheet.Cells(12 + CDbl(PVCnt), "L") = BOR!DIAFinalDiag1 & "/" & BOR!DIAFinalDiag2 & "/" & BOR!DIAFinalDiag3
                    objWorksheet.Cells(12 + CDbl(PVCnt), "M") = BOR!DebitCreditRefNo
                    objWorksheet.Cells(12 + CDbl(PVCnt), "N") = BOR!CLMBordxAmt
                    objWorksheet.Cells(12 + CDbl(PVCnt), "O") = BOR!RecoveryAmt * -1
                    objWorksheet.Cells(12 + CDbl(PVCnt), "P") = BOR!CLMTotalActual
                    tmpTotalActual = CDbl(tmpTotalActual) + CDbl(BOR!CLMTotalActual)
                    tmpTotalBordxAmt = CDbl(tmpTotalBordxAmt) + CDbl(BOR!CLMBordxAmt)
                    tmpTotalRecovery = CDbl(tmpTotalRecovery) + CDbl(BOR!RecoveryAmt)

                
                End If
            BOR.MoveNext
        
        
        'BOR.MoveNext
        Loop
        BOR.Close
        Set BOR = Nothing
        PVCnt = PVCnt + 2
        If tmpRptType = "2" Then

            objWorksheet.Cells(12 + CDbl(PVCnt), "O") = Format(tmpTotalBordxAmt, "#,##0.00")
            objWorksheet.Cells(12 + CDbl(PVCnt), "P") = Format(tmpTotalRecovery * -1, "#,##0.00")
            objWorksheet.Cells(12 + CDbl(PVCnt), "Q") = Format(tmpTotalActual, "#,##0.00")
            objWorksheet.Cells(12 + CDbl(PVCnt), "O").Font.Bold = True
            objWorksheet.Cells(12 + CDbl(PVCnt), "P").Font.Bold = True
            objWorksheet.Cells(12 + CDbl(PVCnt), "Q").Font.Bold = True
        ElseIf tmpRptType = "1" Then
            objWorksheet.Cells(12 + CDbl(PVCnt), "P") = Format(tmpTotalBordxAmt, "#,##0.00")
            objWorksheet.Cells(12 + CDbl(PVCnt), "Q") = Format(tmpTotalRecovery * -1, "#,##0.00")
            objWorksheet.Cells(12 + CDbl(PVCnt), "R") = Format(tmpTotalActual, "#,##0.00")
            objWorksheet.Cells(12 + CDbl(PVCnt), "P").Font.Bold = True
            objWorksheet.Cells(12 + CDbl(PVCnt), "Q").Font.Bold = True
            objWorksheet.Cells(12 + CDbl(PVCnt), "R").Font.Bold = True
        
        End If
        'objWorksheet.Cells(7 + CDbl(PVCnt), "K").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        PVCnt = PVCnt + 1
        objWorksheet.Cells(12 + CDbl(PVCnt), "B") = "Kindly note that payment term is 14 days and payment should be issued to MediExpress (Malaysia) Sdn Bhd"
        objWorksheet.Cells(12 + CDbl(PVCnt), "B").Font.Bold = True
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

Private Function SearchRecords()
Dim sql As String

    strAppend = ""
    
    If Len(Trim(TxtBordxDate1)) > 0 And IsDate(TxtBordxDate1) = True Then
        strAppend = strAppend + " And BordxDate >= '" & Format(TxtBordxDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxDate2)) > 0 And IsDate(TxtBordxDate2) = True Then
        strAppend = strAppend + " And BordxDate <= '" & Format(TxtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtWSDateFr)) > 0 And IsDate(txtWSDateFr) = True Then
        strAppend = strAppend + " And ClaimOpenDate >= '" & Format(txtWSDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtWSDateTo)) > 0 And IsDate(txtWSDateTo) = True Then
        strAppend = strAppend + " And ClaimOpenDate <= '" & Format(txtWSDateTo, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(TxtBordxNo)) Then
        strAppend = strAppend & " AND BordxRefNo >= '" & Trim(TxtBordxNo) & "'"
    End If
    
    If Len(Trim(TxtBordxNoTo)) Then
        strAppend = strAppend & " AND BordxRefNo <= '" & Trim(TxtBordxNoTo) & "'"
    End If
    
    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend & " AND MBMGroupCompany = '" & Trim(CboGrpCom) & "'"
    End If

    'StrAppend = StrAppend & " AND RecoveryAmt < 0"
End Function

Private Sub combolist()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String

    CboGrpCom.Clear
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    strsql = "select CoGRPName from CompanyGrp where BTBGClient = 'Y'"
    PAYGRP.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly

    Do While Not PAYGRP.EOF
       CboGrpCom.AddItem Trim(PAYGRP!CoGRPName)
       PAYGRP.MoveNext
    Loop
    
    PAYGRP.Close
    Set PAYGRP = Nothing
 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

Private Sub SearchGrpCompany()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String

    tmpRptType = ""
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    strsql = "select CoGRPName,BTBGRpt from CompanyGrp where CoGRPName = '" & CboGrpCom & "' And BTBGClient = 'Y'"
    PAYGRP.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly

    Do While Not PAYGRP.EOF
       tmpRptType = Trim(PAYGRP!BTBGRpt)
       PAYGRP.MoveNext
    Loop
    
    PAYGRP.Close
    Set PAYGRP = Nothing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

Private Sub txtWSDateFr_LostFocus()
    If txtWSDateFr <> "__/__/____" Then
        If IsDate(txtWSDateFr) = False Then
            MsgBox "Invalid date format!", vbCritical
        End If
    End If

End Sub

Private Sub txtWSDateTo_LostFocus()
    If txtWSDateTo <> "__/__/____" Then
        If IsDate(txtWSDateTo) = False Then
            MsgBox "Invalid date format!", vbCritical
        End If
    End If

End Sub


