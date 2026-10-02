VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmPrintPVRange 
   Caption         =   "Print PV Range"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   6840
      Width           =   7935
      Begin VB.TextBox txtPrintingStatus 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Text            =   "Printing Not In Progress !"
         Top             =   195
         Width           =   2175
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6840
         TabIndex        =   13
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   12
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Printing Status"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblPymtType 
         Caption         =   "Label5"
         Height          =   255
         Left            =   5520
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotalchqSel 
         Caption         =   "Label11"
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotalChqList 
         Caption         =   "Label11"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      Begin VB.TextBox txtPVNoFr 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txtPVNoTo 
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   180
         Width           =   1455
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         Top             =   1440
         Width           =   3375
      End
      Begin VB.OptionButton optSihat 
         Caption         =   "CA - Sihat Citi1"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optNSMAA 
         Caption         =   "CC - NS MAA"
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   150
         Width           =   735
      End
      Begin VB.OptionButton optNS 
         Caption         =   "CB - NS Citi2"
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "PV No  From"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Account"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lsPVListing 
      Height          =   6015
      Left            =   0
      TabIndex        =   19
      Top             =   840
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10610
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PV No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PV Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Payee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Claim No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cheque No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cheque Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Cheque Amt"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Print PV Range"
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
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmPrintPVRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public chqMode As String
Dim clickStatus As String
Dim strAppend As String
Dim tmpAdd1, tmpAdd2, tmpAdd3 As String
Dim tmpPCode, tmpCity, tmpState As String
Dim objword As Word.Application
Dim objdoc As Word.Document
Dim tmpPayMBMDollar As String
Dim tmpPayMBMCents As String
Dim tmpPostCode, TmpStateCode, TmpCityCode As String
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim i As Integer
Dim strsql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
    Screen.MousePointer = vbHourglass
    If lsPVListing.ListItems.Count > 0 Then
        If MsgBox("Do you want to print now?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
    
            'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
            RecFound = False
            txtPrintingStatus = "Printing In Progress!"
            CmdPrint.Enabled = False
            CmdExit.Enabled = False
            If chqMode = "MedixPayMBM" Then
                Call PVPrintingMBM
            ElseIf chqMode = "MedixPayHOS" Then
                'Call PVPrintingCLN
            End If
        End If
    Else
        MsgBox "Please select PV to print", vbOKOnly + vbCritical
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PVPrintingMBM()
'On Error Resume Next
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim i As Integer
Dim strsql As String
Dim PV As New ADODB.Recordset
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim PatientSql As String
Dim PatientRec As New ADODB.Recordset
'Dim TmpHLtCode As String
Dim tmpPatientName As String
Dim TotBankCharges As Double
Dim TmpPrtAmt As String
Dim tmpBankCharges As Double
Dim RecFound As Boolean
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim RecCnt As Integer
    Screen.MousePointer = vbHourglass
    TotBankCharges = 0
    tmpBankCharges = 0
    'If txtPVNo = "" Then
    '    MsgBox "Please Enter or Generate PV No", vbOKOnly + vbCritical
    '    Screen.MousePointer = vbDefault
    '    Exit Sub
    'End If
    ' Don't allow user to affect workbook
        strsql = "And PVNo >= '" & Trim(txtPVNoFr) & "' And PVNo <= '" & Trim(txtPVNoTo) & "' "
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        RecCnt = 0
        RecFound = False
        Set cmd1.ActiveConnection = medixmasterconnPayment
        cmd1.CommandTimeout = 300
        cmd1.CommandType = adCmdStoredProc
        cmd1.CommandText = "SP_Payment"
        Set Prm1 = cmd1.CreateParameter("PaymentType", adVarChar, adParamInput, 2000, "Pay2MBM")
        cmd1.Parameters.Append Prm1
        Set Prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, strsql)
        cmd1.Parameters.Append Prm2
        PV.CursorLocation = adUseClient
        PV.CursorType = adOpenForwardOnly
        PV.CacheSize = 100
        Set PV = cmd1.Execute
        Do While Not PV.EOF
            RecCnt = 0
            RecCnt = RecCnt + 1
            If objExcel Is Nothing Then
                Set objExcel = CreateObject("Excel.application")
                    objExcel.Interactive = False
                    objExcel.Visible = False
                    objExcel.WindowState = xlMaximized
                Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\PV-Reimbursement.xlt")
                Set objWorksheet = objWorkbook.Sheets(1)
                objExcel.DisplayAlerts = False
                Set objWorksheet = objWorkbook.Worksheets.Item(1)
            End If
                
            'On Error GoTo HANDLE_ERROR

            'If RecCnt = 1 Then
                tmpAdd1 = IIf(Len(PV!MBMAdd1), PV!MBMAdd1, "")
                tmpAdd2 = IIf(Len(PV!MBMAdd2), PV!MBMAdd2, "")
                tmpAdd3 = IIf(Len(PV!MBMAdd3), PV!MBMAdd3, "")
                tmpPCode = IIf(Len(PV!PayeePCode), PV!PayeePCode, "")
                tmpCity = IIf(Len(PV!MBMCity), PV!MBMCity, "")
                tmpState = IIf(Len(PV!MBMState), PV!MBMState, "")
                objWorksheet.Cells(3, "B") = IIf(Len(PV!CLMPayto), PV!CLMPayto, "")
                objWorksheet.Cells(4, "B") = tmpAdd1
                If Len(Trim(tmpAdd2)) Then
                    objWorksheet.Cells(5, "B") = tmpAdd2
                    
                    If Len(Trim(tmpAdd3)) Then
                        objWorksheet.Cells(6, "B") = tmpAdd3
                        objWorksheet.Cells(7, "B") = tmpPCode & " " & tmpCity
                        objWorksheet.Cells(8, "B") = tmpState
                    Else
                        objWorksheet.Cells(6, "B") = tmpPCode & " " & tmpCity
                        objWorksheet.Cells(7, "B") = tmpState
                    End If
                    
                Else
                    objWorksheet.Cells(5, "B") = tmpPCode & " " & tmpCity
                    objWorksheet.Cells(6, "B") = tmpState
                End If
                
                
                
                objWorksheet.Cells(3, "H") = PV!PVNo  '& "/" & Format(PV!PVDate, "MM/YY")
                objWorksheet.Cells(4, "H") = PV!PVChequeNo
                objWorksheet.Cells(5, "H") = PV!PVDate
                objWorksheet.Cells(33, "B") = Logon
            'End If  ' RECCnt = 1
            objWorksheet.Cells(12 + CDbl(RecCnt), "A") = RecCnt
            objWorksheet.Cells(12 + CDbl(RecCnt), "B") = PV!CasNumber & "-" & PV!claimversion
            objWorksheet.Cells(12 + CDbl(RecCnt), "C") = PV!ClaimCloseDate
            objWorksheet.Cells(12 + CDbl(RecCnt), "D") = PV!BordxRefNo
        
            RecSql = "select * from ReceiptfrMNRB where CASNumber = '" & PV!CasNumber & "' And ClaimVersion =  '" & PV!claimversion & "'"
            If Trim(PV!HLTCode) = "S" Then
                RecSql = "Select REceiptNo, CheckAmt From REceiptFrMNRB WHERE CASNumber='" & Trim(PV!CasNumber) & _
                         "' AND ClaimVersion='" & Trim(PV!claimversion) & "'"
            ElseIf Trim(PV!HLTCode) = "N" Then
                RecSql = "Select REceiptNo, CheckAmt From REceiptFrPay WHERE CASNumber='" & Trim(PV!CasNumber) & _
                         "' AND ClaimVersion='" & Trim(PV!claimversion) & "'"
            End If
            ReceiptRec.Open RecSql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            If ReceiptRec.recordCount Then
                objWorksheet.Cells(12 + CDbl(RecCnt), "E") = IIf(Len(ReceiptRec!ReceiptNo), ReceiptRec!ReceiptNo, "")
            End If
            ReceiptRec.Close
            Set ReceiptRec = Nothing
            tmpPatientName = IIf(Len(PV!MBMName), Trim(PV!MBMName), "")
            If PV!MBMPatientCovID <> "00" Then
                PatientSql = "Select MBMCName FROM MBMCoveredPersons WHERE MBMCNumber ='" & Trim(PV!MBMNumber) & _
                             "' And MBMCCoverID ='" & Trim(PV!MBMPatientCovID) & "'"
                PatientRec.Open PatientSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                If PatientRec.recordCount Then
                    tmpPatientName = Trim(PatientRec!MBMCName)
                End If
                PatientRec.Close
                Set PatientRec = Nothing
            End If
            objWorksheet.Cells(12 + CDbl(RecCnt), "F") = tmpPatientName
            TmpPrtAmt = PV!MBMPaidAmt
            tmpBankCharges = IIf(IsNumeric(PV!PVBankCharges), PV!PVBankCharges, 0)
            
            If tmpBankCharges <> 0 Then
                TmpPrtAmt = PV!CLMPayableByMedix2M
                TotBankCharges = TotBankCharges + tmpBankCharges
            End If
            objWorksheet.Cells(12 + CDbl(RecCnt), "H") = TmpPrtAmt
               
            If CDbl(TotBankCharges) > 0 Then
                objWorksheet.Cells.Font.Bold(12 + CDbl(RecCnt), "B") = "Less : Bank Charges"
                objWorksheet.Cells(12 + CDbl(RecCnt), "H") = TotBankCharges * -1
            End If
               
            objWorkbook.PrintOut
            'objWorkbook.SaveAs (LocCardPath & "PaymentVoucher\" & "tmpPaycheque" & ".xlt")
            objWorkbook.SaveAs ("C:\Template\PaymentVoucher" & "tmpPaycheque" & ".xlt")
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
            Set objExcel = Nothing
               
            PV.MoveNext
        Loop
        
ExitFunction:
            txtPrintingStatus = "Printing Not In Progress!"
            CmdPrint.Enabled = False
            CmdExit.Enabled = True
            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
            'medixmasterconnPayment.Close
            'Set medixmasterconnPayment = Nothing
            Exit Sub
'HANDLE_ERROR:
'            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
'            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
'            Set objWorksheet = Nothing
'            Set objWorkbook = Nothing
'            objExcel.Quit
            GoTo ExitFunction
End Sub

Private Sub CmdSearch_Click()
Dim RecSel As Double
    Screen.MousePointer = vbHourglass

    RecSel = 0
    lblTotalchqSel = 0
    If txtPVNoFr = "" Then
        MsgBox "Please Enter PV No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If txtPVNoTo = "" Then
        MsgBox "Please Enter PV No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'lblTotalchqSel = RecSel
    lsPVListing.ListItems.Clear
  '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

        If chqMode = "MedixPayMBM" Then
            Call Payment2MBMListing
        ElseIf chqMode = "MedixPayCLN" Then
        '    Call PVCLNListing
        End If
 '   medixmasterconnPayment.Close
 '   Set medixmasterconnPayment = Nothing
    Screen.MousePointer = Default

End Sub
Private Sub Payment2MBMListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim CasNumber As String
Dim WSVersion As String
Dim Record As ListItem
Dim Check As Boolean
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim ContrSql As String
Dim ContraREc As New ADODB.Recordset
Dim PayCurrRec As New ADODB.Recordset
Dim strsql As String
Dim Pay2MBMAmt As Double
Dim Paid2MBMAmt As Double
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpReceiptNo As String
Dim tmpReceiptAmt As Double
    
    strAppend = " And PVNo >= '" & Trim(txtPVNoFr) & "' And PVNo <= '" & Trim(txtPVNoTo) & "' "

    RecFound = False
    Set cmd1.ActiveConnection = medixmasterconnPayment
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_Payment"
    Set Prm1 = cmd1.CreateParameter("PaymentType", adVarChar, adParamInput, 2000, "Pay2MBM")
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, strAppend)
    cmd1.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd1.Execute

        Do While Not rst2.EOF
        RecFound = True
            With rst2
                Set Record = lsPVListing.ListItems.Add(, , !PVNo)
                If Len(!PVDate) Then
                    Record.SubItems(1) = !PVDate
                End If
                Record.SubItems(2) = IIf(Len(!CLMPayto), !CLMPayto, "")
                If Len(!CasNumber & "-" & !claimversion) Then
                    Record.SubItems(3) = (!CasNumber & "-" & !claimversion)
                End If
                Record.SubItems(4) = Trim(!PVChequeNo)
                Record.SubItems(5) = Trim(!PVChequeDate)
                Record.SubItems(6) = IIf(IsNumeric(!PVChequeAmt), !PVChequeAmt, "")
            End With
            rst2.MoveNext
        Loop
    
        If RecFound = True Then
            CmdPrint.Enabled = True
        Else
            MsgBox "No record found", vbOKOnly + vbCritical
        End If

    rst2.Close
    Set rst2 = Nothing

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

Private Sub txtPVNoFr_Click()
    CmdPrint.Enabled = False
End Sub

Private Sub txtPVNoTo_Click()
    CmdPrint.Enabled = False
End Sub




