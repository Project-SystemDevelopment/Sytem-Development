VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintCN 
   Caption         =   "Print CN Range"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   7935
      Begin VB.OptionButton optNS 
         Caption         =   "CB - NS Citi2"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   4440
         TabIndex        =   14
         Top             =   150
         Width           =   735
      End
      Begin VB.OptionButton optNSMAA 
         Caption         =   "CC - NS MAA"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optSihat 
         Caption         =   "CA - Sihat Citi1"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   2880
         TabIndex        =   11
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtPVNoTo 
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox txtPVNoFr 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Account"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "PV No  From"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   6840
      Width           =   7935
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6840
         TabIndex        =   2
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtPrintingStatus 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "Printing Not In Progress !"
         Top             =   195
         Width           =   2175
      End
      Begin VB.Label lblTotalChqList 
         Caption         =   "Label11"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotalchqSel 
         Caption         =   "Label11"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPymtType 
         Caption         =   "Label5"
         Height          =   255
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Printing Status"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
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
      Caption         =   "Print CN Range"
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
Attribute VB_Name = "frmPrintCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public chqMode As String
Dim clickStatus As String
Dim strAppend As String
Dim tmpMBMNo, tmpPolno, tmpInsName As String
Dim tmpPatientName, tmpDOA, tmpDOD As String
Dim tmpAdd1, tmpAdd2, tmpAdd3 As String
Dim tmpPostCode, TmpStateCode, TmpCityCode, tmpChqNo As String
Dim TmpCLMAttention, tmpPayMode, tmpClaimNo, tmpPVNo, tmpPVDate As String
Dim TmpChqAmt As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    If lsPVListing.ListItems.Count > 0 Then
        If MsgBox("Do you want to print now?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
    
            'MediConnCISPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
            txtPrintingStatus = "Printing In Progress!"
            cmdPrint.Enabled = False
            cmdExit.Enabled = False
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
Dim rst2 As New ADODB.Recordset
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

        strAppend = "And PVNo >= '" & Trim(txtPVNoFr) & "' And PVNo <= '" & Trim(txtPVNoTo) & "' "
        'medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        RecCnt = 0
        RecFound = False
        'tmpCount = 0
        'tmpClaimNo = ""
        Set cmd1.ActiveConnection = medixmasterconnPayment
        cmd1.CommandTimeout = 300
        cmd1.CommandType = adCmdStoredProc
        cmd1.CommandText = "SP_PaymentJul08"
        Set Prm1 = cmd1.CreateParameter("PaymentType", adVarChar, adParamInput, 2000, "Pay2MBM")
        cmd1.Parameters.Append Prm1
        Set Prm2 = cmd1.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, strAppend)
        cmd1.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd1.Execute
            
            Do While Not rst2.EOF
                'tmpCount = tmpCount + 1
               ' If tmpCount > 1 Then
               '     tmpClaimNo = tmpClaimNo & "/" & rst2!CasNumber & "-" & rst2!ClaimVersion
               ' End If
               ' If tmpCount = 1 Then
                    tmpPVNo = rst2!PVNo
                    tmpPVDate = Format(rst2!PVDate, "DD/MM/YYYY")
                    tmpClaimNo = rst2!CasNumber & "-" & rst2!ClaimVersion
                    tmpMBMNo = rst2!MBMNumber
                    tmpPolno = rst2!MBMPolicyNo
                    tmpInsName = rst2!MBMName
                    If Trim(rst2!MBMPatientCovID) = "00" Then
                        tmpPatientName = rst2!MBMName
                    Else
                        tmpPatientName = rst2!MBMCName
                    End If
                    tmpDOA = Format(rst2!CASDateAdmitted, "DD/MM/YYYY")
                    tmpDOD = Format(rst2!CASDischargeDate, "DD/MM/YYYY")
                    tmpAdd1 = IIf(Len(rst2!MBMAdd1), rst2!MBMAdd1, "")
                    tmpAdd2 = IIf(Len(rst2!MBMAdd2), rst2!MBMAdd2, "")
                    tmpAdd3 = IIf(Len(rst2!MBMAdd3), rst2!MBMAdd3, "")
                    tmpPostCode = IIf(Len(rst2!PayeePCode), rst2!PayeePCode, "")
                    TmpStateCode = IIf(Len(rst2!MBMState), rst2!MBMState, "")
                    TmpCityCode = IIf(Len(rst2!MBMCity), rst2!MBMCity, "")
                    TmpCLMAttention = IIf(Len(rst2!CLMAttention), "Attention : " & rst2!CLMAttention, "")
                    tmpPayMode = IIf(Len(rst2!CLMPayMode), rst2!CLMPayMode, "")
                    tmpChqNo = IIf(Len(rst2!PVChequeNo), rst2!PVChequeNo, "")
                    TmpChqAmt = IIf(IsNumeric(rst2!PVChequeAmt), Format(rst2!PVChequeAmt, "#,##0.00"), 0)
                    If objExcel Is Nothing Then
                        Set objExcel = CreateObject("Excel.application")
                            objExcel.Interactive = False
                            objExcel.Visible = False
                            objExcel.WindowState = xlMaximized
                        If rst2!PVPymtMode = "ONLINE" Then
                            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\PaymentMemberOnline1.xlt")
                        Else
                            If Trim(tmpPayMode) = "BY POST" Then
                                Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\PaymentMember1.xlt")
                            Else
                                Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\PaymentMemberBI.xlt")
                            End If
                        End If
                        Set objWorksheet = objWorkbook.Sheets(1)
                        objExcel.DisplayAlerts = False
                        Set objWorksheet = objWorkbook.Worksheets.Item(1)
                    End If
                        
                    On Error GoTo HANDLE_ERROR
                        If Trim(tmpPayMode) = "BY POST" Then
                            'objWorksheet.Cells(4, "F") = tmpPVNo
                            'objWorksheet.Cells(5, "F") = Format(tmpPVDate, "yyyy/mm/dd")
                            
                            objWorksheet.Cells(5, "I") = tmpPVNo
                            objWorksheet.Cells(6, "I") = Format(tmpPVDate, "yyyy/mm/dd")
                            
                        Else
                            objWorksheet.Cells(4, "H") = tmpPVNo
                            objWorksheet.Cells(5, "H") = Format(tmpPVDate, "yyyy/mm/dd")
                        End If
                        objWorksheet.Cells(7, "A") = tmpInsName
                        objWorksheet.Cells(8, "A") = tmpAdd1
                        If Len(Trim(tmpAdd2)) Then
                            objWorksheet.Cells(9, "A") = tmpAdd2
                            If Len(Trim(tmpAdd3)) Then
                                objWorksheet.Cells(10, "A") = tmpAdd3
                                objWorksheet.Cells(11, "A") = tmpPostCode & " " & TmpCityCode
                                objWorksheet.Cells(12, "A") = TmpStateCode
                                objWorksheet.Cells(13, "A") = TmpCLMAttention
                            Else
                                objWorksheet.Cells(10, "A") = tmpPostCode & " " & TmpCityCode
                                objWorksheet.Cells(11, "A") = TmpStateCode
                                objWorksheet.Cells(12, "A") = TmpCLMAttention
                            End If
                            
                        Else
                            objWorksheet.Cells(9, "A") = tmpPostCode & " " & TmpCityCode
                            objWorksheet.Cells(10, "A") = TmpStateCode
                            objWorksheet.Cells(11, "A") = TmpCLMAttention
                        End If
                                                
                        objWorksheet.Cells(19, "B") = tmpClaimNo
                        objWorksheet.Cells(20, "B") = tmpMBMNo
                        objWorksheet.Cells(21, "B") = tmpPolno
                        objWorksheet.Cells(22, "B") = tmpInsName
                        objWorksheet.Cells(23, "B") = tmpPatientName
                        objWorksheet.Cells(24, "B") = Format(tmpDOA, "yyyy/mm/dd")
                        objWorksheet.Cells(25, "B") = Format(tmpDOD, "yyyy/mm/dd")
                        
                        If rst2!PVPymtMode = "ONLINE" Then
                        
                            objWorksheet.Cells(28, "B") = "RM" & " " & TmpChqAmt
                            objWorksheet.Cells(28, "C") = "will be transferred to your " + IIf(Len(rst2!CLMBank), rst2!CLMBank, " bank account") + " "
                            objWorksheet.Cells(31, "A") = "The amount is expected to credit into your account within 3 working days from the date above."
                            objWorksheet.Cells(35, "A") = "Kindly contact " + LogonName + " if the transaction is not Successful."
                            
                        Else
                            If Trim(tmpPayMode) = "BY POST" Then
                                objWorksheet.Cells(28, "C") = tmpChqNo
                                objWorksheet.Cells(28, "I") = TmpChqAmt
                                objWorksheet.Cells(33, "A") = "Kindly contact " + LogonName + " for further information."
                                'objWorksheet.Cells(28, "E") = TmpChqAmt
                            Else
                                objWorksheet.Cells(28, "C") = tmpChqNo
                                objWorksheet.Cells(28, "H") = TmpChqAmt
                                objWorksheet.Cells(33, "A") = "Kindly contact " + LogonName + " for further information."
                            End If
                        End If
                        objWorkbook.PrintOut
                        'objWorkbook.SaveAs (LocCardPath & "PaymentVoucher\" & "tmpPaymentMember" & ".xlt")
                        objWorkbook.SaveAs ("C:\Template\PaymentVoucher\tmpPaymentMember.xlt")
                        Set objWorksheet = Nothing
                        Set objWorkbook = Nothing
                        objExcel.Quit
                        Set objExcel = Nothing
            rst2.MoveNext
            Loop
            rst2.Close
            Set rst2 = Nothing
            
            
        'End If
    
ExitFunction:
            txtPrintingStatus = "Printing Not In Progress!"
            cmdPrint.Enabled = False
            cmdExit.Enabled = True
            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
           ' medixmasterconnPayment.Close
          '  Set medixmasterconnPayment = Nothing
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
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
    'medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

        If chqMode = "MedixPayMBM" Then
            Call Payment2MBMListing
        ElseIf chqMode = "MedixPayCLN" Then
        '    Call PVCLNListing
        End If
   ' medixmasterconnPayment.Close
    'Set medixmasterconnPayment = Nothing
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
    
    strAppend = "And PVNo >= '" & Trim(txtPVNoFr) & "' And PVNo <= '" & Trim(txtPVNoTo) & "' "

    RecFound = False
    Set cmd1.ActiveConnection = medixmasterconnPayment
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_PaymentJul08"
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
                If Len(!CasNumber & "-" & !ClaimVersion) Then
                    Record.SubItems(3) = (!CasNumber & "-" & !ClaimVersion)
                End If
                Record.SubItems(4) = Trim(!PVChequeNo)
                Record.SubItems(5) = Trim(!PVChequeDate)
                Record.SubItems(6) = IIf(IsNumeric(!PVChequeAmt), !PVChequeAmt, "")
            End With
            rst2.MoveNext
        Loop
    
        If RecFound = True Then
            cmdPrint.Enabled = True
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
    cmdPrint.Enabled = False
End Sub

Private Sub txtPVNoTo_Click()
    cmdPrint.Enabled = False
End Sub








