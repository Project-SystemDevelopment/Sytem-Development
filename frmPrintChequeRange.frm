VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintChequeRange 
   Caption         =   "Print Cheque Range"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   8025
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lstChequeListing 
      Height          =   5895
      Left            =   75
      TabIndex        =   13
      Top             =   1200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10398
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cheque No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cheque Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cheque Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Account No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Payment Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   75
      TabIndex        =   4
      Top             =   240
      Width           =   7935
      Begin VB.OptionButton optMX 
         Caption         =   "CG - MX"
         Height          =   255
         Left            =   4320
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   150
         Width           =   735
      End
      Begin VB.OptionButton optNSMAA 
         Caption         =   "CC - NS MAA"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optNS 
         Caption         =   "CB - NS Citi2"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optSihat 
         Caption         =   "CA - Sihat Citi1"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cboAccount 
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtChequeNoTo 
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox txtChequeNoFr 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Account"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Cheque No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   75
      TabIndex        =   8
      Top             =   7080
      Width           =   7935
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   11
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6840
         TabIndex        =   10
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtPrintingStatus 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "Printing Not In Progress !"
         Top             =   195
         Width           =   2175
      End
      Begin VB.Label lblTotalChqList 
         Caption         =   "Label11"
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotalchqSel 
         Caption         =   "Label11"
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPymtType 
         Caption         =   "Label5"
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Printing Status"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Print Cheque Range"
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
      TabIndex        =   3
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmPrintChequeRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public chqMode As String
Dim clickStatus As String

Private Sub cboAccount_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboAccount.Text, cboAccount.SelStart) + Chr(KeyAscii) + Mid(cboAccount.Text, cboAccount.SelStart + cboAccount.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboAccount.ListCount - 1
 
        If NewText = LCase(Left(cboAccount.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboAccount.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboAccount.Text = ValidValue
                cboAccount.SelStart = 0
                cboAccount.SelLength = Len(ValidValue)
            End If
        End If
End Sub
Private Sub CmdExit_Click()
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
    If lstChequeListing.ListItems.Count > 0 Then
        If MsgBox("Do you want to print now?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
    
           ' medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
            RecFound = False
            txtPrintingStatus = "Printing In Progress!"
            cmdPrint.Enabled = False
            cmdExit.Enabled = False
            If chqMode = "MedixPayMBM" Then
                Call ChequePrintingMBM
            ElseIf chqMode = "MedixPayHOS" Then
                Call ChequePrintingHOS
            End If
        End If
    Else
        MsgBox "Please select cheque to print", vbOKOnly + vbCritical
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub ChequePrintingMBM()
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
Dim strsqlUpd As String
Dim RecordFound As Boolean
    RecordFound = False
    strsql = "select PVChequeNo, PVChequeDate,PVChequeAmt, PayCode, " & _
    " PVAmtWording,PVAmtWordingTwo, CLMPayto, PVPayeeACNo from VIEW_Payment2MBM_Chq where (PVChequeNo >= '" & Trim(txtChequeNoFr) & "' And  PVChequeNo <= '" & Trim(txtChequeNoTo) & "') And PymtType = '" & Trim(lblPymtType) & "' order by PVChequeNo"
    chq.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
    Do While Not chq.EOF
        RecordFound = True
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            'If objExcel.Visible = False Then
                objExcel.Visible = False
            'End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\PaymentCheque.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
            
        On Error GoTo HANDLE_ERROR
            
            If Trim(chq!PAYCode) = "IB" Then
                objWorksheet.Cells(2, "A") = chq!CLMPayto
                objWorksheet.Cells(3, "A") = chq!PVPayeeACNo
            Else
                objWorksheet.Cells(3, "A") = chq!CLMPayto
            End If
            objWorksheet.Cells(5, "A") = chq!PVAmtWording
            objWorksheet.Cells(7, "A") = chq!PVAmtWordingTwo
            objWorksheet.Cells(5, "E") = chq!PVChequeAmt
            tmpDD = Format(chq!PVChequeDate, "DD")
            tmpMM = Format(chq!PVChequeDate, "MM")
            tmpYY = Format(chq!PVChequeDate, "YY")
            objWorksheet.Cells(1, "H") = Mid(tmpDD, 1, 1)
            objWorksheet.Cells(1, "I") = Mid(tmpDD, 2, 1)
            objWorksheet.Cells(1, "J") = Mid(tmpMM, 1, 1)
            objWorksheet.Cells(1, "K") = Mid(tmpMM, 2, 1)
            objWorksheet.Cells(1, "L") = Mid(tmpYY, 1, 1)
            objWorksheet.Cells(1, "M") = Mid(tmpYY, 2, 1)
    
            objWorkbook.PrintOut
            'objWorkbook.SaveAs (LocCardPath & "PaymentVoucher\" & "tmpPaycheque" & ".xlt")
            objWorkbook.SaveAs ("C:\Template\PaymentVoucher\tmpPaycheque.xlt")
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
            Set objExcel = Nothing
            
        chq.MoveNext
    Loop
    chq.Close
    Set chq = Nothing
    
    If RecordFound = True Then
        strsqlUpd = "Update Payment2MBM set PVChequeStatus = 'Y' Where (PVChequeNo >= '" & Trim(txtChequeNoFr) & "' And  PVChequeNo <= '" & Trim(txtChequeNoTo) & "') And Substring(PVNo,1,2) = '" & Trim(lblPymtType) & "'"
        medixmasterconnPayment.Execute strsqlUpd
    End If

ExitFunction:
          '  medixmasterconnPayment.Close
           ' Set medixmasterconnPayment = Nothing
            txtPrintingStatus = "Printing Not In Progress!"
            cmdPrint.Enabled = False
            cmdExit.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
           ' medixmasterconnPayment.Close
           ' Set medixmasterconnPayment = Nothing
            objExcel.Quit
            GoTo ExitFunction
            txtPrintingStatus = "Printing Not In Progress!"
            cmdPrint.Enabled = True
            cmdExit.Enabled = True

End Sub


Private Sub ChequePrintingHOS()
'On Error Resume Next
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
Dim strsqlUpd As String

    strsql = "select PVHOSChequeNo, PVHOSChequeDate, PVHOSChequeAmt, " & _
    " PVAmtWording,PVAmtWordingTwo,PVPayee from VIEW_Payment2HOS_Chq where (PVHOSChequeNo >= '" & Trim(txtChequeNoFr) & "' And  PVHOSChequeNo <= '" & Trim(txtChequeNoTo) & "') And AccountType = '" & Trim(lblPymtType) & "' order by PVHOSChequeNo"
    chq.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
    Do While Not chq.EOF
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
                objExcel.Visible = False
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\PaymentCheque.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
            
        On Error GoTo HANDLE_ERROR
    
            objWorksheet.Cells(3, "A") = chq!PVPayee
            objWorksheet.Cells(5, "A") = chq!PVAmtWording
            objWorksheet.Cells(7, "A") = chq!PVAmtWordingTwo
            objWorksheet.Cells(5, "E") = chq!PVHOSChequeAmt
            tmpDD = Format(chq!PVHOSChequeDate, "DD")
            tmpMM = Format(chq!PVHOSChequeDate, "MM")
            tmpYY = Format(chq!PVHOSChequeDate, "YY")
            objWorksheet.Cells(1, "H") = Mid(tmpDD, 1, 1)
            objWorksheet.Cells(1, "I") = Mid(tmpDD, 2, 1)
            objWorksheet.Cells(1, "J") = Mid(tmpMM, 1, 1)
            objWorksheet.Cells(1, "K") = Mid(tmpMM, 2, 1)
            objWorksheet.Cells(1, "L") = Mid(tmpYY, 1, 1)
            objWorksheet.Cells(1, "M") = Mid(tmpYY, 2, 1)

            objWorkbook.PrintOut
            objWorkbook.SaveAs (LocCardPath & "PaymentVoucher\" & "tmpPaycheque" & ".xlt")
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            objExcel.Quit
            Set objExcel = Nothing

        chq.MoveNext
    Loop
    chq.Close
    Set chq = Nothing
    
    strsqlUpd = "Update Payment2HOS set PVChequeStatus = 'Y' where (PVHOSChequeNo >= '" & Trim(txtChequeNoFr) & "' And  PVHOSChequeNo <= '" & Trim(txtChequeNoTo) & "') And SUBSTRING(PVHOSNo, 1, 2) = '" & Trim(lblPymtType) & "'"
    medixmasterconnPayment.Execute strsqlUpd

ExitFunction:
           ' medixmasterconnPayment.Close
           ' Set medixmasterconnPayment = Nothing
            txtPrintingStatus = "Printing Not In Progress!"
            cmdPrint.Enabled = False
            cmdExit.Enabled = True
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            'medixmasterconnPayment.Close
            'Set medixmasterconnPayment = Nothing
            txtPrintingStatus = "Printing Not In Progress!"
            cmdPrint.Enabled = True
            cmdExit.Enabled = True
End Sub

Private Sub cmdSearch_Click()
Dim RecSel As Double
    RecSel = 0
    lblTotalchqSel = 0
    If txtChequeNoFr = "" Then
        MsgBox "Please Enter Cheque No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If txtChequeNoTo = "" Then
        MsgBox "Please Enter Cheque No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If optSihat.Value = False And optNS.Value = False And optNSMAA.Value = False And optMX = False Then
        MsgBox "Please select payment type", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    RecSel = (CDbl(Trim(Mid(txtChequeNoTo, 4, 8))) - CDbl(Trim(Mid(txtChequeNoFr, 4, 8)))) + 1
    If RecSel <= 0 Then
        MsgBox "Please ensure cheque range entered correctly", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        
    lblTotalchqSel = RecSel
    lstChequeListing.ListItems.Clear
    'medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

    If chqMode = "MedixPayMBM" Then
        Call ChequeMBMListing
    ElseIf chqMode = "MedixPayHOS" Then
        Call ChequeHOSListing
    End If
   ' medixmasterconnPayment.Close
   ' Set medixmasterconnPayment = Nothing
    
End Sub
Private Sub ChequeHOSListing()

Dim strsql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim Record As ListItem
Dim RecordFound As Boolean
    RecordFound = False
    RecCnt = 0
    lblTotalChqList = 0
    cmdPrint.Enabled = False
    cmdExit.Enabled = False
    
    strsql = ""
    strsql = "select PVHOSChequeNo, PVHOSChequeDate, PVHOSChequeAmt, " & _
    " PVAmtWording,PVAmtWordingTwo,PVPayee,ChequeNo,AccountType from VIEW_Payment2HOS_Chq where PVHOSChequeNo >= '" & Trim(txtChequeNoFr) & "' And  PVHOSChequeNo <= '" & Trim(txtChequeNoTo) & "' And AccountType = '" & Trim(lblPymtType) & "' order by PVHOSChequeNo"
    chq.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
    Do While Not chq.EOF
        RecordFound = True
        Set Record = lstChequeListing.ListItems.Add(, , chq!PVHOSChequeNo)
        
        If Len(chq!PVHOSChequeDate) Then
            Record.SubItems(1) = (chq!PVHOSChequeDate)
        End If
        If Len(chq!PVHOSChequeAmt) Then
            Record.SubItems(2) = (chq!PVHOSChequeAmt)
        End If
        If Len(chq!ChequeNo) Then
            Record.SubItems(3) = (chq!ChequeNo)
        End If
        If Len(chq!AccountType) Then
            Record.SubItems(4) = (chq!AccountType)
        End If
        RecCnt = RecCnt + 1
        lblTotalChqList = RecCnt
        chq.MoveNext
    Loop
    chq.Close
    Set chq = Nothing
    
    cmdExit.Enabled = True
    
    If RecordFound = True Then
        If CDbl(lblTotalChqList) = CDbl(lblTotalchqSel) Then
            cmdPrint.Enabled = True
        Else
            cmdPrint.Enabled = False
            MsgBox "Cheque range not in series", vbOKOnly + vbCritical
        End If
    Else
        MsgBox "No record found", vbOKOnly + vbCritical
    End If
End Sub

Private Sub ChequeMBMListing()
Dim i As Integer
Dim strsql As String
Dim chq As New ADODB.Recordset
Dim RecFound As Boolean
Dim RecCnt As Integer
Dim tmpDD As String
Dim tmpMM As String
Dim tmpYY As String
Dim Record As ListItem
Dim RecordFound As Boolean
    
    RecordFound = False
    strsql = "select PVChequeNo, PVChequeDate,PVChequeAmt, " & _
    " PVAmtWording,PVAmtWordingTwo, CLMPayto,PymtType,ChequeNo,PVChequeStatus from VIEW_Payment2MBM_Chq where (PVChequeNo >= '" & Trim(txtChequeNoFr) & "' And  PVChequeNo <= '" & Trim(txtChequeNoTo) & "') And PymtType = '" & Trim(lblPymtType) & "' order by PVChequeNo"
    chq.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
    Do While Not chq.EOF
        RecordFound = True
        Set Record = lstChequeListing.ListItems.Add(, , chq!PVChequeNo)
        
        If Len(chq!PVChequeDate) Then
            Record.SubItems(1) = (chq!PVChequeDate)
        End If
        If Len(chq!PVChequeAmt) Then
            Record.SubItems(2) = (chq!PVChequeAmt)
        End If
        If Len(chq!ChequeNo) Then
            Record.SubItems(3) = (chq!ChequeNo)
        End If
        If Len(chq!PymtType) Then
            Record.SubItems(4) = (chq!PymtType)
        End If
        RecCnt = RecCnt + 1
        lblTotalChqList = RecCnt
        chq.MoveNext
    Loop
    chq.Close
    Set chq = Nothing
    
    cmdExit.Enabled = True
    
    If RecordFound = True Then
        If CDbl(lblTotalChqList) = CDbl(lblTotalchqSel) Then
            cmdPrint.Enabled = True
        Else
            cmdPrint.Enabled = False
            MsgBox "Cheque range not in series", vbOKOnly + vbCritical
        End If
    Else
        MsgBox "No record found", vbOKOnly + vbCritical
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

Private Sub ComboListing()
Dim Acc As New ADODB.Recordset
Dim strsql As String

 '   medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    
    strsql = "Select ChequeNo from ChequeType "
    Acc.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    Do Until Acc.EOF
        cboAccount.AddItem Trim(Acc!ChequeNo)
        Acc.MoveNext
    Loop
        Acc.Close
        Set Acc = Nothing
    
 '   medixmasterconnPayment.Close
 '   Set medixmasterconnPayment = Nothing
End Sub

Private Sub optMX_Click()
    lblPymtType = ""
    lblPymtType = Mid(optMX.Caption, 1, 2)
    cmdPrint.Enabled = False
End Sub

Private Sub optNS_Click()
    lblPymtType = ""
    lblPymtType = Mid(optNS.Caption, 1, 2)
    cmdPrint.Enabled = False
End Sub

Private Sub optNSMAA_Click()
    lblPymtType = ""
    lblPymtType = Mid(optNSMAA.Caption, 1, 2)
    cmdPrint.Enabled = False
End Sub

Private Sub optSihat_Click()
    lblPymtType = ""
    lblPymtType = Mid(optSihat.Caption, 1, 2)
    cmdPrint.Enabled = False
End Sub

Private Sub txtChequeNoFr_Click()
    cmdPrint.Enabled = False
End Sub

Private Sub txtChequeNoTo_Click()
    cmdPrint.Enabled = False
End Sub
