VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMCORefundListing2 
   Caption         =   "MCO Refund Listing"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameMCO 
      Caption         =   "Enter Criteria For Searching"
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   11055
      Begin VB.ComboBox CboDateSelection 
         Height          =   315
         ItemData        =   "frmMCORefundListing2.frx":0000
         Left            =   1680
         List            =   "frmMCORefundListing2.frx":000A
         TabIndex        =   1
         Top             =   285
         Width           =   4215
      End
      Begin VB.ComboBox cboPayor 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   570
         Width           =   4215
      End
      Begin MSMask.MaskEdBox TxtDate1 
         Height          =   300
         Left            =   7320
         TabIndex        =   2
         Top             =   285
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDate2 
         Height          =   300
         Left            =   9000
         TabIndex        =   3
         Top             =   285
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboPrintStatus 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   570
         Width           =   1150
      End
      Begin VB.Label Label6 
         Caption         =   "Print Status"
         Height          =   255
         Left            =   6000
         TabIndex        =   28
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   8640
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Date"
         Height          =   255
         Left            =   6000
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Date Selection"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbl20 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   675
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5880
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11055
      Begin MSComctlLib.ListView lvwMCO 
         Height          =   5490
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   9684
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   22
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Member No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Cover ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Policy No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Member Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "I/C No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Bordx Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Eff. Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Exp. Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Days Insured"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Cancellation Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Endt. Bordx Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Days Utilised"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "Month Utilised"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "MCO fees Charged"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "MCO fees Utilised"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Text            =   "MCO fees Refund"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Text            =   "MCO fees Refund (Min.)"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Invoice No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   18
            Text            =   "Invoice Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Remarks"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "Credit Note"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Credit Date"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Width           =   11055
      Begin VB.CommandButton cmdPrintCRNote 
         Caption         =   "&Print CR Note"
         Height          =   300
         Left            =   8520
         TabIndex        =   26
         Top             =   270
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   300
         Left            =   7200
         TabIndex        =   7
         Top             =   270
         Width           =   720
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   300
         Left            =   6480
         TabIndex        =   6
         Top             =   270
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   7920
         TabIndex        =   8
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9600
         TabIndex        =   9
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10200
         TabIndex        =   10
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5400
         TabIndex        =   25
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Total MCO Refund (Min.)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   3480
         TabIndex        =   24
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label Label4 
         Caption         =   "Total MCO Refund"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   3480
         TabIndex        =   23
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label txtMCOTotalSupplementary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5400
         TabIndex        =   22
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label201 
         Caption         =   "Total Supp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   960
      End
      Begin VB.Label txtMCOTotalPrincipal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label200 
         Caption         =   "Total Principal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Refund Fee Listing"
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
      Height          =   225
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmMCORefundListing2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim MBMNo As String
Dim CurrentS As String
Private m_blnDirection(1 To 22) As Boolean
Dim printCRNote As Boolean
Dim MBMRefundFees As Variant
Dim StrAppend As String
Dim strAppend2 As String
Dim MBMStatus As Boolean

Private Sub CboPayor_Change()
    If InStr(cboPayor.Text, "null") Then
       cboPayor.Enabled = False
    End If
End Sub

Private Sub cboPayor_Click()
Dim PAYGRP As New ADODB.Recordset
Dim StrSql As String

 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    StrSql = " Select MBMCancelStatus FROM PAY Where PAYCode = '" & Mid(cboPayor, 1, 2) & "'"
    
    PAYGRP.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    If PAYGRP.recordCount Then
        If PAYGRP!MBMCancelStatus = "Y" Then
            MBMStatus = True
        Else
            MBMStatus = False
        End If
    End If
    PAYGRP.Close
    Set PAYGRP = Nothing
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
End Sub

Private Sub CboPayor_KeyDown(KeyCode As Integer, Shift As Integer)
Dim PAYGRP As New ADODB.Recordset
Dim StrSql As String

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    StrSql = " Select MBMCancelStatus FROM PAY Where PAYCode = '" & Mid(cboPayor, 1, 2) & "'"
    
    PAYGRP.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    If PAYGRP.recordCount Then
        If PAYGRP!MBMCancelStatus = "Y" Then
            MBMStatus = True
        Else
            MBMStatus = False
        End If
    End If
    PAYGRP.Close
    Set PAYGRP = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayor.Text, cboPayor.SelStart) + Chr(KeyAscii) + Mid(cboPayor.Text, cboPayor.SelStart + cboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayor.ListCount - 1
 
        If NewText = LCase(Left(cboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayor.Text = ValidValue
                cboPayor.SelStart = 0
                cboPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboPayor_KeyUp(KeyCode As Integer, Shift As Integer)
Dim PAYGRP As New ADODB.Recordset
Dim StrSql As String

  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    StrSql = " Select MBMCancelStatus FROM PAY Where PAYCode = '" & Mid(cboPayor, 1, 2) & "'"
    
    PAYGRP.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    If PAYGRP.recordCount Then
        If PAYGRP!MBMCancelStatus = "Y" Then
            MBMStatus = True
        Else
            MBMStatus = False
        End If
    End If
    PAYGRP.Close
    Set PAYGRP = Nothing
    
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

Private Sub CboPrintStatus_Change()
    If InStr(CboPrintStatus.Text, "null") Then
       CboPrintStatus.Enabled = False
    End If
End Sub

Private Sub cboPrintStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(CboPrintStatus.Text, CboPrintStatus.SelStart) + Chr(KeyAscii) + Mid(CboPrintStatus.Text, CboPrintStatus.SelStart + CboPrintStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPrintStatus.ListCount - 1
 
        If NewText = LCase(Left(CboPrintStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPrintStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPrintStatus.Text = ValidValue
                CboPrintStatus.SelStart = 0
                CboPrintStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CmdClear_Click()
    
    MCOFee = 0
    TotalAmt = 0
    TotalAmt2 = 0
    Current = 0
    CurrentS = 0
    MBMNo = ""
    MBMNumber = ""
    lvwMCO.listitems.Clear
    txtMCOTotalPrincipal = 0
    txtMCOTotalSupplementary = 0
    CboPrintStatus = "N - No"
    Label3 = 0
    Label5 = 0
    TxtDate1.Text = "__/__/____"
    TxtDate2.Text = "__/__/____"
    cboPayor.Text = ""
    CboDateSelection.Text = ""
End Sub

Private Sub CmdPrint_Click()
    'Call ExportToExcelMBMMCORefund(lvwMCO)
    Screen.MousePointer = vbHourglass
    
    If lvwMCO.listitems.Count <> 0 Then
        Call ExportListViewtoExcel(lvwMCO)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = Default

End Sub

Private Sub cmdPrintCRNote_Click()
Dim sql As String
Dim PAY As New ADODB.Recordset
Dim Sql2 As String
Dim rst2 As New ADODB.Recordset
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim i As Integer
Dim ii As Integer
Dim CRNote As String
Dim MBMNumber As String
Dim CRNoteDate As Date
Dim PAYCompanyName As String
Dim PAYAddress1 As String
Dim PAYAddress2 As String
Dim PAYAddress3 As String
Dim PAYMAINContact As String
Dim PAYTELno As String

    Screen.MousePointer = vbHourglass
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If lvwMCO.listitems.Count = 0 Then
        MsgBox "Please Select Records!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf CboDateSelection.Text = "" Then
        MsgBox "Please Select Date Selection!", vbOKOnly + vbCritical, DialogBoxTitle
        CboDateSelection.SetFocus
    ElseIf cboPayor.Text = "" Then
        MsgBox "Please Select Payor!", vbOKOnly + vbCritical, DialogBoxTitle
        cboPayor.SetFocus
    Else
    
        sql = " Select PAYCode,  PAYCompanyName, PAYAddress1, " & _
              " PAYAddress2, PAYAddress3, PAYMAINContact, PAYTELno " & _
              " FROM PAY where PAYCode = '" & Mid(cboPayor, 1, 2) & "'"
        PAY.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If PAY.recordCount Then
            PAYCompanyName = PAY!PAYCompanyName
            PAYAddress1 = PAY!PAYAddress1
            PAYAddress2 = PAY!PAYAddress2
            PAYAddress3 = PAY!PAYAddress3
            PAYTELno = PAY!PAYTELno
            PAYMAINContact = PAY!PAYMAINContact
        End If
        PAY.Close
        Set PAY = Nothing
    
    If Mid(CboPrintStatus, 1, 1) = "N" Then
        CRNote = GenerateMCOCreditNo
    Else
        Sql2 = " Select MBMCRNote, MBMCRNoteDate " & _
               " FROM VIEW_MBM_Refund_Search Where " & _
               "(MBMStatus = 'C' AND Checkday <= 30 OR MBMStatus = 'P')"
            
        If Len(Trim(StrAppend)) Then
            Sql2 = Sql2 + StrAppend
        End If
       
        rst2.Open Sql2, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        If rst2.EOF = False Then
            CRNote = rst2!MBMCRNote
            CRNoteDate = rst2!MBMCRNoteDate
        End If
        rst2.Close
        Set rst2 = Nothing
    End If
    
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
        End If
        
        'If ERR.Number > 0 Then
        '    MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        '    GoTo ExitFunction
        'End If
        
        On Error GoTo HANDLE_ERROR
        ' Don't allow user to affect workbook
        objExcel.Interactive = False
    
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
        
            objExcel.WindowState = xlMaximized
    
        'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MCO\MCOCreditNote.xlt")
        Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Tracking\MCOCreditNote.xlt")
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
            
            objWorksheet.Cells(8, "C") = PAYCompanyName
            objWorksheet.Cells(9, "C") = PAYAddress1
            objWorksheet.Cells(10, "C") = PAYAddress2
            objWorksheet.Cells(11, "C") = PAYAddress3
            objWorksheet.Cells(12, "C") = PAYTELno
            objWorksheet.Cells(12, "E") = PAYMAINContact
            objWorksheet.Cells(8, "J") = CRNote
            If Mid(CboPrintStatus, 1, 1) = "N" Then
                objWorksheet.Cells(9, "J") = Date
            Else
                objWorksheet.Cells(9, "J") = CRNoteDate
            End If
            objWorksheet.Cells(17, "H") = MBMRefundFees
            objWorksheet.Cells(17, "J") = MBMRefundFees
            
            If Mid(CboPrintStatus, 1, 1) = "N" Then
                
                For ii = 1 To lvwMCO.listitems.Count
                    
                    MBMNumber = Trim(lvwMCO.listitems(ii).Text)
                    Call UpdateCreditNote(CRNote, MBMNumber)
                        
                Next ii
                
            End If
            
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            Screen.MousePointer = vbDefault
            cmdPrintCRNote.Enabled = True
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
            
           ' medixmasterconnMaint.Close
           ' Set medixmasterconnMaint = Nothing
            
            Exit Sub
ExitFunction:
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            objExcel.Quit
            Screen.MousePointer = vbDefault
            cmdPrintCRNote.Enabled = True
         '   medixmasterconn.Close
        '    Set medixmasterconn = Nothing
            
         '   medixmasterconnMaint.Close
         '   Set medixmasterconnMaint = Nothing
            
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            GoTo ExitFunction
    'End If
    
    End If

    Screen.MousePointer = vbDefault
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    cmdPrint.Enabled = True
    cmdPrintCRNote.Enabled = True
    cmdSearch.Enabled = True
    
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
    'medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub cmdSearch_Click()
    
    cmdStop.Enabled = True
    cmdClear.Enabled = False
    cmdExit.Enabled = False
    cmdPrint.Enabled = False
    cmdPrintCRNote.Enabled = False
    cmdSearch.Enabled = False
    txtMCOTotalPrincipal = 0
    txtMCOTotalSupplementary = 0
    CurrentS = "0"
    printCRNote = False
    Screen.MousePointer = vbHourglass

    If CboDateSelection.Text = "" Then
        MsgBox "Please Select Date Selection!", vbOKOnly + vbCritical, DialogBoxTitle
        CboDateSelection.SetFocus
    ElseIf cboPayor.Text = "" Then
        MsgBox "Please Select Payor!", vbOKOnly + vbCritical, DialogBoxTitle
        cboPayor.SetFocus
    Else
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call SearchRecord
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If

    Screen.MousePointer = vbDefault
    
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    cmdPrint.Enabled = True
    cmdPrintCRNote.Enabled = True
    cmdSearch.Enabled = True
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
    cmdClear.Enabled = True
    cmdPrint.Enabled = True
    cmdPrintCRNote.Enabled = True
    cmdSearch.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CboDateSelection_Click()
    
    Select Case CboDateSelection.ListIndex
    Case 0
        Label8.Caption = "Endt. Eff Date"
        lvwMCO.listitems.Clear
        txtMCOTotalSupplementary = 0
        txtMCOTotalPrincipal = 0
    Case 1
        Label8.Caption = "Endt. Bordx Date"
        lvwMCO.listitems.Clear
        txtMCOTotalSupplementary = 0
        txtMCOTotalPrincipal = 0
    End Select

End Sub

Sub ComboListing()
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
            cboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    
    PAY.Close
    Set PAY = Nothing
    
    CboPrintStatus.AddItem "N - No"
    CboPrintStatus.AddItem "Y - Yes"
    
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
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    CboPrintStatus = "N - No"
End Sub

Private Sub SearchRecord()
     
    StrAppend = ""
    strAppend2 = ""
     
    
    If Len(Mid(cboPayor, 1, 2)) Then
        StrAppend = StrAppend & " AND PAYCode= '" & Mid(cboPayor, 1, 2) & "'"
    End If
    
    If CboDateSelection.ListIndex = 0 Then
    
        If IsDate(TxtDate1) = True Then
            StrAppend = StrAppend + " And MBMENDtEffDate >= '" & Format(TxtDate1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(TxtDate2) = True Then
            StrAppend = StrAppend + " And MBMENDtEffDate <= '" & Format(TxtDate2, "mm/dd/yyyy") & "'"
        End If
    
    ElseIf CboDateSelection.ListIndex = 1 Then
    
        If IsDate(TxtDate1) = True Then
            StrAppend = StrAppend + " And MBMENDtBordxDate >= '" & Format(TxtDate1, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(TxtDate2) = True Then
            StrAppend = StrAppend + " And MBMENDtBordxDate <= '" & Format(TxtDate2, "mm/dd/yyyy") & "'"
        End If
    
    End If
    
    If MBMStatus = True Then
        Call MCORefundListing(StrAppend)
    Else
        Call MCORefundListing2(StrAppend)
    End If
    
End Sub


Private Sub MCORefundListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql1 As String
Dim TotalAmt As String
Dim Current As String
Dim TotalAmt2 As Variant
Dim MCOFee As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String


    MCOFee = 0
    TotalAmt = 0
    TotalAmt2 = 0
    Current = 0
    CurrentS = 0
    MBMNo = ""
    MBMRefundFees = 0
    lvwMCO.listitems.Clear
    
    
    strAppendCase = "1"
    
    If Mid(CboPrintStatus, 1, 1) = "N" Then
        strAppend2 = strAppend2 & " AND (MBMCRNote IS NULL OR MBMCRNote = '') "
    Else
        strAppend2 = strAppend2 & " AND MBMCRNote <> '' "
    End If
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + strAppend2 + sql1
    End If
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "MBMRefundListing"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

        
    'Sql = " Select MBMNumber, MBMPolicyNo, MBMName, MBMMCOUtilised, " & _
          " MBMStatus, MBMCOVID, MBMENDtEffDate, MBMENDtBordxDate," & _
          " MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMMCORefund, " & _
          " MBMBordxDate, DayInsured, MBMIcBcPp, DayUtilised, INSCode, MBMCRNote, MBMCRNoteDate, " & _
          " MBMInvoiceNo, MBMInvoiceDate, TotalCase, CheckDay, MonthUtilised From VIEW_MBM_Refund_Search Where MBMStatus = 'C' AND Checkday <= 30 "
    
    'Sql1 = "ORDER BY MBMNumber"
    
    'If Mid(CboPrintStatus, 1, 1) = "N" Then
        'strAppend2 = strAppend2 & " AND (MBMCRNote IS NULL OR MBMCRNote = '') "
    'Else
        'strAppend2 = strAppend2 & " AND MBMCRNote <> '' "
    'End If
    
    'If Len(Trim(StrAppend)) Then
        'Sql = Sql + StrAppend + strAppend2 + Sql1
    'End If

    'rst2.Open Sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
    'If rst2.EOF = True Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'cmdClear.Enabled = True
        'cmdExit.Enabled = True
        'cmdPrint.Enabled = True
        'cmdPrintCRNote.Enabled = True
    'Else
         
    Do
        Do Until rst2.EOF
         
            With rst2
                Set Record = lvwMCO.listitems.Add(, , !MBMNumber)
                    MBMNo = Trim(!MBMNumber)
                
                    If Len(!MBMCOVID) Then
                        Record.SubItems(1) = !MBMCOVID
                    End If
                      
                    If Len(!MBMPolicyNo) <> "" Then
                        Record.SubItems(2) = Trim(!MBMPolicyNo)
                    End If
                    
                    If !MBMStatus = "C" Then
                        Record.SubItems(3) = Trim(!MBMName) & " (Cancelled)"
                    ElseIf !MBMStatus = "P" Then
                        Record.SubItems(3) = Trim(!MBMName) & " (Pending)"
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                        Record.SubItems(4) = Trim(!MBMIcBcPp)
                    End If
                      
                    If Len(Trim(!MBMBordxDate)) Then
                        Record.SubItems(5) = Format(!MBMBordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(6) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                      
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(7) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!DayInsured) Then
                        Record.SubItems(8) = !DayInsured
                    End If
                    
                    If Len(!MBMENDtEffDate) Then
                        Record.SubItems(9) = Format(!MBMENDtEffDate, "dd/mm/yyyy")
                    End If
                      
                    If Len(!MBMENDtBordxDate) Then
                        Record.SubItems(10) = Format(!MBMENDtBordxDate, "dd/mm/yyyy")
                    End If
                    
                    If !DayUtilised > 0 Then
                        Record.SubItems(11) = !DayUtilised
                    Else
                        Record.SubItems(11) = "0"
                    End If
                    Record.SubItems(12) = IIf(IsNumeric(!MonthUtilised), !MonthUtilised, 0)
                    'If !MonthUtilised Then
                    '    Record.SubItems(12) = !MonthUtilised
                    'Else
                    '    Record.SubItems(12) = "0"
                    'End If
                                        
                    If Len(!MBMMCOFee) Then
                        Record.SubItems(13) = Format(!MBMMCOFee, "#,##0.00")
                    Else
                        Record.SubItems(13) = "0.00"
                    End If
                    
                    If Len(!MBMMCOUtilised) Then
                        Record.SubItems(14) = Format(!MBMMCOUtilised, "#,##0.00")
                    Else
                        Record.SubItems(14) = "0.00"
                    End If
                    
                    If Len(!MBMMCORefund) Then
                        Record.SubItems(15) = Format(!MBMMCORefund, "#,##0.00")
                        MBMRefundFees = CDbl(MBMRefundFees) + CDbl(!MBMMCORefund)
                    Else
                        Record.SubItems(15) = "0.00"
                    End If
                                        
                    If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                    
                        If Len(!MBMMCOUtilised) Then
                        If Format(!MBMMCOUtilised, "#,##0.00") <= 50# Then
                            MCOFee = (!MBMMCOFee - 50)
                            Record.SubItems(16) = Format(MCOFee, "#,##0.00")
                        Else
                            Record.SubItems(16) = Format(!MBMMCOUtilised, "#,##0.00")
                        End If
                    Else
                        Record.SubItems(16) = "0.00"
                    End If
                    
                    Else
                    
                        If Len(!MBMMCOUtilised) Then
                            If Format(!MBMMCOUtilised, "#,##0.00") <= 25# Then
                                MCOFee = (!MBMMCOFee - 25)
                                Record.SubItems(16) = Format(MCOFee, "#,##0.00")
                            Else
                                Record.SubItems(16) = Format(!MBMMCOUtilised, "#,##0.00")
                            End If
                        Else
                            Record.SubItems(16) = "0.00"
                        End If
                    
                    End If
                    
                    If Len(!MBMInvoiceNo) Then
                        Record.SubItems(17) = Trim(!MBMInvoiceNo)
                    End If
                    
                    If Len(!MBMInvoiceDate) Then
                        Record.SubItems(18) = Format(!MBMInvoiceDate, "dd/mm/yyyy")
                    End If
                    
                    If !TotalCase > 0 Then
                        Record.SubItems(19) = "This member had" & " " & !TotalCase & " " & "Case(s)"
                    End If
                    
                    If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                        Call MCORefundListingSupp(StrAppend)
                    End If
                    
                    If Len(!MBMCRNote) Then
                        Record.SubItems(20) = Format(!MBMCRNote)
                    End If
                    
                    If Len(!MBMCRNoteDate) Then
                        Record.SubItems(21) = Format(!MBMCRNoteDate, "dd/mm/yyyy")
                    End If
                    
                    Current = Current + 1
                    txtMCOTotalPrincipal = Current
                    If Len(rst2!MBMMCORefund) Then
                        TotalAmt = TotalAmt + rst2!MBMMCORefund
                    End If
                    TotalAmt2 = TotalAmt2 + CDbl(Record.SubItems(15))
                    
                    Label3 = Format(TotalAmt, "#,##0.00")
                    Label5 = Format(TotalAmt2, "#,##0.00")
                    
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    'End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    'cmdClear.Enabled = True
    'cmdExit.Enabled = True
    'cmdPrint.Enabled = True
    'cmdPrintCRNote.Enabled = True
End Sub

Private Sub MCORefundListing2(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql1 As String
Dim TotalAmt As String
Dim Current As String
Dim TotalAmt2 As Variant
Dim MCOFee As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

    MCOFee = 0
    TotalAmt = 0
    TotalAmt2 = 0
    Current = 0
    CurrentS = 0
    MBMNo = ""
    MBMRefundFees = 0
    lvwMCO.listitems.Clear
    
    
    strAppendCase = "2"
    
    If Mid(CboPrintStatus, 1, 1) = "N" Then
        strAppend2 = strAppend2 & " AND (MBMCRNote IS NULL OR MBMCRNote = '') "
    Else
        strAppend2 = strAppend2 & " AND MBMCRNote <> '' "
    End If
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + strAppend2 + sql1
    End If
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "MBMRefundListing"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute

         
    Do
        Do Until rst2.EOF
         
            With rst2
                Set Record = lvwMCO.listitems.Add(, , !MBMNumber)
                    MBMNo = Trim(!MBMNumber)
                
                    If Len(!MBMCOVID) Then
                        Record.SubItems(1) = !MBMCOVID
                    End If
                      
                    If Len(!MBMPolicyNo) <> "" Then
                        Record.SubItems(2) = Trim(!MBMPolicyNo)
                    End If
                    
                    If !MBMStatus = "C" Then
                        Record.SubItems(3) = Trim(!MBMName) & " (Cancelled)"
                    ElseIf !MBMStatus = "P" Then
                        Record.SubItems(3) = Trim(!MBMName) & " (Pending)"
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                        Record.SubItems(4) = Trim(!MBMIcBcPp)
                    End If
                      
                    If Len(Trim(!MBMBordxDate)) Then
                        Record.SubItems(5) = Format(!MBMBordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(6) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                      
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(7) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!DayInsured) Then
                        Record.SubItems(8) = !DayInsured
                    End If
                    
                    If Len(!MBMENDtEffDate) Then
                        Record.SubItems(9) = Format(!MBMENDtEffDate, "dd/mm/yyyy")
                    End If
                      
                    If Len(!MBMENDtBordxDate) Then
                        Record.SubItems(10) = Format(!MBMENDtBordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!DayUtilised) Then
                        Record.SubItems(11) = !DayUtilised
                        'Record.SubItems(12) = Month(!DayUtilised)
                    Else
                        Record.SubItems(11) = "0"
                        'Record.SubItems(12) = "0"
                    End If
                    Record.SubItems(12) = IIf(IsNumeric(!MonthUtilised), !MonthUtilised, 0)
                
                    If Len(!MBMMCOFee) Then
                        Record.SubItems(13) = Format(!MBMMCOFee, "#,##0.00")
                    Else
                        Record.SubItems(13) = "0.00"
                    End If
                    
                    If Len(!MBMMCOUtilised) Then
                        Record.SubItems(14) = Format(!MBMMCOUtilised, "#,##0.00")
                    Else
                        Record.SubItems(14) = "0.00"
                    End If
                    
                    If Len(!MBMMCORefund) Then
                        Record.SubItems(15) = Format(!MBMMCORefund, "#,##0.00")
                        MBMRefundFees = CDbl(MBMRefundFees) + CDbl(!MBMMCORefund)
                    Else
                        Record.SubItems(15) = "0.00"
                    End If
                                        
                    If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                    
                        If Len(!MBMMCOUtilised) Then
                        If Format(!MBMMCOUtilised, "#,##0.00") <= 50# Then
                            MCOFee = (!MBMMCOFee - 50)
                            Record.SubItems(16) = Format(MCOFee, "#,##0.00")
                        Else
                            Record.SubItems(16) = Format(!MBMMCOUtilised, "#,##0.00")
                        End If
                    Else
                        Record.SubItems(16) = "0.00"
                    End If
                    
                    Else
                    
                        If Len(!MBMMCOUtilised) Then
                            If Format(!MBMMCOUtilised, "#,##0.00") <= 25# Then
                                MCOFee = (!MBMMCOFee - 25)
                                Record.SubItems(16) = Format(MCOFee, "#,##0.00")
                            Else
                                Record.SubItems(16) = Format(!MBMMCOUtilised, "#,##0.00")
                            End If
                        Else
                            Record.SubItems(16) = "0.00"
                        End If
                    
                    End If
                    
                    If Len(!MBMInvoiceNo) Then
                        Record.SubItems(17) = Trim(!MBMInvoiceNo)
                    End If
                    
                    If Len(!MBMInvoiceDate) Then
                        Record.SubItems(18) = Format(!MBMInvoiceDate, "dd/mm/yyyy")
                    End If
                    
                    If !TotalCase > 0 Then
                        Record.SubItems(19) = "This member had" & " " & !TotalCase & " " & "Case(s)"
                    End If
                    
                    If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                        Call MCORefundListingSupp(StrAppend)
                    End If
                    
                    If Len(!MBMCRNote) Then
                        Record.SubItems(20) = Format(!MBMCRNote)
                    End If
                    
                    If Len(!MBMCRNoteDate) Then
                        Record.SubItems(21) = Format(!MBMCRNoteDate, "dd/mm/yyyy")
                    End If
                    
                    Current = Current + 1
                    txtMCOTotalPrincipal = Current
                    If Len(rst2!MBMMCORefund) Then
                        TotalAmt = TotalAmt + rst2!MBMMCORefund
                    End If
                    TotalAmt2 = TotalAmt2 + CDbl(Record.SubItems(15))
                    
                    Label3 = Format(TotalAmt, "#,##0.00")
                    Label5 = Format(TotalAmt2, "#,##0.00")
                    
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
    
End Sub

Private Sub MCORefundListingSupp(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strAppendCase As String

            
    strAppendCase = "3"
    
    strAppend2 = " AND MBMNumber = '" & Trim(MBMNo) & "'"
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend + strAppend2
    End If
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "MBMRefundListing"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, sql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    'Sql = " Select MBMNumber, MBMPolicyNo, MBMName, MBMInvoiceNo, MBMInvoiceDate, " & _
          " MBMStatus, MBMCOVID, MBMENDtEffDate, MBMENDtBordxDate," & _
          " MBMPayorEffDate, MBMPayorExpDate, TotalCase, MBMCRNote, MBMCRNoteDate, " & _
          " MBMBordxDate, DayInsured, MBMIcBcPp, DayUtilised, MonthUtilised " & _
          " From View_MBM_Supplementary_Refund Where MBMNumber = '" & Trim(MBMNo) & "'"
    
    'If Len(Trim(StrAppend)) Then
        'Sql = Sql + StrAppend
    'End If

    'rst2.Open Sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
         
    Do
        Do Until rst2.EOF
         
            With rst2
                Set Record = lvwMCO.listitems.Add(, , !MBMNumber)
                                    
                    If Len(!MBMCOVID) Then
                        Record.SubItems(1) = !MBMCOVID
                    End If
                      
                    If Len(!MBMPolicyNo) <> "" Then
                        Record.SubItems(2) = Trim(!MBMPolicyNo)
                    End If
                    
                    If !MBMStatus = "C" Then
                        Record.SubItems(3) = Trim(!MBMName) & " (Cancelled)"
                    ElseIf !MBMStatus = "P" Then
                        Record.SubItems(3) = Trim(!MBMName) & " (Pending)"
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                        Record.SubItems(4) = Trim(!MBMIcBcPp)
                    End If
                      
                    If Len(Trim(!MBMBordxDate)) Then
                        Record.SubItems(5) = Format(!MBMBordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(6) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                      
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(7) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!DayInsured) Then
                        Record.SubItems(8) = !DayInsured
                    End If
                    
                    If Len(!MBMENDtEffDate) Then
                        Record.SubItems(9) = Format(!MBMENDtEffDate, "dd/mm/yyyy")
                    End If
                      
                    If Len(!MBMENDtBordxDate) Then
                        Record.SubItems(10) = Format(!MBMENDtBordxDate, "dd/mm/yyyy")
                    End If
                    
                    If !DayUtilised > 0 Then
                        Record.SubItems(11) = !DayUtilised
                        'Record.SubItems(12) = Month(!DayUtilised)
                    Else
                        Record.SubItems(11) = "0"
                        'Record.SubItems(12) = "0"
                    End If
                    Record.SubItems(12) = IIf(IsNumeric(!MonthUtilised), !MonthUtilised, 0)

                    'If Len(!MonthUtilised) Then
                     '   Record.SubItems(12) = !MonthUtilised
                    'Else
                    '    Record.SubItems(12) = "0"
                    'End If
                    
                    If Len(!MBMMCOFee) Then
                        Record.SubItems(13) = Format(!MBMMCOFee, "#,##0.00")
                    End If
                    
                    If Len(!MBMMCOUtilised) Then
                        Record.SubItems(14) = Format(!MBMMCOUtilised, "#,##0.00")
                    End If
                    
                    If Len(!MBMMCORefund) Then
                        Record.SubItems(15) = Format(!MBMMCORefund, "#,##0.00")
                        MBMRefundFees = CDbl(MBMRefundFees) + CDbl(!MBMMCORefund)
                    End If
                    
                    If Len(!MBMMCOUtilised) Then
                        If Format(!MBMMCOUtilised, "#,##0.00") <= 25# Then
                            MCOFee = (!MBMMCOFee - 25)
                            Record.SubItems(16) = Format(MCOFee, "#,##0.00")
                        Else
                            Record.SubItems(16) = Format(!MBMMCOUtilised, "#,##0.00")
                        End If
                    Else
                        Record.SubItems(16) = "0.00"
                    End If
                    
                    'Record.SubItems(13) = "0.00"
                    'Record.SubItems(14) = "0.00"
                    'Record.SubItems(15) = "0.00"
                    'Record.SubItems(16) = "0.00"
                    
                    If Len(!MBMInvoiceNo) Then
                        Record.SubItems(17) = Trim(!MBMInvoiceNo)
                    End If
                    
                    If Len(!MBMInvoiceDate) Then
                        Record.SubItems(18) = Format(!MBMInvoiceDate, "dd/mm/yyyy")
                    End If
                    
                    If !TotalCase > 0 Then
                        Record.SubItems(19) = "This member had" & " " & !TotalCase & " " & "Case(s)"
                    End If
                    
                    If Len(!MBMCRNote) Then
                        Record.SubItems(20) = Format(!MBMCRNote)
                    End If
                    
                    If Len(!MBMCRNoteDate) Then
                        Record.SubItems(21) = Format(!MBMCRNoteDate, "dd/mm/yyyy")
                    End If
                                                    
                    CurrentS = CurrentS + 1
                    txtMCOTotalSupplementary = CurrentS
            
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    'End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing

End Sub

Private Sub lvwMCO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
    Case 3
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    Case 14
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    Case 15
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
    Case 16
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
    Case 17
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
    Case 18
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(18)
    Case 19
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(19)
    Case 20
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(20)
    Case 21
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(21)
    Case 22
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(22)
    End Select
End Sub

Private Sub TxtDate1_LostFocus()
    If IsDate(TxtDate1) Then
    Else
        If TxtDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDate2_LostFocus()
    If IsDate(TxtDate2) Then
    Else
        If TxtDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDate2.SetFocus
        End If
    End If
End Sub

Private Sub MCORefundCRNote(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim MCOFeeRef As Variant

    MCOFeeRef = 0
    
    sql = " Select MBMStatus, MBMCOVID, MBMENDtEffDate, MBMENDtBordxDate," & _
          " MBMPayorEffDate, MBMPayorExpDate, MBMMCOFee, MBMMCORefund, " & _
          " MBMInvoiceNo, MBMInvoiceDate, TotalCase From VIEW_MBM_MCORefund_Inv Where 0 = 0 "
          
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdClear.Enabled = True
        cmdExit.Enabled = True
        cmdPrint.Enabled = True
        cmdPrintCRNote.Enabled = True
    Else
         
        Do Until rst2.EOF
            With rst2
                    If Len(!MBMMCORefund) Then
                        Record.SubItems(14) = Format(!MBMMCORefund, "#,##0.00")
                    End If
                                        
                    If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                    
                        If Len(!MBMMCOUtilised) Then
                        If Format(!MBMMCOUtilised, "#,##0.00") <= 50# Then
                            MCOFee = (!MBMMCOFee - 50)
                            Record.SubItems(15) = Format(MCOFee, "#,##0.00")
                        Else
                            Record.SubItems(15) = Format(!MBMMCOUtilised, "#,##0.00")
                        End If
                    Else
                        Record.SubItems(15) = "0.00"
                    End If
                    
                    Else
                    
                        If Len(!MBMMCOUtilised) Then
                            If Format(!MBMMCOUtilised, "#,##0.00") <= 25# Then
                                MCOFee = (!MBMMCOFee - 25)
                                Record.SubItems(15) = Format(MCOFee, "#,##0.00")
                            Else
                                Record.SubItems(15) = Format(!MBMMCOUtilised, "#,##0.00")
                            End If
                        Else
                            Record.SubItems(15) = "0.00"
                        End If
                    
                    End If
                    
                    If Len(!MBMInvoiceNo) Then
                        Record.SubItems(16) = Trim(!MBMInvoiceNo)
                    End If
                    
                    If Len(!MBMInvoiceDate) Then
                        Record.SubItems(17) = Format(!MBMInvoiceDate, "dd/mm/yyyy")
                    End If
                    
                    If !TotalCase > 0 Then
                        Record.SubItems(18) = "This member had" & " " & !TotalCase & " " & "Case(s)"
                    End If
                    
                    If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                        Call MCORefundListingSupp(StrAppend)
                    End If
                                
                    Current = Current + 1
                    txtMCOTotalPrincipal = Current
                    If Len(rst2!MBMMCORefund) Then
                        TotalAmt = TotalAmt + rst2!MBMMCORefund
                    End If
                    TotalAmt2 = TotalAmt2 + CDbl(Record.SubItems(15))
                    
                    Label3 = Format(TotalAmt, "#,##0.00")
                    Label5 = Format(TotalAmt2, "#,##0.00")
                    
            End With
            rst2.MoveNext
        Loop
    End If
    rst2.Close
    Set rst2 = Nothing
    cmdClear.Enabled = True
    cmdExit.Enabled = True
    cmdPrint.Enabled = True
    cmdPrintCRNote.Enabled = True
End Sub

Private Sub UpdateCreditNote(CRNote As String, MBMNumber As String)
'Dim MBM As New ADODB.Recordset
Dim strsqlCoveredPersons As String
'Dim MBMOthers As New ADODB.Recordset
Dim strsqlMBMOthers As String
Dim CRNoteDate As String

    CRNoteDate = Format(Date, "YYYY/MM/DD")
  
    ' update MBMOthers Table
    strsqlMBMOthers = " update MBMOthers set " & _
                      " MBMCRNoteDate = '" & CRNoteDate & "', " & _
                      " MBMCRNote = '" & Trim(CRNote) & "', " & _
                      " MBMCRNoteby = '" & Logon & "' " & _
                      " Where MBMNumber = '" & Trim(MBMNumber) & "'"
    
        
    medixmasterconn.Execute strsqlMBMOthers
    
    ' update MBMCoveredPersons Table
    strsqlCoveredPersons = " update MBMCoveredPersonsTwo set " & _
                           " MBMCRNoteDate = '" & CRNoteDate & "', " & _
                           " MBMCRNote = '" & Trim(CRNote) & "', " & _
                           " MBMCRNoteby = '" & Logon & "' " & _
                           " Where MBMCNumber = '" & Trim(MBMNumber) & "'"
    
    medixmasterconn.Execute strsqlCoveredPersons


    'strsqlMBM = " Select MBMNumber, MBMENDtEffDate, MBMENDtBordxDate" & _
                "  From MBM Where 0 = 0 "
    
    'If Len(Trim(strAppend)) Then
        'strsqlMBM = strsqlMBM + strAppend
    'End If

    'MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'If MBM.recordCount Then
        'Do Until MBM.EOF
            'strsqlMBMOthers = " Select MBMNumber, MBMCRNote, MBMCRNoteby, MBMCRNoteDate " & _
                              " From MBMOthers Where MBMNumber = '" & Trim(MBM!MBMNumber) & "'"
        
            'MBMOthers.Open strsqlMBMOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
    
            'If MBMOthers.recordCount Then
                'MBMOthers!MBMCRNote = CRNote
                'MBMOthers!MBMCRNoteDate = Date
                'MBMOthers!MBMCRNoteby = Logon
                'MBMOthers.Update
            'End If
            'MBMOthers.Close
            'Set MBMOthers = Nothing
        'MBM.MoveNext
        'Loop
    'End If
    'MBM.Close
    'Set MBM = Nothing
    
End Sub


