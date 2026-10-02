VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMBMClinical 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Clinic"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTempMCOFees 
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtTempPrem 
      Height          =   315
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   26
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtWarranty 
         Height          =   885
         Left            =   120
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1320
         Width           =   2850
      End
      Begin VB.TextBox txtRemarks 
         Height          =   885
         Left            =   3360
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1320
         Width           =   2850
      End
      Begin VB.ComboBox cboCisPln 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   200
         Width           =   1740
      End
      Begin VB.ComboBox cboCisINS 
         Height          =   315
         Left            =   4440
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   200
         Width           =   1740
      End
      Begin MSMask.MaskEdBox txtCLNEffDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtClnExpDate 
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   480
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboCardType 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   770
         Width           =   1740
      End
      Begin VB.Label Label4 
         Caption         =   "Card Type"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   780
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Warranty / Exclusion"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label40 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   1080
         Width           =   1590
      End
      Begin VB.Label Label21 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label16 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   520
         Width           =   1035
      End
      Begin VB.Label Label20 
         Caption         =   "Ins Type"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label15 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   520
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "DOB"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   525
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   6375
      Begin VB.Label Label3 
         Caption         =   "Visit Limit"
         Height          =   255
         Left            =   3360
         TabIndex        =   29
         Top             =   495
         Width           =   840
      End
      Begin VB.Label LblVisitLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   28
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label lblMBMPremium 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label lblMCOFees 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label lblAnnualLimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label27 
         Caption         =   "Ann Limit"
         Height          =   255
         Left            =   3360
         TabIndex        =   12
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label26 
         Caption         =   "MCO Fees"
         Height          =   240
         Left            =   480
         TabIndex        =   11
         Top             =   500
         Width           =   870
      End
      Begin VB.Label Label25 
         Caption         =   "Premium"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   2880
      Width           =   6375
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   5280
         TabIndex        =   7
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   160
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtTempEffDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTempExpDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMBMBirthdate 
         Height          =   315
         Left            =   2040
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.Label txtAGECode 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1080
      TabIndex        =   30
      Top             =   3120
      Width           =   1635
   End
End
Attribute VB_Name = "frmMBMClinical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer

Private Sub CboCardType_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboCardType, KeyAscii)
End Sub

Private Sub CboCisINS_LostFocus()
    If Source <> 3 Then
        If Len(cboCisPln) >= 2 Then
            Call CallCLNLimit
            Call MemberCLNCallFees
        End If
    End If
End Sub

Private Sub ComboListing()
Dim PLN As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim CheckItems As New ADODB.Recordset
Dim cmd As New ADODB.Command

    'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
    
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "SearchPlan"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PLN.CursorLocation = adUseClient
    PLN.CursorType = adOpenForwardOnly
    PLN.CacheSize = 100
    Set PLN = cmd.Execute

    Do Until PLN.EOF
        With PLN
            cboCisPln.AddItem !PLNCode & " - " & !PLNDesc
            PLN.MoveNext
        End With
    Loop
    PLN.Close
    Set PLN = Nothing
    
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "SearchINS"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    INS.CursorLocation = adUseClient
    INS.CursorType = adOpenForwardOnly
    INS.CacheSize = 100
    Set INS = cmd.Execute
    
    Do Until INS.EOF
        With INS
            cboCisINS.AddItem !INSCode & " - " & !InsDesc
            INS.MoveNext
        End With
    Loop
    INS.Close
    Set INS = Nothing
    
    Set cmd.ActiveConnection = MediConnCISDB
    cmd.CommandText = "SearchCheckItems"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    CheckItems.CursorLocation = adUseClient
    CheckItems.CursorType = adOpenForwardOnly
    CheckItems.CacheSize = 100
    Set CheckItems = cmd.Execute

    Do Until CheckItems.EOF
        With CheckItems
            CboCardType.AddItem Trim(!CatName)
            CheckItems.MoveNext
        End With
    Loop
    CheckItems.Close
    Set CheckItems = Nothing
    
    'MediConnCISDB.Close
    'Set MediConnCISDB = Nothing
End Sub

Private Sub CboCisINS_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCisINS, KeyAscii)
End Sub

Private Sub CboCisPln_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCisPln, KeyAscii)
End Sub

Private Sub CboCisPln_LostFocus()
    If Len(cboCisPln) >= 2 And txtCLNEffDate <> "__/__/____" Then
        Call CallCLNLimit
        Call MemberCLNCallFees
    End If
End Sub

Public Sub cmdClose_Click()
Dim RecordExit
    
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
       Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    If Trim(cboCisPln) = "" Then
        MsgBox "Please Enter Plan Code!", vbCritical
        cboCisPln.SetFocus
    ElseIf Trim(cboCisINS) = "" Then
        MsgBox "Please Enter Insured Type!", vbCritical
        cboCisINS.SetFocus
    ElseIf txtCLNEffDate = "__/__/____" Then
        MsgBox "Please Enter Effective Date!", vbCritical
        txtCLNEffDate.SetFocus
    ElseIf txtClnExpDate = "__/__/____" Then
        MsgBox "Please Enter ExpiryDate!", vbCritical
        txtClnExpDate.SetFocus
    ElseIf Trim(CboCardType) = "" Then
        MsgBox "Please Enter Card Type !", vbCritical
        CboCardType.SetFocus
    End If
    frmMBMClinical.Visible = False
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
    Call ComboListing
End Sub

Private Sub txtCLNEffDate_LostFocus()
    If txtCLNEffDate <> "__/__/____" Then
        If IsDate(txtCLNEffDate) Then
            Call CheckEffDateRange("Eff", txtCLNEffDate)
        Else
            MsgBox "Invalid Date format ! ", vbOKOnly + vbCritical
        End If
    End If
End Sub

Private Sub txtCLNExpDate_LostFocus()
    If txtCLNEffDate <> "__/__/____" Then
        If IsDate(txtClnExpDate) Then
            Call CheckEffDateRange("Exp", txtClnExpDate)
        Else
            MsgBox "Invalid Date format ! ", vbOKOnly + vbCritical
        End If
    End If
End Sub

Public Sub CallCLNLimit()
Dim ChkAnnLimit As New ADODB.Recordset
Dim StrSql As String
Dim ChkVisitLimit As New ADODB.Recordset
Dim tmpPln As String
Dim tmpINS As String

    lblAnnualLimit = ""
    LblVisitLimit = ""
    tmpPln = Trim(Mid(cboCisPln, 1, IIf(InStr(1, cboCisPln, "-") = 0, 4, InStr(1, cboCisPln, "-") - 1)))
    tmpINS = Mid(cboCisINS, 1, 1)
    If Trim(tmpPln) <> "" And IsDate(txtCLNEffDate) Then
       ' MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

        StrSql = " Select PLNCode, ANNLmtInd From PLN Where PLNCode = '" & tmpPln & "'"
        ChkAnnLimit.Open StrSql, MediConnCISDB, adOpenForwardOnly, adLockReadOnly
        Do While Not ChkAnnLimit.EOF
        
            If ChkAnnLimit!ANNLmtInd = "N" Then
                lblAnnualLimit = "0.00"
            Else
                lblAnnualLimit = Format(GetCLNAnnualLimit(tmpINS, tmpPln, Trim(frmNewMBMReg05.txtHLTcode), txtCLNEffDate), "#,###,##0.00")
            End If
            ChkAnnLimit.MoveNext
        Loop
        ChkAnnLimit.Close
        Set ChkAnnLimit = Nothing
        
        StrSql = " Select PLNCode, VISLmtInd From PLN Where PLNCode = '" & Trim(tmpPln) & "'"
        ChkVisitLimit.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        
        If ChkVisitLimit.recordCount Then
        
            If ChkVisitLimit!VISLmtInd = "N" Then
                LblVisitLimit = "0.00"
            Else
                LblVisitLimit = Format(GetVisitLimit(tmpINS, tmpPln, Trim(frmNewMBMReg05.txtHLTcode), txtCLNEffDate), "#,###,##0.00")
            End If
        End If
        
        ChkVisitLimit.Close
        Set ChkVisitLimit = Nothing
'
       ' MediConnCISDB.Close
        'Set MediConnCISDB = Nothing

    End If
End Sub

Private Sub MemberCLNCallFees()
Dim chkPLN As New ADODB.Recordset
Dim StrSql As String
Dim tmpPremINd As String
Dim tmpMCOInd As String
Dim tmpRoundUp As String
Dim tmpProRate As String
Dim tmpPln As String
Dim XPos As Integer
Dim tmpMCO As String
    
    tmpPln = Trim(Mid(cboCisPln, 1, IIf(InStr(1, cboCisPln, "-") = 0, 4, InStr(1, cboCisPln, "-") - 1)))
    
    If Trim(frmNewMBMReg05.txtPAYcode) <> "" And Trim(txtAGECode) <> "" And Trim(tmpPln) <> "" _
        And Trim(frmNewMBMReg05.txtHLTcode) <> "" Then
        lblMBMPremium = "0"
        lblMCOFees = "0"
        txtTempPrem = "0.00"
        txtTempMCOFees = "0.00"
        If IsDate(txtCLNEffDate) Then
            'MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"
            StrSql = " Select PLNCode, MCOInd, PRMInd, PLNRoundUpStatus, PLNProRate " & _
                    " From PLN Where PLNCode = '" & tmpPln & "'"
            chkPLN.Open StrSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            tmpPremINd = "N"
            tmpMCOInd = "N"
            tmpRoundUp = "Y"
            tmpProRate = "Y"
            If chkPLN.recordCount Then
                tmpPremINd = IIf(Len(Trim(chkPLN!PRMInd)), Trim(chkPLN!PRMInd), "N")
                tmpMCOInd = IIf(Len(Trim(chkPLN!MCOInd)), Trim(chkPLN!MCOInd), "N")
                tmpRoundUp = IIf(Len(Trim(chkPLN!PLNRoundUpStatus)), Trim(chkPLN!PLNRoundUpStatus), "Y")
                tmpProRate = IIf(Len(Trim(chkPLN!PLNProRate)), Trim(chkPLN!PLNProRate), "Y")
            End If
            chkPLN.Close
            Set chkPLN = Nothing
            
            If tmpPremINd = "Y" Then
                lblMBMPremium = Format(GetCLNPremium(Mid(cboCisINS, 1, 1), Trim(frmNewMBMReg05.txtPAYcode), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), _
                tmpPln, Trim(frmNewMBMReg05.txtHLTcode), txtCLNEffDate), "#,##0.00")
                txtTempPrem = Format(GetCLNPremium(Mid(cboCisINS, 1, 1), Trim(frmNewMBMReg05.txtPAYcode), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), _
                tmpPln, Trim(frmNewMBMReg05.txtHLTcode), txtCLNEffDate), "#,##0.00")
            End If
            If tmpMCOInd = "Y" Then
                tmpMCO = Format(GetCLNMCO(Trim(Mid(cboCisINS, 1, 1)), Trim(frmNewMBMReg05.txtHLTcode), frmNewMBMReg05.txtMBMAdultChild, txtCLNEffDate, tmpPln), "#,##0.00")
                If tmpProRate <> "N" Then
                    tmpMCO = InsertMCOFees(CDate(txtClnExpDate), CDate(txtCLNEffDate), Format(tmpMCO, "#,##0.00"))
                End If
                tmpMCO = Format(tmpMCO, "#,##0.00")
                If tmpRoundUp <> "N" Then
                    tmpMCO = Round(tmpMCO)
                End If
                txtTempMCOFees = tmpMCO
                lblMCOFees = tmpMCO
            End If
          '  MediConnCISDB.Close
           ' Set MediConnCISDB = Nothing
        End If
    End If
        
End Sub

Private Sub CLNMemberCallAge()
Dim SelectAge As String
Dim tmpPlnCode As String
Dim tmpInsType As String
Dim tmpDOb As String
    
    tmpPlnCode = frmNewMBMReg05.CboPlnCode
    tmpInsType = Mid(frmNewMBMReg05.cboINSCode, 1, 1)
    tmpDOb = frmNewMBMReg05.txtMBMBirthdate
    
    If Trim(tmpPlnCode) = "" Then
        tmpPlnCode = Trim(Mid(cboCisPln, 1, IIf(InStr(1, cboCisPln, "-") = 0, 4, InStr(1, cboCisPln, "-") - 1)))
        tmpInsType = Mid(cboCisINS, 1, 1)
    End If
    If Trim(txtCLNEffDate) <> "__/__/____" And tmpDOb <> "__/__/____" _
        And Trim(frmNewMBMReg05.txtHLTcode) <> "" Then
        If IsDate(tmpDOb) And IsDate(txtCLNEffDate) Then
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

            SelectAge = GetAge(txtCLNEffDate, CDate(tmpDOb))
            txtAGECode = GetAgeBand(Trim(frmNewMBMReg05.txtHLTcode), SelectAge, tmpPlnCode, tmpInsType)
           '' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
            If Trim(frmNewMBMReg05.txtMBMAdultChild) = "" Then
                If Len(Trim(txtAGECode)) <> 0 Then
                    If Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)) = "01" Then
                        frmNewMBMReg05.txtMBMAdultChild = "C"
                    Else
                        frmNewMBMReg05.txtMBMAdultChild = "A"
                    End If
                End If
            End If
            If Trim(frmNewMBMReg05.txtAGECode) = "" Then
                frmNewMBMReg05.txtAGECode = Trim(txtAGECode)
            End If
        End If
    End If
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

Private Sub CheckEffDateRange(DateStatus As String, PolDate As String)
Dim StrSql As String
Dim tmpSDate As String
Dim tmpEDate As String
Dim TempDate As String
Dim MBMValidDate As New ADODB.Recordset

    If DateStatus = "Eff" Then
        StrSql = "Select ValidID , MBMEffStart as StartDate, MBMEffEnd as EndDate From MBMValidDateRange "
    Else
        StrSql = "Select ValidID , MBMExpStart as StartDate, MBMExpEnd as EndDate From MBMValidDateRange "
    End If
    
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    MBMValidDate.Open StrSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    Do While Not MBMValidDate.EOF
        tmpSDate = MBMValidDate!StartDate
        tmpEDate = MBMValidDate!EndDate
        MBMValidDate.MoveNext
    Loop
    MBMValidDate.Close
    Set MBMValidDate = Nothing
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
    If Len(tmpSDate) Then
         If DateDiff("d", tmpSDate, PolDate) >= 0 And DateDiff("d", tmpEDate, PolDate) <= 0 Then
             If DateStatus = "Eff" Then
                TempDate = Format(DateAdd("yyyy", 1, txtCLNEffDate), "dd/mm/yyyy")
                txtClnExpDate = Format(DateAdd("d", -1, TempDate), "dd/mm/yyyy")
             End If
             If txtMBMBirthdate <> "__/__/____" Then
                 Call CLNMemberCallAge
                 Call MemberCLNCallFees
                 Call CallCLNLimit
             End If
             txtTempEffDate = txtCLNEffDate
             txtTempExpDate = txtClnExpDate
         Else
             MsgBox "Policy Effective Date is out of acceptable range! " & vbNewLine & _
                     "It must be in between " & tmpSDate & " and " & tmpEDate & "! ", vbCritical + vbOKOnly
         End If
    End If
End Sub
