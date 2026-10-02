VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimsUploadAXA 
   Caption         =   "Upload Claims File - AXA Affin General"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   8550
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   8415
      Begin VB.CommandButton cmdFileSearch 
         Caption         =   ">"
         Height          =   255
         Left            =   7890
         TabIndex        =   7
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox txtFileName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   150
         Width           =   7695
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload"
         Height          =   375
         Left            =   6360
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   7320
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSComctlLib.ListView lstSendDate 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SendDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstErrorList 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Approval Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Approval Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ReturnDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Upload Claims File"
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
      TabIndex        =   5
      Top             =   0
      Width           =   8445
   End
End
Attribute VB_Name = "frmClaimsUploadAXA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordArray() As String
Dim startTrans As Boolean
Dim TextLine
Dim tmpRecord As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFileSearch_Click()
  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  'CommonDialog1.InitDir = "\\MARS\" & "Claims\" & "ASCII_File"
  CommonDialog1.InitDir = "\\svr-doc\" & "Claims\" & "ASCII_File"
  CommonDialog1.Flags = cdlOFNHideReadOnly
  CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
  CommonDialog1.ShowOpen
  txtFileName = CommonDialog1.FileTitle
  Exit Sub
  
ErrHandler:
  Exit Sub
End Sub

Private Sub cmdUpload_Click()
On Err GoTo ErrorSave
Dim RecordSaved
Dim PAY As New ADODB.Recordset
Dim strsql As String


    If Trim(txtFileName) = "" Then
        MsgBox "Please Insert File", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Do you want to upload this file?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Screen.MousePointer = vbHourglass
            cmdUpload.Enabled = False
            cmdExit.Enabled = False
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
            'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans
            medixmasterconnWS.beginTrans
            startTrans = True

                'If Trim(Mid(txtFileName, 6, 2)) = "CH" Then
                    Call SaveClaimsDetails
                'End If
            
            medixmasterconn.commitTrans
            medixmasterconnMaint.commitTrans
            medixmasterconnWS.commitTrans
            startTrans = False
            MsgBox "Record Import Successfully...", vbOKOnly + vbInformation, DialogBoxTitle

            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
            'medixmasterconnMaint.Close
            'Set medixmasterconnMaint = Nothing
            'medixmasterconnWS.Close
            'Set medixmasterconnWS = Nothing
        End If
    End If
    Screen.MousePointer = vbDefault
    cmdExit.Enabled = True
    Exit Sub
ErrorSave:
        Screen.MousePointer = vbDefault
        If startTrans = True Then
            medixmasterconn.RollbackTrans
            medixmasterconnMaint.RollbackTrans
            medixmasterconnWS.RollbackTrans
        End If
        MsgBox Err.Description
    
End Sub

Private Sub SaveClaimsDetails()
Dim MBMWSql As String
Dim VHospCode, VInsurerCode, VInsPolicyID, VInsFWID, VPolicyNo As String
Dim VPolEffDate, VPassportNo, VFWName, VESPRefNo As String
Dim VDOA, VDOD, VAdmissionReason, VBillNo, VBillDate As String
Dim VBilledAmt, VTotalHospBilAmt, VTotalHospCollMBM As Double
Dim VBillLodgedDate, VBillLodgedby, VDiagnosisCode, VTreamentCode, VMRNNo As String
Dim VCaseNumber, VMBMNumber As String
Dim MBMrst As New ADODB.Recordset
Dim NoOfRecLine As Integer
Dim strsqlMBM, strsqlCLM As String
Dim TempDOR As String
Dim strsqlCAS, tmpBordxRefNo, tmpPayto, lblLastVersion, strsql As String
Dim CASStatus As Boolean
Dim intIndex As String
Dim tmpRcrdType As String
Dim tmpAXAClaimNo As String
Dim tmpAXAClaimID As String
Dim tmpAXAMBMID As String
Dim tmpAXAPrinName As String
Dim tmpAXADepName As String
Dim tmpPlan As String
Dim tmpHospDays As String
Dim tmpClaimType As String
Dim tmpRemarks As String
Dim tmpDoctor As String
Dim tmptreament As String
Dim tmpDiagCode1 As String
Dim tmpDiagCode2 As String
Dim tmpDiagCode3 As String
Dim tmpRejectreason As String
Dim tmpclaimstatus As String
Dim tmpAmtpaidtombm As String
Dim tmpAmtpaidbymbm As String
Dim tmpStaytype As String
Dim tmpMedicalHistory As String
Dim tmpChqType As String
Dim tmpCoypay As String
Dim tmpPayeeName As String
Dim tmpApprovedby As String
Dim tmpPostAct As Double
Dim tmpPostApp As Double
Dim tmpPostNotPay As Double
Dim tmpCancerAct As Double
Dim tmpCancerApp As Double
Dim tmpCancerNotPay As Double
Dim tmpRBAct As Double
Dim tmpRBApp As Double
Dim tmpRBNotPay As Double
Dim tmpHSSAct As Double
Dim tmpHSSApp As Double
Dim tmpHSSNotPay As Double
Dim tmpPhyVisitAct As Double
Dim tmpPhyVisitApp As Double
Dim tmpPhyVisitNotPay As Double
Dim tmpMedRptAct As Double
Dim tmpMedRptApp As Double
Dim tmpMedRptNotPay As Double
Dim tmpPreHospAct As Double
Dim tmpPreHospApp As Double
Dim tmpPreHospNotpay As Double
Dim tmpRelstatus As String
Dim tmpHISPLNCode As String
Dim tmpBalLmt As Double
Dim tmpGovAlwAct As Double
Dim tmpGovAlwApp As Double
Dim tmpGovAlwNotPay As Double
Dim tmpGovTaxAct As Double
Dim tmpGovTaxApp As Double
Dim tmpGovTaxNotPay As Double

Dim tmpRBLmt As Double
Dim tmpGovAlwLmt As Double
Dim tmpMedRptLmt As Double
Dim tmpCancerLmt As Double
Dim tmpOPAccTreamentLmt As String

Dim strsqlchk As String
Dim CLM As New ADODB.Recordset
Dim CLMRecord As Boolean

Dim i, Count As Integer

    If txtFileName <> "" Then
        tmpPreHospAct = 0
        tmpPreHospApp = 0
        tmpPreHospNotpay = 0
        
        Erase RecordArray
        tmpBordxRefNo = ""
        tmpPayto = ""
        Open txtFileName For Input As #1   ' Open file.
        Do While Not EOF(1)  ' Loop until end of file.
           Line Input #1, TextLine   ' Read line into variable.
            NoOfRecLine = NoOfRecLine + 1
                RecordArray() = Split(TextLine, "|")
                    tmpRcrdType = Trim(RecordArray(0))
                    If tmpRcrdType = "PH" Then
                        tmpAXAClaimNo = Trim(RecordArray(1))
                        tmpAXAClaimID = Trim(RecordArray(2))
                        tmpAXAMBMID = Trim(RecordArray(3))
                        tmpAXAPrinName = Trim(RecordArray(4))
                        tmpAXADepName = Trim(RecordArray(5))
                        
                        
                        VPolicyNo = Trim(RecordArray(6))
                        VDOA = Mid(RecordArray(7), 7, 2) & "/" & Mid(RecordArray(7), 5, 2) & "/" & Mid(RecordArray(7), 1, 4)
                        VDOD = Mid(RecordArray(8), 7, 2) & "/" & Mid(RecordArray(8), 5, 2) & "/" & Mid(RecordArray(8), 1, 4)
                        If VDOD = "99/99/9999" Then
                            VDOD = VDOA
                        End If
                        tmpPlan = Trim(RecordArray(9))
                        tmpHospDays = Trim(RecordArray(10))
                        VBillNo = Trim(RecordArray(11))
                        tmpClaimType = Trim(RecordArray(12))
                        tmpRemarks = Trim(RecordArray(13))
                        tmpDoctor = Trim(RecordArray(14))
                        
                        
                        If Trim(RecordArray(15)) = "A&S" Then
                            tmptreament = "B"
                        ElseIf Trim(RecordArray(15)) = "S" Then
                            tmptreament = "S"
                        ElseIf Trim(RecordArray(15)) = "A" Then
                            tmptreament = "N"
                        ElseIf Trim(RecordArray(15)) = "M" Then
                            tmptreament = "N"
                        Else
                            tmptreament = "N"
                        End If
                        tmpDiagCode1 = Trim(RecordArray(16))
                        tmpDiagCode2 = Trim(RecordArray(18))
                        tmpDiagCode3 = Trim(RecordArray(20))
                        tmpRejectreason = Trim(RecordArray(23))
                        If Trim(RecordArray(24)) = "D" Then
                            tmpclaimstatus = "RR"
                        Else
                            tmpclaimstatus = "AA"
                        End If
                        
                        VHospCode = "GP"
                        'VInsurerCode = Trim(RecordArray(2))
                        'VInsPolicyID = Trim(RecordArray(3))
                        'VInsFWID = Trim(RecordArray(4))
                        'tPolDate = Mid(RecordArray(4), 7, 2) & "/" & Mid(RecordArray(4), 5, 2) & "/" & Mid(RecordArray(4), 1, 4)
                        
                        'VPolEffDate = Mid(RecordArray(6), 7, 2) & "/" & Mid(RecordArray(6), 5, 2) & "/" & Mid(RecordArray(6), 1, 4)
                        'tmpEmpType = Trim(RecordArray(7))
                        'VPassportNo = Trim(RecordArray(8))
                        'VFWName = Trim(RecordArray(9))
                        'tmpVAdd1 = Trim(RecordArray(10))
                        'tmpVAdd2 = Trim(RecordArray(11))
                        'VESPRefNo = Trim(RecordArray(12))
                        'VAdmissionReason = Trim(RecordArray(15))
                        'tmpclaimstatus = Trim(RecordArray(16))
                    ElseIf tmpRcrdType = "PD" Then
                        If Trim(RecordArray(2)) = "D007" Then
                            tmpCancerAct = IIf(IsNumeric(RecordArray(4)), Format(RecordArray(4), "#,##0.00"), 0)
                            tmpCancerApp = IIf(IsNumeric(RecordArray(5)), Format(RecordArray(5), "#,##0.00"), 0)
                            tmpCancerNotPay = IIf(IsNumeric(RecordArray(6)), Format(RecordArray(6), "#,##0.00"), 0)
                        ElseIf Trim(RecordArray(2)) = "D008" Then
                            tmpPreHospAct = IIf(IsNumeric(RecordArray(4)), Format(RecordArray(4), "#,##0.00"), 0)
                            tmpPreHospApp = IIf(IsNumeric(RecordArray(5)), Format(RecordArray(5), "#,##0.00"), 0)
                            tmpPreHospNotpay = IIf(IsNumeric(RecordArray(6)), Format(RecordArray(6), "#,##0.00"), 0)
                        ElseIf Trim(RecordArray(2)) = "D009" Then
                            tmpPreHospAct = tmpPreHospAct + IIf(IsNumeric(RecordArray(4)), Format(RecordArray(4), "#,##0.00"), 0)
                            tmpPreHospApp = tmpPreHospApp + IIf(IsNumeric(RecordArray(5)), Format(RecordArray(5), "#,##0.00"), 0)
                            tmpPreHospNotpay = tmpPreHospNotpay + IIf(IsNumeric(RecordArray(6)), Format(RecordArray(6), "#,##0.00"), 0)
                        
                        ElseIf Trim(RecordArray(2)) = "D011" Then
                            tmpPostAct = IIf(IsNumeric(RecordArray(4)), Format(RecordArray(4), "#,##0.00"), 0)
                            tmpPostApp = IIf(IsNumeric(RecordArray(5)), Format(RecordArray(5), "#,##0.00"), 0)
                            tmpPostNotPay = IIf(IsNumeric(RecordArray(6)), Format(RecordArray(6), "#,##0.00"), 0)
                            
                        ElseIf Trim(RecordArray(2)) = "H001" Then
                            tmpRBAct = IIf(IsNumeric(RecordArray(4)), Format(RecordArray(4), "#,##0.00"), 0)
                            tmpRBApp = IIf(IsNumeric(RecordArray(5)), Format(RecordArray(5), "#,##0.00"), 0)
                            tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), Format(RecordArray(6), "#,##0.00"), 0)
                        ElseIf Trim(RecordArray(2)) = "H002" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H003" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        
                        ElseIf Trim(RecordArray(2)) = "H004" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H005" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H006" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H007" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H008" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H009" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H010" Then
                            tmpPhyVisitAct = IIf(IsNumeric(RecordArray(4)), Format(RecordArray(4), "#,##0.00"), 0)
                            tmpPhyVisitApp = IIf(IsNumeric(RecordArray(5)), Format(RecordArray(5), "#,##0.00"), 0)
                            tmpPhyVisitNotPay = IIf(IsNumeric(RecordArray(6)), Format(RecordArray(6), "#,##0.00"), 0)
                        ElseIf Trim(RecordArray(2)) = "H011" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H012" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H013" Then
                            tmpHSSAct = IIf(IsNumeric(RecordArray(4)), Format(RecordArray(4), "#,##0.00"), 0)
                            tmpHSSApp = IIf(IsNumeric(RecordArray(5)), Format(RecordArray(5), "#,##0.00"), 0)
                            tmpHSSNotPay = IIf(IsNumeric(RecordArray(6)), Format(RecordArray(6), "#,##0.00"), 0)
                        ElseIf Trim(RecordArray(2)) = "H014" Then
                            tmpGovAlwAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            tmpGovAlwApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            tmpGovAlwNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H015" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H016" Then
                            'tmpRBAct = IIf(IsNumeric(RecordArray(4)), RecordArray(4), 0)
                            'tmpRBApp = IIf(IsNumeric(RecordArray(5)), RecordArray(5), 0)
                            'tmpRBNotPay = IIf(IsNumeric(RecordArray(6)), RecordArray(6), 0)
                        ElseIf Trim(RecordArray(2)) = "H017" Then
                            tmpMedRptAct = IIf(IsNumeric(RecordArray(4)), Format(RecordArray(4), "#,##0.00"), 0)
                            tmpMedRptApp = IIf(IsNumeric(RecordArray(5)), Format(RecordArray(5), "#,##0.00"), 0)
                            tmpMedRptNotPay = IIf(IsNumeric(RecordArray(6)), Format(RecordArray(6), "#,##0.00"), 0)
                        ElseIf Trim(RecordArray(2)) = "H018" Then
                            tmpGovTaxAct = IIf(IsNumeric(RecordArray(4)), Format(RecordArray(4), "#,##0.00"), 0)
                            tmpGovTaxApp = IIf(IsNumeric(RecordArray(5)), Format(RecordArray(4), "#,##0.00"), 0)
                            tmpGovTaxNotPay = IIf(IsNumeric(RecordArray(6)), Format(RecordArray(4), "#,##0.00"), 0)
                        End If
                    
                    ElseIf tmpRcrdType = "PT" Then
                        'VBillDate = Mid(RecordArray(17), 7, 2) & "/" & Mid(RecordArray(17), 5, 2) & "/" & Mid(RecordArray(17), 1, 4)
                        VTotalHospBilAmt = Trim(Format(RecordArray(2), "#,##0.00"))
                        tmpAmtpaidtombm = Trim(Format(RecordArray(3), "#,##0.00"))
                        tmpAmtpaidbymbm = Trim(Format(RecordArray(4), "#,##0.00"))
                        'tmpAmtpaidbymbm = Trim(RecordArray(3))
                        
                        
                        If Trim(RecordArray(5)) = "PRE" Then
                            tmpStaytype = "P"
                        ElseIf Trim(RecordArray(5)) = "POS" Then
                            tmpStaytype = "F"
                        ElseIf Trim(RecordArray(5)) = "IP" Then
                            tmpStaytype = "I"
                        ElseIf Trim(RecordArray(5)) = "OP" Then
                            tmpStaytype = "O"
                        End If
                        
                        tmpMedicalHistory = Trim(RecordArray(6))
                        
                        tmpPayeeName = Trim(RecordArray(7))
                        
                        tmpChqType = Trim(RecordArray(8))
                        
                        tmpApprovedby = Trim(RecordArray(9))
                        
                        tmpCoypay = Trim(Format(RecordArray(10), "#,##0.00"))
                        'If RecordArray(19) = "" Then
                        '    VBillLodgedDate = ""
                        'Else
                        '    VBillLodgedDate = Mid(RecordArray(19), 7, 2) & "/" & Mid(RecordArray(19), 5, 2) & "/" & Mid(RecordArray(19), 1, 4)
                        'End If
                        'VBillLodgedby = Trim(RecordArray(20))
                        'VDiagnosisCode = Trim(RecordArray(21))
                        'VTreamentCode = Trim(RecordArray(22))
                        'VMRNNo = Trim(RecordArray(23))
                        'VTotalHospBilAmt = Trim(RecordArray(24))
                        'VTotalHospCollMBM = Trim(RecordArray(25))
                        i = 0
                        'Count = 0
                        i = InStr(i + 1, tmpPayeeName, "HOSPITAL")
                        
                    End If
                        TempDOR = Date
                        tmpRecord = False
                        
                        'VFWName = Trim(Replace(VFWName, "'", "''"))
                        'If i = 0 Then
                        'MsgBox 0
                        'Exit Sub
                        'End If
                        'Do While i > 0
                        'i = InStr(i + 1, Text1.Text, "How")
                        'Count = Count + 1
                        'Loop
                        'MsgBox Count
                        
                    'If tmpRcrdType = "PT" And i = 0 Then
                    If tmpRcrdType = "PT" Then
                        strsqlMBM = ""
                        VMBMNumber = ""
                        CLMRecord = False
                            If tmpAXAClaimNo <> "" Then
                                strsqlchk = "select * from Claim where  CASNumber = '" & Mid(tmpAXAClaimNo, 1, 7) & "' And  CLMTotalActual = '" & VTotalHospBilAmt & "'"
                                CLM.Open strsqlchk, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                                
                                If CLM.recordCount Then
                                    CLMRecord = True
                                Else
                                    CLMRecord = False
                                End If
                                CLM.Close
                                Set CLM = Nothing
                            
                            Else
                                CLMRecord = False
                            End If
                        If CLMRecord = False Then
                                strsqlMBM = "Select MBMNumber,MBMPrevMemNo from View_MBM_Search2 where MBMPrevMemNo = '" & Trim(tmpAXAMBMID) & "' and  MBMPolicyNo = '" & VPolicyNo & "'"
                                MBMrst.Open strsqlMBM, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                                
                                Do While Not MBMrst.EOF
                                    VMBMNumber = MBMrst!MBMNumber
                                    tmpRelstatus = "00"
                                    tmpRecord = True
                                MBMrst.MoveNext
                                Loop
                                
                                MBMrst.Close
                                Set MBMrst = Nothing
                                
                                
                                
                                If tmpRecord = True Then
                                    tmpRelstatus = "00"
                                    
                                    strsqlMBM = ""
                                    strsqlMBM = "Select MBMNumber,PLNCode,MBMAvailableLimit from MBM where MBMNumber = '" & Trim(VMBMNumber) & "'"
                                    MBMrst.Open strsqlMBM, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                                    
                                    Do While Not MBMrst.EOF
                                        tmpRelstatus = "00"
                                        tmpHISPLNCode = MBMrst!PLNCode
                                        tmpBalLmt = MBMrst!MBMAvailableLimit
                                    MBMrst.MoveNext
                                    Loop
                                    
                                    MBMrst.Close
                                    Set MBMrst = Nothing
    
                                Else
                                
                                    strsqlMBM = ""
                                    strsqlMBM = "Select MBMCNumber,MBMCCoverID,MBMName,MBMAXANumber,MBMPolicyNo,MBMCAvailableLimit,MBMCPLNCode from View_MBM_Supplementary2 where MBMAXANumber = '" & Trim(tmpAXAMBMID) & "' and  MBMPolicyNo = '" & VPolicyNo & "'"
                                    MBMrst.Open strsqlMBM, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                                    
                                    Do While Not MBMrst.EOF
                                        VMBMNumber = MBMrst!MBMCNumber
                                        tmpRelstatus = MBMrst!MBMCCoverID
                                        tmpHISPLNCode = MBMrst!MBMCPLNCode
                                        tmpBalLmt = MBMrst!MBMCAvailableLimit
                                        tmpRecord = True
                                    MBMrst.MoveNext
                                    Loop
                                    MBMrst.Close
                                    Set MBMrst = Nothing
    
                                'strsqlMBM = ""
                                'strsqlMBM = "Select MBMCNumber,MBMCCoverID,MBMCName,MBMCAvailableLimit from MBMCoveredPersons where MBMNumber = '" & Trim(VMBMNumber) & "' and  MBMCName = '" & Trim(tmpAXADepName) & "'"
                                'MBMrst.Open strsqlMBM, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                               '
                               ' Do While Not MBMrst.EOF
                               '     tmpRelstatus = MBMrst!MBMCCoverID
                               '     tmpHISPLNCode = MBMrst!MBMCPLNCode
                               '     tmpBalLmt = MBMrst!MBMCAvailableLimit
                               ' MBMrst.MoveNext
                               ' Loop
                               '
                               ' MBMrst.Close
                               ' Set MBMrst = Nothing
                                End If
                            
                            
                                strsqlMBM = ""
                                strsqlMBM = "Select BNFCode , BNFNoOFDays, BNFClaimableAmount , BNFRMPerDis  , BNFEffDate from BNF  " & _
                                " where PLNCode = '" & tmpHISPLNCode & "'"
                                strsqlMBM = strsqlMBM & " order by BNFCOde "
                                MBMrst.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                            
                                    Do While Not MBMrst.EOF
                                        If MBMrst!BNFCode = "101" Then
                                            tmpRBLmt = IIf(IsNumeric(MBMrst!BNFClaimAbleAmount), MBMrst!BNFClaimAbleAmount, 0)
                                            If tmpHospDays <> 0 Then
                                                tmpRBLmt = CDbl(tmpRBLmt) * CDbl(tmpHospDays)
                                            End If
                                            
                                        End If
                                        
                                        
                                        If MBMrst!BNFCode = "401" Then
                                            tmpOPAccTreamentLmt = IIf(IsNumeric(MBMrst!BNFClaimAbleAmount), MBMrst!BNFClaimAbleAmount, 0)
                                        End If
                                        
                                        
                                        If MBMrst!BNFCode = "404" Then
                                            tmpCancerLmt = IIf(IsNumeric(MBMrst!BNFClaimAbleAmount), MBMrst!BNFClaimAbleAmount, 0)
                                        End If
                                        
                                        If MBMrst!BNFCode = "601" Then
                                            tmpGovAlwLmt = IIf(IsNumeric(MBMrst!BNFClaimAbleAmount), MBMrst!BNFClaimAbleAmount, 0)
                                        End If
                                        
                                        If MBMrst!BNFCode = "801" Then
                                            tmpMedRptLmt = IIf(IsNumeric(MBMrst!BNFClaimAbleAmount), MBMrst!BNFClaimAbleAmount, 0)
                                        End If
                                        
                                        MBMrst.MoveNext
                                    Loop
    
                                MBMrst.Close
                                Set MBMrst = Nothing
                                
                            
                            If tmpRecord = True Then
                                VCaseNumber = GenerateCaseNo("AX")
                                lblLastVersion = 1
    
                                If CASStatus = False Then
                                    strsqlCAS = ""
                                    'If VDOD = "99/99/9999" Then
                                    '    VDOD = VDOA
                                    'End If
                                    strsqlCAS = "EXEC Case_RegAXACAS '" & VCaseNumber & "', '" & VMBMNumber & "', '" & tmpRelstatus & "', " & _
                                                "'O', '" & Format(VDOA, "yyyy/mm/dd") & "', '" & Format(VDOD, "yyyy/mm/dd") & "', 'AA', " & _
                                                "'I' , 'R', " & _
                                                "'" & VTreamentCode & "', '" & Logon & "' ,'" & Format(TempDOR, "yyyy/mm/dd") & "', '" & VHospCode & "', " & _
                                                "'', '' ,'', '" & VMRNNo & "'"
                                                medixmasterconn.Execute strsqlCAS
                                    
                                    strsqlCAS = ""
                                    strsqlCAS = " EXEC Case_RegInsertCASOthers '" & VCaseNumber & "', " & _
                                                    "'', '' , '', " & _
                                                    "'', '', '', " & _
                                                    "'', '', '', '', " & _
                                                    "" & 0 & ", '', '', '', '', '', " & _
                                                    "'', '', '', '', '','',  " & _
                                                    "'', '', '', '', '',  ''," & _
                                                    "'', '', '', '', '', ''," & _
                                                    "'', '', '', '', '', '', '', ''"
                                                    
                                    medixmasterconn.Execute strsqlCAS
            
                                    strsqlCAS = ""
                                    strsqlCAS = "EXEC Case_RegAXADiagFinal '" & VCaseNumber & "', '" & tmpDiagCode1 & "', '" & tmpDiagCode2 & "', '" & tmpDiagCode3 & "'"
                                    medixmasterconn.Execute strsqlCAS
                                
                                    strsqlCAS = ""
                                    strsqlCAS = "EXEC Case_RegAXADiagFirst '" & VCaseNumber & "', '" & tmpDiagCode1 & "', '" & tmpDiagCode2 & "', '" & tmpDiagCode3 & "'"
                                    medixmasterconn.Execute strsqlCAS
                                    
                                    strsqlCAS = ""
                                    strsqlCAS = "EXEC Case_RegInsertCASRemark '" & VCaseNumber & "', '" & tmpMedicalHistory & "', ''"
                                    medixmasterconn.Execute strsqlCAS
                                    
                                    strsqlCAS = ""
                                    strsqlCAS = " EXEC Case_RegInsertCASRef1 '" & VCaseNumber & "', '" & VMBMNumber & "', '" & tmpRelstatus & "', " & _
                                                "'O', '', '', " & _
                                                "'" & Format(VDOA, "yyyy/mm/dd") & "', 'AA', " & _
                                                "'', '', ''"
                                    medixmasterconnMaint.Execute strsqlCAS
                                End If
                                                                                    
                                strsql = ""
                                strsql = "EXEC Insert_ClaimInvoice_WO '" & lblLastVersion & "', '1', '1', " & _
                                         "'" & VBillNo & "', '" & VCaseNumber & "' , '" & Format(VBillDate, "yyyy/mm/dd") & "', " & _
                                         "'" & Format(TempDOR, "YYYY/MM/DD") & "', '" & Logon & "' , '" & VHospCode & "', " & _
                                         "'I', 'H' , 'M', " & _
                                         "'HS', '" & VBilledAmt & "', " & 0 & " , " & 0 & "," & 0 & " ," & 0 & " ,  '" & VBilledAmt & "'"
                                medixmasterconnWS.Execute strsql
                                
                                Dim ClaimInvoiceLastIndex  As New ADODB.Recordset
                                ClaimInvoiceLastIndex.Open "Select INVINdex from ClaimInvoice Order by INVIndex DESC", medixmasterconnWS, adOpenKeyset, adLockOptimistic
                                If Not ClaimInvoiceLastIndex.EOF Then
                                    intIndex = ClaimInvoiceLastIndex!INVIndex
                                Else
                                    intIndex = 1
                                End If
                                ClaimInvoiceLastIndex.Close
                                Set ClaimInvoiceLastIndex = Nothing
                        
                                strsql = ""
                                strsql = "EXEC Insert_ClaimInvActual2010 '" & intIndex & "', " & 0 & " , " & tmpRBAct & ",  " & 0 & " ," & _
                                            "" & 0 & ", " & 0 & "," & 0 & ", " & tmpHSSAct & ", " & 0 & ", " & _
                                            "" & tmpPreHospAct & ", " & tmpPhyVisitAct & ", " & 0 & ", " & 0 & ", " & 0 & ", " & _
                                            "" & 0 & ", " & 0 & ", " & tmpPostAct & ", " & 0 & ", " & 0 & ", " & 0 & ", " & _
                                            "" & 0 & ", " & tmpCancerAct & ", " & 0 & ", " & tmpMedRptAct & " , " & 0 & ", " & 0 & ", " & _
                                            "" & 0 & ", " & 0 & ", " & 0 & "," & 0 & ", " & 0 & ", " & 0 & "," & 0 & ",  " & 0 & "," & 0 & "," & 0 & ", " & _
                                            "" & 0 & ", " & 0 & ",  " & 0 & "," & tmpGovAlwAct & ", " & 0 & ", " & 0 & "," & tmpGovTaxAct & " ," & 0 & "," & 0 & "," & 0 & "," & 0 & ""
                                           ' "" & VTotalHospBilAmt & ", " & VTotalHospCollMBM & ""
        
                                medixmasterconnWS.Execute strsql
                                
                                strsql = ""
                                strsql = "EXEC Insert_ClaimInvoiceActOthers_New " & intIndex & ", " & 0 & " , " & 0 & ", " & _
                                     "" & 0 & ", " & 0 & " , " & 0 & ", " & _
                                     "" & 0 & ", " & 0 & " , " & 0 & ", " & _
                                     "" & 0 & ", " & 0 & " , " & 0 & ", " & _
                                     "" & 0 & ", " & 0 & ", " & 0 & ", " & _
                                     "" & 0 & ", " & 0 & ", " & 0 & ", " & _
                                     "" & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", " & _
                                     "" & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ",  " & 0 & ", " & _
                                     "'" & 0 & "', '" & 0 & "', '" & 0 & "', '" & 0 & "' , '" & 0 & "', " & _
                                     "" & 0 & ", " & 0 & ", " & 0 & ", " & _
                                     "" & 0 & ", " & 0 & ", " & 0 & ", " & 0 & "," & _
                                     "" & 0 & ", " & 0 & ""
                                medixmasterconnWS.Execute strsql
                                
                                strsql = ""
                                strsql = "EXEC Insert_ClaimInvoiceNotCov08 " & intIndex & ", " & 0 & " , " & 0 & ", " & _
                                     "" & 0 & ", " & 0 & " , " & 0 & ", " & _
                                     "" & 0 & ", " & 0 & " , " & 0 & " , " & 0 & ", " & 0 & ""
                                medixmasterconnWS.Execute strsql
                                
                                
                                If CASStatus = False Then
                                    strsqlCLM = ""
                                    strsqlCLM = "EXEC Insert_WS_Claim_2010 '" & VCaseNumber & "', " & lblLastVersion & ", '" & Format(TempDOR, "yyyy/mm/dd") & "', '" & Logon & "', " & _
                                                "'" & Logon & "', '" & Format(TempDOR, "yyyy/mm/dd") & "', " & _
                                                "'1', 'N'," & _
                                                " 'N' ,'N', '', " & _
                                                "'C', '" & Format(TempDOR, "yyyy/mm/dd") & "', 'N', " & _
                                                "'I' , '" & VTotalHospBilAmt & "' , '" & tmpAmtpaidtombm & "', " & _
                                                "" & 0 & ", " & 0 & ", " & _
                                                "'H'," & 0 & " , " & 0 & ",  " & _
                                                "" & 0 & ", " & 0 & ", " & _
                                                "'" & Format(TempDOR, "yyyy/mm/dd") & "','" & tmpBalLmt & "' , " & 0 & ", " & 0 & ", " & _
                                                "" & 0 & ", 'VOID', '" & tmpAmtpaidtombm & "' , " & 0 & ", " & 0 & ", '" & tmpAmtpaidtombm & "', 'O', '" & tmpCoypay & "' , " & 0 & ", " & 0 & ", " & 0 & ", 'O'," & 0 & " , '" & tmpAmtpaidtombm & "', " & 0 & ", '" & tmpAmtpaidtombm & "', " & _
                                                "'" & tmpPayeeName & "', " & 0 & ", '" & tmpChqType & "', " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ""
                                    medixmasterconnWS.Execute strsqlCLM
                                    
                                    strsqlCLM = ""
                                    strsqlCLM = "EXEC Insert_WS_ClaimLimit '" & VCaseNumber & "', '" & lblLastVersion & "', '" & tmpRBLmt & "','" & tmpHospDays & "', 'FULL', " & _
                                                "'0', 'FULL', 'FULL', 'FULL', 'FULL', " & _
                                                "'FULL' , 'FULL', 'FULL' , 'FULL', 'FULL', " & _
                                                "'FULL' , 'FULL', 'FULL', 'FULL', 'FULL', " & _
                                                "'FULL', 'FULL','FULL' ,'" & tmpCancerLmt & "' , 'FULL' , " & _
                                                "'0', '0', '0', '0', '0' , " & _
                                                "'" & tmpMedRptLmt & "' , '0', '0', '0', '0' , " & _
                                                "'0'  ,'" & tmpGovTaxAct & "', '0', '0'"
                                    medixmasterconnWS.Execute strsqlCLM
                                    
                                    strsqlCLM = ""
                                    strsqlCLM = "EXEC Insert_WS_ClaimActual '" & VCaseNumber & "', '" & lblLastVersion & "', '" & tmpRBAct & "','0', " & _
                                                "'" & tmpHSSAct & "', '0', '0', '" & tmpPreHospAct & "', " & _
                                                "'" & tmpPhyVisitAct & "' , '0', '0' , '0', '0', " & _
                                                "'" & tmpPostAct & "', '0', '0', '0', " & _
                                                "'0', '0','0' ,'" & tmpCancerAct & "'," & _
                                                "'0', '0', '0', " & _
                                                "'" & tmpMedRptAct & "' , '0', '0', '0', '0' , " & _
                                                "'0'  ,'" & tmpGovTaxAct & "', '0', '0'"
                                    medixmasterconnWS.Execute strsqlCLM
    
                                    strsqlCLM = ""
                                    strsqlCLM = "EXEC Insert_WS_ClaimApprove '" & VCaseNumber & "', '" & lblLastVersion & "', '" & tmpRBApp & "','0', " & _
                                                "'" & tmpHSSApp & "', '0', '0', '" & tmpPreHospApp & "', " & _
                                                "'" & tmpPhyVisitApp & "' , '0', '0' , '0', '0', " & _
                                                "'" & tmpPostApp & "', '0', '0', '0', " & _
                                                "'0', '0','0' ,'" & tmpCancerApp & "'," & _
                                                "'0', '0', '0', " & _
                                                "'" & tmpMedRptApp & "' , '0', '0', '0', '0' , " & _
                                                "'0'  ,'" & tmpGovTaxApp & "', '0', '0'"
                                    medixmasterconnWS.Execute strsqlCLM
                                    
                                    strsqlCLM = ""
                                    strsqlCLM = "EXEC Insert_WS_ClaimHOS '" & VCaseNumber & "', '" & lblLastVersion & "', '0','0', " & _
                                                "'0', '0', '0', '0', " & _
                                                "'0' , '0', '0' , '0', '0', " & _
                                                "'0', '0', '0', '0', " & _
                                                "'0', '0','0' ,'0'," & _
                                                "'0', '0 ', '0', " & _
                                                "'0' , '0', '0', '0', '0' , " & _
                                                "'0'  , '0', '0', '0'"
                                    medixmasterconnWS.Execute strsqlCLM
                                    
                                    strsqlCLM = ""
                                    strsqlCLM = "EXEC Insert_WS_ClaimMBM '" & VCaseNumber & "', '" & lblLastVersion & "', '" & tmpRBApp & "','0', " & _
                                                "'" & tmpHSSApp & "', '0', '0', '" & tmpPreHospApp & "', " & _
                                                "'" & tmpPhyVisitApp & "' , '0', '0' , '0', '0', " & _
                                                "'" & tmpPostApp & "', '0', '0', '0', " & _
                                                "'0', '0','0' ,'" & tmpCancerApp & "'," & _
                                                "'0', '0', '0', " & _
                                                "'" & tmpMedRptApp & "' , '0', '0', '0', '0' , " & _
                                                "'0'  ,'" & tmpGovTaxApp & "', '0', '0'"
                                    medixmasterconnWS.Execute strsqlCLM
                                End If
                            End If
                        End If
                    End If
                    
                    If tmpRcrdType = "PT" Then
                        tmpCancerAct = 0
                        tmpCancerApp = 0
                        tmpCancerNotPay = 0
                        tmpPreHospAct = 0
                        tmpPreHospApp = 0
                        tmpPreHospNotpay = 0
                        tmpPostAct = 0
                        tmpPostApp = 0
                        tmpPostNotPay = 0
                        tmpRBAct = 0
                        tmpRBApp = 0
                        tmpRBNotPay = 0
                        tmpPhyVisitAct = 0
                        tmpPhyVisitApp = 0
                        tmpPhyVisitNotPay = 0
                        tmpMedRptAct = 0
                        tmpMedRptApp = 0
                        tmpMedRptNotPay = 0
                        tmpHSSAct = 0
                        tmpHSSApp = 0
                        tmpHSSNotPay = 0
                        
                    End If
                DoEvents
            Loop
        Close #1   ' Close file.
        
        End If

End Sub



