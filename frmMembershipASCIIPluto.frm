VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMembershipASCIIPluto 
   Caption         =   "Membership ASCII"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   7065
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   80
      TabIndex        =   5
      Top             =   240
      Width           =   6930
      Begin VB.ComboBox cboHLTCode 
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
         ItemData        =   "frmMembershipASCIIPluto.frx":0000
         Left            =   2205
         List            =   "frmMembershipASCIIPluto.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   4305
      End
      Begin VB.ComboBox cboPAYCode 
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
         Left            =   2205
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   680
         Width           =   4305
      End
      Begin MSMask.MaskEdBox txtMBMBordxDate 
         Height          =   285
         Left            =   2205
         TabIndex        =   2
         Top             =   960
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text1 
         Height          =   195
         Left            =   5400
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Health Type"
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
         Left            =   180
         TabIndex        =   11
         Top             =   450
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "File on Server"
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
         Left            =   180
         TabIndex        =   10
         Top             =   150
         Width           =   1680
      End
      Begin VB.Label txtPath 
         BackColor       =   &H00FFFFC0&
         Caption         =   "\\Svr-Doc\Membership\Bor_in\Temp\"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2205
         TabIndex        =   9
         Top             =   150
         Width           =   2850
      End
      Begin VB.Label Label17 
         Caption         =   "Bordx. Date (dd/mm/yyyy)"
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
         Left            =   180
         TabIndex        =   7
         Top             =   1000
         Width           =   2265
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
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   720
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   80
      TabIndex        =   13
      Top             =   1560
      Width           =   6930
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate"
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
         Left            =   5040
         TabIndex        =   3
         Top             =   180
         Width           =   840
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
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
         Left            =   5880
         TabIndex        =   4
         Top             =   180
         Width           =   960
      End
      Begin VB.Label lblTotalSupp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label33 
         Caption         =   "Total Supplementary"
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
         Index           =   6
         Left            =   2280
         TabIndex        =   16
         Top             =   225
         Width           =   1740
      End
      Begin VB.Label lblTotalPrin 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label33 
         Caption         =   "Total Principal"
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
         Index           =   9
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Membership ASCII"
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
      TabIndex        =   8
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmMembershipASCIIPluto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub combolist()
    cboHLTCode.AddItem "S - Sihat"
    cboHLTCode.AddItem "N - Non-Sihat"
End Sub

Private Sub cboHLTCode_Click()
    Dim PAY As New ADODB.Recordset
    Dim strsql As String
    cboPAYCode.Clear
    
    Dim strAppend As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    
    strAppend = strAppend & "And HLTCOde = '" & Mid(Trim(cboHLTCode), 1, 1) & "'"
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayorHLTType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 255, strAppend)
    cmd.Parameters.Append Prm1
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    
    cboPAYCode.Clear
    Do Until PAY.EOF
       cboPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
'    medixmasterconn.Close
 '   Set medixmasterconn = Nothing
End Sub

Private Sub cboPAYCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPAYCode.Text, cboPAYCode.SelStart) + Chr(KeyAscii) + Mid(cboPAYCode.Text, cboPAYCode.SelStart + cboPAYCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPAYCode.ListCount - 1
 
        If NewText = LCase(Left(cboPAYCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPAYCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPAYCode.Text = ValidValue
                cboPAYCode.SelStart = 0
                cboPAYCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub CmdGenerate_Click()
Dim tempPath As String
Dim rst As New ADODB.Recordset
Dim BATNo As String
Dim strAppend As String
Dim cntPrincipal As String
Dim cntSupplementary As String
Dim Header As String
Dim Trailer As String
Dim txtFilName As String
        cntPrincipal = 0
        cntSupplementary = 0
    If Len(Trim(cboHLTCode)) = 0 Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboHLTCode.SetFocus
    ElseIf Len(Trim(cboPAYCode)) = 0 Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPAYCode.SetFocus
    ElseIf Len(Trim(txtMBMBordxDate)) = 0 Then
        MsgBox "Please Enter Bordx Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMBordxDate.SetFocus
    ElseIf IsDate(Trim(txtMBMBordxDate)) = False Then
        MsgBox "Please Enter a Valid Bordx Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMBordxDate.SetFocus
    Else
        tempPath = "\\PLUTO\HISSaturn\Conn\MediConnectPluto.UDL"
        medixmasterconnPluto.Open "File Name=" & tempPath
        strAppend = ""
        strAppend = " Select * from MBMExcellUploadShelly where MBMBordxDate = '" & Format(CDate(txtMBMBordxDate), "MM/DD/YYYY") & "'" & _
                 " AND PAYCode = '" & Mid(cboPAYCode, 1, 2) & "'  ORDER BY No "
        'rst.Open strAppend, medixmasterconnPluto, adOpenForwardOnly, adLockOptimistic
        rst.Open strAppend, medixmasterconnPluto, adOpenKeyset, adLockOptimistic
        If rst.EOF Then
        'If rstCount!totalrecord = 0 Then
            MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
                'BATNo = GetCLMAsciiBatch(Mid(cboPAYCode, 1, 2), Date)
                'textField = Trim(Date)
                      'txtFilName = TxtBordxNoFr & "(M-MEDIX)" & ".txt"
                txtFilName = Trim(Mid(cboPAYCode, 1, 2) & Format(rst!MBMBordxDate, "DDMMYYYY")) & ".txt"
                Open txtPath & "\" & txtFilName For Output As #1
                
                Header = Trim("F" & Mid(cboPAYCode, 1, 2) & Mid(rst!MBMBordxDate, 1, 2) & Mid(rst!MBMBordxDate, 4, 2) & Mid(rst!MBMBordxDate, 7, 4) & l("", 4788))
                
                Print #1, Trim(Header)
            With rst
                Do Until rst.EOF
                    Text1 = ""
                    Text1 = l(!RELCode, 1)
                    Text1 = Text1 & l(IIf(Len(!MBMPolicyNo), !MBMPolicyNo, ""), 25)
                    Text1 = Text1 & l(IIf(Len(!MBMAgentCode), !MBMAgentCode, ""), 15)
                    Text1 = Text1 & l(IIf(Len(!MBMSalutation), !MBMSalutation, ""), 5)
                    Text1 = Text1 & l(IIf(Len(!MBMName), !MBMName, ""), 60)
                    Text1 = Text1 & l(IIf(Len(!MBMAddress1), !MBMAddress1, ""), 30)
                    Text1 = Text1 & l(IIf(Len(!MBMAddress2), !MBMAddress2, ""), 30)
                    Text1 = Text1 & l(IIf(Len(!MBMAddress3), !MBMAddress3, ""), 30)
                    Text1 = Text1 & l(IIf(Len(!CTYCode), !CTYCode, ""), 20)
                    Text1 = Text1 & l(IIf(Len(!MBMPostCode), !MBMPostCode, ""), 5)
                    Text1 = Text1 & l(IIf(Len(!STACode), !STACode, ""), 15)
                    Text1 = Text1 & l(IIf(Len(!MBMTelnoH), !MBMTelnoH, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMTelNoO), !MBMTelNoO, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMTelNoM), !MBMTelNoM, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMIcBcPp), !MBMIcBcPp, ""), 20)
                    Text1 = Text1 & l(Format(!MBMDOB, "DDMMYYYY"), 8)
                    Text1 = Text1 & l(IIf(Len(!MBMSex), !MBMSex, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!NATCode), !NATCode, ""), 10)
                    Text1 = Text1 & l(IIf(Len(!MBMWeight), !MBMWeight, ""), 3)
                    Text1 = Text1 & l(IIf(Len(!MBMHeight), !MBMHeight, ""), 3)
                    Text1 = Text1 & l(IIf(Len(!LGNCode), !LGNCode, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!RACCode), !RACCode, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!MBMOccupation), !MBMOccupation, ""), 30)
                    Text1 = Text1 & l(IIf(Len(!MBMWorkNature), !MBMWorkNature, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!MBMBlood), !MBMBlood, ""), 3)
                    Text1 = Text1 & l(IIf(Len(!MBMAllergic), !MBMAllergic, ""), 20)
                    Text1 = Text1 & l(Format(!MBMPayorEffDate, "DDMMYYYY"), 8)
                    Text1 = Text1 & l(Format(!MBMPayorEXPDate, "DDMMYYYY"), 8)
                    Text1 = Text1 & l(IIf(Len(!PLNCode), !PLNCode, ""), 4)
                    Text1 = Text1 & l(IIf(Len(!INSCode), !INSCode, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!MBMGroupCompany), !MBMGroupCompany, ""), 40)
                    Text1 = Text1 & l(IIf(Len(!MBMRenewFlag), !MBMRenewFlag, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!MBMExclusion), !MBMExclusion, ""), 2500)
                    Text1 = Text1 & l(IIf(Len(!MBMSmoking), !MBMSmoking, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!MBMAlcohol), !MBMAlcohol, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!MBMMaritalStatus), !MBMMaritalStatus, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!BRNCode), !BRNCode, ""), 15)
                    Text1 = Text1 & l(IIf(Len(!MBMPrevPolNo), !MBMPrevPolNo, ""), 25)
                    Text1 = Text1 & l(IIf(Len(!MBMPrevMemNo), !MBMPrevMemNo, ""), 18)
                    Text1 = Text1 & l(IIf(Len(!MBMEmployeeNo), !MBMEmployeeNo, ""), 15)
                    Text1 = Text1 & l(IIf(Len(!MBMJoinedDate), Format(!MBMJoinedDate, "DDMMYYYY"), ""), 8)
                    Text1 = Text1 & l(IIf(Len(!Clinical), !Clinical, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!ClinicName1), !ClinicName1, ""), 50)
                    Text1 = Text1 & l(IIf(Len(!ClinicName2), !ClinicName2, ""), 50)
                    Text1 = Text1 & l(IIf(Len(!ClinicName3), !ClinicName3, ""), 50)
                    Text1 = Text1 & l(IIf(Len(!ClinicName4), !ClinicName4, ""), 50)
                    Text1 = Text1 & l(IIf(Len(!MBMDepartment), !MBMDepartment, ""), 30)
                    Text1 = Text1 & l(IIf(Len(!MBMGRPCompanyAdd1), !MBMGRPCompanyAdd1, ""), 30)
                    Text1 = Text1 & l(IIf(Len(!MBMGRPCompanyAdd2), !MBMGRPCompanyAdd2, ""), 30)
                    Text1 = Text1 & l(IIf(Len(!MBMGRPCompanyAdd3), !MBMGRPCompanyAdd3, ""), 30)
                    Text1 = Text1 & l(IIf(Len(!MBMGRPCompanyAdd4), !MBMGRPCompanyAdd4, ""), 30)
                    Text1 = Text1 & l(IIf(Len(!MBMPrevTakeOverPlnCode), !MBMPrevTakeOverPlnCode, ""), 4)
                    Text1 = Text1 & l(IIf(Len(!MBMInstallment), !MBMInstallment, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!MBMPolSubNo1), !MBMPolSubNo1, ""), 10)
                    Text1 = Text1 & l(IIf(Len(!MBMPOLSubNo2), !MBMPOLSubNo2, ""), 10)
                    Text1 = Text1 & l(IIf(Len(!MBMPOLIssuedDate), Format(!MBMPOLIssuedDate, "ddmmyyyy"), ""), 8)
                    Text1 = Text1 & l(IIf(Len(!MBMRiskNo), !MBMRiskNo, ""), 4)
                    Text1 = Text1 & l(IIf(Len(!MBMTransNo), !MBMTransNo, ""), 5)
                    Text1 = Text1 & l(IIf(Len(!MBMPayNo), !MBMPayNo, ""), 8)
                    Text1 = Text1 & l(IIf(Len(!MBMPAYSeqNo), !MBMPAYSeqNo, ""), 2)
                    Text1 = Text1 & l(IIf(Len(!MBMClientNo), !MBMClientNo, ""), 8)
                    Text1 = Text1 & l(IIf(Len(!MBMClientID), !MBMClientID, ""), 2)
                    Text1 = Text1 & l(IIf(Len(!MBMPrevPLNCode), !MBMPrevPLNCode, ""), 15)
                    Text1 = Text1 & l(IIf(Len(!MBMRBAmt), !MBMRBAmt, ""), 8)
                    Text1 = Text1 & l(IIf(Len(!MBMPrevInsCom), !MBMPrevInsCom, ""), 60)
                    Text1 = Text1 & l(IIf(Len(!MBMPAYRemarks), !MBMPAYRemarks, ""), 60)
                    Text1 = Text1 & l(IIf(Len(!MBMCLMNonPayFr), Format(!MBMCLMNonPayFr, "ddmmyyyy"), ""), 8)
                    Text1 = Text1 & l(IIf(Len(!MBMCLMNonPayTo), Format(!MBMCLMNonPayTo, "ddmmyyyy"), ""), 8)
                    Text1 = Text1 & l(IIf(Len(!MBMUpgradedPLN), !MBMUpgradedPLN, ""), 4)
                    Text1 = Text1 & l(IIf(Len(!MBMNewEffDate), Format(!MBMNewEffDate, "ddmmyyyy"), ""), 8)
                    Text1 = Text1 & l(IIf(Len(!MBMEndorsementNo), !MBMEndorsementNo, ""), 20)
                    Text1 = Text1 & l(IIf(Len(!MBMAppName), !MBMAppName, ""), 60)
                    Text1 = Text1 & l(IIf(Len(!MBMPayorPremNet), !MBMPayorPremNet, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMPayorPremGross), !MBMPayorPremGross, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMPayorMCONet), !MBMPayorMCONet, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMPayorMCOGross), !MBMPayorMCOGross, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMFamPolicyDisc), !MBMFamPolicyDisc, ""), 5)
                    Text1 = Text1 & l(IIf(Len(!MBMRenewDisc), !MBMRenewDisc, ""), 5)
                    Text1 = Text1 & l(IIf(Len(!MBMIcBcPp2nd), !MBMIcBcPp2nd, ""), 20)
                    Text1 = Text1 & l(IIf(Len(!MBMRnBStatus), !MBMRnBStatus, ""), 2)
                    Text1 = Text1 & l(IIf(Len(!MBMLoadPrem), !MBMLoadPrem, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMFamDiscAmt), !MBMFamDiscAmt, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMRenewDiscAmt), !MBMRenewDiscAmt, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMBlackListed), !MBMBlackListed, ""), 5)
                    Text1 = Text1 & l(IIf(Len(!MBMBankACNo), !MBMBankACNo, ""), 25)
                    Text1 = Text1 & l(IIf(Len(!MBMUnderwritingExcess), !MBMUnderwritingExcess, ""), 12)
                    Text1 = Text1 & l(IIf(Len(!MBMGuaRenewal), !MBMGuaRenewal, ""), 5)
                    Text1 = Text1 & l("", 2861)
                    Text1 = Text1 & l(IIf(Len(!MBMCLNPlanCode), !MBMCLNPlanCode, ""), 4)
                    Text1 = Text1 & l(IIf(Len(!MBMCLNInsType), !MBMCLNInsType, ""), 1)
                    Text1 = Text1 & l(IIf(Len(!MBMCLNEffDate), Format(!MBMCLNEffDate, "ddmmyyyy"), ""), 8)
                    Text1 = Text1 & l(IIf(Len(!MBMCLNExpDate), Format(!MBMCLNExpDate, "ddmmyyyy"), ""), 8)
                    Text1 = Text1 & l(IIf(Len(!MBMCLNExclusion), !MBMCLNExclusion, ""), 2500)
                    Text1 = Text1 & l("", 20)
                    Text1 = Text1 & l("", 752)
                    If Trim(Mid(cboPAYCode, 6, 14)) = Trim("TAKAFUL IKHLAS") Then
                        Text1 = Text1 & l(IIf(Len(!MBMTINumber), !MBMTINumber, ""), 20)
                        Text1 = Text1 & l(IIf(Len(!MBMTICoverID), !MBMTICoverID, ""), 5)
                    End If
                    If Trim(!RELCode) = "P" Then
                        cntPrincipal = CDbl(cntPrincipal) + 1
                    Else
                        cntSupplementary = CDbl(cntSupplementary) + 1
                    End If
                    Text1 = Text1 & l(IIf(Len(!Subsidiary), !Subsidiary, ""), 50)
                    Text1 = Text1 & l(IIf(Len(!costcentre), !costcentre, ""), 50)
                    Text1 = Text1 & l(IIf(Len(!Branch), !Branch, ""), 50)
                    Print #1, Trim(Text1)
                rst.MoveNext
                Loop
            End With
                Trailer = "T" & l(cntPrincipal, 10) & l(cntSupplementary, 10) & l("", 4778)
                    Print #1, Trailer
                    cmdGenerate.Enabled = False
                    Close #1
                    lblTotalPrin = cntPrincipal
                    lblTotalSupp = cntSupplementary
                    MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                    cmdGenerate.Enabled = True
             End If
            rst.Close
            Set rst = Nothing
            medixmasterconnPluto.Close
            Set medixmasterconnPluto = Nothing
        End If

End Sub

Private Sub Form_Load()
    Call combolist
End Sub


Private Sub txtMBMBordxDate_LostFocus()
    If IsDate(txtMBMBordxDate) Then
    Else
        If txtMBMBordxDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMBMBordxDate.SetFocus
        End If
    End If

End Sub

