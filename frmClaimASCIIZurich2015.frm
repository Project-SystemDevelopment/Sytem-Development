VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmClaimASCIIZurich2015 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   7845
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      Caption         =   "Rejected Cases"
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   600
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Accepted Cases"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   7095
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         TabIndex        =   20
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         TabIndex        =   19
         Top             =   240
         Width           =   825
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   18
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7095
      Begin VB.ComboBox cboBordxWSStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4980
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1920
         Width           =   1995
      End
      Begin VB.ComboBox CboExportStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4980
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Width           =   1995
      End
      Begin VB.ComboBox cboPAYCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   5535
      End
      Begin MSMask.MaskEdBox txtReceiptDateFr 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1560
         Width           =   1065
         _ExtentX        =   1879
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
      Begin MSMask.MaskEdBox txtReceiptDateTo 
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   1560
         Width           =   1065
         _ExtentX        =   1879
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
      Begin VB.Label txtFilName 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   840
         Width           =   1905
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2600
         TabIndex        =   15
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "Endt Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Export "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   13
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label txtLocalFileLocation 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1440
         TabIndex        =   12
         Top             =   540
         Width           =   1905
      End
      Begin VB.Label Label26 
         Caption         =   "Case Reg Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1545
         Width           =   1155
      End
      Begin VB.Label Label24 
         Caption         =   "File on Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label txtPath 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label28 
         Caption         =   "File on PC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   615
         Width           =   765
      End
      Begin VB.Label Label18 
         Caption         =   "Payor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1185
         Width           =   915
      End
      Begin VB.Label Label25 
         Caption         =   "Export File Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   1245
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Generate Claims ASCII (Zurich 2015)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmClaimASCIIZurich2015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Exportsql As String
Dim textField As String
Dim TotalRecord As String
Dim Header As String
Dim DataString As String
Dim tmpActGST As Double
Dim tmpAppGST As Double
Dim CLMBordx As New ADODB.Recordset

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGenerate_Click()
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim RecordSaved

    TotalRecord = "0"
    
    If cboPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPAYCode.SetFocus
    ElseIf Option1.Value = True And txtReceiptDateFr = "__/__/____" Then
        MsgBox "Please Enter Bordx Date Fr ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDateFr.SetFocus
    ElseIf Option1.Value = True And txtReceiptDateTo = "__/__/____" Then
        MsgBox "Please Enter Bordx Date To ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDateTo.SetFocus
    ElseIf Option2.Value = True And txtReceiptDateFr = "__/__/____" Then
        MsgBox "Please Enter Case Reg. Date Fr ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDateFr.SetFocus
    ElseIf Option2.Value = True And txtReceiptDateTo = "__/__/____" Then
        MsgBox "Please Enter Case Reg. Date To ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDateTo.SetFocus
        
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation, DialogBoxTitle)
        If RecordSaved = vbYes Then
            Screen.MousePointer = vbHourglass
                    Call PrincipalASCII
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub PrincipalASCII()
Dim strsql As String
Dim BordxNo As String
Dim BordxDateFr As String
Dim BordxDateTo As String
Dim PAYCode As String
Dim ii As Integer
Dim sql As String
Dim ClaimNo As String
Dim VersionNo As String
Dim STRBordxArc As String
Dim BordxArc As New ADODB.Recordset
Dim BATNo As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim BenAppAmt As Double
Dim BordxDate As String
Dim Trailer As String
Dim PatientAge As String
Dim tmpRBTax As Double
Dim tmpICUtax As Double
Dim tmpLodgerTax As Double
Dim tmpMRI As Double
Dim tmpOthers As Double
Dim tmpMedicalRpt As Double
Dim tmpHSS As Double
Dim tmpCLMNSPreHos As Double
Dim tmpCLMPreHSAmt As Double
Dim tmpCLMPreHosp As Double
Dim tmpCLMMPreHSAmt As Double
Dim tmpCopay As Double
Dim tmpNoCoPay As Double
Dim tmpCLMTax As Double
Dim tmpCLMMonthly As Double
Dim tmpCLMIndKidneyDial As Double
Dim tmpCLMExGratiaAmt As Double
Dim tmpLmtExceeded As Double
Dim tmoCoPayRB As Double
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlsheet As Excel.WORKSHEET
Dim Rcounter As Double
Dim row  As Double
        Rcounter = 0
        row = 0
        BordxDateFr = txtReceiptDateFr
        BordxDateTo = txtReceiptDateTo
        PAYCode = Mid(cboPAYCode, 1, 2)
        strAppend = ""
        If Option1.Value = True Then
            If Mid(CboExportStatus, 1, 1) = "N" Then
                strAppend = " AND BordxDate >= '" & Format(CDate(BordxDateFr), "yyyy/mm/dd") & "'" & _
                                    " AND BordxDate <= '" & Format(CDate(BordxDateTo), "yyyy/mm/dd") & "'" & _
                                    " AND PAYCode = '" & PAYCode & "' "
            Else
                strAppend = " AND BordxDate >= '" & Format(CDate(BordxDateFr), "MM/DD/YYYY") & "'" & _
                                    " AND BordxDate <= '" & Format(CDate(BordxDateTo), "MM/DD/YYYY") & "'" & _
                                    " AND PAYCode = '" & PAYCode & "' AND CLMBordxWSASCII = 'Y' "
            End If
        Else
                strAppend = " AND CaseRegDate >= '" & Format(CDate(BordxDateFr), "yyyy/mm/dd") & "'" & _
                                    " AND CaseRegDate <= '" & Format(CDate(BordxDateTo), "yyyy/mm/dd") & "'" & _
                                    " AND PAYCode = '" & PAYCode & "' "
        End If
        If Option1.Value = True Then
            strsql = "SELECT * From View_Claims_ASCII1 where 0 = 0 "
        Else
            strsql = "select * from View_ClaimsRejected_ASCII where 0 = 0"
        End If
        If Len(Trim(strAppend)) Then
            strsql = strsql + strAppend + "ORDER BY BordxRefNo"
        End If
        CLMBordx.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If CLMBordx.EOF Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
            BATNo = GetCLMAsciiMedBatch(Mid(cboPAYCode, 1, 2), Date)
            textField = Trim(Date)
                  
            txtFilName = Trim("MX" & Format(Date, "YYMMDD"))
            Open txtLocalFileLocation & "\" & txtFilName & "(" & BATNo & ")" For Output As #1
            Header = Trim("F" & "MX" & BATNo & Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 7, 4) & l("", 174))
            Print #1, Trim(Header)
        Do Until CLMBordx.EOF
            'DataString = ""
            BenAppAmt = 0
            ClaimNo = CLMBordx!CasNumber
            VersionNo = IIf(Len(CLMBordx!claimversion), CLMBordx!claimversion, "")
            BordxNo = IIf(Len(CLMBordx!BordxRefNo), CLMBordx!BordxRefNo, "")
            BordxDate = IIf(Len(CLMBordx!BordxDate), CLMBordx!BordxDate, "")
            Rcounter = Rcounter + 1
            row = row + 1
            DataString = l(BordxNo, 15)                            '15  Bordx Number
            If Option1.Value = True Then
                DataString = DataString & l(Format(BordxDate, "ddmmyyyy"), 8)      ' Len 8 Bordx Date
            Else
                DataString = DataString & l(Format(CLMBordx!CaseRegDate, "ddmmyyyy"), 8)   '23  Bordx Date
            End If
            
            Call GSTCount
            
            DataString = DataString & l(CLMBordx!PAYCode, 2) '25    Payor (2)
            
            DataString = DataString & l(CLMBordx!MBMNumber, 15) '40   MBMNumber (15)
            
            DataString = DataString & l("00", 2) '42       MBMcoverID (2)
            
            DataString = DataString & l(CLMBordx!MBMPolicyNo, 25)  '67    Policy Number (25)
            
            If Len(CLMBordx!MBMPolSubNo1) Then
                DataString = DataString & l(CLMBordx!MBMPolSubNo1, 25)  '92   Policy Sub Number1 (25)
            Else
                DataString = DataString & l("", 25)                      '92
            End If

            If Len(CLMBordx!MBMPOLSubNo2) Then
                DataString = DataString & l(CLMBordx!MBMPOLSubNo2, 25)   '117  Policy Sub Number 2 (25)
            Else
                DataString = DataString & l("", 25)                       '117
            End If
            DataString = DataString & l(Format(CLMBordx!MBMPayorEffDate, "ddmmyyyy"), 8) '125 Policy Eff Date (8)
            DataString = DataString & l(Format(CLMBordx!MBMPayorEXPDate, "ddmmyyyy"), 8) '133 Policy Exp Date (8)
            If Len(CLMBordx!MBMAgentCode) Then
                DataString = DataString & l(CLMBordx!MBMAgentCode, 15)       '148   Agent Code (15)
            Else
                DataString = DataString & l("", 15)                          '148
            End If
            DataString = DataString & l(CLMBordx!MBMName, 80)                '228    Principle Name  (80)
            DataString = DataString & l(CLMBordx!MBMIcBcPp, 20)              '248    Principle IC(20)
            DataString = DataString & l(CLMBordx!MBMSex, 1)                  '249   Sex(1)
            If Len(CLMBordx!RACCode) Then
                DataString = DataString & l(CLMBordx!RACCode, 1)            'Principle Race (1)
            Else
                DataString = DataString & l("", 1)                            '250  Principle Race (1)
            End If
            If Len(CLMBordx!MBMDOB) Then
                DataString = DataString & l(Format(CLMBordx!MBMDOB, "ddmmyyyy"), 8)    '258 Principle DOB (8)
            Else
                DataString = DataString & l("", 8)                     '258   Principle DOB (8)
            End If
            If CLMBordx!MBMRenewFlag <> "" Then
                DataString = DataString & l(CLMBordx!MBMRenewFlag, 1)               '259 Renewal (1)
            Else
                DataString = DataString & l("", 1)                                  '259 Renewal (1)
            End If
            If Len(CLMBordx!MBMTakeover) Then
                DataString = DataString & l(CLMBordx!MBMTakeover, 1)                ' 260 Takeover Code (1)
            Else
                DataString = DataString & l("", 1)                                  '260 Takeover Code (1)
            End If
            DataString = DataString & l(CLMBordx!INSCode, 2)     '262  Insurance Type (2)
            DataString = DataString & l(IIf(Len(CLMBordx!PLNPayorPlan), CLMBordx!PLNPayorPlan, ""), 7)  '269  Plan Code (7)
            If Len(CLMBordx!MBMGroupCompany) Then
                DataString = DataString & l(CLMBordx!MBMGroupCompany, 47)             '316  Group Company (47)
            Else
                DataString = DataString & l("", 47)
            End If
            DataString = DataString & l("", 30)               '336    Branch (30)
            If Len(CLMBordx!AGECode) Then
                DataString = DataString & l(CLMBordx!AGECode, 2)  '338  Principle Age Band (2)
            Else
                DataString = DataString & l("", 2)
            End If
            If Len(CLMBordx!PatientName) Then
                DataString = DataString & l(CLMBordx!PatientName, 80)    '418  Patient Name (80)
            Else
                DataString = DataString & l("", 80)
            End If
            If Len(CLMBordx!PatientIC) Then
                DataString = DataString & l(CLMBordx!PatientIC, 20)    'Patient IC (20)
            Else
                DataString = DataString & l("", 20)    'Patient IC (20)
            End If
            If Len(CLMBordx!PatientSex) Then
                DataString = DataString & l(CLMBordx!PatientSex, 1)   'Patient Sex (1)
            Else
                DataString = DataString & l("", 1)
            End If
            If Len(CLMBordx!PatientAgeband) Then
                DataString = DataString & l(CLMBordx!PatientAgeband, 2)  'Patient Age Band (2)
            Else
                DataString = DataString & l("", 2)
            End If
            If Len(CLMBordx!PatientDOB) Then
                DataString = DataString & l(Format(CLMBordx!PatientDOB, "ddmmyyyy"), 8)   'Patient DOB (8)
            Else
                DataString = DataString & l("", 8)
            End If
                       
            If Len(CLMBordx!HosName) Then
                DataString = DataString & l(CLMBordx!HosName, 80)     ' Hospital  (80)
            Else
                DataString = DataString & l("", 80)
            End If
            If CLMBordx!FinalDiag1 <> "" Then
                DataString = DataString & l(CLMBordx!FinalDiag1, 150)    ' Final Diagnosis1 (150)
            Else
                DataString = DataString & l("", 150)
            End If
            If CLMBordx!FinalDiag2 <> "" Then
                DataString = DataString & l(CLMBordx!FinalDiag2, 150)    ' Final Diagnosis2 (150)
            Else
                DataString = DataString & l("", 150)
            End If
            If CLMBordx!FinalDiag3 <> "" Then
                DataString = DataString & l(CLMBordx!FinalDiag3, 150)     ' Final Diagnosis3 (150)
            Else
                DataString = DataString & l("", 150)
            End If
            If CLMBordx!CASDateAdmitted <> "" Then
                DataString = DataString & l(Format(CLMBordx!CASDateAdmitted, "ddmmyyyy"), 8)  ' DOA (8)
            Else
                DataString = DataString & l("", 8)
            End If
            If CLMBordx!CASDischargeDate <> "" Then
                DataString = DataString & l(Format(CLMBordx!CASDischargeDate, "ddmmyyyy"), 8) '995  DOD (8)
            Else
                DataString = DataString & l("", 8)
            End If
            If CLMBordx!CASAdmissionReason <> "" Then
                DataString = DataString & l(CLMBordx!CASAdmissionReason, 200)  '1195 Admission Reasons (200)
            Else
                DataString = DataString & l("", 200)          'Admission Reasons (200)
            End If
            If CLMBordx!CASAttendDoctor1 <> "" Then
                DataString = DataString & l(CLMBordx!CASAttendDoctor1, 80)   '1275   Attending Doc1 80
            Else
                DataString = DataString & l("", 80)
            End If
            If CLMBordx!CASAttendDoctor2 <> "" Then
                DataString = DataString & l(CLMBordx!CASAttendDoctor2, 80)   '1355   Attending Doc 2 80
            Else
                DataString = DataString & l("", 80)
            End If
            If CLMBordx!CASAttendDoctor3 <> "" Then
                DataString = DataString & l(CLMBordx!CASAttendDoctor3, 80)    '1435  Attending Doc 3 80
            Else
                DataString = DataString & l("", 80)
            End If
            If CLMBordx!CASAttendDoctor4 <> "" Then
                DataString = DataString & l(CLMBordx!CASAttendDoctor4, 80)    '1515  Attending Doc 4 80
            Else
                DataString = DataString & l("", 80)
            End If
            If CLMBordx!CASAttendDoctor5 <> "" Then
                DataString = DataString & l(CLMBordx!CASAttendDoctor5, 80)    '1595  Attending Doc 5 80
            Else
                DataString = DataString & l("", 80)
            End If
            If CLMBordx!ClaimNatureCode <> "" Then
                DataString = DataString & l(CLMBordx!ClaimNatureCode, 5)       '1600  Nature of Claim 5
            Else
                DataString = DataString & l("", 5)
            End If
            If CLMBordx!ModalityCode <> "" Then
                DataString = DataString & l(CLMBordx!ModalityCode, 1)            '1601 Treatment Mode 1
            Else
                DataString = DataString & l("", 1)
            End If
            If CLMBordx!CASStayType <> "" Then
                DataString = DataString & l(CLMBordx!CASStayType, 1)              '1602 Stay Type 1
            Else
                DataString = DataString & l("", 1)
            End If
            If CLMBordx!InvPayTo = "H" Then
                DataString = DataString & l("C", 1)                 '1603   cash type 1
            Else
                DataString = DataString & l("R", 1)
            End If
            DataString = DataString & l(CLMBordx!MBMPatientCovID, 2)        '1605  Patient Cov ID 2
            DataString = DataString & l(CLMBordx!CasNumber, 15)             '1630  Case Number 15
            If Len(CLMBordx!claimversion) Then
                DataString = DataString & l(CLMBordx!claimversion, 3)       '1633   Claim Version 3
            Else
                DataString = DataString & l("", 3)
            End If
            DataString = DataString & l(Format(CLMBordx!ClaimOpenDate, "ddmmyyyy"), 8)      '1641  Claim Reg Date 8
            'Bill amount 11
            If IsNumeric(CLMBordx!CLMTotalActual) Then
                If IsNumeric(CLMBordx!CLMTotalActualGST) Then
                    DataString = DataString & l(Format(CDbl(CLMBordx!CLMTotalActual) + CDbl(CLMBordx!CLMTotalActualGST), "#,##0.00"), 11)   '1652 Bill amount
                Else
                    DataString = DataString & l(Format(CDbl(CLMBordx!CLMTotalActual), "#,##0.00"), 11)
                End If
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'Approve amount 11
            If IsNumeric(CLMBordx!CLMBordxAmt) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMBordxAmt), "#,##0.00"), 11)     '1663 Approve Amount 11
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'Final Approve amount 11
            If IsNumeric(CLMBordx!CLMBordxAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMBordxAmt, "#,##0.00"), 11)           '1674 Final Approve Amount11
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'Limit exceeded 11
            If IsNumeric(CLMBordx!LimitExc) Then
                DataString = DataString & l(Format(CLMBordx!LimitExc, "#,##0.00"), 11)                  '1685 Limit exceded 11
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'Previous Balance 11
            If IsNumeric(CLMBordx!CLMPreviousBalance) Then
                DataString = DataString & l(Format(CLMBordx!CLMPreviousBalance, "#,##0.00"), 11)        '1696 Previos Balance 11
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'UndExcess 11
            If IsNumeric(CLMBordx!CLMPreviousUndExcess) Then
                DataString = DataString & l(Format(CLMBordx!CLMPreviousUndExcess, "#,##0.00"), 11)          '1707 UndExcess11
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'Exgratia Amount 11
            If IsNumeric(CLMBordx!CLMExGratiaAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMExGratiaAmt, "#,##0.00"), 11)        '1718 Exgratia Amount 11
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'Bordx Amount 11
            If IsNumeric(CLMBordx!CLMBordxAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMBordxAmt, "#,##0.00"), 11)           '1729 Bordx amount
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'Pay Group Code 11
            DataString = DataString & l("", 11)         '1740 Pay Group Code 11
            
            'If IsNumeric(CLMBordx!CLMCoPayRB) Then
            '    tmoCoPayRB = CLMBordx!CLMCoPayRB * -1
            'Else
            '    tmoCoPayRB = 0
            'End If
            
            'CLMCoPayRB 11
            If IsNumeric(CLMBordx!CLMCoPayRB) Then
                DataString = DataString & l(Format(tmoCoPayRB, "#,##0.00"), 11)     '1751 CLMCoPayRB 11
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            DataString = DataString & l(Format(CLMBordx!CLMExcess, "#,##0.00"), 11)     '1762 Total Excess 11
            DataString = DataString & l(Format(CLMBordx!INVPaidByMem, "#,##0.00"), 11)  '1773 Excess Coll By Hosp 11
            DataString = DataString & l(Format(CLMBordx!CLMHosPayColl, "#,##0.00"), 11)  '1784 CLM HOS Pay Coll 11
            'Limit
            DataString = DataString & l(Format(CLMBordx!CLMBoardAmtLmt, "#,##0.00"), 11)    '1795 Limit Board Amt
            DataString = DataString & l(Format(CLMBordx!CLMICUAmtLmt, "#,##0.00"), 11)      '1806 Limit ICU amount
            DataString = DataString & l(Format(CLMBordx!CLMServicesLmt, "#,##0.00"), 11)    '1817 Limit HSS
            DataString = DataString & l(Format(CLMBordx!CLMTheatreLmt, "#,##0.00"), 11)     '1828 Limit Theatre
            DataString = DataString & l(Format(CLMBordx!CLMPreHSAmtLmt, "#,##0.00"), 11)     '1839 Lmt Pre-Hosp
            DataString = DataString & l(Format(CLMBordx!CLMSurgicalAmtLmt, "#,##0.00"), 11)   '1850 Lmt  surgical
            DataString = DataString & l(Format(CLMBordx!CLMAnaLmt, "#,##0.00"), 11)             '1861 Lmt Anae
            DataString = DataString & l(Format(CLMBordx!CLMNSPreHosLmt, "#,##0.00"), 11)        '1872 Lmt NS Pres-Hosp
            DataString = DataString & l(Format(CLMBordx!CLMPhysicianLmt, "#,##0.00"), 11)       '1883 Lmt Physician
            DataString = DataString & l(Format(CLMBordx!CLMPostLmt, "#,##0.00"), 11)            '1894  Lmt Post
            'If Trim(Mid(cboPAYCode, 1, 2)) = Trim("MM") Then
            '    DataString = DataString & l(Format(0, "#,##0.00"), 11)       '2016
            'Else
            '    DataString = DataString & l(Format(CLMBordx!CLMEmergencyLmt, "#,##0.00"), 11)       '2005
            'End If
            DataString = DataString & l(Format(CLMBordx!CLMEmergencyLmt, "#,##0.00"), 11)       '2005   Lmt Emergency
            DataString = DataString & l(Format(CLMBordx!CLMAmbulanceLmt, "#,##0.00"), 11)       '2016   Lmt Ambulance
            DataString = DataString & l(Format(CLMBordx!CLMOutpatientLmt, "#,##0.00"), 11)      '2027   Lmt Outpatient
            DataString = DataString & l(Format(CLMBordx!CLMMonthlyLmt, "#,##0.00"), 11)         '2038   Lmt Kidney Dialysis Amt
            DataString = DataString & l(Format(CLMBordx!CLMOrganLmt, "#,##0.00"), 11)           '2049   Lmt Organ Trans
            DataString = DataString & l(Format(CLMBordx!CLMMedicalLmt, "#,##0.00"), 11)         '2060   Lmt Medical Rpt
            DataString = DataString & l(Format(CLMBordx!CLMPhoneLmt, "#,##0.00"), 11)           '2071   Lmt Phone
            DataString = DataString & l(Format(CLMBordx!CLMGuardianAmtLmt, "#,##0.00"), 11)     '2082   Lmt Gurdian
            DataString = DataString & l(Format(CLMBordx!CLMAdminLmt, "#,##0.00"), 11)           '2093   Lmt admin
            DataString = DataString & l(Format(CLMBordx!CLMTaxLmt, "#,##0.00"), 11)             '2104   Lmt Tax
            If Trim(Mid(cboPAYCode, 1, 2)) = Trim("MM") Then
                DataString = DataString & l(Format(0, "#,##0.00"), 11)                          '2016   Lmt Others1
            Else
                DataString = DataString & l(Format(CLMBordx!CLMOthers1Lmt, "#,##0.00"), 11)     '2016   Lmt Others1
            End If
            DataString = DataString & l(Format(CLMBordx!CLMOthers2Lmt, "#,##0.00"), 11)         '2126   Lmt Others2
            DataString = DataString & l(Format(CLMBordx!CLMCashAmtLmt, "#,##0.00"), 11)         '2137   Lmt Cash
            DataString = DataString & l(Format(CLMBordx!CLMBoardTaxAmtLmt, "#,##0.00"), 11)     '2148   Lmt Board tax
            DataString = DataString & l(Format(CLMBordx!CLMICUTaxAmtLmt, "#,##0.00"), 11)       '2159   Lmt ICU tax
            DataString = DataString & l(Format(CLMBordx!CLMLodgerTaxAmtLmt, "#,##0.00"), 11)    '2170   Lmt Lodger Tax
            
            'If Trim(Mid(cboPAYCode, 1, 2)) = Trim("MM") Then
            '    DataString = DataString & l(Format(CLMBordx!CLMEmergencyLmt, "#,##0.00"), 11)       '2181
            'Else
            
            'End If
            DataString = DataString & l(Format(CLMBordx!CLMBoardDaysLmt, "#,##0.00"), 11)           '2181 Lmt Board Days
            DataString = DataString & l(Format(CLMBordx!CLMICUDaysLmt, "#,##0.00"), 11)             '2192 Lmt ICU Days
            DataString = DataString & l(Format(CLMBordx!CLMCashDaysLmt, "#,##0.00"), 11)            '2203 Lmt Cash Days
            DataString = DataString & l(Format(CLMBordx!CLMMRILmt, "#,##0.00"), 11)                 '2214 Lmt MRI amt
            DataString = DataString & l(Format(CLMBordx!CLMMPreHSAmtLmt, "#,##0.00"), 11)           '2225 Lmt Sur Pre Hosp Amt
            DataString = DataString & l(Format(CLMBordx!CLMCOPreHSSumLmt, "#,##0.00"), 11)          '2236 Lmt NS Pre-Hosp Sum Amt
            DataString = DataString & l(Format(CLMBordx!CLMCOMPreHSSumLmt, "#,##0.00"), 11)         '2247 Lmt Surg Pre-Hops Sum amt
            DataString = DataString & l(Format("", "#,##0.00"), 11)                                 '2258 Lmt AccPost Amt
            DataString = DataString & l(Format(CLMBordx!CLMPreHospLmt, "#,##0.00"), 11)             '2269 Lmt Pre-Hosp Amt
            If Trim(Mid(cboPAYCode, 1, 2)) = Trim("MM") Then
                DataString = DataString & l(Format(CLMBordx!CLMOthers1Lmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPayLmt, "#,##0.00"), 11)  '2280 Lmt others Not Co-Pay
            End If
            DataString = DataString & l(Format(CLMBordx!CLMHomeNCAmtLmt, "#,##0.00"), 11)           '2291 Lmt CLM Home NC Amt
            
            'Actual
            'Act Board Amount
            If IsNumeric(CLMBordx!CLMBoardAmtAct) Then
                If IsNumeric(CLMBordx!CLMBoardAmtActGst) Then
                    DataString = DataString & l(Format(CDbl(CLMBordx!CLMBoardAmtAct) + CDbl(CLMBordx!CLMBoardAmtActGst), "#,##0.00"), 11)       '2302  Act Board Amount
                Else
                    DataString = DataString & l(Format(CLMBordx!CLMBoardAmtAct, "#,##0.00"), 11)
                End If
            Else
                DataString = DataString & l(Format(CLMBordx!CLMBoardAmtAct, "#,##0.00"), 11)
            End If
            
            'Act ICU Amount
            If IsNumeric(CLMBordx!CLMICUAmtAct) And IsNumeric(CLMBordx!CLMICUAmtActGST) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMICUAmtAct) + CDbl(CLMBordx!CLMICUAmtActGST), "#,##0.00"), 11)       '2313 Act ICU Amount
            Else
                DataString = DataString & l(Format(CLMBordx!CLMICUAmtAct, "#,##0.00"), 11)
            End If
                
            'Act HSS amt
            If IsNumeric(CLMBordx!CLMServicesAct) Then
                If IsNumeric(CLMBordx!CLMServicesActGST) Then
                    DataString = DataString & l(Format(CDbl(CLMBordx!CLMServicesAct) + CDbl(CLMBordx!CLMServicesActGST), "#,##0.00"), 11)       '2324 Act HSS amt
                Else
                    DataString = DataString & l(Format(CLMBordx!CLMServicesAct, "#,##0.00"), 11)
                End If
            Else
                DataString = DataString & l(Format(CLMBordx!CLMServicesAct, "#,##0.00"), 11)
            End If
            
            'Act Theatre Amt
            If IsNumeric(CLMBordx!CLMTheatreAct) And IsNumeric(CLMBordx!CLMTheatreActGST) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMTheatreAct) + CDbl(CLMBordx!CLMTheatreActGST), "#,##0.00"), 11)             '2335 Act Theatre Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMTheatreAct, "#,##0.00"), 11)
            End If
            
            'Act Pre-Hosp Amt
            If IsNumeric(CLMBordx!CLMPreHSAmtAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMPreHSAmtAct) + CDbl(IIf(Len(CLMBordx!CLMPreHSAmtActGST), CLMBordx!CLMPreHSAmtActGST, 0)), "#,##0.00"), 11)  '2346 Act Pre-Hosp Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMPreHSAmtAct, "#,##0.00"), 11)
            End If
            
            'Act Surgical Amt
            If IsNumeric(CLMBordx!CLMSurgicalAmtAct) Then
                If IsNumeric(CLMBordx!CLMSurgicalAmtActGST) Then
                    DataString = DataString & l(Format(CDbl(CLMBordx!CLMSurgicalAmtAct) + CDbl(CLMBordx!CLMSurgicalAmtActGST), "#,##0.00"), 11)         '2357 Act Surgical Amt
                Else
                    DataString = DataString & l(Format(CLMBordx!CLMSurgicalAmtAct, "#,##0.00"), 11)
                End If
            Else
                DataString = DataString & l(Format(CLMBordx!CLMSurgicalAmtAct, "#,##0.00"), 11)
            End If
            
            'Act Anae amt
            If IsNumeric(CLMBordx!CLMAnaAct) Then
                If IsNumeric(CLMBordx!CLMAnaActGST) Then
                    DataString = DataString & l(Format(CDbl(CLMBordx!CLMAnaAct) + CDbl(CLMBordx!CLMAnaActGST), "#,##0.00"), 11)                             '2368 Act Anae amt
                Else
                    DataString = DataString & l(Format(CLMBordx!CLMAnaAct, "#,##0.00"), 11)
                End If
            Else
                DataString = DataString & l(Format(CLMBordx!CLMAnaAct, "#,##0.00"), 11)
            End If
            'Act NS Pres-Hosp Amt
            If IsNumeric(CLMBordx!CLMNSPreHosAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMNSPreHosAct) + CDbl(IIf(Len(CLMBordx!CLMNSPreHosActGST), CLMBordx!CLMNSPreHosActGST, 0)), "#,##0.00"), 11)              '2379 Act NS Pres-Hosp Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMNSPreHosAct, "#,##0.00"), 11)
            End If
            'Act Physician Amt
            If IsNumeric(CLMBordx!CLMPhysicianAct) Then
                If IsNumeric(CLMBordx!CLMPhysicianActGST) Then
                    DataString = DataString & l(Format(CDbl(CLMBordx!CLMPhysicianAct) + CDbl(IIf(Len(CLMBordx!CLMPhysicianActGST), CLMBordx!CLMPhysicianActGST, 0)), "#,##0.00"), 11)           '2390 Act Physician Amt
                Else
                    DataString = DataString & l(Format(CLMBordx!CLMPhysicianAct, "#,##0.00"), 11)
                End If
            Else
                DataString = DataString & l(Format(CLMBordx!CLMPhysicianAct, "#,##0.00"), 11)
            End If
            'Act Post Amt
            If IsNumeric(CLMBordx!CLMPostActGST) And IsNumeric(CLMBordx!CLMPostAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMPostAct) + CDbl(IIf(Len(CLMBordx!CLMPostActGST), CLMBordx!CLMPostActGST, 0)), "#,##0.00"), 11)                              '2401 Act Post Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMPostAct, "#,##0.00"), 11)
            End If
            'Act Emergency Amt
            If IsNumeric(CLMBordx!CLMEmergencyAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMEmergencyAct) + CDbl(IIf(Len(CLMBordx!CLMEmergencyActGST), CLMBordx!CLMEmergencyActGST, 0)), "#,##0.00"), 11)               '2412 Act Emergency Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMEmergencyAct, "#,##0.00"), 11)
            End If
            'Act Ambulance Amt
            If IsNumeric(CLMBordx!CLMAmbulanceAct) And IsNumeric(CLMBordx!CLMAmbulanceActGST) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMAmbulanceAct) + CDbl(IIf(Len(CLMBordx!CLMAmbulanceActGST), CLMBordx!CLMAmbulanceActGST, 0)), "#,##0.00"), 11)               '2423 Act Ambulance Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMAmbulanceAct, "#,##0.00"), 11)
            End If
            'Act Outpatient amt
            If IsNumeric(CLMBordx!CLMOutpatientAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMOutpatientAct) + CDbl(IIf(Len(CLMBordx!CLMOutpatientActGST), CLMBordx!CLMOutpatientActGST, 0)), "#,##0.00"), 11)            '2434 Act Outpatient amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMOutpatientAct, "#,##0.00"), 11)
            End If
            'Act Kidney Dialysis Amt
            If IsNumeric(CLMBordx!CLMMonthlyAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMMonthlyAct) + CDbl(IIf(Len(CLMBordx!CLMMonthlyActGST), CLMBordx!CLMMonthlyActGST, 0)), "#,##0.00"), 11)                     '2445 Act Kidney Dialysis Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMMonthlyAct, "#,##0.00"), 11)
            End If
            'Act Organ Trans Amt
            If IsNumeric(CLMBordx!CLMOrganAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMOrganAct) + CDbl(IIf(Len(CLMBordx!CLMOrganActGST), CLMBordx!CLMOrganActGST, 0)), "#,##0.00"), 11)                            '2456 Act Organ Trans Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMOrganAct, "#,##0.00"), 11)
            End If
            'act Medical Report
            If IsNumeric(CLMBordx!CLMMedicalAct) Then
                If IsNumeric(CLMBordx!CLMMedicalActGST) Then
                    DataString = DataString & l(Format(CDbl(CLMBordx!CLMMedicalAct) + CDbl(IIf(Len(CLMBordx!CLMMedicalActGST), CLMBordx!CLMMedicalActGST, 0)), "#,##0.00"), 11)                     '2467 Act Medical Report
                Else
                    DataString = DataString & l(Format(CLMBordx!CLMMedicalAct, "#,##0.00"), 11)
                End If
            Else
                DataString = DataString & l(Format(CLMBordx!CLMMedicalAct, "#,##0.00"), 11)
            End If
            'Act Phone Amt
            If IsNumeric(CLMBordx!CLMPhoneAct) And IsNumeric(CLMBordx!CLMPhoneActGST) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMPhoneAct) + CDbl(IIf(Len(CLMBordx!CLMPhoneActGST), CLMBordx!CLMPhoneActGST, 0)), "#,##0.00"), 11)                               '2478 Act Phone Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMPhoneAct, "#,##0.00"), 11)
            End If
            'Act Gurdian amt
            If IsNumeric(CLMBordx!CLMGuardianAmtAct) And IsNumeric(CLMBordx!CLMGuardianAmtActGST) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMGuardianAmtAct) + CDbl(IIf(Len(CLMBordx!CLMGuardianAmtActGST), CLMBordx!CLMGuardianAmtActGST, 0)), "#,##0.00"), 11)             '2489 Act Gurdian amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMGuardianAmtAct, "#,##0.00"), 11)
            End If
            'Act Admin Amt
            If IsNumeric(CLMBordx!CLMAdminAct) And IsNumeric(CLMBordx!CLMAdminActGST) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMAdminAct) + CDbl(IIf(Len(CLMBordx!CLMAdminActGST), CLMBordx!CLMAdminActGST, 0)), "#,##0.00"), 11)                               '2500 Act Admin Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMAdminAct, "#,##0.00"), 11)
            End If
            'Act Tax Amt
            DataString = DataString & l(Format(IIf(Len(CLMBordx!CLMTaxAct), CLMBordx!CLMTaxAct, 0), "#,##0.00"), 11)                                                                                '2511 Act Tax Amt
            
            'Act Others1 Amt
            If IsNumeric(CLMBordx!CLMOthers1Act) Then
                If IsNumeric(CLMBordx!CLMOthers1ActGST) Then
                    DataString = DataString & l(Format(CDbl(CLMBordx!CLMOthers1Act) + CDbl(IIf(Len(CLMBordx!CLMOthers1ActGST), CLMBordx!CLMOthers1ActGST, 0)), "#,##0.00"), 11)                       '2522 Act Others1 Amt
                Else
                    DataString = DataString & l(Format(CLMBordx!CLMOthers1Act, "#,##0.00"), 11)
                End If
            Else
                DataString = DataString & l(Format(CLMBordx!CLMOthers1Act, "#,##0.00"), 11)
            End If
            'Act Others2 Amt
            If IsNumeric(CLMBordx!CLMOthers2Act) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMOthers2Act) + CDbl(IIf(Len(CLMBordx!CLMOthers2ActGST), CLMBordx!CLMOthers2ActGST, 0)), "#,##0.00"), 11)                     '2533 Act Others2 Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMOthers2Act, "#,##0.00"), 11)
            End If
            'Act Cash Amt
            If IsNumeric(CLMBordx!CLMCashAmtAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMCashAmtAct) + CDbl(IIf(Len(CLMBordx!CLMCashAmtActGST), CLMBordx!CLMCashAmtActGST, 0)), "#,##0.00"), 11)                 '2544
            Else
                DataString = DataString & l(Format(CLMBordx!CLMCashAmtAct, "#,##0.00"), 11)
            End If
            
            DataString = DataString & l(Format(CLMBordx!CLMBoardTaxAmtAct, "#,##0.00"), 11)     '2555 Act Board Tax Amt
            DataString = DataString & l(Format(CLMBordx!CLMICUTaxAmtAct, "#,##0.00"), 11)       '2566 Act ICU Tax amt
            DataString = DataString & l(Format(CLMBordx!CLMLodgerTaxAmtAct, "#,##0.00"), 11)     '2577 Act Lodger Tax Amt
            DataString = DataString & l(Format(CLMBordx!CLMMRIAct, "#,##0.00"), 11)             '2588 Act MRI Amt
            'Act Sur Pre-Hosp Amt
            If IsNumeric(CLMBordx!CLMMPreHSAmtGST) Then
                Dim dHSAmt As Double
                dHSAmt = 0
                If IsNumeric(CLMBordx!CLMMPreHSAmtAct) Then
                    dHSAmt = CDbl(CLMBordx!CLMMPreHSAmtAct)
                Else
                    dHSAmt = 0
                End If
                DataString = DataString & l(Format(dHSAmt + CDbl(CLMBordx!CLMMPreHSAmtGST), "#,##0.00"), 11)    '2599 Act Sur Pre-Hosp Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMMPreHSAmtAct, "#,##0.00"), 11)
            End If
            
            'Act NS Pre-Hosp SUM Amt
            If IsNumeric(CLMBordx!CLMCOPreHSSumAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMCOPreHSSumAct) + CDbl(CLMBordx!CLMCOPreHSSumActGST), "#,##0.00"), 11)       '2610 Act NS Pre-Hosp SUM Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMCOPreHSSumAct, "#,##0.00"), 11)
            End If
            'Act Surg Pre-Hosp SUM Amt
            If IsNumeric(CLMBordx!CLMCOMPreHSSumAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMCOMPreHSSumAct) + CDbl(CLMBordx!CLMCOMPreHSSumActGST), "#,##0.00"), 11)     '2621 Act Surg Pre-Hosp SUM Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMCOMPreHSSumAct, "#,##0.00"), 11)
            End If
            
            DataString = DataString & l(Format("", "#,##0.00"), 11)     '2632 Act Acc Post Amt
            DataString = DataString & l(Format("", "#,##0.00"), 11)     '2643 Acc Post Treatment
            
            'Act Pre-Hosp Amt
            If IsNumeric(CLMBordx!CLMPreHospAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMPreHospAct) + CDbl(IIf(Len(CLMBordx!CLMPreHospActGST), CLMBordx!CLMPreHospActGST, 0)), "#,##0.00"), 11) '2654 Act Pre-Hosp Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMPreHospAct, "#,##0.00"), 11)
            End If
            
            'Act Others Not Co-Pay amt
            If IsNumeric(CLMBordx!CLMOthersNotCoPayAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMOthersNotCoPayAct) + CDbl(CLMBordx!CLMOthersNotCoPayActGST), "#,##0.00"), 11)       '2665 Act Others Not Co-Pay amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPayAct, "#,##0.00"), 11)
            End If
            'Act CLM Home NC Amt
            If IsNumeric(CLMBordx!CLMHomeNCAmtAct) Then
                DataString = DataString & l(Format(CDbl(CLMBordx!CLMHomeNCAmtAct) + CDbl(CLMBordx!CLMHomeNCAmtActGST), "#,##0.00"), 11)     '2676 Act CLM Home NC Amt
            Else
                DataString = DataString & l(Format(CLMBordx!CLMHomeNCAmtAct, "#,##0.00"), 11)
            End If
            
            'Approved
            'App Board amt
            If IsNumeric(CLMBordx!CLMBoardAmt) Then
                    DataString = DataString & l(CLMBordx!CLMBoardAmt, 11)           '2687 App Board amt
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            'App ICU amt
            If IsNumeric(CLMBordx!CLMICUAmt) Then
                    DataString = DataString & l(Format(CLMBordx!CLMICUAmt, "#,##0.00"), 11)     '2698 App ICU amt
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'App HSS amt
            If IsNumeric(CLMBordx!CLMServices) Then
                    DataString = DataString & l(Format(CLMBordx!CLMServices, "#,##0.00"), 11)       '2709 App HSS amt
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'App Theatre Amt
            If IsNumeric(CLMBordx!CLMTheatre) Then
                    DataString = DataString & l(Format(CLMBordx!CLMTheatre, "#,##0.00"), 11)        '2720 App Theatre Amt
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'App Pre-Hosp amt
            If IsNumeric(CLMBordx!CLMPreHSAmt) Then
                    DataString = DataString & l(CLMBordx!CLMPreHSAmt, 11)                       '2731 App Pre-Hosp amt
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'App Surgical Amt
            If IsNumeric(CLMBordx!CLMSurgicalAmt) Then
                    DataString = DataString & l(Format(CLMBordx!CLMSurgicalAmt, "#,##0.00"), 11)        '2742 App Surgical Amt
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'app anae Amt
            If IsNumeric(CLMBordx!CLMAna) Then
                    DataString = DataString & l(Format(CLMBordx!CLMAna, "#,##0.00"), 11)            '2753 app anae Amt
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            'Location 2643 App NS Pre-Hosp Amt
            If IsNumeric(CLMBordx!CLMNSPreHos) Then
                    DataString = DataString & l(Format(CLMBordx!CLMNSPreHos, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
             
            '2654 App Physician Amt
            If IsNumeric(CLMBordx!CLMPhysician) Then
                    DataString = DataString & l(Format(CLMBordx!CLMPhysician, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2678 App Post amt
            If IsNumeric(CLMBordx!CLMPost) Then
                    DataString = DataString & l(Format(CLMBordx!CLMPost, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2689 App Emergency amt
            If IsNumeric(CLMBordx!CLMEmergency) Then
                    DataString = DataString & l(Format(CLMBordx!CLMEmergency, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            'App Ambulance amt 2687
            If IsNumeric(CLMBordx!CLMAmbulance) Then
                    DataString = DataString & l(Format(CLMBordx!CLMAmbulance, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            '2711 App Outpatient amt
            If IsNumeric(CLMBordx!CLMOutpatient) Then
                    DataString = DataString & l(Format(CLMBordx!CLMOutpatient, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2722 App Kidney Dialysis amt
            If IsNumeric(CLMBordx!CLMMonthly) Then
                    DataString = DataString & l(Format(CLMBordx!CLMMonthly, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2733 App Organ Trans Amt
            If IsNumeric(CLMBordx!CLMOrgan) Then
                    DataString = DataString & l(Format(CLMBordx!CLMOrgan, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2744 App Medical Report Amt
            If IsNumeric(CLMBordx!CLMMedical) Then
                    DataString = DataString & l(Format(CLMBordx!CLMMedical, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            '2755 App Phone Amt
            If IsNumeric(CLMBordx!CLMPhone) Then
                    DataString = DataString & l(Format(CLMBordx!CLMPhone, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            '2766 App Gurdian amt
            If IsNumeric(CLMBordx!CLMGuardianAmt) Then
                    DataString = DataString & l(Format(CLMBordx!CLMGuardianAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            '2777  App Admin Amt
            If IsNumeric(CLMBordx!CLMAdmin) Then
                    DataString = DataString & l(Format(CLMBordx!CLMAdmin, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            ''App Tax Amt
            'If IsNumeric(CLMBordx!CLMTaxActGST) Then
            '        DataString = DataString & l(Format(CLMBordx!CLMTaxActGST, "#,##0.00"), 11)
            'Else
            '    DataString = DataString & l("0.00", 11)
            'End If
            
            'App Tax Amt
            If IsNumeric(CLMBordx!CLMTax) Then
                DataString = DataString & l(Format(CLMBordx!CLMTax, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2799 App Others 1 amt
            If IsNumeric(CLMBordx!CLMOthers1) Then
                If Mid(cboPAYCode, 1, 2) = "MM" Then
                    If CLMBordx!CLMOthers1 < 0 Then
                        DataString = DataString & l(Format("0", "#,##0.00"), 11)
                    Else
                        DataString = DataString & l(Format(CLMBordx!CLMOthers1, "#,##0.00"), 11)
                    End If
                Else
                    DataString = DataString & l(Format(CLMBordx!CLMOthers1, "#,##0.00"), 11)
                End If
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2810 App Others 2 amt
            If IsNumeric(CLMBordx!CLMOthers2) Then
                    DataString = DataString & l(Format(CLMBordx!CLMOthers2, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2821 App Cash Amt
            If IsNumeric(CLMBordx!CLMCashAmt) Then
                    DataString = DataString & l(Format(CLMBordx!CLMCashAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2832 App Board Tax Amt
            If IsNumeric(CLMBordx!CLMBoardTaxAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMBoardTaxAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            '2843 App ICU Tax Amt
            If IsNumeric(CLMBordx!CLMICUTaxAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMICUTaxAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            '2854 App Lodger Amt
            If IsNumeric(CLMBordx!CLMLodgerTaxAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMLodgerTaxAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            '2865 App MRI amt
            If IsNumeric(CLMBordx!CLMMRI) Then
                DataString = DataString & l(Format(CLMBordx!CLMMRI, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            '2876 App Sur Pre Hosp Amt
            If IsNumeric(CLMBordx!CLMMPreHSAmt) Then
                    DataString = DataString & l(Format(CLMBordx!CLMMPreHSAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            'If IsNumeric(CLMBordx!CLMCOPreHSSum) Then
            '        'DataString = DataString & l(Format(CLMBordx!CLMCOPreHSSum, "#,##0.00"), 11)
            '    DataString = DataString & l("0.00", 11)
            'Else
            '    DataString = DataString & l("0.00", 11)
            'End If
            
            '2887  App NSPreHos
            If IsNumeric(CLMBordx!CLMNSPreHos) Then
                    DataString = DataString & l(Format(CLMBordx!CLMNSPreHos, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2898 App Surg Pre-Hosp sum Amt
            If IsNumeric(CLMBordx!CLMCOMPreHSSum) Then
                    DataString = DataString & l(Format(CLMBordx!CLMCOMPreHSSum, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2909 App Acc Post amt
            If IsNumeric(CLMBordx!CLMAccPostAmt) Then
                    DataString = DataString & l(Format(CLMBordx!CLMAccPostAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2920 Acc Pre-Hosp Amt
            If IsNumeric(CLMBordx!CLMPreHosp) Then
                    DataString = DataString & l(Format(CLMBordx!CLMPreHosp, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '2931 App Others Not- Co Pay amt
            If Trim(Mid(cboPAYCode, 1, 2)) = Trim("MM") Then
                DataString = DataString & l(Format(CLMBordx!CLMOthers1Lmt, "#,##0.00"), 11)       '2181
            Else
                    If IsNumeric(CLMBordx!CLMOthersNotCoPay) Then
                            DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPay, "#,##0.00"), 11)
                    Else
                        DataString = DataString & l("0.00", 11)
                    End If
            End If
            
            '2942 App CLM Home NC Amt
            If IsNumeric(CLMBordx!CLMHomeNCAmt) Then
                    DataString = DataString & l(Format(CLMBordx!CLMHomeNCAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            
            DataString = DataString & l("0.00", 11)  ' Not cover Pre- Hosp
            DataString = DataString & l("0.00", 11) '2964 Not Cover GP Fee
            DataString = DataString & l("0.00", 11) '2975 Not Cover Pre-Medicine
            DataString = DataString & l("0.00", 11) '2986 Not Covered Medical Report
            DataString = DataString & l("0.00", 11) '2997 Not Covered Specialist Fee
            DataString = DataString & l("0.00", 11) '3008 Not Covered Post-Hosp
            DataString = DataString & l("0.00", 11) '3019 Not Covered Phycisian
            DataString = DataString & l("0.00", 11) '3030 Not Covered Others
            
            '3035 ICD Code 1
            If CLMBordx!ICDCode1 <> "" Then
                DataString = DataString & l(CLMBordx!ICDCode1, 5)
            Else
                DataString = DataString & l("", 5)
            End If
            
            '3040  ICD Code 2
            If CLMBordx!ICDCode2 <> "" Then
                DataString = DataString & l(CLMBordx!ICDCode2, 5)
            Else
                DataString = DataString & l("", 5)
            End If
            
            '3045 ICD Code 3
            If CLMBordx!ICDCode3 <> "" Then
                DataString = DataString & l(CLMBordx!ICDCode3, 5)
            Else
                DataString = DataString & l("", 5)
            End If
            
            
            DataString = DataString & l("", 8)  '3053 Other Remarks 1
            DataString = DataString & l("", 8)  '3061 Other Remarks 2
            
            DataString = DataString & l(Format(0, "#,##0.00"), 11)  '3072 Lmt Cancer amt
            DataString = DataString & l(Format(0, "#,##0.00"), 11)  '3083 Act Cancer Amt
            DataString = DataString & l(Format(0, "#,##0.00"), 11)  '3094 App Cancer amt
            DataString = DataString & l("KLL", 8)  '3102  KLL
            DataString = DataString & l("0", 11)   '3113  Kidney Dialysis Lifetime Lmt
            DataString = DataString & l("CLL", 8)  '3121    CLL
            DataString = DataString & l("0", 11)   '3132Cancer Lifetime Limit
                
            If Mid(cboPAYCode, 1, 2) = "MM" Then
                If IsNumeric(CLMBordx!CLMOthers1) Then
                    If CLMBordx!CLMOthers1 < 0 Then
                        DataString = DataString & l("HS", 8)   '3140
                        DataString = DataString & l(Format(CLMBordx!CLMOthers1 * -1, "#,##0.00"), 11) '3151
                    Else
                        DataString = DataString & l("HS", 8)   '3140
                        DataString = DataString & l("0", 11)   '3151
                    End If
                Else
                    DataString = DataString & l("HS", 8)   '3140
                    DataString = DataString & l("0", 11)   '3151
                End If
            Else
                DataString = DataString & l("HS", 8)   '3140    HS
                DataString = DataString & l("0", 11)   '3151    H7S Co-payment
            End If
            
            DataString = DataString & l("OP", 8)    '3159   OP
            DataString = DataString & l("0", 11)    '3170   O/P Co-payment
            
        'New Addition for Zurich
            If Option1.Value = True Then
                DataString = DataString & l("", 50)  '3220  Repudiation Reason 1
                DataString = DataString & l("", 50)  '3270  Repudiation Reason 2
                DataString = DataString & l("", 50)  '3320  Repudiation Reason 3
            Else
                If CLMBordx!RepDescription1 <> "" Then
                    DataString = DataString & l(CLMBordx!RepDescription1, 50)
                Else
                    DataString = DataString & l("", 50)
                End If
                DataString = DataString & l("", 50)
                DataString = DataString & l("", 50)
            End If
            
            DataString = DataString & l(Mid(CLMBordx!CaseClaim, 2, 1), 1) '3321  Claimability
            
            Call GSTActApp
            Call NewZurichRequest
            
            DataString = DataString & Chr(13)
            TotalRecord = TotalRecord + 1
            CLMBordx.MoveNext
            Print #1, DataString
        Loop
                Trailer = "T" & l(TotalRecord, 10) & l("", 164)
                Print #1, Trailer
                cmdGenerate.Enabled = False
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
    End If
    CLMBordx.Close
    Set CLMBordx = Nothing
    Close #1
End Sub
Private Sub NewZurichRequest()
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Deductible Amt
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'NCB Clamback Amount
    
    DataString = DataString & l("", 100)   'NCB ID
    DataString = DataString & l("", 40)    'NCB Clawback Period/ Policy yr
    DataString = DataString & l(Format(0, "#,##0.00"), 20)    'OLL
    DataString = DataString & l(Format(0, "#,##0.00"), 15)    'OAL
    DataString = DataString & l(Format(0, "#,##0.00"), 20)    'Kidney OLL
    DataString = DataString & l(Format(0, "#,##0.00"), 20)    'Cancer OLL
    
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Rooms & Board
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins ICU/HDU
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Services
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Surgical Fees
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Anaeshetist fee
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Operating  Theatre
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Day Surgery
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Test
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Pre HS Consultation
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Physician Visit
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Post Hospitalization
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Ambulance Fee
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Out-patient Kidney Dialysis
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins OP Cancer Dialysis
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Co-Ins Physiotherapy Treatment
    DataString = DataString & l(Format(0, "#,##0.00"), 11)    'Emenrgency Accidental OP
    
End Sub
Private Sub GSTActApp()
Dim GSTBillAmt As Double
Dim GSTAppAmt As Double
GSTBillAmt = 0
GSTAppAmt = 0
            'Actual GST
            DataString = DataString & l(Format(CLMBordx!CLMBoardAmtActGst, "#,##0.00"), 11)    '3330 GST Act Board Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMBoardAmtActGst), CDbl(CLMBordx!CLMBoardAmtActGst), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMICUAmtActGST, "#,##0.00"), 11)      '3343 GST Act ICU Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMICUAmtActGST), CDbl(CLMBordx!CLMICUAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMServicesActGST, "#,##0.00"), 11)    '3354 GST Act HSS Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMServicesActGST), CDbl(CLMBordx!CLMServicesActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMTheatreActGST, "#,##0.00"), 11)     '3365 GST Act Theatre Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMTheatreActGST), CDbl(CLMBordx!CLMTheatreActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMPreHSAmtActGST, "#,##0.00"), 11)    '3376 GST Act Pre-Hosp Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMPreHSAmtActGST), CDbl(CLMBordx!CLMPreHSAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMSurgicalAmtActGST, "#,##0.00"), 11) '3387 GST Act Surgical Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMSurgicalAmtActGST), CDbl(CLMBordx!CLMSurgicalAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMAnaActGST, "#,##0.00"), 11)         '3398 GST Act Anae Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMAnaActGST), CDbl(CLMBordx!CLMAnaActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMNSPreHosActGST, "#,##0.00"), 11)    '3409 GST Act NS-Pre hosp Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMNSPreHosActGST), CDbl(CLMBordx!CLMNSPreHosActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMPhysicianActGST, "#,##0.00"), 11)   '3420 GST Act Physician Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMPhysicianActGST), CDbl(CLMBordx!CLMPhysicianActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMPostActGST, "#,##0.00"), 11)        '3431 GST Act Post Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMPostActGST), CDbl(CLMBordx!CLMPostActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMEmergencyActGST, "#,##0.00"), 11)   '3442 GST Act Emergency Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMEmergencyActGST), CDbl(CLMBordx!CLMEmergencyActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMAmbulanceActGST, "#,##0.00"), 11)    '3453 GST Act Ambulance Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMAmbulanceActGST), CDbl(CLMBordx!CLMAmbulanceActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMOutpatientActGST, "#,##0.00"), 11)   '3464 GST Act OutpatientAmt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMOutpatientActGST), CDbl(CLMBordx!CLMOutpatientActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMMonthlyActGST, "#,##0.00"), 11)      '3475 GST Act Kidney Dialysis Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMMonthlyActGST), CDbl(CLMBordx!CLMMonthlyActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMOrganActGST, "#,##0.00"), 11)        '3486 GST Act Organ trans Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMOrganActGST), CDbl(CLMBordx!CLMOrganActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMMedicalActGST, "#,##0.00"), 11)      '3497 GST Act Medical Report Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMMedicalActGST), CDbl(CLMBordx!CLMMedicalActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMPhoneActGST, "#,##0.00"), 11)        '3508 GST Act Phone Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMPhoneActGST), CDbl(CLMBordx!CLMPhoneActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMGuardianAmtActGST, "#,##0.00"), 11)  '3519 GST Act Gurdian Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMGuardianAmtActGST), CDbl(CLMBordx!CLMGuardianAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMAdminActGST, "#,##0.00"), 11)        '3530 GST Act Admin Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMAdminActGST), CDbl(CLMBordx!CLMAdminActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMTaxActGST, "#,##0.00"), 11)          '3541 GST Act Tax Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMTaxActGST), CDbl(CLMBordx!CLMTaxActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthers1ActGST, "#,##0.00"), 11)      '3552 GST Act Others1 Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMOthers1ActGST), CDbl(CLMBordx!CLMOthers1ActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthers2ActGST, "#,##0.00"), 11)      '3563 GST Act Others2 Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMOthers2ActGST), CDbl(CLMBordx!CLMOthers2ActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMCashAmtActGST, "#,##0.00"), 11)      '3574 GST Act cash Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMCashAmtActGST), CDbl(CLMBordx!CLMCashAmtActGST), 0)
            
            DataString = DataString & l(Format(0, "#,##0.00"), 11)                              '3585 GST Act Borad Tax Amt
            DataString = DataString & l(Format(0, "#,##0.00"), 11)                              '3596 GST Act ICU Tax Amt
            DataString = DataString & l(Format(0, "#,##0.00"), 11)                              '3607 GST Act Lodger Tax Amt
            
            DataString = DataString & l(Format(CLMBordx!CLMMRIActGST, "#,##0.00"), 11)           '3618 GST Act MRI Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMMRIActGST), CDbl(CLMBordx!CLMMRIActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMMPreHosAmtActGST, "#,##0.00"), 11)     '3629 GST Act Sur Pre-HOSP Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMMPreHosAmtActGST), CDbl(CLMBordx!CLMMPreHosAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMCHospSumActGST, "#,##0.00"), 11)       '3630 GST Act NS Pre-Hosp SUM Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMCHospSumActGST), CDbl(CLMBordx!CLMCHospSumActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMMHospSumActGST, "#,##0.00"), 11)       '3641 GST Act Surg Pre-Hosp Sum Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMMHospSumActGST), CDbl(CLMBordx!CLMMHospSumActGST), 0)
            
            DataString = DataString & l(Format("", "#,##0.00"), 11)                               '3652 GST Act AccPost Amt
            
            DataString = DataString & l(Format(CLMBordx!CLMPreHospActGST, "#,##0.00"), 11)        '3663 GST Act Pre-Hosp Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMPreHospActGST), CDbl(CLMBordx!CLMPreHospActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPayActGST, "#,##0.00"), 11)  '3674 GST Act Others Not Co-Pay Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMOthersNotCoPayActGST), CDbl(CLMBordx!CLMOthersNotCoPayActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMHomeNCAmtActGST, "#,##0.00"), 11)       '3685 GST Act CLM Home NC Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMHomeNCAmtActGST), CDbl(CLMBordx!CLMHomeNCAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMCancerAmtActGST, "#,##0.00"), 11)       '3696 GST Act Cancer Amt
            GSTBillAmt = GSTBillAmt + IIf(IsNumeric(CLMBordx!CLMCancerAmtActGST), CDbl(CLMBordx!CLMCancerAmtActGST), 0)
             
            'APPROVE GST
            tmpAppGST = 0
            'DataString = DataString & l(Format(CLMBordx!CLMBoardAmtGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMICUAmtGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMServicesGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMTheatreGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMPreHSAmtGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMSurgicalAmtGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMAnaGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMNSPreHosGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMPhysicianGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMPostGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMEmergencyGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMAmbulanceGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMOutpatientGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMMonthlyGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMOrganGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMMedicalGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMPhoneGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMGuardianAmtGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMAdminGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMTaxGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMOthers1GST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMOthers2GST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMCashAmtGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMBoardTaxAmtGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMICUTaxAmtGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMLodgerTaxAmtGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMMRIGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMMPreHSAmtGST, "#,##0.00"), 11)
            'dataString = DataString & l(Format(CLMBordx!CLMCOPreHSSumGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMCOMPreHSSumGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format("", "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMPreHospGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPayGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMHomeNCAmtGST, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMCancerAmtGST, "#,##0.00"), 11)

            DataString = DataString & l(Format(CLMBordx!CLMBoardAmtActGst, "#,##0.00"), 11)  '3707  GST App Board Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMBoardAmtActGst), CDbl(CLMBordx!CLMBoardAmtActGst), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMICUAmtActGST, "#,##0.00"), 11)    '3718  GST App ICU Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMICUAmtActGST), CDbl(CLMBordx!CLMICUAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMServicesActGST, "#,##0.00"), 11)  '3729  GST App HSS Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMServicesActGST), CDbl(CLMBordx!CLMServicesActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMTheatreActGST, "#,##0.00"), 11)   '3740  GST App theatre Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMTheatreActGST), CDbl(CLMBordx!CLMTheatreActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMPreHSAmtActGST, "#,##0.00"), 11)  '3751  GST App Pre-Hosp Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMPreHSAmtActGST), CDbl(CLMBordx!CLMPreHSAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMSurgicalAmtActGST, "#,##0.00"), 11) '3762 GST App Surgical Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMSurgicalAmtActGST), CDbl(CLMBordx!CLMSurgicalAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMAnaActGST, "#,##0.00"), 11)    '3773 GST App anae Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMAnaActGST), CDbl(CLMBordx!CLMAnaActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMNSPreHosActGST, "#,##0.00"), 11)  '3784 GST App NS Pres-Hosp Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMNSPreHosActGST), CDbl(CLMBordx!CLMNSPreHosActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMPhysicianActGST, "#,##0.00"), 11) '3795 GST App Physician Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMPhysicianActGST), CDbl(CLMBordx!CLMPhysicianActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMPostActGST, "#,##0.00"), 11)  '3806 GST App Post Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMPostActGST), CDbl(CLMBordx!CLMPostActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMEmergencyActGST, "#,##0.00"), 11) '3817 GST App Emergency Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMEmergencyActGST), CDbl(CLMBordx!CLMEmergencyActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMAmbulanceActGST, "#,##0.00"), 11) '3828 GST App Ambulance Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMAmbulanceActGST), CDbl(CLMBordx!CLMAmbulanceActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMOutpatientActGST, "#,##0.00"), 11) '3839 GST App Outpatient Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMOutpatientActGST), CDbl(CLMBordx!CLMOutpatientActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMMonthlyActGST, "#,##0.00"), 11)   '3850 GST App Kidney Dialysis Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMMonthlyActGST), CDbl(CLMBordx!CLMMonthlyActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMOrganActGST, "#,##0.00"), 11)    '3861  GST App Organ Trans Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMOrganActGST), CDbl(CLMBordx!CLMOrganActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMMedicalActGST, "#,##0.00"), 11)  '3872 GST App Medical Report Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMMedicalActGST), CDbl(CLMBordx!CLMMedicalActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMPhoneActGST, "#,##0.00"), 11)    '3883 GST App Phone Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMPhoneActGST), CDbl(CLMBordx!CLMPhoneActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMGuardianAmtActGST, "#,##0.00"), 11)  '3894 GST App Gurdian Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMGuardianAmtActGST), CDbl(CLMBordx!CLMGuardianAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMAdminActGST, "#,##0.00"), 11)   '3905 GST App Admin Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMAdminActGST), CDbl(CLMBordx!CLMAdminActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMTaxActGST, "#,##0.00"), 11)  '3916  GST App Tax Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMTaxActGST), CDbl(CLMBordx!CLMTaxActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthers1ActGST, "#,##0.00"), 11)   '3927 GST App Others1 Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMOthers1ActGST), CDbl(CLMBordx!CLMOthers1ActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthers2ActGST, "#,##0.00"), 11)   '3938 GST App Others2 Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMOthers2ActGST), CDbl(CLMBordx!CLMOthers2ActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMCashAmtActGST, "#,##0.00"), 11)   '3949 GST App Cash Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMCashAmtActGST), CDbl(CLMBordx!CLMCashAmtActGST), 0)
            
            DataString = DataString & l(Format(0, "#,##0.00"), 11)                           '3960  GST App Board Tax Amt
            DataString = DataString & l(Format(0, "#,##0.00"), 11)                           '3971  GST App ICU Tax Amt
            DataString = DataString & l(Format(0, "#,##0.00"), 11)                           '3982  GST App Lodger Tax Amt
            
            DataString = DataString & l(Format(CLMBordx!CLMMRIActGST, "#,##0.00"), 11)       '3993  GST App MRI Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMMRIActGST), CDbl(CLMBordx!CLMMRIActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMMPreHosAmtActGST, "#,##0.00"), 11) '4004 GST App Sur Pre-Hosp Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMMPreHosAmtActGST), CDbl(CLMBordx!CLMMPreHosAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMCHospSumActGST, "#,##0.00"), 11)  '4023 GST App NS Pre-Hosp SUM Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMCHospSumActGST), CDbl(CLMBordx!CLMCHospSumActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMMHospSumActGST, "#,##0.00"), 11)  '4034 GST App Surg Pre-Hosp Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMMHospSumActGST), CDbl(CLMBordx!CLMMHospSumActGST), 0)
            
            DataString = DataString & l(Format("", "#,##0.00"), 11)          '4045 GST App AccPost Amt
            
            
            DataString = DataString & l(Format(CLMBordx!CLMPreHospActGST, "#,##0.00"), 11)  '4056 GST Apppre-Hosp Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMPreHospActGST), CDbl(CLMBordx!CLMPreHospActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPayActGST, "#,##0.00"), 11) '4067 GST App Others Not Co-pay Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMOthersNotCoPayActGST), CDbl(CLMBordx!CLMOthersNotCoPayActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMHomeNCAmtActGST, "#,##0.00"), 11)  '4078 GST App CLM Home NC Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMHomeNCAmtActGST), CDbl(CLMBordx!CLMHomeNCAmtActGST), 0)
            
            DataString = DataString & l(Format(CLMBordx!CLMCancerAmtActGST, "#,##0.00"), 11)  '4089 GST App Cancer Amt
            GSTAppAmt = GSTAppAmt + IIf(IsNumeric(CLMBordx!CLMCancerAmtActGST), CDbl(CLMBordx!CLMCancerAmtActGST), 0)
            
            'New Added for Zurich
             DataString = DataString & l(Format(GSTBillAmt, "#,##0.00"), 11)  '4089 GST App Cancer Amt
             DataString = DataString & l(Format(GSTAppAmt, "#,##0.00"), 11)  '4089 GST App Cancer Amt
            

End Sub
Private Sub GSTCount()
            tmpActGST = 0
            'Actual GST
          tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMBoardAmtActGst), CLMBordx!CLMBoardAmtActGst, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMICUAmtActGST), CLMBordx!CLMICUAmtActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMServicesActGST), CLMBordx!CLMServicesActGST, 0)
           tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMTheatreActGST), CLMBordx!CLMTheatreActGST, 0)
           tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMPreHSAmtActGST), CLMBordx!CLMPreHSAmtActGST, 0)
           tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMSurgicalAmtActGST), CLMBordx!CLMSurgicalAmtActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMAnaActGST), CLMBordx!CLMAnaActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMNSPreHosActGST), CLMBordx!CLMNSPreHosActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMPhysicianActGST), CLMBordx!CLMPhysicianActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMPostActGST), CLMBordx!CLMPostActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMEmergencyActGST), CLMBordx!CLMEmergencyActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMAmbulanceActGST), CLMBordx!CLMAmbulanceActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMOutpatientActGST), CLMBordx!CLMOutpatientActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMMonthlyActGST), CLMBordx!CLMMonthlyActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMOrganActGST), CLMBordx!CLMOrganActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMMedicalActGST), CLMBordx!CLMMedicalActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMPhoneActGST), CLMBordx!CLMPhoneActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMGuardianAmtActGST), CLMBordx!CLMGuardianAmtActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMAdminActGST), CLMBordx!CLMAdminActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMTaxActGST), CLMBordx!CLMTaxActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMOthers1ActGST), CLMBordx!CLMOthers1ActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMOthers2ActGST), CLMBordx!CLMOthers2ActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMCashAmtActGST), CLMBordx!CLMCashAmtActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMMRIActGST), CLMBordx!CLMMRIActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMMPreHosAmtActGST), CLMBordx!CLMMPreHosAmtActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMCHospSumActGST), CLMBordx!CLMCHospSumActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMMHospSumActGST), CLMBordx!CLMMHospSumActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMPreHospActGST), CLMBordx!CLMPreHospActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMOthersNotCoPayActGST), CLMBordx!CLMOthersNotCoPayActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMHomeNCAmtActGST), CLMBordx!CLMHomeNCAmtActGST, 0)
            tmpActGST = tmpActGST + IIf(IsNumeric(CLMBordx!CLMCancerAmtActGST), CLMBordx!CLMCancerAmtActGST, 0)
            
             'APPROVE GST
            tmpAppGST = 0
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMBoardAmtGST), CLMBordx!CLMBoardAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMICUAmtGST), CLMBordx!CLMICUAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMServicesGST), CLMBordx!CLMServicesGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMTheatreGST), CLMBordx!CLMTheatreGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMPreHSAmtGST), CLMBordx!CLMPreHSAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMSurgicalAmtGST), CLMBordx!CLMSurgicalAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMAnaGST), CLMBordx!CLMAnaGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMNSPreHosGST), CLMBordx!CLMNSPreHosGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMPhysicianGST), CLMBordx!CLMPhysicianGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMPostGST), CLMBordx!CLMPostGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMEmergencyGST), CLMBordx!CLMEmergencyGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMAmbulanceGST), CLMBordx!CLMAmbulanceGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMOutpatientGST), CLMBordx!CLMOutpatientGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMMonthlyGST), CLMBordx!CLMMonthlyGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMOrganGST), CLMBordx!CLMOrganGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMMedicalGST), CLMBordx!CLMMedicalGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMPhoneGST), CLMBordx!CLMPhoneGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMGuardianAmtGST), CLMBordx!CLMGuardianAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMAdminGST), CLMBordx!CLMAdminGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMTaxGST), CLMBordx!CLMTaxGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMOthers1GST), CLMBordx!CLMOthers1GST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMOthers2GST), CLMBordx!CLMOthers2GST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMCashAmtGST), CLMBordx!CLMCashAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMBoardTaxAmtGST), CLMBordx!CLMBoardTaxAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMICUTaxAmtGST), CLMBordx!CLMICUTaxAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMLodgerTaxAmtGST), CLMBordx!CLMLodgerTaxAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMMRIGST), CLMBordx!CLMMRIGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMMPreHSAmtGST), CLMBordx!CLMMPreHSAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMCOPreHSSumGST), CLMBordx!CLMCOPreHSSumGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMCOMPreHSSumGST), CLMBordx!CLMCOMPreHSSumGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMPreHospGST), CLMBordx!CLMPreHospGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMOthersNotCoPayGST), CLMBordx!CLMOthersNotCoPayGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMHomeNCAmtGST), CLMBordx!CLMHomeNCAmtGST, 0)
            tmpAppGST = tmpAppGST + IIf(IsNumeric(CLMBordx!CLMCancerAmtGST), CLMBordx!CLMCancerAmtGST, 0)

End Sub

Private Sub Form_Load()
    txtPath = SSRVPath & "ClaimASCI"
    txtLocalFileLocation = "C:\MedicaGenClaimASCII\"
    Call ComboListing
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command

  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
        
    Do Until PAY.EOF
        With PAY
            cboPAYCode.AddItem !PAYCode & " - " & Trim(!PAYCompanyName)
        End With
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
        
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
    CboExportStatus.AddItem "N - New"
    CboExportStatus.AddItem "R - Reprint"
    CboExportStatus.Text = "N - New"
    
End Sub





