VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmClaimASCIIRejZurich 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   8865
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option2 
      Caption         =   "Reject Cases"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   360
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Pending Cases"
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame8 
      Height          =   2295
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   7095
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
         TabIndex        =   8
         Top             =   1200
         Width           =   5535
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
         TabIndex        =   7
         Top             =   1560
         Width           =   1995
      End
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
         TabIndex        =   6
         Top             =   1920
         Width           =   1995
      End
      Begin MSMask.MaskEdBox txtReceiptDateFr 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   21
         Top             =   900
         Width           =   1245
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
         TabIndex        =   20
         Top             =   1185
         Width           =   915
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
         TabIndex        =   19
         Top             =   615
         Width           =   765
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
         TabIndex        =   18
         Top             =   240
         Width           =   1905
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
         TabIndex        =   17
         Top             =   315
         Width           =   1125
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
         TabIndex        =   16
         Top             =   1545
         Width           =   1155
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
         TabIndex        =   15
         Top             =   540
         Width           =   1905
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
         TabIndex        =   14
         Top             =   1560
         Width           =   555
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
         TabIndex        =   13
         Top             =   1920
         Width           =   1035
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
         TabIndex        =   12
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label txtFilName 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   840
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   7095
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   825
      End
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
         TabIndex        =   2
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Accepted Cases"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Generate Claims Reject ASCII (Zurich 2015)"
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
      TabIndex        =   22
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmClaimASCIIRejZurich"
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
    ElseIf Option3.Value = True And txtReceiptDateFr = "__/__/____" Then
        MsgBox "Please Enter Case Reg. Date Fr ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDateFr.SetFocus
    ElseIf Option3.Value = True And txtReceiptDateTo = "__/__/____" Then
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
        ElseIf Option2.Value = True Then
                strAppend = " AND CaseRegDate >= '" & Format(CDate(BordxDateFr), "yyyy/mm/dd") & "'" & _
                            " AND CaseRegDate <= '" & Format(CDate(BordxDateTo), "yyyy/mm/dd") & "'" & _
                            " AND PAYCode = '" & PAYCode & "' "
        ElseIf Option3.Value = True Then
                strAppend = " AND CaseRegDate >= '" & Format(CDate(BordxDateFr), "yyyy/mm/dd") & "'" & _
                            " AND CaseRegDate <= '" & Format(CDate(BordxDateTo), "yyyy/mm/dd") & "'" & _
                            " AND PAYCode = '" & PAYCode & "' "
        End If
        
        If Option1.Value = True Then
            strsql = "SELECT * From View_Claims_ASCII1 where 0 = 0 "
        ElseIf Option2.Value = True Then
            strsql = "select * from View_ClaimsRejected_ASCII where CaseClaim='RR' and  0 = 0"
        ElseIf Option3.Value = True Then
            strsql = "select * from View_ClaimsRejected_ASCII where CaseClaim='PP' and 0 = 0"
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
            If Option1.Value = True Then
                txtFilName = Trim("MXACC" & Format(Date, "YYMMDD"))
            ElseIf Option2.Value = True Then
                txtFilName = Trim("MXREJ" & Format(Date, "YYMMDD"))
            ElseIf Option3.Value = True Then
                txtFilName = Trim("MXPEN" & Format(Date, "YYMMDD"))
            End If
            'txtFilName = Trim("MX" & Format(Date, "YYMMDD"))
            Open txtLocalFileLocation & "\" & txtFilName & "(" & BATNo & ")" For Output As #1
            Header = Trim("F" & Format(Date, "DDMMYYYY"))
            If Option2.Value = True Then
                Header = Header & l("REJECT", 7)
            ElseIf Option3.Value = True Then
                Header = Header & l("PENDING", 7)
            End If
            Print #1, Trim(Header)
            
        Do Until CLMBordx.EOF
        
            DataString = ""
            DataString = DataString & l(CLMBordx!MBMPolicyNo, 25)                       'Policy Number
            DataString = DataString & l(Format(CLMBordx!MBMPayorEffDate, "ddmmyyyy"), 8)         'Policy Eff Date
            DataString = DataString & l(Format(CLMBordx!MBMPayorEXPDate, "ddmmyyyy"), 8)         'Policy Exp Date
            DataString = DataString & l(CLMBordx!PLNCode, 7)                            'Plan code
            DataString = DataString & l(CLMBordx!PatientName, 60)                       'Patient Name
            DataString = DataString & l(Format(CLMBordx!PatientDOB, "ddmmyyyy"), 8)             'Patient DOB
            DataString = DataString & l(CLMBordx!PatientSex, 1)                         'Patient Sex
            DataString = DataString & l(CLMBordx!PatientIC, 14)                         'Patient IC
            DataString = DataString & l(CLMBordx!HosName, 80)                           'Hospital
            DataString = DataString & l(Format(CLMBordx!CASDateAdmitted, "ddmmyyyy"), 8)        'DOA
            DataString = DataString & l((IIf(Len(CLMBordx!CasNumber), CLMBordx!CasNumber, "")), 15)                        'Case No
            If Option2.Value = True Then
                Dim RepDesc As String
                RepDesc = IIf(Len(CLMBordx!RepDescription1), CLMBordx!RepDescription1, "") & IIf(Len(CLMBordx!RepDescription2), CLMBordx!RepDescription2, "") & IIf(Len(CLMBordx!RepDescription3), CLMBordx!RepDescription3, "") & IIf(Len(CLMBordx!RepDescription4), CLMBordx!RepDescription4, "") & " " & IIf(Len(CLMBordx!CASREPReasons1), CLMBordx!CASREPReasons1, "")
                DataString = DataString & l(IIf(Len(Trim(RepDesc)), RepDesc, "No reject reason"), 60) 'Reject Reason RepDescription1
            ElseIf Option3.Value = True Then
                DataString = DataString & l(IIf(Len(CLMBordx!DocRemark), CLMBordx!DocRemark, ""), 60)             'Pending Reason
            End If
            DataString = DataString & l(Format(CLMBordx!CASDateAdmitted, "ddmmyyyy"), 8)        'DOA
            
            'BenAppAmt = 0
            'ClaimNo = CLMBordx!CasNumber
            'VersionNo = IIf(Len(CLMBordx!claimversion), CLMBordx!claimversion, "")
            'BordxNo = IIf(Len(CLMBordx!BordxRefNo), CLMBordx!BordxRefNo, "")
            'BordxDate = IIf(Len(CLMBordx!BordxDate), CLMBordx!BordxDate, "")
            'Rcounter = Rcounter + 1
            'row = row + 1
            'Call GSTActApp
            'Call NewZurichRequest
            
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







