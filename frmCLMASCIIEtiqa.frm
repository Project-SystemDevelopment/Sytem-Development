VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCLMASCIIEtiqa 
   Caption         =   "Claim ASCII Etiqa"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   7335
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   7095
      Begin VB.TextBox txtpolnumber 
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   1960
         Width           =   5535
      End
      Begin VB.ComboBox cboCashType 
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
         TabIndex        =   23
         Top             =   1575
         Width           =   5535
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
         TabIndex        =   6
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
         ItemData        =   "frmCLMASCIIEtiqa.frx":0000
         Left            =   4980
         List            =   "frmCLMASCIIEtiqa.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2280
         Width           =   1995
      End
      Begin MSMask.MaskEdBox txtReceiptDateFr 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   2280
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
         Left            =   3000
         TabIndex        =   8
         Top             =   2280
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
      Begin VB.Label Label26 
         Caption         =   "Policy Number"
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
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label Label18 
         Caption         =   "Cash Type"
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
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   1560
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
         TabIndex        =   18
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
         Index           =   0
         Left            =   120
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label Label26 
         Caption         =   "Bordx Date"
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
         Left            =   120
         TabIndex        =   13
         Top             =   2265
         Width           =   1275
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
         Left            =   4200
         TabIndex        =   11
         Top             =   2280
         Width           =   555
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
         Left            =   2640
         TabIndex        =   10
         Top             =   2280
         Width           =   195
      End
      Begin VB.Label txtFilName 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   840
         Width           =   1905
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Rejected Cases"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Accepted Cases"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Generate Claims ASCII"
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
         Left            =   180
         TabIndex        =   1
         Top             =   60
         Width           =   6615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   3360
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   240
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmCLMASCIIEtiqa"
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
Dim BnfCnt As Integer

Private Sub cboCashType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboCashType.Text, cboCashType.SelStart) + Chr(KeyAscii) + Mid(cboCashType.Text, cboCashType.SelStart + cboCashType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboCashType.ListCount - 1
 
        If NewText = LCase(Left(cboCashType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCashType.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboCashType.Text = ValidValue
                cboCashType.SelStart = 0
                cboCashType.SelLength = Len(ValidValue)
            End If
        End If
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
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim RecordSaved

    TotalRecord = "0"
    BnfCnt = 0
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
    ElseIf cboCashType = "" Then
        MsgBox "Please Select Cash Type", vbOKOnly + vbCritical, DialogBoxTitle
        cboCashType.SetFocus
    ElseIf Trim(txtpolnumber) = "" Then
        MsgBox "Please Enter Policy Number", vbOKOnly + vbCritical, DialogBoxTitle
        txtpolnumber.SetFocus
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation, DialogBoxTitle)
        If RecordSaved = vbYes Then
    
            Screen.MousePointer = vbHourglass
            
          '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                If Option1.Value = True Then
                    Call PrincipalASCII
                'Else
                '    Call PrincipalASCIIRejectCase
                End If
       '     medixmasterconnWS.Close
        '    Set medixmasterconnWS = Nothing
                        
            Screen.MousePointer = vbDefault
            
        End If
    End If

End Sub

Private Sub PrincipalASCII()
Dim strsql As String
Dim CLMBordx As New ADODB.Recordset
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
Dim tmpICUTax As Double
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
Dim dashString As String
Dim tmpLimitExc As Double
Dim strsqlPost As String
Dim rstCLMPost As New ADODB.Recordset
Dim tmpDOD As String
Dim tmpCLMAdminAct As Double
Dim tmpCLMServices As Double

        dashString = ""
        BordxDateFr = txtReceiptDateFr
        BordxDateTo = txtReceiptDateTo
        PAYCode = Mid(cboPAYCode, 1, 2)
        tmpLimitExc = 0
        strAppendCase = 1
        '------ Start (Do not Delete this Code) -----'
        'If Mid(cboBordxWSStatus, 1, 1) = "0" Then
        '    strAppendCase = "1"
        'End If
        strAppend = ""
        If Mid(CboExportStatus, 1, 1) = "N" Then
            strAppend = " AND BordxDate >= '" & Format(CDate(BordxDateFr), "yyyy/mm/dd") & "'" & _
                                " AND BordxDate <= '" & Format(CDate(BordxDateTo), "yyyy/mm/dd") & "'" & _
                                " AND PAYCode = '" & PAYCode & "' And INVPayTo = '" & Mid(cboCashType, 1, 1) & "'" & _
                                " And MBMPolicyNo = '" & Trim(txtpolnumber) & "'"
        Else
            strAppend = " AND BordxDate >= '" & Format(CDate(BordxDateFr), "MM/DD/YYYY") & "'" & _
                                " AND BordxDate <= '" & Format(CDate(BordxDateTo), "MM/DD/YYYY") & "'" & _
                                " AND PAYCode = '" & PAYCode & "' AND CLMBordxWSASCII = 'Y' And INVPayTo = '" & Mid(cboCashType, 1, 1) & "'" & _
                                " And MBMPolicyNo = '" & Trim(txtpolnumber) & "'"
        End If
        
       'If Mid(cboBordxWSStatus, 1, 1) = "0" Then
        '    StrAppend = StrAppend + " AND (CLMBordxWSStatus = 'A' or CLMBordxWSStatus = 'N')"
        'End If
              
        '----- End -----'
                    
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_ClaimAsciiEtiqa"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append Prm2
        CLMBordx.CursorLocation = adUseClient
        CLMBordx.CursorType = adOpenForwardOnly
        CLMBordx.CacheSize = 100
        Set CLMBordx = cmd.Execute
        
         
        'strsql = "SELECT * From View_Claims_ASCII where 0 = 0 "
        
        'If Len(Trim(StrAppend)) Then
            'strsql = strsql + StrAppend + "ORDER BY BordxRefNo"
        'End If
    
    'CLMBordx.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If CLMBordx.EOF Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
            BATNo = GetCLMAsciiEtiqaBatch(Mid(cboPAYCode, 1, 2), Date)
            textField = Trim(Date)
            
            If Mid(cboCashType, 1, 1) = "H" Then
                txtFilName = Trim(Mid(cboPAYCode, 1, 2) & BATNo & "(" & Format(Date, "YYMMDD") & ")" & "-" & "Cashless")
            Else
                txtFilName = Trim(Mid(cboPAYCode, 1, 2) & BATNo & "(" & Format(Date, "YYMMDD") & ")" & "-" & "Reimbursement")
            End If
                
            Open txtLocalFileLocation & "\" & txtFilName For Output As #1
            'Open txtLocalFileLocation & txtFilName & "(" & BATNo & ")" For Output As #2
            
            Header = "T" & vbTab
            'Header = Header & Mid(cboPAYCode, 1, 2) & vbTab
            Header = Header & BATNo & vbTab
            Header = Header & Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 7, 4)
            
            Print #1, Trim(Header)
            'Print #2, Trim(Header)
            'Debug.Print Header
        
        
        Do Until CLMBordx.EOF
            'DataString = ""
            BenAppAmt = 0
            ClaimNo = CLMBordx!CasNumber
            VersionNo = CLMBordx!claimversion
            BordxNo = CLMBordx!BordxRefNo
            BordxDate = CLMBordx!BordxDate
            tmpLimitExc = IIf(IsNumeric(CLMBordx!LimitExc), CLMBordx!LimitExc, 0)
            tmpCLMAdminAct = IIf(IsNumeric(CLMBordx!CLMAdminAct), CLMBordx!CLMAdminAct, 0)
            tmpCLMServices = IIf(IsNumeric(CLMBordx!CLMServices), CLMBordx!CLMServices, 0)
            'tmpCLMServices = CDbl(tmpCLMServices) + CDbl(tmpCLMAdminAct)
            
            DataString = CLMBordx!INVNo & vbTab
            
            DataString = DataString & "A" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & Format(BordxDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & CLMBordx!PAYCode & vbTab
            
            DataString = DataString & CLMBordx!CasNumber & vbTab
            
            DataString = DataString & CLMBordx!CasNumber & "-" & CLMBordx!claimversion & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & CLMBordx!MBMPolicyNo & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
                        
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
                        
            DataString = DataString & CLMBordx!MBMName & vbTab
            
            dashString = Replace(CLMBordx!MBMIcBcPp, "-", "")
            dashString = Replace(dashString, " ", "")
            'dashString = dashString.Replace("-", " ")
            dashString = Trim(dashString)
            'DataString = DataString & Mid(CLMBordx!MBMIcBcPp, 1, 6) & Mid(CLMBordx!MBMIcBcPp, 8, 2) & Mid(CLMBordx!MBMIcBcPp, 10, 4) & vbTab
            
            If Len(dashString) >= 12 Then
                DataString = DataString & dashString & vbTab
            Else
                DataString = DataString & "" & vbTab
            End If
            
            If Len(dashString) < 12 Then
                DataString = DataString & "Old IC" & vbTab
            Else
                DataString = DataString & "" & vbTab
            End If


            'DataString = DataString & "" & vbTab
            
            If Len(dashString) < 12 Then
                DataString = DataString & dashString & vbTab
            Else
                DataString = DataString & "" & vbTab
            End If
            'DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            If Len(CLMBordx!PLNPayorPlan) And Not IsNull(CLMBordx!PLNPayorPlan) Then
                DataString = DataString & CLMBordx!PLNPayorPlan & vbTab
            Else
                DataString = DataString & "" & vbTab
            End If
            'DataString = DataString & CLMBordx!PLNDescription & vbTab
            
            'DataString = DataString & "PLAN1" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & CLMBordx!PatientName & vbTab
                                
            DataString = DataString & "" & vbTab
        
            'DataString = DataString & Mid(CLMBordx!PatientIC, 1, 6) & Mid(CLMBordx!PatientIC, 7, 2) & Mid(CLMBordx!PatientIC, 10, 4) & vbTab
            dashString = Replace(CLMBordx!PatientIc, "-", "")
            dashString = Replace(dashString, " ", "")
            'dashString = dashString.Replace("-", " ")
            dashString = Trim(dashString)
            
            If Len(dashString) >= 12 Then
                DataString = DataString & dashString & vbTab
            Else
                DataString = DataString & "" & vbTab
            End If
            
            'DataString = DataString & CLMBordx!PatientIC & vbTab
            'ClaimNo = Mid(CLMBordx!PatientIC, 1, 6) & Mid(CLMBordx!PatientIC, 6, 2) & Mid(CLMBordx!PatientIC, 10, 4) & vbTab
            If Len(dashString) < 12 Then
                DataString = DataString & "Old IC" & vbTab
            Else
                DataString = DataString & "" & vbTab
            End If
            
            'DataString = DataString & "" & vbTab
            
            If Len(dashString) < 12 Then
                DataString = DataString & dashString & vbTab
            Else
                DataString = DataString & "" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & CLMBordx!EtiqaHOSCode & vbTab
            
            DataString = DataString & CLMBordx!HosName & vbTab
            'DataString = DataString & "PH27" & vbTab
            'DataString = DataString & "KUANTAN MEDICAL CENTRE" & vbTab
            
            DataString = DataString & "" & vbTab

            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & CLMBordx!ICDCode1 & vbTab
            
            DataString = DataString & CLMBordx!ICDCode2 & vbTab
           
            DataString = DataString & CLMBordx!ICDCode3 & vbTab
            
            'DataString = DataString & "A03" & vbTab
            'DataString = DataString & "" & vbTab
            'DataString = DataString & "" & vbTab
            
            If CLMBordx!CASDateAdmitted <> "" Then
                DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
            Else
                DataString = DataString & "" & vbTab
            End If
            
            If CLMBordx!CASDischargeDate <> "" Then
                DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            Else
                DataString = DataString & "" & vbTab
            End If
                        
            If IsNumeric(CLMBordx!CLMTotalActual) Then
                DataString = DataString & Format(CLMBordx!CLMTotalActual, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
                        
            If IsNumeric(CLMBordx!CLMBordxAmt) Then
                DataString = DataString & CLMBordx!CLMBordxAmt & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
                        
            If IsNumeric(CLMBordx!CLMBordxAmt) Then
                DataString = DataString & CLMBordx!CLMBordxAmt & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
                        
            DataString = DataString & "" & vbTab

            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
                        
            'DataString = DataString & CLMBordx!ClaimNatureCode & vbTab
            
            'DataString = DataString & CLMBordx!ModalityCode & vbTab
            
            DataString = DataString & "HS" & vbTab
            
            DataString = DataString & CLMBordx!CASCptCode1 & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            
            DataString = DataString & "" & vbTab
            

            'Limit
            
            
        BnfCnt = 0
        If CLMBordx!CLMBoardAmtAct <> 0 And IsNumeric(CLMBordx!CLMBoardAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If
         
        If CLMBordx!CLMICUAmtAct <> 0 And IsNumeric(CLMBordx!CLMICUAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If
            
        If CLMBordx!CLMServicesAct <> 0 And IsNumeric(CLMBordx!CLMServicesAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMTheatreAct <> 0 And IsNumeric(CLMBordx!CLMTheatreAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMPreHSAmtAct <> 0 And IsNumeric(CLMBordx!CLMPreHSAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If


        If CLMBordx!CLMSurgicalAmtAct <> 0 And IsNumeric(CLMBordx!CLMSurgicalAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If


        If CLMBordx!CLMAnaAct <> 0 And IsNumeric(CLMBordx!CLMAnaAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMNSPreHosAct <> 0 And IsNumeric(CLMBordx!CLMNSPreHosAct) Then
            BnfCnt = BnfCnt + 1
        End If


        If CLMBordx!CLMPhysicianAct <> 0 And IsNumeric(CLMBordx!CLMPhysicianAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMPostAct <> 0 And IsNumeric(CLMBordx!CLMPostAct) Then
            BnfCnt = BnfCnt + 1
        End If


        If CLMBordx!CLMEmergencyAct <> 0 And IsNumeric(CLMBordx!CLMEmergencyAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMAmbulanceAct <> 0 And IsNumeric(CLMBordx!CLMAmbulanceAct) Then
            BnfCnt = BnfCnt + 1
        End If


        If CLMBordx!CLMOutpatientAct <> 0 And IsNumeric(CLMBordx!CLMOutpatientAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMMonthlyAct <> 0 And IsNumeric(CLMBordx!CLMMonthlyAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMOrganAct <> 0 And IsNumeric(CLMBordx!CLMOrganAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMMedicalAct <> 0 And IsNumeric(CLMBordx!CLMMedicalAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMPhoneAct <> 0 And IsNumeric(CLMBordx!CLMPhoneAct) Then
            BnfCnt = BnfCnt + 1
            
        End If

        If CLMBordx!CLMGuardianAmtAct <> 0 And IsNumeric(CLMBordx!CLMGuardianAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If


        If CLMBordx!CLMAdminAct <> 0 And IsNumeric(CLMBordx!CLMAdminAct) Then
            BnfCnt = BnfCnt + 1
        End If


        If CLMBordx!CLMTaxAct <> 0 And IsNumeric(CLMBordx!CLMTaxAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMOthers1Act <> 0 And IsNumeric(CLMBordx!CLMOthers1Act) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMCashAmtAct <> 0 And IsNumeric(CLMBordx!CLMCashAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMBoardTaxAmtAct <> 0 And IsNumeric(CLMBordx!CLMBoardTaxAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMICUTaxAmtAct <> 0 And IsNumeric(CLMBordx!CLMICUTaxAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMLodgerTaxAmtAct <> 0 And IsNumeric(CLMBordx!CLMLodgerTaxAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMMRIAct <> 0 And IsNumeric(CLMBordx!CLMMRIAct) Then
            BnfCnt = BnfCnt + 1
        End If

        If CLMBordx!CLMMPreHSAmtAct <> 0 And IsNumeric(CLMBordx!CLMMPreHSAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If
 
        If CLMBordx!CLMMPreHSAmtAct <> 0 And IsNumeric(CLMBordx!CLMMPreHSAmtAct) Then
            BnfCnt = BnfCnt + 1
        End If
 
            
        If CLMBordx!CLMPreHospAct <> 0 And IsNumeric(CLMBordx!CLMPreHospAct) Then
            BnfCnt = BnfCnt + 1
        End If
        
        If CLMBordx!CLMMajorMedFee <> 0 And IsNumeric(CLMBordx!CLMMajorMedFee) Then
            BnfCnt = BnfCnt + 1
        End If
            
        DataString = DataString & BnfCnt & vbTab
            
                        
        If CLMBordx!CLMBoardAmtAct <> 0 And IsNumeric(CLMBordx!CLMBoardAmtAct) Then
            
            DataString = DataString & "H001" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMBoardAmtAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab
            
            If IsNumeric(CLMBordx!CLMBoardAmt) Then
                DataString = DataString & Format(CLMBordx!CLMBoardAmt, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            'DataString = DataString & CLMBordx!CLMBoardDay & vbTab
            DataString = DataString & (CDate(CLMBordx!CASDischargeDate) - CDate(CLMBordx!CASDateAdmitted)) + 1 & vbTab
        End If
         
    If CLMBordx!CLMICUAmtAct <> 0 And IsNumeric(CLMBordx!CLMICUAmtAct) Then
            DataString = DataString & "H003" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMICUAmtAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab
            
            If IsNumeric(CLMBordx!CLMICUAmt) Then
                DataString = DataString & Format(CLMBordx!CLMICUAmt, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If
        
        'CLMBordx!CLMAdminAct
    If tmpCLMServices <> 0 Then
    
            DataString = DataString & "H004" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(tmpCLMServices, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(tmpCLMServices) Then
                If CDbl(tmpLimitExc) < 0 Then
                    If CDbl(tmpCLMServices) > (tmpLimitExc * -1) Then
                        DataString = DataString & Format(CDbl(tmpCLMServices) + CDbl(tmpLimitExc), "###0.00") & vbTab
                        tmpLimitExc = 0
                    Else
                        DataString = DataString & Format("0", "###0.00") & vbTab
                        tmpLimitExc = CDbl(tmpCLMServices) + CDbl(tmpLimitExc)
                    End If
                Else
                    DataString = DataString & Format(CLMBordx!CLMServices, "###0.00") & vbTab
                End If
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

    If CLMBordx!CLMTheatreAct <> 0 And IsNumeric(CLMBordx!CLMTheatreAct) Then
            DataString = DataString & "H011" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMTheatreAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMTheatre) Then
                DataString = DataString & Format(CLMBordx!CLMTheatre, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If


    If CLMBordx!CLMPreHSAmtAct <> 0 And IsNumeric(CLMBordx!CLMPreHSAmtAct) Then
            DataString = DataString & "H005" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMPreHSAmtAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMPreHSAmt) Then
                DataString = DataString & Format(CLMBordx!CLMPreHSAmt, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            DataString = DataString & "" & vbTab
            
        End If


    If CLMBordx!CLMSurgicalAmtAct <> 0 And IsNumeric(CLMBordx!CLMSurgicalAmtAct) Then
            DataString = DataString & "H008" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMSurgicalAmtAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMSurgicalAmt) Then
                DataString = DataString & Format(CLMBordx!CLMSurgicalAmt, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            DataString = DataString & "" & vbTab
            
        End If


    If CLMBordx!CLMAnaAct <> 0 And IsNumeric(CLMBordx!CLMAnaAct) Then
            DataString = DataString & "H010" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMAnaAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMAna) Then
                DataString = DataString & Format(CLMBordx!CLMAna, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            DataString = DataString & "" & vbTab
            
        End If

    If CLMBordx!CLMNSPreHosAct <> 0 And IsNumeric(CLMBordx!CLMNSPreHosAct) Then
            DataString = DataString & "H005" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMNSPreHosAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMNSPreHos) Then
                DataString = DataString & Format(CLMBordx!CLMNSPreHos, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If


    If CLMBordx!CLMPhysicianAct <> 0 And IsNumeric(CLMBordx!CLMPhysicianAct) Then
            
            DataString = DataString & "H012" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMPhysicianAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMPhysician) Then
                DataString = DataString & Format(CLMBordx!CLMPhysician, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

        If CLMBordx!CLMPostAct <> 0 And IsNumeric(CLMBordx!CLMPostAct) Then
            strsqlPost = ""
            strsqlPost = " Select CASNUmber, INVWSversion, CLMPost,InvDate " & _
                    " From View_Search_PrePost where CASNumber ='" & Trim(ClaimNo) & "' " & _
                    " AND INVWSversion = " & VersionNo
            rstCLMPost.Open strsqlPost, medixmasterconnWS, adOpenKeyset, adLockOptimistic

            Do While Not rstCLMPost.EOF
                DataString = DataString & "H013" & vbTab
                'DataString = DataString & (CDate(CLMBordx!CASDischargeDate) - CDate(CLMBordx!CASDateAdmitted)) + 1 & vbTab
                tmpDOD = CDate(CLMBordx!CASDischargeDate) + 1
                DataString = DataString & Format(rstCLMPost!InvDate, "ddmmyyyy") & vbTab
                            
                DataString = DataString & Format(CLMBordx!CLMPostAct, "###0.00") & vbTab
                
                DataString = DataString & Format(rstCLMPost!InvDate, "ddmmyyyy") & vbTab
                
                DataString = DataString & "" & vbTab
    
                If IsNumeric(CLMBordx!CLMPostAct) Then
                    DataString = DataString & Format(CLMBordx!CLMPostAct, "###0.00") & vbTab
                Else
                    DataString = DataString & "0.00" & vbTab
                End If
                
                DataString = DataString & "" & vbTab
                rstCLMPost.MoveNext
            Loop
            
            rstCLMPost.Close
            Set rstCLMPost = Nothing
        End If


    If CLMBordx!CLMEmergencyAct <> 0 And IsNumeric(CLMBordx!CLMEmergencyAct) Then
            
            DataString = DataString & "H014" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMEmergencyAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMEmergency) Then
                DataString = DataString & Format(CLMBordx!CLMEmergency, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

    If CLMBordx!CLMAmbulanceAct <> 0 And IsNumeric(CLMBordx!CLMAmbulanceAct) Then
            
            DataString = DataString & "H016" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMAmbulanceAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMAmbulance) Then
                DataString = DataString & Format(CLMBordx!CLMAmbulance, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If


    If CLMBordx!CLMOutpatientAct <> 0 And IsNumeric(CLMBordx!CLMOutpatientAct) Then
            
            DataString = DataString & "H029" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMOutpatientAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMOutpatient) Then
                DataString = DataString & Format(CLMBordx!CLMOutpatient, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

    If CLMBordx!CLMMonthlyAct <> 0 And IsNumeric(CLMBordx!CLMMonthlyAct) Then
            
            DataString = DataString & "H030" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMMonthlyAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMMonthly) Then
                DataString = DataString & Format(CLMBordx!CLMMonthly, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

    If CLMBordx!CLMOrganAct <> 0 And IsNumeric(CLMBordx!CLMOrganAct) Then
            
            DataString = DataString & "H028" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMOrganAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMOrgan) Then
                DataString = DataString & Format(CLMBordx!CLMOrgan, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

    If CLMBordx!CLMMedicalAct <> 0 And IsNumeric(CLMBordx!CLMMedicalAct) Then
            
            DataString = DataString & "H024" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMMedicalAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab


            If IsNumeric(CLMBordx!CLMMedical) Then
                DataString = DataString & Format(CLMBordx!CLMMedical, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

    If CLMBordx!CLMPhoneAct <> 0 And IsNumeric(CLMBordx!CLMPhoneAct) Then
            
            DataString = DataString & "H036" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMPhoneAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab


            If IsNumeric(CLMBordx!CLMPhone) Then
                DataString = DataString & CLMBordx!CLMPhone & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

        If CLMBordx!CLMGuardianAmtAct <> 0 And IsNumeric(CLMBordx!CLMGuardianAmtAct) Then
            
            
            DataString = DataString & "H062" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMGuardianAmtAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            
            If IsNumeric(CLMBordx!CLMGuardianAmt) Then
                DataString = DataString & Format(CLMBordx!CLMGuardianAmt, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If


        If CLMBordx!CLMAdminAct <> 0 And IsNumeric(CLMBordx!CLMAdminAct) Then
            
            DataString = DataString & "H056" & vbTab
           
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
        
            DataString = DataString & Format(CLMBordx!CLMAdminAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab


             
            If IsNumeric(CLMBordx!CLMAdmin) Then
                DataString = DataString & Format(CLMBordx!CLMAdmin, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
           
            DataString = DataString & "" & vbTab
            
        End If



        If CLMBordx!CLMOthers1Act <> 0 And IsNumeric(CLMBordx!CLMOthers1Act) Then
        
            
            DataString = DataString & "H036" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMOthers1Act, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMOthers1) Then
                DataString = DataString & Format(CLMBordx!CLMOthers1, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If


           ' DataString = DataString & CLMBordx!CLMOthers2Act & vbTab
           '
           ' DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
           '

           ' If IsNumeric(CLMBordx!CLMOthers2) Then
           '     DataString = DataString & CLMBordx!CLMOthers2 & vbTab
           ' Else
           '     DataString = DataString & "0.00" & vbTab
           ' End If
            
        'If CLMBordx!CLMCashAmtAct <> 0 And IsNumeric(CLMBordx!CLMCashAmtAct) Then
        '
        '    DataString = DataString & "H017" & vbTab
        '
        '    DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
        '
        '    DataString = DataString & Format(CLMBordx!CLMCashAmtAct, "###0.00") & vbTab
        '
        '    DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
        '
        '    DataString = DataString & "" & vbTab

        '    If IsNumeric(CLMBordx!CLMCashAmt) Then
        '        DataString = DataString & Format(CLMBordx!CLMCashAmt, "###0.00") & vbTab
        '    Else
        '        DataString = DataString & "0.00" & vbTab
        '    End If
        
         '   DataString = DataString & "" & vbTab
            
       ' End If
        
        If CLMBordx!CLMTaxAct <> 0 And IsNumeric(CLMBordx!CLMTaxAct) Then
            
            
            DataString = DataString & "H017" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMTaxAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMTax) Then
                DataString = DataString & Format(CLMBordx!CLMTax, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
                        
            DataString = DataString & "" & vbTab
                        
        End If
        

        If CLMBordx!CLMBoardTaxAmtAct <> 0 And IsNumeric(CLMBordx!CLMBoardTaxAmtAct) Then
            
            DataString = DataString & "H025" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMBoardTaxAmtAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

             If IsNumeric(CLMBordx!CLMBoardTaxAmt) Then
                DataString = DataString & Format(CLMBordx!CLMBoardTaxAmt, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

        If CLMBordx!CLMICUTaxAmtAct <> 0 And IsNumeric(CLMBordx!CLMICUTaxAmtAct) Then
        
            DataString = DataString & "H058" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMICUTaxAmtAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMICUTaxAmt) Then
                DataString = DataString & Format(CLMBordx!CLMICUTaxAmt, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
        
            DataString = DataString & "" & vbTab
            
        End If

        If CLMBordx!CLMLodgerTaxAmtAct <> 0 And IsNumeric(CLMBordx!CLMLodgerTaxAmtAct) Then
            
            DataString = DataString & "H057" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMLodgerTaxAmtAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

           If IsNumeric(CLMBordx!CLMLodgerTaxAmt) Then
                DataString = DataString & Format(CLMBordx!CLMLodgerTaxAmt, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

        If CLMBordx!CLMMRIAct <> 0 And IsNumeric(CLMBordx!CLMMRIAct) Then
            
            DataString = DataString & "H055" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMMRIAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMMRI) Then
                DataString = DataString & Format(CLMBordx!CLMMRI, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If

        If CLMBordx!CLMMPreHSAmtAct <> 0 And IsNumeric(CLMBordx!CLMMPreHSAmtAct) Then
            
            DataString = DataString & "H005" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMMPreHSAmtAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            DataString = DataString & Format(CLMBordx!CLMMPreHSAmtAct, "###0.00") & vbTab
                        
           ' DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab
            
        End If
 
        If CLMBordx!CLMMPreHSAmtAct <> 0 And IsNumeric(CLMBordx!CLMMPreHSAmtAct) Then
            
            DataString = DataString & "H005" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & CLMBordx!CLMMPreHSAmtAct & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMMPreHSAmt) Then
                DataString = DataString & CLMBordx!CLMMPreHSAmt & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            
            DataString = DataString & "" & vbTab
            
        End If
 
            'DataString = DataString & CLMBordx!CLMCOMPreHSSumAct & vbTab
            
            'DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
 
           'If IsNumeric(CLMBordx!CLMCOPreHSSum) Then
            '    DataString = DataString & CLMBordx!CLMCOPreHSSum & vbTab
            'Else
            '    DataString = DataString & "0.00" & vbTab
            'End If

        If CLMBordx!CLMPreHospAct <> 0 And IsNumeric(CLMBordx!CLMPreHospAct) Then
            
            DataString = DataString & "H005" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & Format(CLMBordx!CLMPreHospAct, "###0.00") & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMPreHosp) Then
                DataString = DataString & Format(CLMBordx!CLMPreHosp, "###0.00") & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            DataString = DataString & "" & vbTab
            
        End If
        
        If CLMBordx!CLMMajorMedFee <> 0 And IsNumeric(CLMBordx!CLMMajorMedFee) Then
            DataString = DataString & "H009" & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "ddmmyyyy") & vbTab
                        
            DataString = DataString & CLMBordx!CLMMajorMedFee & vbTab
            
            DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            DataString = DataString & "" & vbTab

            If IsNumeric(CLMBordx!CLMMajorMedFee) Then
                DataString = DataString & CLMBordx!CLMMajorMedFee & vbTab
            Else
                DataString = DataString & "0.00" & vbTab
            End If
            DataString = DataString & "" & vbTab
            
        End If


            'DataString = DataString & CLMBordx!CLMOthersNotCoPayAct & vbTab
                        
            'DataString = DataString & Format(CLMBordx!CASDischargeDate, "ddmmyyyy") & vbTab
            
            'DataString = DataString & CLMBordx!CLMHomeNCAmtAct & vbTab
            
            
            'Approved
 

            'If IsNumeric(CLMBordx!CLMOthersNotCOPay) Then
            '    DataString = DataString & CLMBordx!CLMOthersNotCOPay & vbTab
            'Else
            '    DataString = DataString & "0.00" & vbTab
            'End If
            
            'If IsNumeric(CLMBordx!CLMHomeNCAmt) Then
            '    DataString = DataString & CLMBordx!CLMHomeNCAmt & vbTab
            'Else
            '    DataString = DataString & "0.00" & vbTab
            'End If
            
           ' DataString = DataString & l("0.00", 11)
           ' DataString = DataString & l("0.00", 11)
           ' DataString = DataString & l("0.00", 11)
           ' DataString = DataString & l("0.00", 11)
           ' DataString = DataString & l("0.00", 11)
           ' DataString = DataString & l("0.00", 11)
           ' DataString = DataString & l("0.00", 11)
           ' DataString = DataString & l("0.00", 11)
            
            'If CLMBordx!ICDCode1 <> "" Then
            '    DataString = DataString & l(CLMBordx!ICDCode1, 5)
            'Else
            '    DataString = DataString & l("", 5)
            'End If
            
           ' If CLMBordx!ICDCode2 <> "" Then
           '     DataString = DataString & l(CLMBordx!ICDCode2, 5)
           ' Else
           '     DataString = DataString & l("", 5)
           ' End If
            
           ' If CLMBordx!ICDCode3 <> "" Then
           '     DataString = DataString & l(CLMBordx!ICDCode3, 5)
           ' Else
           '     DataString = DataString & l("", 5)
           ' End If
            
            'DataString = DataString & l("", 100)
            'DataString = DataString & l("", 8)
            
              
            'DataString = DataString & Chr(13)
            
            '--- Start (Do not delete this)--'
                'Call UpdateWSAscii(ClaimNo, VersionNo)
            '-- end --'
            
            TotalRecord = TotalRecord + 1
            
            CLMBordx.MoveNext
            
            'Debug.Print DataString
            Print #1, DataString
            'Print #2, DataString
        
        Loop
            
            Trailer = "T" & l(TotalRecord, 10) & l("", 164)
                'Debug.Print Trailer
                Print #1, Trailer
                'Print #2, Trailer
                cmdGenerate.Enabled = False
                Close #1
                'Close #2
                
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
    End If
    CLMBordx.Close
    Set CLMBordx = Nothing
    Close #1
    Close #2
End Sub
Private Sub Form_Load()
    txtPath = SSRVPath & "ClaimASCI"
    txtLocalFileLocation = "C:\MedicaGenClaimASCII\"
    Call ComboListing
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
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
        
 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
    CboExportStatus.AddItem "N - New"
    CboExportStatus.AddItem "R - Reprint"
    CboExportStatus.Text = "N - New"
    
    cboCashType.AddItem "H - Hospital"
    cboCashType.AddItem "M - Member"
    
End Sub


Private Sub txtReceiptDatefr_LostFocus()
    If IsDate(txtReceiptDateFr) Then
    Else
        If txtReceiptDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDateFr.SetFocus
        End If
    End If

End Sub

Private Sub txtReceiptDateTo_LostFocus()
    If IsDate(txtReceiptDateTo) Then
    Else
        If txtReceiptDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDateTo.SetFocus
        End If
    End If
End Sub
