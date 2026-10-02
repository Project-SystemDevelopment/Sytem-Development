VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmClaimASCIIMedGen 
   Caption         =   "Generate Claims ASCII"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   7275
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      TabIndex        =   17
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
         TabIndex        =   18
         Top             =   60
         Width           =   6615
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Accepted Cases"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Rejected Cases"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame8 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   7095
      Begin VB.ComboBox cboPLAN 
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
         Left            =   4800
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   1560
         Width           =   2175
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
         ItemData        =   "frmClaimASCIIMedGen.frx":0000
         Left            =   4500
         List            =   "frmClaimASCIIMedGen.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1920
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
         TabIndex        =   3
         Top             =   1200
         Width           =   5535
      End
      Begin MSMask.MaskEdBox txtReceiptDateFr 
         Height          =   285
         Left            =   1440
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
      Begin MSMask.MaskEdBox txtReceiptDateTo 
         Height          =   285
         Left            =   3000
         TabIndex        =   6
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
      Begin VB.Label Label18 
         Caption         =   "Plan"
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
         Left            =   4200
         TabIndex        =   24
         Top             =   1560
         Width           =   915
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
         Left            =   2640
         TabIndex        =   15
         Top             =   1560
         Width           =   195
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
         Left            =   3720
         TabIndex        =   14
         Top             =   1920
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
         TabIndex        =   13
         Top             =   540
         Width           =   1905
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
         Left            =   120
         TabIndex        =   12
         Top             =   1545
         Width           =   1275
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         Index           =   0
         Left            =   120
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   900
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   2760
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
   End
End
Attribute VB_Name = "frmClaimASCIIMedGen"
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

Private Sub cboPAYCode_LostFocus()
Dim StrSql As String
Dim PLN As New Adodb.Recordset

    StrSql = "Select PLNDescription From PLN where PAYCode = '" & Mid(cboPAYCode, 1, 2) & "'"
    PLN.Open StrSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        Do While Not PLN.EOF
            cboPlan.AddItem PLN!PLNDescription
        PLN.MoveNext
        Loop
        PLN.Close
        Set PLN = Nothing
End Sub

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
            
            'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                'If Option1.Value = True Then
                    Call PrincipalASCII
                'Else
                '    Call PrincipalASCIIRejectCase
                'End If
           ' medixmasterconnWS.Close
           ' Set medixmasterconnWS = Nothing
                        
            Screen.MousePointer = vbDefault
            
        End If
    End If

End Sub

Private Sub PrincipalASCII()
Dim StrSql As String
Dim CLMBordx As New Adodb.Recordset
Dim BordxNo As String
Dim BordxDateFr As String
Dim BordxDateTo As String
Dim PAYCode As String
Dim ii As Integer
Dim sql As String
Dim ClaimNo As String
Dim VersionNo As String
Dim STRBordxArc As String
Dim BordxArc As New Adodb.Recordset
Dim BATNo As String
Dim cmd As New Adodb.Command
Dim prm1 As Adodb.Parameter
Dim prm2 As Adodb.Parameter
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
Dim xlApp As excel.Application
Dim xlBook As excel.Workbook
Dim xlsheet As excel.WORKSHEET
Dim Rcounter As Double
Dim row  As Double
Dim tmpTotalGST As Double
        
        tmpTotalGST = 0
        Rcounter = 0
        row = 0

        BordxDateFr = txtReceiptDateFr
        BordxDateTo = txtReceiptDateTo
        'PAYCode = "XL"
        PAYCode = Mid(cboPAYCode, 1, 2)
        '------ Start (Do not Delete this Code) -----'
        'If Mid(cboBordxWSStatus, 1, 1) = "0" Then
        '    strAppendCase = "1"
        'End If
        strAppend = ""
        
        'PLNDescription
        If Option1.Value = True Then
            If Mid(CboExportStatus, 1, 1) = "N" Then
                strAppend = " AND BordxDate >= '" & Format(CDate(BordxDateFr), "yyyy/mm/dd") & "'" & _
                                    " AND BordxDate <= '" & Format(CDate(BordxDateTo), "yyyy/mm/dd") & "'" & _
                                    " AND PAYCode = '" & PAYCode & "' And PLNDescription = '" & cboPlan & "' "
            Else
                strAppend = " AND BordxDate >= '" & Format(CDate(BordxDateFr), "MM/DD/YYYY") & "'" & _
                                    " AND BordxDate <= '" & Format(CDate(BordxDateTo), "MM/DD/YYYY") & "'" & _
                                    " AND PAYCode = '" & PAYCode & "' AND CLMBordxWSASCII = 'Y' "
            End If
        Else
                strAppend = " AND CaseRegDate >= '" & Format(CDate(BordxDateFr), "yyyy/mm/dd") & "'" & _
                                    " AND CaseRegDate <= '" & Format(CDate(BordxDateTo), "yyyy/mm/dd") & "'" & _
                                    " AND PAYCode = '" & PAYCode & "'  And PLNDescription = '" & cboPlan & "' "

            
        End If
       'If Mid(cboBordxWSStatus, 1, 1) = "0" Then
        '    StrAppend = StrAppend + " AND (CLMBordxWSStatus = 'A' or CLMBordxWSStatus = 'N')"
        'End If
              
        '----- End -----'
       ' If Option1.Value = True Then
       '     strAppendCase = 1
       '     Set cmd.ActiveConnection = medixmasterconnWS
       '     cmd.CommandText = "Case_ClaimAsciiMedGen_New"
       '     cmd.CommandType = adCmdStoredProc
       '     cmd.CommandTimeout = 300
       '     Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
       '     cmd.Parameters.Append Prm1
       '     Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
       '     cmd.Parameters.Append Prm2
       '     CLMBordx.CursorLocation = adUseClient
       '     CLMBordx.CursorType = adOpenForwardOnly
       '     CLMBordx.CacheSize = 100
       '     Set CLMBordx = cmd.Execute
       ' Else
       '     strAppendCase = 2
       '     Set cmd.ActiveConnection = medixmasterconnWS
       '     cmd.CommandText = "Case_ClaimAsciiMedGen_New"
       '     cmd.CommandType = adCmdStoredProc
       '     cmd.CommandTimeout = 300
       '     Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
       '     cmd.Parameters.Append Prm1
       '     Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
       '     cmd.Parameters.Append Prm2
       '     CLMBordx.CursorLocation = adUseClient
       '     CLMBordx.CursorType = adOpenForwardOnly
       '     CLMBordx.CacheSize = 100
       '     Set CLMBordx = cmd.Execute
       ' End If
        
        
          If Option1.Value = True Then
            StrSql = "SELECT * From View_Claims_ASCII where 0 = 0 "
            'select * from View_Claims_ASCII where 0 = 0
        Else
            StrSql = "select * from View_ClaimsRejected_ASCII where 0 = 0"
            'select * from View_ClaimsRejected_ASCII where 0 = 0
        End If
        'strsql = "SELECT * From View_Claims_ASCII where 0 = 0 "
        
        If Len(Trim(strAppend)) Then
            StrSql = StrSql + strAppend + "ORDER BY BordxRefNo"
        End If
        CLMBordx.Open StrSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If CLMBordx.EOF Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
            BATNo = GetCLMAsciiMedBatch(Mid(cboPAYCode, 1, 2), Date)
            textField = Trim(Date)
                  
            txtFilName = Trim("MX" & Format(Date, "YYMMDD"))
                
            Open txtLocalFileLocation & "\" & txtFilName & "(" & BATNo & ")" For Output As #1
            'Open txtLocalFileLocation & txtFilName & "(" & BATNo & ")" For Output As #2
            
            Header = Trim("F" & "MX" & BATNo & Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 7, 4) & l("", 174))
            
            Print #1, Trim(Header)
            'Print #2, Trim(Header)
            'Debug.Print Header
        
            'Set xlApp = New Excel.Application
            'Set xlBook = xlApp.Workbooks.Add
            'Set xlsheet = xlBook.Sheets(1)

        Do Until CLMBordx.EOF
            'DataString = ""
            BenAppAmt = 0
            ClaimNo = CLMBordx!CasNumber
            VersionNo = IIf(Len(CLMBordx!claimversion), CLMBordx!claimversion, "")
            BordxNo = IIf(Len(CLMBordx!BordxRefNo), CLMBordx!BordxRefNo, "")
            BordxDate = IIf(Len(CLMBordx!BordxDate), CLMBordx!BordxDate, "")
            
            
            Rcounter = Rcounter + 1
            row = row + 1
            
            DataString = l(BordxNo, 15)
            
            If Option1.Value = True Then
                DataString = DataString & l(Format(BordxDate, "ddmmyyyy"), 8)
            Else
                DataString = DataString & l(Format(CLMBordx!CaseRegDate, "ddmmyyyy"), 8)
            End If
           
            DataString = DataString & l(CLMBordx!PAYCode, 2)
            
            DataString = DataString & l(CLMBordx!MBMNumber, 15)
            
            DataString = DataString & l("00", 2)
            
            DataString = DataString & l(CLMBordx!MBMPolicyNo, 25)
            
            If Len(CLMBordx!MBMPolSubNo1) Then
                DataString = DataString & l(CLMBordx!MBMPolSubNo1, 25)
            Else
                DataString = DataString & l("", 25)
            End If

            If Len(CLMBordx!MBMPOLSubNo2) Then
                DataString = DataString & l(CLMBordx!MBMPOLSubNo2, 25)
            Else
                DataString = DataString & l("", 25)
            End If
            
            DataString = DataString & l(Format(CLMBordx!MBMPayorEffDate, "ddmmyyyy"), 8)
            
            DataString = DataString & l(Format(CLMBordx!MBMPayorEXPDate, "ddmmyyyy"), 8)
            

            
            If Len(CLMBordx!MBMAgentCode) Then
                DataString = DataString & l(CLMBordx!MBMAgentCode, 15)
            Else
                DataString = DataString & l("", 15)
            End If
            
            'DataString = DataString & l(CLMBordx!MBMName, 80)
            DataString = DataString & l(CLMBordx!MBMName, 60)
            
            DataString = DataString & l(CLMBordx!MBMIcBcPp, 20)
            
            DataString = DataString & l(CLMBordx!MBMSex, 1)
            
            If Len(CLMBordx!RACCode) Then
                DataString = DataString & l(CLMBordx!RACCode, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            
            If Len(CLMBordx!MBMDOB) Then
                DataString = DataString & l(Format(CLMBordx!MBMDOB, "ddmmyyyy"), 8)
            Else
                DataString = DataString & l("", 8)
            End If
                        
            'DataString = DataString & l("", 1)
                        
            If CLMBordx!MBMRenewFlag <> "" Then
                DataString = DataString & l(CLMBordx!MBMRenewFlag, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            
            If Len(CLMBordx!MBMTakeover) Then
                DataString = DataString & l(CLMBordx!MBMTakeover, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            
            DataString = DataString & l(CLMBordx!INSCode, 2)
            
            DataString = DataString & l(CLMBordx!PLNPayorPlan, 7)
            
            If Len(CLMBordx!MBMGroupCompany) Then
                DataString = DataString & l(CLMBordx!MBMGroupCompany, 47)
            Else
                DataString = DataString & l("", 47)
            End If
            
                'DataString = DataString & l("", 30)
                DataString = DataString & l("", 50)
                        
            If Len(CLMBordx!AGECode) Then
                DataString = DataString & l(CLMBordx!AGECode, 2)
            Else
                DataString = DataString & l("", 2)
            End If

            If Len(CLMBordx!PatientName) Then
                'DataString = DataString & l(CLMBordx!PatientName, 80)
                DataString = DataString & l(CLMBordx!PatientName, 60)
            Else
                'DataString = DataString & l("", 80)
                DataString = DataString & l("", 60)
            End If
            
            DataString = DataString & l(CLMBordx!PatientIc, 20)
            
            If Len(CLMBordx!PatientSex) Then
                DataString = DataString & l(CLMBordx!PatientSex, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            
            If Len(CLMBordx!PatientIc) Then
                DataString = DataString & l(CLMBordx!PatientAgeband, 2)
            Else
                DataString = DataString & l("", 2)
            End If
            
            If Len(CLMBordx!PatientAgeband) Then
                DataString = DataString & l(Format(CLMBordx!PatientDOB, "ddmmyyyy"), 8)
            Else
                DataString = DataString & l("", 8)
            End If

            
            If Len(CLMBordx!HosName) Then
                DataString = DataString & l(CLMBordx!HosName, 80)
            Else
                DataString = DataString & l("", 80)
            End If
            
            If CLMBordx!FinalDiag1 <> "" Then
                DataString = DataString & l(CLMBordx!FinalDiag1, 150)
            Else
                DataString = DataString & l("", 150)
            End If
            
            If CLMBordx!FinalDiag2 <> "" Then
                DataString = DataString & l(CLMBordx!FinalDiag2, 150)
            Else
                DataString = DataString & l("", 150)
            End If
            
            If CLMBordx!FinalDiag3 <> "" Then
                DataString = DataString & l(CLMBordx!FinalDiag3, 150)
            Else
                DataString = DataString & l("", 150)
            End If
            
            If CLMBordx!CASDateAdmitted <> "" Then
                DataString = DataString & l(Format(CLMBordx!CASDateAdmitted, "ddmmyyyy"), 8)
            Else
                DataString = DataString & l("", 8)
            End If
            
            If CLMBordx!CASDischargeDate <> "" Then
                DataString = DataString & l(Format(CLMBordx!CASDischargeDate, "ddmmyyyy"), 8)
            Else
                DataString = DataString & l("", 8)
            End If
                        
            If CLMBordx!CASAdmissionReason <> "" Then
                DataString = DataString & l(CLMBordx!CASAdmissionReason, 200)
            Else
                DataString = DataString & l("", 200)
            End If
            
            If CLMBordx!CASAttendDoctor1 <> "" Then
                'DataString = DataString & l(CLMBordx!CASAttendDoctor1, 80)
                DataString = DataString & l(CLMBordx!CASAttendDoctor1, 100)
            Else
                'DataString = DataString & l("", 80)
                DataString = DataString & l("", 100)
            End If
            
            If CLMBordx!CASAttendDoctor2 <> "" Then
                DataString = DataString & l(CLMBordx!CASAttendDoctor2, 100)
            Else
                DataString = DataString & l("", 100)
            End If
            
            If CLMBordx!CASAttendDoctor3 <> "" Then
                DataString = DataString & l(CLMBordx!CASAttendDoctor3, 100)
            Else
                DataString = DataString & l("", 100)
            End If

            If CLMBordx!CASAttendDoctor4 <> "" Then
                DataString = DataString & l(CLMBordx!CASAttendDoctor4, 100)
            Else
                DataString = DataString & l("", 100)
            End If
            
            If CLMBordx!CASAttendDoctor5 <> "" Then
                DataString = DataString & l(CLMBordx!CASAttendDoctor5, 100)
            Else
                DataString = DataString & l("", 100)
            End If

            If CLMBordx!ClaimNatureCode <> "" Then
                DataString = DataString & l(CLMBordx!ClaimNatureCode, 5)
            Else
                DataString = DataString & l("", 5)
            End If
            
            If CLMBordx!ModalityCode <> "" Then
                DataString = DataString & l(CLMBordx!ModalityCode, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            
            If CLMBordx!CASStayType <> "" Then
                DataString = DataString & l(CLMBordx!CASStayType, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            
            If CLMBordx!InvPayTo = "H" Then
                DataString = DataString & l("C", 1)
            Else
                DataString = DataString & l("R", 1)
            End If
            
            DataString = DataString & l(CLMBordx!MBMPatientCovID, 2)
            
            DataString = DataString & l(CLMBordx!CasNumber, 15)
            
            If Len(CLMBordx!claimversion) Then
                DataString = DataString & l(CLMBordx!claimversion, 3)
            Else
                DataString = DataString & l("", 3)
            End If
            DataString = DataString & l(Format(CLMBordx!ClaimOpenDate, "ddmmyyyy"), 8)
            
            If IsNumeric(CLMBordx!CLMTotalActual) Then
                DataString = DataString & l(Format(CLMBordx!CLMTotalActual, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMTotalApproved) Then
                DataString = DataString & l(Format(CLMBordx!CLMTotalApproved, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMFinalApprove) Then
                DataString = DataString & l(Format(CLMBordx!CLMFinalApprove, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!LimitExc) Then
                DataString = DataString & l(Format(CLMBordx!LimitExc, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMPreviousBalance) Then
                DataString = DataString & l(Format(CLMBordx!CLMPreviousBalance, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If

            If IsNumeric(CLMBordx!CLMPreviousUndExcess) Then
                DataString = DataString & l(Format(CLMBordx!CLMPreviousUndExcess, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If

            If IsNumeric(CLMBordx!CLMExGratiaAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMExGratiaAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If

            If IsNumeric(CLMBordx!CLMBordxAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMBordxAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            DataString = DataString & l("", 11)
            
            If IsNumeric(CLMBordx!CLMCoPayRB) Then
                tmoCoPayRB = CLMBordx!CLMCoPayRB * -1
            Else
                tmoCoPayRB = 0
            End If
            
            If IsNumeric(CLMBordx!CLMCoPayRB) Then
                DataString = DataString & l(Format(tmoCoPayRB, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            DataString = DataString & l(Format(CLMBordx!CLMExcess, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!INVPaidByMem, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMHosPayColl, "#,##0.00"), 11)
            
            'Limit
            
            DataString = DataString & l(Format(CLMBordx!CLMBoardAmtLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMICUAmtLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMServicesLmt, "#,##0.00"), 11)
                       
            DataString = DataString & l(Format(CLMBordx!CLMTheatreLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMPreHSAmtLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMSurgicalAmtLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMAnaLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMNSPreHosLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMPhysicianLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMPostLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMEmergencyLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMAmbulanceLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMOutpatientLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMMonthlyLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMOrganLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMMedicalLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMPhoneLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMGuardianAmtLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMAdminLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMTaxLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthers1Lmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthers2Lmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMCashAmtLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMBoardTaxAmtLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMICUTaxAmtLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMLodgerTaxAmtLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMBoardDaysLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMICUDaysLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMCashDaysLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMMRILmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMMPreHSAmtLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMCOPreHSSumLmt, "#,##0.00"), 11)
            
            
            DataString = DataString & l(Format(CLMBordx!CLMCOMPreHSSumLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format("", "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMPreHospLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPayLmt, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMHomeNCAmtLmt, "#,##0.00"), 11)
            
            

         
            'Actual
            DataString = DataString & l(Format(CLMBordx!CLMBoardAmtAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMICUAmtAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMServicesAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMTheatreAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMPreHSAmtAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMSurgicalAmtAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMAnaAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMNSPreHosAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMPhysicianAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMPostAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMEmergencyAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMAmbulanceAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMOutpatientAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMMonthlyAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMServicesAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMOrganAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMMedicalAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMPhoneAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMGuardianAmtAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMAdminAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMTaxAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthers1Act, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthers2Act, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMCashAmtAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMBoardTaxAmtAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMICUTaxAmtAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMLodgerTaxAmtAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMMRIAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMMPreHSAmtAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMCOPreHSSumAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMCOMPreHSSumAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format("", "#,##0.00"), 11)
            
            'edit --act.post.treatm
            DataString = DataString & l(Format("", "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMPreHospAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPayAct, "#,##0.00"), 11)
            
            DataString = DataString & l(Format(CLMBordx!CLMHomeNCAmtAct, "#,##0.00"), 11)
            
            
            'Approved
            If IsNumeric(CLMBordx!CLMBoardAmt) Then
                DataString = DataString & l(CLMBordx!CLMBoardAmt, 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMICUAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMICUAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMServices) Then
                DataString = DataString & l(Format(CLMBordx!CLMServices, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'DataString = DataString & l("0.00", 11)
            
            
            If IsNumeric(CLMBordx!CLMTheatre) Then
                DataString = DataString & l(Format(CLMBordx!CLMTheatre, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            '    DataString = DataString & l("0.00", 11)
            
            If IsNumeric(CLMBordx!CLMPreHSAmt) Then
                DataString = DataString & l(CLMBordx!CLMPreHSAmt, 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMSurgicalAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMSurgicalAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMAna) Then
                DataString = DataString & l(Format(CLMBordx!CLMAna, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMNSPreHos) Then
                DataString = DataString & l(Format(CLMBordx!CLMNSPreHos, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMPhysician) Then
                DataString = DataString & l(Format(CLMBordx!CLMPhysician, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMPost) Then
                DataString = DataString & l(Format(CLMBordx!CLMPost, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMEmergency) Then
                DataString = DataString & l(Format(CLMBordx!CLMEmergency, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMAmbulance) Then
                DataString = DataString & l(Format(CLMBordx!CLMAmbulance, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMOutpatient) Then
                DataString = DataString & l(Format(CLMBordx!CLMOutpatient, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMMonthly) Then
                DataString = DataString & l(Format(CLMBordx!CLMMonthly, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'If IsNumeric(CLMBordx!CLMServices) Then
             '   DataString = DataString & l(Format(CLMBordx!CLMServices, "#,##0.00"), 11)
            'Else
            '    DataString = DataString & l("0.00", 11)
            'End If
            
            If IsNumeric(CLMBordx!CLMOrgan) Then
                DataString = DataString & l(Format(CLMBordx!CLMOrgan, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMMedical) Then
                DataString = DataString & l(Format(CLMBordx!CLMMedical, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMPhone) Then
                DataString = DataString & l(Format(CLMBordx!CLMPhone, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMGuardianAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMGuardianAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMAdmin) Then
                DataString = DataString & l(Format(CLMBordx!CLMAdmin, "#,##0.00"), 11)
                
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMTax) Then
                DataString = DataString & l(Format(CLMBordx!CLMTax, "#,##0.00"), 11)
                
                If IsNumeric(CLMBordx!CLMTax) Then
                    tmpTotalGST = CDbl(tmpTotalGST) + CDbl(CLMBordx!CLMTax)
                End If
                
            
            Else
                DataString = DataString & l("0.00", 11)
            End If
                        
            If IsNumeric(CLMBordx!CLMOthers1) Then
                DataString = DataString & l(Format(CLMBordx!CLMOthers1, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMOthers2) Then
                DataString = DataString & l(Format(CLMBordx!CLMOthers2, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMCashAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMCashAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMBoardTaxAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMBoardTaxAmt, "#,##0.00"), 11)
                
                If IsNumeric(CLMBordx!CLMBoardTaxAmt) Then
                    tmpTotalGST = CDbl(tmpTotalGST) + CDbl(CLMBordx!CLMBoardTaxAmt)
                End If
                
            
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMICUTaxAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMICUTaxAmt, "#,##0.00"), 11)
            
                If IsNumeric(CLMBordx!CLMICUTaxAmt) Then
                    tmpTotalGST = CDbl(tmpTotalGST) + CDbl(CLMBordx!CLMICUTaxAmt)
                End If
                            
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMLodgerTaxAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMLodgerTaxAmt, "#,##0.00"), 11)
            
                If IsNumeric(CLMBordx!CLMLodgerTaxAmt) Then
                    tmpTotalGST = CDbl(tmpTotalGST) + CDbl(CLMBordx!CLMLodgerTaxAmt)
                End If
                
            
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMMRI) Then
                DataString = DataString & l(Format(CLMBordx!CLMMRI, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMMPreHSAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMMPreHSAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMCOPreHSSum) Then
                DataString = DataString & l(Format(CLMBordx!CLMCOPreHSSum, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMCOMPreHSSum) Then
                DataString = DataString & l(Format(CLMBordx!CLMCOMPreHSSum, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMAccPostAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMAccPostAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            'DataString = DataString & l("0.00", 11)
            
            If IsNumeric(CLMBordx!CLMPreHosp) Then
                DataString = DataString & l(Format(CLMBordx!CLMPreHosp, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMOthersNotCoPay) Then
                DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPay, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            If IsNumeric(CLMBordx!CLMHomeNCAmt) Then
                DataString = DataString & l(Format(CLMBordx!CLMHomeNCAmt, "#,##0.00"), 11)
            Else
                DataString = DataString & l("0.00", 11)
            End If
            
            DataString = DataString & l("0.00", 11)
            DataString = DataString & l("0.00", 11)
            DataString = DataString & l("0.00", 11)
            DataString = DataString & l("0.00", 11)
            DataString = DataString & l("0.00", 11)
            DataString = DataString & l("0.00", 11)
            DataString = DataString & l("0.00", 11)
            DataString = DataString & l("0.00", 11)
            
            If CLMBordx!ICDCode1 <> "" Then
                DataString = DataString & l(CLMBordx!ICDCode1, 5)
            Else
                DataString = DataString & l("", 5)
            End If
            
            If CLMBordx!ICDCode2 <> "" Then
                DataString = DataString & l(CLMBordx!ICDCode2, 5)
            Else
                DataString = DataString & l("", 5)
            End If
            
            If CLMBordx!ICDCode3 <> "" Then
                DataString = DataString & l(CLMBordx!ICDCode3, 5)
            Else
                DataString = DataString & l("", 5)
            End If
            
            DataString = DataString & l("", 100)
            DataString = DataString & l("", 8)
            
            DataString = DataString & l(Format(tmpTotalGST, "#,##0.00"), 11)
            
            'DataString = DataString & l("50000.00", 11)
            'DataString = DataString & l(Format(CLMBordx!CLMTheatreAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMTheatre, "#,##0.00"), 11)
            
            'DataString = DataString & l("KLL", 8)
            'DataString = DataString & l("0", 11)
            'DataString = DataString & l("CLL", 8)
            'DataString = DataString & l("0", 11)
            'DataString = DataString & l("HS", 8)
            'DataString = DataString & l("0", 11)
            'DataString = DataString & l("OP", 8)
            'DataString = DataString & l("0", 11)
        
        'New Addition for Zurich
           ' If Option1.Value = True Then
           '     DataString = DataString & l("", 50)
           '     DataString = DataString & l("", 50)
           '     DataString = DataString & l("", 50)
           ' Else
           '     If CLMBordx!RepDescription1 <> "" Then
           '         DataString = DataString & l(CLMBordx!RepDescription1, 50)
           '     Else
           '         DataString = DataString & l("", 50)
           '     End If
           '
           '     DataString = DataString & l("", 50)
           '     DataString = DataString & l("", 50)
           ' End If
                

            'DataString = DataString & l(Mid(CLMBordx!CaseClaim, 2, 1), 1)
            
            DataString = DataString & Chr(13)
            '--- Start (Do not delete this)--'
                'Call UpdateWSAscii(ClaimNo, VersionNo)
            '-- end --'
            
            TotalRecord = TotalRecord + 1
            
            CLMBordx.MoveNext
            
            'Debug.Print DataString
            Print #1, DataString
            'Print #2, DataString
        
        Loop
                'DataString = DataString & Chr(13)
                Trailer = "T" & l(TotalRecord, 10) & l("", 164)
                Print #1, Trailer
                cmdGenerate.Enabled = False
                
                'xlBook.SaveAs (Trim(txtLocalFileLocation) & Trim(txtFilName & ".csv"))
                
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
    End If
    CLMBordx.Close
    Set CLMBordx = Nothing
    Close #1
    'Close #2
End Sub
Private Sub Form_Load()
    txtPath = SSRVPath & "ClaimASCI"
    txtLocalFileLocation = "C:\MedicaGenClaimASCII\"
    Call ComboListing
End Sub

Private Sub ComboListing()
Dim PAY As New Adodb.Recordset
Dim cmd As New Adodb.Command

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

