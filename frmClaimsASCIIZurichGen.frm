VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmClaimsASCIIZurichGen 
   Caption         =   "General Claims ASCII Zurich (General Insurance)"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   7290
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   7095
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
         ItemData        =   "frmClaimsASCIIZurichGen.frx":0000
         Left            =   4980
         List            =   "frmClaimsASCIIZurichGen.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   1200
         Width           =   5535
      End
      Begin MSMask.MaskEdBox txtReceiptDateFr 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         Left            =   4200
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   900
         Width           =   1245
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
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "General Claims ASCII Zurich (General Insurance)"
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
Attribute VB_Name = "frmClaimsASCIIZurichGen"
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
           ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
           ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                'If Option1.Value = True Then
                    Call PrincipalASCII
                    Call CaseASCII
                'Else
                '    Call PrincipalASCIIRejectCase
                'End If
          '  medixmasterconn.Close
         '   Set medixmasterconn = Nothing
          '  medixmasterconnWS.Close
          '  Set medixmasterconnWS = Nothing
                        
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
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlsheet As Excel.WORKSHEET
Dim Rcounter, row As Double
Dim txtSVRfile As String
Dim BNF As New ADODB.Recordset
Dim strsqlbnf As String
Dim tmpRBBnf As String
Dim tmpICUBnf As String
Dim tmpHSSBnf As String
Dim tmpOTBnf As String
Dim tmpPreHosBnf As String
Dim tmpSurgFeeBnf As String
Dim tmpAnaeFeeBnf As String
Dim tmpPreHosSpeBnf As String
Dim tmpDailyInBnf As String
Dim tmpPostBnf As String
Dim tmpOPAccBnf As String
Dim tmpOPAmbulanceBnf As String
Dim tmpPhysioBnf As String
Dim tmpKidneyBnf As String
Dim tmpOrganBnf As String
Dim tmpGovHosIncBnf As String
Dim tmpGovHosTaxBnf As String
Dim tmpGuardianAllBnf As String
Dim tmpMedRptBnf As String
Dim tmpOthersAct As String
Dim tmpOthersApp As String
Dim tmpReimPreH60 As String
Dim tmpReimPreGP As String
Dim tmpReimPreMedic As String
Dim tmpReimPreMR As String
Dim tmpReimPreSpec As String
Dim tmpReimPostHosp As String
Dim tmpReimPostNotSame As String
Dim tmpReimOthers As String
Dim tmpReimActOthers As String
Dim tmpCLMCoPayRB As String
Dim tmpTaxAct As String
Dim tmpTaxApp As String
Dim tmpNotCoPayAct As String
Dim tmpNotCoPayApp As String
Dim tmpICUTaxAct As String
Dim tmpICUTaxApp As String

        BordxDateFr = txtReceiptDateFr
        BordxDateTo = txtReceiptDateTo
        'PAYCode = "XL"
        PAYCode = Mid(cboPAYCode, 1, 2)
        '------ Start (Do not Delete this Code) -----'
        'If Mid(cboBordxWSStatus, 1, 1) = "0" Then
        '    strAppendCase = "1"
        'End If
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
       'If Mid(cboBordxWSStatus, 1, 1) = "0" Then
        '    StrAppend = StrAppend + " AND (CLMBordxWSStatus = 'A' or CLMBordxWSStatus = 'N')"
        'End If
              
        '----- End -----'
        If Option1.Value = True Then
            strAppendCase = 1
            Set cmd.ActiveConnection = medixmasterconnWS
            cmd.CommandText = "Case_ClaimAsciiMedGen_New"
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
        Else
            strAppendCase = 2
            Set cmd.ActiveConnection = medixmasterconnWS
            cmd.CommandText = "Case_ClaimAsciiMedGen_New"
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
        End If
         
        'strsql = "SELECT * From View_Claims_ASCII where 0 = 0 "
        
        'If Len(Trim(StrAppend)) Then
            'strsql = strsql + StrAppend + "ORDER BY BordxRefNo"
        'End If
    
    'CLMBordx.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If CLMBordx.EOF Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
            Rcounter = 0
            row = 0
            
            BATNo = GetCLMAsciiMedBatch(Mid(cboPAYCode, 1, 2), Date)
            textField = Trim(Date)
                  
            txtFilName = Trim("R001" & Mid(cboPAYCode, 1, 2) & Format(Date, "YYYYMMDD") & "001")
            Open txtLocalFileLocation & txtFilName & ".csv" For Output As #1
            'Open txtLocalFileLocation & "\" & txtFilName & "(" & BATNo & ")" For Output As #1
            'Open txtLocalFileLocation & txtFilName & "(" & BATNo & ")" For Output As #2
            
            'Header = Trim("F" & "MX" & BATNo & Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 7, 4) & l("", 174))
            
            txtSVRfile = (Trim(txtLocalFileLocation) & Trim(txtFilName & ".csv"))
            
            'Set xlApp = New Excel.Application
            'Set xlBook = xlApp.Workbooks.Add
            'Set xlsheet = xlBook.Sheets(1)

        Do Until CLMBordx.EOF
        
            strsqlbnf = "Select *  from BNF where INSCode = '" & Trim(CLMBordx!INSCode) & "' AND PLNCode = '" & Trim(CLMBordx!PLNCode) & "'"
            BNF.Open strsqlbnf, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                      
            Do While Not BNF.EOF
                If BNF!BNFCode = "101" Then
                    tmpRBBnf = BNF!BNFClaimAbleAmount
                End If
                If BNF!BNFCode = "102" Then
                    tmpICUBnf = IIf(Len(BNF!BNFClaimAbleAmount), BNF!BNFClaimAbleAmount, "0")
                End If
                If BNF!BNFCode = "103" Then
                    tmpHSSBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "104" Then
                    tmpOTBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "201" Then
                    tmpPreHosBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "203" Then
                    tmpSurgFeeBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "204" Then
                    tmpAnaeFeeBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "301" Then
                    tmpPreHosSpeBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "302" Then
                    tmpDailyInBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "303" Then
                    tmpPostBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "401" Then
                    tmpOPAccBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "402" Then
                    tmpOPAmbulanceBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "403" Then
                    tmpPhysioBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "404" Then
                    tmpKidneyBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "501" Then
                    tmpOrganBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "601" Then
                    tmpGovHosIncBnf = BNF!BNFClaimAbleAmount
                End If
                If BNF!BNFCode = "602" Then
                    tmpGovHosTaxBnf = IIf(BNF!BNFRMperDis = "FULL", "As Charge", BNF!BNFRMperDis)
                End If
                If BNF!BNFCode = "701" Then
                    tmpGuardianAllBnf = IIf(Len(BNF!BNFClaimAbleAmount), BNF!BNFClaimAbleAmount, "0")
                End If
                If BNF!BNFCode = "801" Then
                    tmpMedRptBnf = BNF!BNFRMperDis
                End If
                
                
            BNF.MoveNext
            Loop
            BNF.Close
            Set BNF = Nothing
            
            'DataString = ""
            BenAppAmt = 0
            ClaimNo = CLMBordx!CasNumber
            VersionNo = IIf(Len(CLMBordx!claimversion), CLMBordx!claimversion, "")
            BordxNo = IIf(Len(CLMBordx!BordxRefNo), CLMBordx!BordxRefNo, "")
            BordxDate = IIf(Len(CLMBordx!BordxDate), CLMBordx!BordxDate, "")
            
            Rcounter = Rcounter + 1
            row = row + 1
            
            'xlApp.DisplayAlerts = False
                       
            DataString = " " & ","
            DataString = DataString & CLMBordx!CasNumber & ","
            DataString = DataString & ClaimNo & "-" & VersionNo & ","
            DataString = DataString & BordxNo & ","
            DataString = DataString & Format(CLMBordx!AnnLimit, "###0.00") & ","
            DataString = DataString & "" & ","
            DataString = DataString & "" & ","
            
            If Len(tmpRBBnf) Then
                DataString = DataString & tmpRBBnf & " PER DAY" & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMBoardAmtAct) Then
                DataString = DataString & Format(CLMBordx!CLMBoardAmtAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMBoardAmt) Then
                DataString = DataString & Format(CLMBordx!CLMBoardAmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If CLMBordx!CLMBoardAmt < CLMBordx!CLMBoardAmtAct Then
                DataString = DataString & Format(CLMBordx!CLMBoardAmtAct, "###0.00") - Format(CLMBordx!CLMBoardAmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMICUAmtLmt) Then
                DataString = DataString & tmpICUBnf & " PER DAY" & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMICUAmtAct) Then
                DataString = DataString & Format(CLMBordx!CLMICUAmtAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMICUAmt) Then
                DataString = DataString & Format(CLMBordx!CLMICUAmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMICUAmt) = False And IsNumeric(CLMBordx!CLMICUAmtAct) = True Then
                DataString = DataString & Format(CLMBordx!CLMICUAmtAct, "###0.00") & ","
            ElseIf IsNumeric(CLMBordx!CLMICUAmt) = True And IsNumeric(CLMBordx!CLMICUAmtAct) = True Then
                If CLMBordx!CLMICUAmt < CLMBordx!CLMICUAmtAct Then
                    If IsNumeric(CLMBordx!CLMICUAmt) And IsNumeric(CLMBordx!CLMICUAmtAct) Then
                        DataString = DataString & Format(CLMBordx!CLMICUAmtAct, "###0.00") - Format(CLMBordx!CLMICUAmt, "###0.00") & ","
                    ElseIf IsNumeric(CLMBordx!CLMICUAmt) = False And IsNumeric(CLMBordx!CLMICUAmtAct) = True Then
                        DataString = DataString & Format(CLMBordx!CLMICUAmtAct, "###0.00") & ","
                    End If
                Else
                    DataString = DataString & "0" & ","
                End If
            Else
                DataString = DataString & "0" & ","
            End If
            If tmpHSSBnf <> "" Then
                DataString = DataString & tmpHSSBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMServicesAct) Then
                DataString = DataString & Format(CLMBordx!CLMServicesAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMServices) Then
                DataString = DataString & Format(CLMBordx!CLMServices, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMServices) = False And IsNumeric(CLMBordx!CLMServicesAct) = True Then
                DataString = DataString & Format(CLMBordx!CLMServicesAct, "###0.00") & ","
            ElseIf IsNumeric(CLMBordx!CLMServices) = True And IsNumeric(CLMBordx!CLMServicesAct) = True Then
                If CDbl(CLMBordx!CLMServices) < CDbl(CLMBordx!CLMServicesAct) Then
                    DataString = DataString & Format(CLMBordx!CLMServicesAct, "###0.00") - Format(CLMBordx!CLMServices, "###0.00") & ","
                Else
                    DataString = DataString & "0" & ","
                End If
            Else
                DataString = DataString & "0" & ","
            End If
            If tmpPreHosBnf <> "" Then
                DataString = DataString & tmpPreHosBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMPreHSAmtAct) Then
                DataString = DataString & Format(CLMBordx!CLMPreHSAmtAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMPreHSAmt) Then
                DataString = DataString & Format(CLMBordx!CLMPreHSAmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If CLMBordx!CLMPreHSAmt < CLMBordx!CLMPreHSAmtAct Then
                DataString = DataString & Format(CLMBordx!CLMPreHSAmtAct, "###0.00") - Format(CLMBordx!CLMPreHSAmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If tmpSurgFeeBnf <> "" Then
                DataString = DataString & tmpSurgFeeBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMSurgicalAmtAct) Then
                DataString = DataString & Format(CLMBordx!CLMSurgicalAmtAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMSurgicalAmt) Then
                DataString = DataString & Format(CLMBordx!CLMSurgicalAmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If CLMBordx!CLMSurgicalAmt < CLMBordx!CLMSurgicalAmtAct Then
                DataString = DataString & Format(CLMBordx!CLMSurgicalAmtAct, "###0.00") - Format(CLMBordx!CLMSurgicalAmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            If tmpAnaeFeeBnf <> "" Then
                DataString = DataString & tmpAnaeFeeBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMAnaAct) Then
                DataString = DataString & Format(CLMBordx!CLMAnaAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMAna) Then
                DataString = DataString & Format(CLMBordx!CLMAna, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            'If CLMBordx!CLMAna < CLMBordx!CLMAnaAct Then
            '    DataString = DataString & Format(CLMBordx!CLMAnaAct, "###0.00") - Format(CLMBordx!CLMAna, "###0.00") & ","
            'Else
            '    DataString = DataString & "0" & ","
            'End If
            
            If IsNumeric(CLMBordx!CLMAna) = False And IsNumeric(CLMBordx!CLMAnaAct) = True Then
                DataString = DataString & Format(CLMBordx!CLMAnaAct, "###0.00") & ","
            ElseIf IsNumeric(CLMBordx!CLMAna) = True And IsNumeric(CLMBordx!CLMAnaAct) = True Then
                If CDbl(CLMBordx!CLMAna) < CDbl(CLMBordx!CLMAnaAct) Then
                    DataString = DataString & Format(CDbl(CLMBordx!CLMAnaAct), "###0.00") - Format(CDbl(CLMBordx!CLMAna), "###0.00") & ","
                Else
                    DataString = DataString & "0" & ","
                End If
            Else
                DataString = DataString & "0" & ","
            End If
            
            If tmpOTBnf <> "" Then
                DataString = DataString & tmpOTBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMTheatreAct) Then
                DataString = DataString & Format(CLMBordx!CLMTheatreAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMTheatre) Then
                DataString = DataString & Format(CLMBordx!CLMTheatre, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If CLMBordx!CLMTheatre < CLMBordx!CLMTheatreAct Then
                DataString = DataString & Format(CLMBordx!CLMTheatreAct, "###0.00") - Format(CLMBordx!CLMTheatre, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            If tmpPreHosSpeBnf <> "" Then
                DataString = DataString & tmpPreHosSpeBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMNSPreHosAct) Then
                DataString = DataString & Format(CLMBordx!CLMNSPreHosAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMNSPreHos) Then
                DataString = DataString & Format(CLMBordx!CLMNSPreHos, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If CLMBordx!CLMNSPreHos < CLMBordx!CLMNSPreHosAct Then
                DataString = DataString & Format(CLMBordx!CLMNSPreHosAct, "###0.00") - Format(CLMBordx!CLMNSPreHos, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If tmpDailyInBnf <> "" Then
                DataString = DataString & tmpDailyInBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMPhysicianAct) Then
                DataString = DataString & Format(CLMBordx!CLMPhysicianAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMPhysician) Then
                DataString = DataString & Format(CLMBordx!CLMPhysician, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If CLMBordx!CLMPhysician < CLMBordx!CLMPhysicianAct Then
                DataString = DataString & Format(CLMBordx!CLMPhysicianAct, "###0.00") - Format(CLMBordx!CLMPhysician, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If tmpPostBnf <> "" Then
                DataString = DataString & tmpPostBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMPostAct) Then
                DataString = DataString & Format(CLMBordx!CLMPostAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMPost) Then
                DataString = DataString & Format(CLMBordx!CLMPost, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If CLMBordx!CLMPost < CLMBordx!CLMPostAct Then
                DataString = DataString & Format(CLMBordx!CLMPostAct, "###0.00") - Format(CLMBordx!CLMPost, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If tmpOPAccBnf <> "" Then
                DataString = DataString & Format(CLMBordx!CLMEmergencyLmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMEmergencyAct) Then
                DataString = DataString & Format(CLMBordx!CLMEmergencyAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMEmergency) Then
                DataString = DataString & Format(CLMBordx!CLMEmergency, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If CLMBordx!CLMEmergency < CLMBordx!CLMEmergencyAct Then
                DataString = DataString & Format(CLMBordx!CLMEmergencyAct, "###0.00") - Format(CLMBordx!CLMEmergency, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If tmpOPAccBnf <> "" Then
                DataString = DataString & tmpOPAccBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMOutpatientAct) Then
                DataString = DataString & Format(CLMBordx!CLMOutpatientAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMOutpatientAct) Then
                DataString = DataString & Format(CLMBordx!CLMOutpatient, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If CLMBordx!CLMOutpatient < CLMBordx!CLMOutpatientAct Then
                DataString = DataString & Format(CLMBordx!CLMOutpatientAct, "###0.00") - Format(CLMBordx!CLMOutpatient, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If tmpKidneyBnf <> "" Then
                DataString = DataString & tmpKidneyBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMMonthlyAct) Then
                DataString = DataString & Format(CLMBordx!CLMMonthlyAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMMonthly) Then
                DataString = DataString & Format(CLMBordx!CLMMonthly, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If CLMBordx!CLMMonthly < CLMBordx!CLMMonthlyAct Then
                DataString = DataString & Format(CLMBordx!CLMMonthlyAct, "###0.00") - Format(CLMBordx!CLMMonthly, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If tmpKidneyBnf <> "" Then
                DataString = DataString & tmpKidneyBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            If tmpOPAmbulanceBnf <> "" Then
                DataString = DataString & tmpOPAmbulanceBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMAmbulanceAct) Then
                DataString = DataString & Format(CLMBordx!CLMAmbulanceAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMAmbulance) Then
                DataString = DataString & Format(CLMBordx!CLMAmbulance, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If CLMBordx!CLMAmbulance < CLMBordx!CLMAmbulanceAct Then
                DataString = DataString & Format(CLMBordx!CLMAmbulanceAct, "###0.00") - Format(CLMBordx!CLMAmbulance, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            
            If tmpGuardianAllBnf <> "" Then
                DataString = DataString & tmpGuardianAllBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMGuardianAmtAct) Then
                DataString = DataString & Format(CLMBordx!CLMGuardianAmtAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMGuardianAmt) Then
                DataString = DataString & Format(CLMBordx!CLMGuardianAmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If CLMBordx!CLMGuardianAmt < CLMBordx!CLMGuardianAmtAct Then
                DataString = DataString & Format(CLMBordx!CLMGuardianAmtAct, "###0.00") - Format(CLMBordx!CLMGuardianAmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            'If IsNumeric(CLMBordx!CLMHomeNCAmtLmt) Then
            '    DataString = DataString & Format(CLMBordx!CLMHomeNCAmtLmt, "###0.00") & ","
            'Else
                DataString = DataString & "0" & ","
            'End If
            
            'If IsNumeric(CLMBordx!CLMHomeNCAmtAct) Then
            '    DataString = DataString & Format(CLMBordx!CLMHomeNCAmtAct, "###0.00") & ","
            'Else
                DataString = DataString & "0" & ","
            'End If
            'If IsNumeric(CLMBordx!CLMHomeNCAmt) Then
            '    DataString = DataString & Format(CLMBordx!CLMHomeNCAmt, "###0.00") & ","
            'Else
                DataString = DataString & "0" & ","
            'End If
            'If CLMBordx!CLMHomeNCAmtAct < CLMBordx!CLMHomeNCAmt Then
            '    DataString = DataString & Format(CLMBordx!CLMHomeNCAmtAct, "###0.00") - Format(CLMBordx!CLMHomeNCAmt, "###0.00") & ","
            'Else
                DataString = DataString & "0" & ","
            'End If
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            If tmpMedRptBnf <> "" Then
                DataString = DataString & tmpMedRptBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(CLMBordx!CLMMedicalAct) Then
                DataString = DataString & Format(CLMBordx!CLMMedicalAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMMedical) Then
                DataString = DataString & Format(CLMBordx!CLMMedical, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If CLMBordx!CLMMedical < CLMBordx!CLMMedicalAct Then
                If IsNumeric(CLMBordx!CLMMedical) Then
                    DataString = DataString & Format(CLMBordx!CLMMedicalAct, "###0.00") - Format(CLMBordx!CLMMedical, "###0.00") & ","
                Else
                    DataString = DataString & Format(CLMBordx!CLMMedicalAct, "###0.00") & ","
                End If
            Else
                DataString = DataString & "0" & ","
            End If
            
            If tmpGovHosTaxBnf <> "" Then
                DataString = DataString & tmpGovHosTaxBnf & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
           
            tmpTaxAct = IIf(IsNumeric(CLMBordx!CLMBoardTaxAmtAct), CLMBordx!CLMBoardTaxAmtAct, 0)
            tmpTaxApp = IIf(IsNumeric(CLMBordx!CLMBoardTaxAmt), CLMBordx!CLMBoardTaxAmt, 0)
            
            tmpICUTaxAct = IIf(IsNumeric(CLMBordx!CLMICUTaxAmtAct), CLMBordx!CLMICUTaxAmtAct, 0)
            tmpICUTaxApp = IIf(IsNumeric(CLMBordx!CLMICUTaxAmt), CLMBordx!CLMICUTaxAmt, 0)
            
            tmpTaxAct = CDbl(tmpTaxAct) + CDbl(tmpICUTaxAct)
            tmpTaxApp = CDbl(tmpTaxApp) + CDbl(tmpICUTaxApp)
            
            'tmpTaxAct = CDbl(tmpTaxAct) + IIf(IsNumeric(CLMBordx!CLMICUTaxAmtAct), CDbl(CLMBordx!CLMICUTaxAmtAct), 0)
            'tmpTaxApp = CDbl(tmpTaxApp) + IIf(IsNumeric(CLMBordx!CLMICUTaxAmt), CDbl(CLMBordx!CLMICUTaxAmt), 0)
            
            If IsNumeric(tmpTaxAct) Then
                DataString = DataString & tmpTaxAct & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            
            If IsNumeric(tmpTaxApp) Then
                DataString = DataString & tmpTaxApp & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            If CDbl(tmpTaxApp) < CDbl(tmpTaxAct) Then
                DataString = DataString & Format(CDbl(tmpTaxAct), "###0.00") - Format(CDbl(tmpTaxApp), "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            'If IsNumeric(CLMBordx!CLMOthers1Lmt) Then
            '    DataString = DataString & Format(CLMBordx!CLMOthers1Lmt, "###0.00") & ","
            'Else
            DataString = DataString & "0" & ","
            'End If
            

            
            
            
            

            tmpOthersAct = 0
            tmpOthersApp = 0
            
            tmpReimPreH60 = 0
            tmpReimPreGP = 0
            tmpReimPreMedic = 0
            tmpReimPreMR = 0
            tmpReimPreSpec = 0
            tmpReimPostHosp = 0
            tmpReimPostNotSame = 0
            tmpReimOthers = 0
             
            tmpReimActOthers = 0
            tmpNotCoPayAct = 0
            tmpNotCoPayApp = 0
            
            tmpNotCoPayAct = IIf(IsNumeric(CLMBordx!CLMOthersNotCoPayAct), CLMBordx!CLMOthersNotCoPayAct, 0)
            tmpNotCoPayApp = IIf(IsNumeric(CLMBordx!CLMOthersNotCoPay), CLMBordx!CLMOthersNotCoPay, 0)
            
            If IsNumeric(CLMBordx!PreH60) Then
                tmpReimPreH60 = CLMBordx!PreH60
            Else
                tmpReimPreH60 = 0
            End If

            If IsNumeric(CLMBordx!PreGP) Then
                tmpReimPreGP = CLMBordx!PreGP
            Else
                tmpReimPreGP = 0
            End If

            If IsNumeric(CLMBordx!PreMedic) Then
                tmpReimPreMedic = CLMBordx!PreMedic
            Else
                tmpReimPreMedic = 0
            End If
                        
            

            If IsNumeric(CLMBordx!PreMR) Then
                tmpReimPreMR = CLMBordx!PreMR
            Else
                tmpReimPreMR = 0
            End If

            If IsNumeric(CLMBordx!PreSpec) Then
                tmpReimPreSpec = CLMBordx!PreSpec
            Else
                tmpReimPreSpec = 0
            End If

            If IsNumeric(CLMBordx!PostHosp) Then
                tmpReimPostHosp = CLMBordx!PostHosp
            Else
                tmpReimPostHosp = 0
            End If

            If IsNumeric(CLMBordx!PostNotSame) Then
                tmpReimPostNotSame = CLMBordx!PostNotSame
            Else
                tmpReimPostNotSame = 0
            End If

            If IsNumeric(CLMBordx!Others) Then
                tmpReimOthers = CLMBordx!Others
            Else
                tmpReimOthers = 0
            End If

            If IsNumeric(CLMBordx!CLMOthers1Act) Then
                tmpOthersAct = CLMBordx!CLMOthers1Act
            Else
                tmpOthersAct = 0
            End If
            
            If IsNumeric(CLMBordx!CLMOthers1) Then
                tmpOthersApp = CLMBordx!CLMOthers1
            Else
                tmpOthersApp = 0
            End If
            'tmpOthersAct = 0
            If IsNumeric(CLMBordx!CLMAdminAct) Then
                tmpOthersAct = CDbl(tmpOthersAct) + CDbl(CLMBordx!CLMAdminAct) + CDbl(tmpNotCoPayAct)
            End If
            
            tmpReimActOthers = CDbl(tmpReimPreH60) + CDbl(tmpReimPreGP) + CDbl(tmpReimPreMedic) + CDbl(tmpReimPreMR) + CDbl(tmpReimPreSpec) + CDbl(tmpReimPostHosp) + CDbl(tmpReimPostNotSame) + CDbl(tmpReimOthers)
           'tmpOthersAct = 0
            tmpOthersAct = CDbl(tmpOthersAct) + CDbl(tmpReimActOthers)

            
            If IsNumeric(CLMBordx!CLMAdmin) Then
                tmpOthersApp = CDbl(tmpOthersApp) + CDbl(CLMBordx!CLMAdmin) + CDbl(tmpNotCoPayApp)
            End If
                                  
            If IsNumeric(tmpOthersAct) Then
                DataString = DataString & Format(tmpOthersAct, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If IsNumeric(tmpOthersApp) Then
                DataString = DataString & Format(tmpOthersApp, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            If tmpOthersApp < tmpOthersAct Then
                If IsNumeric(tmpOthersAct) Then
                    DataString = DataString & tmpOthersAct - Format(tmpOthersApp, "###0.00") & ","
                Else
                    DataString = DataString & Format(tmpOthersApp, "###0.00") & ","
                End If
            Else
                DataString = DataString & "0" & ","
            End If
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","

            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","

            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            'SKHPPA In-Hosp Specialist
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","

            'SKHPPA Ambulance
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","

            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","

            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","

            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","

            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","

            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
           
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","

            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
           
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","

            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
           
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
                        
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            DataString = DataString & "0" & ","
            
            DataString = DataString & "0" & ","
            
            tmpCLMCoPayRB = 0
            If IsNumeric(CLMBordx!CLMCoPayRB) Then
                DataString = DataString & Format(CLMBordx!CLMCoPayRB, "###0.00") & ","
                tmpCLMCoPayRB = IIf(IsNumeric(CLMBordx!CLMCoPayRB), CLMBordx!CLMCoPayRB, 0)
            Else
                DataString = DataString & "0" & ","
            End If
            
            If IsNumeric(CLMBordx!CLMExGratiaAmt) Then
                DataString = DataString & Format(CLMBordx!CLMExGratiaAmt, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            DataString = DataString & "0" & ","
            
            If IsNumeric(CLMBordx!LimitExc) Then
                DataString = DataString & Format(CLMBordx!LimitExc, "###0.00") & ","
            Else
                DataString = DataString & "0" & ","
            End If
            
            DataString = DataString & "0" & ","
            
            DataString = DataString & "0" & ","
            
            DataString = DataString & Format(CDbl(CLMBordx!CLMTotalActual) + CDbl(tmpReimActOthers), "###0.00") & ","
            
            DataString = DataString & Format(CLMBordx!CLMBordxAmt, "###0.00") + CDbl(tmpCLMCoPayRB) & ","
                        
            DataString = DataString & Format(CDbl((CLMBordx!CLMTotalActual) + CDbl(tmpReimActOthers)), "###0.00") - (Format(CLMBordx!CLMBordxAmt, "###0.00") + CDbl(tmpCLMCoPayRB))
            
            'DataString = DataString & l(Format(CLMBordx!CLMOrganLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMPhoneLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMAdminLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMOthers2Lmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMCashAmtLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMBoardTaxAmtLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMICUTaxAmtLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMLodgerTaxAmtLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMBoardDaysLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMICUDaysLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMCashDaysLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMMRILmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMMPreHSAmtLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMCOPreHSSumLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMCOMPreHSSumLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format("", "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMPreHospLmt, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPayLmt, "#,##0.00"), 11)
        
            'Actual
            'DataString = DataString & l(Format(CLMBordx!CLMOrganAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMPhoneAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMAdminAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMOthers2Act, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMCashAmtAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMBoardTaxAmtAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMICUTaxAmtAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMLodgerTaxAmtAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMMRIAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMMPreHSAmtAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMCOPreHSSumAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMCOMPreHSSumAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format("", "#,##0.00"), 11)
            ''edit --act.post.treatm
            'DataString = DataString & l(Format("", "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMPreHospAct, "#,##0.00"), 11)
            'DataString = DataString & l(Format(CLMBordx!CLMOthersNotCoPayAct, "#,##0.00"), 11)
            
            
            'Approved
            'If IsNumeric(CLMBordx!CLMOrgan) Then
            '    DataString = DataString & l(Format(CLMBordx!CLMOrgan, "#,##0.00"), 11)
            'Else
            '    DataString = DataString & l("0.00", 11)
            'End If
            'If IsNumeric(CLMBordx!CLMPhone) Then
            '    DataString = DataString & l(Format(CLMBordx!CLMPhone, "#,##0.00"), 11)
            'Else
            '    DataString = DataString & l("0.00", 11)
            'End If
            'If IsNumeric(CLMBordx!CLMAdmin) Then
            '    DataString = DataString & l(Format(CLMBordx!CLMAdmin, "#,##0.00"), 11)
            'Else
            '    DataString = DataString & l("0.00", 11)
            'End If
            'If IsNumeric(CLMBordx!CLMOthers2) Then
            '    DataString = DataString & l(Format(CLMBordx!CLMOthers2, "#,##0.00"), 11)
            'Else
            '    DataString = DataString & l("0.00", 11)
            'End If
           '
           ' If IsNumeric(CLMBordx!CLMCashAmt) Then
           '     DataString = DataString & l(Format(CLMBordx!CLMCashAmt, "#,##0.00"), 11)
           ' Else
           '     DataString = DataString & l("0.00", 11)
           ' End If
           '
           ' If IsNumeric(CLMBordx!CLMBoardTaxAmt) Then
           '     DataString = DataString & l(Format(CLMBordx!CLMBoardTaxAmt, "#,##0.00"), 11)
           ' Else
           '     DataString = DataString & l("0.00", 11)
           ' End If
           '
           ' If IsNumeric(CLMBordx!CLMICUTaxAmt) Then
           '     DataString = DataString & l(Format(CLMBordx!CLMICUTaxAmt, "#,##0.00"), 11)
           ' Else
           '     DataString = DataString & l("0.00", 11)
           ' End If
           '
           ' If IsNumeric(CLMBordx!CLMLodgerTaxAmt) Then
           '     DataString = DataString & l(Format(CLMBordx!CLMLodgerTaxAmt, "#,##0.00"), 11)
           ' Else
           '     DataString = DataString & l("0.00", 11)
           ' End If
           '
           ' If IsNumeric(CLMBordx!CLMMRI) Then
           '     DataString = DataString & l(Format(CLMBordx!CLMMRI, "#,##0.00"), 11)
           ' Else
           '     DataString = DataString & l("0.00", 11)
           ' End If
           '
           ' If IsNumeric(CLMBordx!CLMMPreHSAmt) Then
           '     DataString = DataString & l(Format(CLMBordx!CLMMPreHSAmt, "#,##0.00"), 11)
            'Else
           '     DataString = DataString & l("0.00", 11)
           ' End If
           '
           ' If IsNumeric(CLMBordx!CLMCOPreHSSum) Then
           '     DataString = DataString & l(Format(CLMBordx!CLMCOPreHSSum, "#,##0.00"), 11)
           ' Else
           '     DataString = DataString & l("0.00", 11)
           ' End If
           '
           ' If IsNumeric(CLMBordx!CLMCOMPreHSSum) Then
            '    DataString = DataString & l(Format(CLMBordx!CLMCOMPreHSSum, "#,##0.00"), 11)
            ''Else
            '    DataString = DataString & l("0.00", 11)
            'End If
            
            'If IsNumeric(CLMBordx!CLMAccPostAmt) Then
            '    DataString = DataString & l(Format(CLMBordx!CLMAccPostAmt, "#,##0.00"), 11)
            'Else
            '    DataString = DataString & l("0.00", 11)
            'End If
            'If IsNumeric(CLMBordx!CLMPreHosp) Then
            '    DataString = DataString & l(Format(CLMBordx!CLMPreHosp, "#,##0.00"), 11)
            'Else
            '    DataString = DataString & l("0.00", 11)
            'End If
            
            'If IsNumeric(CLMBordx!CLMOthersNotCOPay) Then
            '    DataString = DataString & l(Format(CLMBordx!CLMOthersNotCOPay, "#,##0.00"), 11)
            'Else
            '    DataString = DataString & l("0.00", 11)
            'End If
            'DataString = DataString & l(Format(CLMBordx!CLMExcess, "#,##0.00"), 11)
            
            'DataString = DataString & l(Format(CLMBordx!INVPaidByMem, "#,##0.00"), 11)
            
            'DataString = DataString & l(Format(CLMBordx!CLMHosPayColl, "#,##0.00"), 11)
            '-- end --'
            Print #1, DataString
            
            CLMBordx.MoveNext
        
        Loop
            
        'xlBook.SaveAs (Trim(txtLocalFileLocation) & Trim(txtFilName & ".csv"))
        
        'Call cmdExpSearch_Click
                                            
        
        'Set xlBook = Nothing
        'xlApp.Quit
        'Set xlApp = Nothing
        cmdGenerate.Enabled = False
        'MsgBox "Export successful", vbOKOnly + vbInformation, DialogBoxTitle
        
    End If
    CLMBordx.Close
    Set CLMBordx = Nothing
    Close #1
End Sub

Private Sub CaseASCII()
Dim strsql As String
Dim CLMBordx As New ADODB.Recordset
Dim BordxNo As String
Dim BordxDateFr As String
Dim BordxDateTo As String
Dim PAYCode As String
Dim ii As Integer
Dim sql As String
Dim CASNo As String
Dim VersionNo As String
Dim STRBordxArc As String
Dim BordxArc As New ADODB.Recordset
Dim BATNo As String
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim CLMCount As String
Dim BatchDate As String
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlsheet As Excel.WORKSHEET
Dim tmpString1 As String
Dim tmpString2 As String
Dim txtSVRfile As String
Dim Rcounter, row As Double
Dim tmpAdmissionReason As String
Dim tmpDiag1 As String
Dim tmpDiag2 As String
Dim tmpDiag3 As String
Dim tmpSurgicalProc1 As String
Dim tmpSurgicalProc2 As String
Dim tmpSurgicalProc3 As String
Dim tmpSurgicalProc4 As String
Dim tmpSurgicalProc5 As String
Dim tmpAdmDays As String

        tmpAdmDays = 0
        CLMCount = 0
        'lblTotalCases = 0
        BordxDateFr = txtReceiptDateFr
        BordxDateTo = txtReceiptDateTo
        PAYCode = Mid(cboPAYCode, 1, 2)
        
        'strsql = " Select * from View_CLMReport where BordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'" & _
        '         " AND PAYCode = '" & PAYCode & "'"
       '
       ' strCount = " Select count(*) as totalRecord From View_CLMReport where BordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'" & _
                   " AND PAYCode = '" & PAYCode & "'"
       ' strCount = strCount
              
        '------ Start (Do not Delete this Code) -----'
       ' If Mid(cboBordxWSStatus, 1, 1) = "0" Then
            'sql = " Select * from View_CLMReport5_New where 0=0"
            strAppendCase = "6"
       ' End If
        'StrAppend = ""
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
        'Else
        '        strAppend = " AND CaseRegDate >= '" & Format(CDate(BordxDateFr), "yyyy/mm/dd") & "'" & _
        '                            " AND CaseRegDate <= '" & Format(CDate(BordxDateTo), "yyyy/mm/dd") & "'" & _
        '                            " AND PAYCode = '" & PAYCode & "' "'

            
        End If
        
        'If Mid(cboBordxWSStatus, 1, 1) = "0" Then
        '    strAppend = strAppend + " AND (CLMBordxWSStatus = 'A' or CLMBordxWSStatus = 'N')"
        'End If
                    
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_ClaimAscii"
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
    
    If CLMBordx.EOF Then
    'If rstCount!totalrecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
            'BATNo = GetCASAsciiBatch(Mid(cboPAYCode, 1, 2), Date)
            BATNo = "001"
            textField = Trim(Date)
                  
            txtFilName = Trim("C001" & Mid(cboPAYCode, 1, 2) & Format(Date, "YYYYMMDD"))
            Open txtLocalFileLocation & txtFilName & "001" & ".csv" For Output As #2
            txtSVRfile = (Trim(txtLocalFileLocation) & Trim(txtFilName & ".csv"))
           'Open txtLocalFileLocation & "\" & txtFilName & " & BATNo & " & ".csv" For Output As #1
            'Open txtLocalFileLocation & txtFilName & "(" & BATNo & ")" For Output As #2
           ' Set fso = CreateObject("Scripting.FileSystemObject")
           ' Set txtSVRfile = fso.CreateTextFile((Trim(txtLocalFileLocation) & Trim(txtFilName & ".xls")), True)
            'Set txtfile = fso.CreateTextFile((Trim(txtPath) & Trim(txtFilName & ".csv")), True)
            'txtfile = (Trim(txtPath) & Trim(txtFilName & ".xls"))
            
            'et xl = New Excel.Application
            'Set xlwbook = xl.Workbooks.Add
            'Set xlsheet = xlwbook.Sheets.Item(1)

            txtSVRfile = (Trim(txtLocalFileLocation) & Trim(txtFilName & ".csv"))
            
            'Set xlApp = New Excel.Application
            'Set xlBook = xlApp.Workbooks.Add
            'Set xlsheet = xlBook.Sheets(1)
            
            'Header = Trim("F" & Mid(cboPAYCode, 1, 2) & BATNo & Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 7, 4) & l("", 174))
            BatchDate = Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 7, 4)
            'Print #1, Trim(Header)
            'Print #2, Trim(Header)
            'Debug.Print Header
        Rcounter = 0
        row = 0
        tmpString1 = "MEDIX"
        tmpString2 = "N"
        Do Until CLMBordx.EOF
        
            Rcounter = Rcounter + 1
            row = row + 1
            
            'xlApp.DisplayAlerts = False
                       'DataString = DataString &
            DataString = tmpString1 & "" & ","
            DataString = DataString & IIf(Len(CLMBordx!CASASCIIAmendStatus), CLMBordx!CASASCIIAmendStatus, tmpString2) & ","
            DataString = DataString & " " & ","
            DataString = DataString & CLMBordx!CasNumber & ","
            DataString = DataString & CLMBordx!CasNumber & "-" & CLMBordx!claimversion & ","
            DataString = DataString & CLMBordx!BordxRefNo & ","
            DataString = DataString & CLMBordx!MBMPolicyNo & ","
            DataString = DataString & "" & ","
            DataString = DataString & CLMBordx!MBMGroupCompany & ","
            DataString = DataString & CLMBordx!MBMIcBcPp & "," '10
            DataString = DataString & CLMBordx!MBMName & ","
            DataString = DataString & CLMBordx!ICNo & ","
            DataString = DataString & CLMBordx!PatientName & ","
            DataString = DataString & "H010GHM" & ","
            DataString = DataString & "HGHM" & ","
            DataString = DataString & Format(Date, "dd/mm/yyyy") & ","
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "dd/mm/yyyy") & ","
            'DataString = DataString & CLMBordx!CASHOSCode & ","
            DataString = DataString & CLMBordx!HOSZurichCode & ","
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "dd/mm/yyyy") & ","
            If CLMBordx!CASDischargeDate <> "" Then '20
                DataString = DataString & Format(CLMBordx!CASDischargeDate, "dd/mm/yyyy") & ","
            Else
                DataString = DataString & Format(CLMBordx!CASDischargeDate, "dd/mm/yyyy") & ","
            End If
            If CLMBordx!CASDischargeDate <> "" Then
                tmpAdmDays = CDate(CLMBordx!CASDischargeDate) - CDate(CLMBordx!CASDateAdmitted)
                If tmpAdmDays = 0 Then
                    tmpAdmDays = 1
                End If
            End If
            If tmpAdmDays <> 0 Then
                DataString = DataString & tmpAdmDays & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASAdmissionReason <> "" Then
                tmpAdmissionReason = Replace(CLMBordx!CASAdmissionReason, ",", "")
                DataString = DataString & tmpAdmissionReason & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CLMBordxAmt <> "" Then
                DataString = DataString & Format(CLMBordx!CLMBordxAmt, "###0.00") & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            DataString = DataString & "N" & ","
            
            
            If CLMBordx!FinalDiag <> "" Then
                DataString = DataString & CLMBordx!FinalDiag & ","
            Else
                DataString = DataString & " " & ","
            End If
            If CLMBordx!FinalDiag2 <> "" Then
                DataString = DataString & CLMBordx!FinalDiag2 & ","
            Else
                DataString = DataString & " " & ","
            End If
            If CLMBordx!FinalDiag3 <> "" Then
                DataString = DataString & CLMBordx!FinalDiag3 & ","
            Else
                DataString = DataString & " " & ","
            End If
                DataString = DataString & " " & ","
                DataString = DataString & " " & ","
            
            If CLMBordx!DIAFinalDiag1 <> "" Then
            
                tmpDiag1 = Replace(CLMBordx!DIAFinalDiag1, ",", "")
                DataString = DataString & tmpDiag1 & ","
            Else
                DataString = DataString & " " & ","
            End If
            If CLMBordx!DIAFinalDiag2 <> "" Then
                tmpDiag2 = Replace(CLMBordx!DIAFinalDiag2, ",", "")
                DataString = DataString & tmpDiag2 & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!DIAFinalDiag3 <> "" Then
                tmpDiag3 = Replace(CLMBordx!DIAFinalDiag3, ",", "")
                DataString = DataString & tmpDiag3 & ","
            Else
                DataString = DataString & " " & ","
            End If
                
                DataString = DataString & " " & ","
                
                DataString = DataString & " " & ","
                
            If CLMBordx!CASCptCode1 <> "" And Len(CLMBordx!CASCptCode1) = "5" Then
                tmpSurgicalProc1 = Replace(CLMBordx!CASCptCode1, ",", "")
                DataString = DataString & tmpSurgicalProc1 & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASCptCode2 <> "" And Len(CLMBordx!CASCptCode2) = "5" Then
                tmpSurgicalProc2 = Replace(CLMBordx!CASCptCode2, ",", "")
                DataString = DataString & tmpSurgicalProc2 & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASCptCode3 <> "" And Len(CLMBordx!CASCptCode3) = "5" Then
                tmpSurgicalProc3 = Replace(CLMBordx!CASCptCode3, ",", "")
                DataString = DataString & tmpSurgicalProc3 & ","
            Else
                DataString = DataString & " " & ","
            End If
                
            If CLMBordx!CASCptCode4 <> "" And Len(CLMBordx!CASCptCode4) = "5" Then
                tmpSurgicalProc4 = Replace(CLMBordx!CASCptCode4, ",", "")
                DataString = DataString & tmpSurgicalProc4 & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASCptCode5 <> "" And Len(CLMBordx!CASCptCode5) = "5" Then
                tmpSurgicalProc5 = Replace(CLMBordx!CASCptCode5, ",", "")
                DataString = DataString & tmpSurgicalProc5 & ","
            Else
                DataString = DataString & " " & ","
            End If
                
            If CLMBordx!AttendDoctor <> "" Then
                DataString = DataString & CLMBordx!AttendDoctor & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASAttendDoctor2 <> "" Then
                DataString = DataString & CLMBordx!CASAttendDoctor2 & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASAttendDoctor3 <> "" Then
                DataString = DataString & CLMBordx!CASAttendDoctor3 & ","
            Else
                DataString = DataString & " " & ","
            End If
                
            If CLMBordx!CASAttendDoctor4 <> "" Then
                DataString = DataString & CLMBordx!CASAttendDoctor4 & ","
            Else
                DataString = DataString & " " & ","
            End If
            If CLMBordx!CASAttendDoctor5 <> "" Then
                DataString = DataString & CLMBordx!CASAttendDoctor5 & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            'DataString = DataString & Mid(CLMBordx!CASClaimability, 1, 1) & ","
            
            DataString = DataString & "A" & ","
            
            DataString = DataString & " "
            'DataString = l("Medix", 20)
            'DataString = DataString & l(IIf(Len(CLMBordx!CASASCIIAmendStatus), CLMBordx!CASASCIIAmendStatus, "N"), 3)
            'ESP Ref Number
            'DataString = DataString & l("", 10)
            'DataString = DataString & l(CLMBordx!CasNumber, 20)
            'DataString = DataString & l(CLMBordx!CasNumber, 20)
            'DataString = DataString & l(CLMBordx!BordxNumber, 20)
            'DataString = DataString & l(CLMBordx!MBMPolicyNo, 20)
            'DataString = DataString & l(CLMBordx!MBMNumber, 30)
            'DataString = DataString & l(CLMBordx!MBMName, 500)
            'DataString = DataString & l("", 30)
            'DataString = DataString & l("", 500)
            'DataString = DataString & l(CLMBordx!MBMIcBcPp, 30)
            'DataString = DataString & l(CLMBordx!PatientName, 500)
            'DataString = DataString & l("H010GHM", 10)
            'DataString = DataString & l(CLMBordx!PLNCode, 10)
            'DataString = DataString & l(Format(CLMBordx!CASDateAdmitted, "dd/mm/yyyy"), 10)
            'DataString = DataString & l(Format(CLMBordx!CASDateAdmitted, "dd/mm/yyyy"), 10)
            'DataString = DataString & l(Format(CLMBordx!CASHOSCode, "dd/mm/yyyy"), 10)
            'DataString = DataString & l(Format(CLMBordx!CASDateAdmitted, "dd/mm/yyyy"), 10)
            'If CLMBordx!CASDischargeDate <> "" Then
            '    DataString = DataString & l(Format(CLMBordx!CASDischargeDate, "ddmmyyyy"), 10)
            'Else
            '    DataString = DataString & l("", 10)
            'End If
            'DataString = DataString & l("", 10)
            'If CLMBordx!CASAdmissionReason <> "" Then
            '    DataString = DataString & l(CLMBordx!CASAdmissionReason, 500)
            'Else
            '    DataString = DataString & l("", 500)
            'End If
            'If CLMBordx!CASReserveAmount <> "" Then
            '    DataString = DataString & l(CLMBordx!CASReserveAmount, 8)
            'Else
            '    DataString = DataString & l("", 8)
            'End If
            'DataString = DataString & l("", 5)
            
            'If CLMBordx!FinalDiag <> "" Then
            '    DataString = DataString & l(CLMBordx!FinalDiag, 10)
            'Else
            '    DataString = DataString & l("", 10)
            'End If
            'If CLMBordx!FinalDiag2 <> "" Then
            '    DataString = DataString & l(CLMBordx!FinalDiag2, 10)
            'Else
            '    DataString = DataString & l("", 10)
            'End If
            'If CLMBordx!FinalDiag3 <> "" Then
            '    DataString = DataString & l(CLMBordx!FinalDiag3, 10)
            'Else
            '    DataString = DataString & l("", 10)
            'End If
            '    DataString = DataString & l("", 10)
            '    DataString = DataString & l("", 10)
                
            'If CLMBordx!DIAFinalDiag1 <> "" Then
            '    DataString = DataString & l(CLMBordx!DIAFinalDiag1, 500)
            'Else
            '    DataString = DataString & l("", 500)
            'End If
            'If CLMBordx!FinalDiag2 <> "" Then
            '    DataString = DataString & l(CLMBordx!DIAFinalDiag2, 500)
            'Else
            '    DataString = DataString & l("", 500)
            'End If
            'If CLMBordx!FinalDiag3 <> "" Then
            '    DataString = DataString & l(CLMBordx!DIAFinalDiag3, 500)
            'Else
            '    DataString = DataString & l("", 500)
            'End If
            '    DataString = DataString & l("", 500)
            '    DataString = DataString & l("", 500)
            'If CLMBordx!CASOPROCRequired1 <> "" Then
            '    DataString = DataString & l(CLMBordx!CASOPROCRequired1, 10)
            'Else
            '    DataString = DataString & l("", 10)
            'End If
            'If CLMBordx!CASOPROCRequired2 <> "" Then
            '    DataString = DataString & l(CLMBordx!CASOPROCRequired2, 10)
            'Else
            '    DataString = DataString & l("", 10)
            'End If
            'If CLMBordx!CASOPROCRequired3 <> "" Then
            '    DataString = DataString & l(CLMBordx!CASOPROCRequired3, 10)
            'Else
            '    DataString = DataString & l("", 10)
            'End If
            'If CLMBordx!CASOPROCRequired4 <> "" Then
            '    DataString = DataString & l(CLMBordx!CASOPROCRequired4, 10)
            'Else
            '    DataString = DataString & l("", 10)
            'End If
            'If CLMBordx!CASOPROCRequired5 <> "" Then
            '    DataString = DataString & l(CLMBordx!CASOPROCRequired5, 10)
            'Else
            '    DataString = DataString & l("", 10)
            'End If
            
            'If CLMBordx!AttendDoctor <> "" Then
            '    DataString = DataString & l(CLMBordx!AttendDoctor, 100)
            'Else
            '    DataString = DataString & l("", 100)
            'End If
            
            'If CLMBordx!CASAttendDoctor2 <> "" Then
            '    DataString = DataString & l(CLMBordx!CASAttendDoctor2, 100)
            'Else
            '    DataString = DataString & l("", 100)
            'End If
            'If CLMBordx!CASAttendDoctor3 <> "" Then
            '    DataString = DataString & l(CLMBordx!CASAttendDoctor3, 100)
            'Else
            '    DataString = DataString & l("", 100)
            'End If
            'If CLMBordx!CASAttendDoctor4 <> "" Then
            '    DataString = DataString & l(CLMBordx!CASAttendDoctor4, 100)
            'Else
            '    DataString = DataString & l("", 100)
            'End If
            'If CLMBordx!CASAttendDoctor5 <> "" Then
            '    DataString = DataString & l(CLMBordx!CASAttendDoctor5, 100)
            'Else
            '    DataString = DataString & l("", 100)
            'End If
            

            'If CLMBordx!CASClaimability <> "" Then
            '    DataString = DataString & l(CLMBordx!CASClaimability, 5)
            'Else
            '    DataString = DataString & l("", 1)
            'End If
            
            'DataString = DataString & l("", 5)
            '--- Start (Do not delete this)--'
                'Call UpdateCASAscii(CASNo)
            '-- end --'
            Print #2, DataString
            
            CLMBordx.MoveNext
            
            
            
        Loop
            'xl.ActiveWorkbook.Save ("c:\myfilename.xls")

        'xlBook.ActiveWorkbook.Save (txtSVRfile)
        
        'xlBook.Save
        'xlBook.SaveAs (Trim(txtLocalFileLocation) & Trim(txtFilName & ".csv"))
        'xlBook.SaveAs (Trim(txtLocalFileLocation) & Trim(txtFilName & ".csv"))
        MsgBox "Export successful", vbOKOnly + vbInformation, DialogBoxTitle
        'Call cmdExpSearch_Click
                                            
        
        'Set xlBook = Nothing
        'xlApp.Quit
        'Set xlApp = Nothing

        cmdGenerate.Enabled = True
    End If
    Close #2
    CLMBordx.Close
    Set CLMBordx = Nothing
    
End Sub
Private Sub Form_Load()
    txtPath = SSRVPath & "ClaimASCI"
    txtLocalFileLocation = "C:\MedicaGenClaimASCII\"
    Call ComboListing
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
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
    
End Sub



