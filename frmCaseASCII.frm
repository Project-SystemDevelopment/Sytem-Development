VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCaseASCII 
   Caption         =   "Case ASCII"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   7125
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   240
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   1920
         Width           =   1995
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label txtFilName 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   2520
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
      Begin VB.Label Label17 
         Caption         =   "Total Cases"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblTotalCases 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Case ASCII"
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
      Index           =   1
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmCaseASCII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Exportsql As String
Private textField As String
Private strCount
Private strCount2

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

Private Sub CmdClear_Click()
    Call ClearForm(Me)
End Sub

Private Sub CmdGenerate_Click()
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim RecordSaved

    If cboPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPAYCode.SetFocus
    ElseIf txtReceiptDateFr = "__/__/____" Then
        MsgBox "Please Enter Bordx Date Fr ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDateFr.SetFocus
    ElseIf txtReceiptDateTo = "__/__/____" Then
        MsgBox "Please Enter Bordx Date To ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDateTo.SetFocus
    ElseIf CboExportStatus = "" Then
        MsgBox "Please Enter Export Status ", vbOKOnly + vbCritical, DialogBoxTitle
        CboExportStatus.SetFocus
    ElseIf cboBordxWSStatus = "" Then
        MsgBox "Please Enter Endt Status ", vbOKOnly + vbCritical, DialogBoxTitle
        cboBordxWSStatus.SetFocus
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation, DialogBoxTitle)
        If RecordSaved = vbYes Then
    
            Screen.MousePointer = vbHourglass
            
            'If Mid(CboExportStatus, 1, 1) = "N" Then
            'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                Call PrincipalASCII
            'medixmasterconnWS.Close
            Set medixmasterconnWS = Nothing
            medixmasterconn.Close
            Set medixmasterconn = Nothing
            'Else
            '    Call RePrintASCII
            'End If
            
            Screen.MousePointer = vbDefault
            
        End If
    End If
End Sub

Private Sub PrincipalASCII()
Dim strsql As String
Dim CLMBordx As New ADODB.Recordset
Dim BORDXNO As String
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
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend As String
Dim CLMCount As String
Dim BatchDate As String

        CLMCount = 0
        lblTotalCases = 0
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
        If Mid(cboBordxWSStatus, 1, 1) = "0" Then
            'sql = " Select * from View_CLMReport5_New where 0=0"
            strAppendCase = "5"
        End If
        'StrAppend = ""
        If Mid(CboExportStatus, 1, 1) = "N" Then
            StrAppend = " AND CaseRegDate >= '" & Format(CDate(BordxDateFr), "MM/DD/YYYY") & "'" & _
                                " AND CaseRegDate <= '" & Format(CDate(BordxDateTo), "MM/DD/YYYY") & "'" & _
                                " AND PAYCode = '" & PAYCode & "'"
        Else
            StrAppend = " AND CaseRegDate >= '" & Format(CDate(BordxDateFr), "MM/DD/YYYY") & "'" & _
                                " AND CaseRegDate <= '" & Format(CDate(BordxDateTo), "MM/DD/YYYY") & "'" & _
                                " AND PAYCode = '" & PAYCode & "' AND CLMBordxWSASCII = 'Y'"
        End If
        
        'If Mid(cboBordxWSStatus, 1, 1) = "0" Then
        '    strAppend = strAppend + " AND (CLMBordxWSStatus = 'A' or CLMBordxWSStatus = 'N')"
        'End If
                    
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_ClaimAscii"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        CLMBordx.CursorLocation = adUseClient
        CLMBordx.CursorType = adOpenForwardOnly
        CLMBordx.CacheSize = 100
        Set CLMBordx = cmd.Execute
    
    If CLMBordx.EOF Then
    'If rstCount!totalrecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
            BATNo = GetCASAsciiBatch(Mid(cboPAYCode, 1, 2), Date)
            textField = Trim(Date)
                  
            txtFilName = Trim(Mid(cboPAYCode, 1, 2) & Format(Date, "YYMMDD"))
                
            Open txtPath & "\" & txtFilName & "(" & BATNo & ")" For Output As #1
            Open txtLocalFileLocation & txtFilName & "(" & BATNo & ")" For Output As #2
            
            'Header = Trim("F" & Mid(cboPAYCode, 1, 2) & BATNo & Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 7, 4) & l("", 174))
            BatchDate = Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 7, 4)
            'Print #1, Trim(Header)
            'Print #2, Trim(Header)
            'Debug.Print Header
        
        
        Do Until CLMBordx.EOF
        
            CASNo = CLMBordx!CasNumber
            DataString = l(IIf(Len(CLMBordx!CASASCIIAmendStatus), CLMBordx!CASASCIIAmendStatus, "N"), 1)
            DataString = DataString & l(CLMBordx!PAYCode, 2)
            DataString = DataString & l(CLMBordx!CasNumber, 15)
            DataString = DataString & l(CLMBordx!MBMNumber, 15)
            DataString = DataString & l(CLMBordx!MBMPatientCovID, 2)
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
                        
            DataString = DataString & l(CLMBordx!MBMName, 60)
            DataString = DataString & l(CLMBordx!MBMIcBcPp, 20)
            If Len(CLMBordx!MBMDOB) Then
                DataString = DataString & l(Format(CLMBordx!MBMDOB, "ddmmyyyy"), 8)
            Else
                DataString = DataString & l("", 8)
            End If
            DataString = DataString & l(CLMBordx!PLNCode, 4)
            
            If Len(CLMBordx!MBMGroupCompany) Then
                DataString = DataString & l(CLMBordx!MBMGroupCompany, 50)
            Else
                DataString = DataString & l("", 50)
            End If
            If Len(CLMBordx!PatientName) Then
                DataString = DataString & l(CLMBordx!PatientName, 60)
            Else
                DataString = DataString & l("", 60)
            End If

            If Len(CLMBordx!ICNo) Then
                DataString = DataString & l(CLMBordx!ICNo, 20)
            Else
                DataString = DataString & l("", 30)
            End If
            DataString = DataString & l(Format(CLMBordx!DOB, "ddmmyyyy"), 8)
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
            If CLMBordx!CASClaimability <> "" Then
                DataString = DataString & l(CLMBordx!CASClaimability, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            
            If CLMBordx!CASAdmissionReason <> "" Then
                DataString = DataString & l(CLMBordx!CASAdmissionReason, 200)
            Else
                DataString = DataString & l("", 200)
            End If
            If CLMBordx!CASStayType <> "" Then
                DataString = DataString & l(CLMBordx!CASStayType, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            
            If CLMBordx!CASPAYType <> "" Then
                DataString = DataString & l(CLMBordx!CASPAYType, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            
            If CLMBordx!CASReserveAmount <> "" Then
                DataString = DataString & l(CLMBordx!CASReserveAmount, 8)
            Else
                DataString = DataString & l("", 8)
            End If
            If CLMBordx!CASLGAmt <> "" Then
                DataString = DataString & l(CLMBordx!CASLGAmt, 8)
            Else
                DataString = DataString & l("", 8)
            End If
            DataString = DataString & l(BatchDate, 8)
            DataString = DataString & l(BATNo, 5)
            
            DataString = DataString & l("", 4370) & Chr(13)
            
            '--- Start (Do not delete this)--'
                'Call UpdateCASAscii(CASNo)
            '-- end --'
            
            CLMBordx.MoveNext
            
            'Debug.Print DataString
            Print #1, DataString
            Print #2, DataString
            
   '     Next ii
            CLMCount = CDbl(CLMCount) + 1
            lblTotalCases = lblTotalCases + 1
        Loop
            
            'Trailer = "T" & l(CLMCount, 10) & l("", 164)
                'Debug.Print Trailer
                'Print #1, Trailer
                'Print #2, Trailer
                cmdGenerate.Enabled = False
                Close #1
                Close #2
                
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
    End If
    CLMBordx.Close
    Set CLMBordx = Nothing
    Close #1
    Close #2
End Sub

Private Sub UpdateCASAscii(CaseNo As String)
Dim WSAscii As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend As String
Dim strsqlWSAscii As String
    
    strsqlWSAscii = "EXEC Case_AscUpdate '" & CaseNo & "', Y"
    medixmasterconnWS.Execute strsqlWSAscii
End Sub

Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub Form_Load()
        
    txtPath = SSRVPath & "CaseASCII"
    txtLocalFileLocation = "C:\CaseASCII\"
    Call ComboListing
    
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

Private Sub cboPAYCode_Change()
If cboPAYCode.Text = "" Then
    If InStr(cboPAYCode.Text, "null") Then
       cboPAYCode.Enabled = False
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

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command

 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
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
        
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
    CboExportStatus.AddItem "N - New"
    CboExportStatus.AddItem "R - Reprint"
    CboExportStatus.Text = "N - New"
    
    cboBordxWSStatus.AddItem "0 - New"
    cboBordxWSStatus.AddItem "1 - Deleted"
    cboBordxWSStatus.AddItem "2 - Amendment"
    cboBordxWSStatus.Text = "0 - New"
End Sub


