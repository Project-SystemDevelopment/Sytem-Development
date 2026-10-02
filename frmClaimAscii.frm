VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmClaimAscii 
   Caption         =   "Claim Ascii"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   7260
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11295
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Generate ASCII"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   180
         TabIndex        =   22
         Top             =   60
         Width           =   6615
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   480
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
         Style           =   2  'Dropdown List
         TabIndex        =   23
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
         Style           =   2  'Dropdown List
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1200
         Width           =   5535
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   5280
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox txtReceiptDateFr 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
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
         TabIndex        =   25
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
         TabIndex        =   12
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
         TabIndex        =   26
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "Bordx WS"
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
         TabIndex        =   24
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
         Left            =   4200
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   1545
         Width           =   915
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   900
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   240
         Width           =   840
      End
      Begin VB.Label txtTotalSupplementary1 
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
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   960
      End
      Begin VB.Label txtTotalPrincipals 
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
         TabIndex        =   6
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label16 
         Caption         =   "Total Supp"
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
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label17 
         Caption         =   "Total Principal"
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
         TabIndex        =   4
         Top             =   270
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmClaimAscii"
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

Private Sub cmdClear_Click()
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
        
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation, DialogBoxTitle)
        If RecordSaved = vbYes Then
    
            Screen.MousePointer = vbHourglass
            
            'If Mid(CboExportStatus, 1, 1) = "N" Then
            'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                Call PrincipalASCII
           ' medixmasterconnWS.Close
           ' Set medixmasterconnWS = Nothing
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
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String


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
            strAppendCase = "1"
        End If
        
        If Mid(CboExportStatus, 1, 1) = "N" Then
            strAppend = " AND BordxDate >= '" & Format(CDate(BordxDateFr), "MM/DD/YYYY") & "'" & _
                                " AND BordxDate <= '" & Format(CDate(BordxDateTo), "MM/DD/YYYY") & "'" & _
                                " AND PAYCode = '" & PAYCode & "' AND CLMBordxWSASCII = 'N'"
        Else
            strAppend = " AND BordxDate >= '" & Format(CDate(BordxDateFr), "MM/DD/YYYY") & "'" & _
                                " AND BordxDate <= '" & Format(CDate(BordxDateTo), "MM/DD/YYYY") & "'" & _
                                " AND PAYCode = '" & PAYCode & "' AND CLMBordxWSASCII = 'Y'"
        End If
        
        If Mid(cboBordxWSStatus, 1, 1) = "0" Then
            strAppend = strAppend + " AND (CLMBordxWSStatus = 'A' or CLMBordxWSStatus = 'N')"
        End If
        'sql = sql + strsql
        
        'If Mid(cboBordxWSStatus, 1, 1) = "0" Then
            'strCount = " Select count(*) as totalRecord From View_CLMReport5 where 0=0"
        'Else
            'strCount = " Select count(*) as totalRecord From View_CLMReportARC where 0=0"
        'End If
        
        'strCount = strCount + strsql
        
        '----- End -----'
                    
        'CLMBordx.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        'Set rstCount = medixmasterconnWS.Execute(strCount)
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Case_ClaimAscii"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
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
            BATNo = GetCLMAsciiBatch(Mid(cboPAYCode, 1, 2), Date)
            textField = Trim(Date)
                  
            txtFilName = Trim(Mid(cboPAYCode, 1, 2) & Format(Date, "YYMMDD"))
                
            Open txtPath & "\" & txtFilName & "(" & BATNo & ")" For Output As #1
            Open txtLocalFileLocation & txtFilName & "(" & BATNo & ")" For Output As #2
            
            Header = Trim("F" & Mid(cboPAYCode, 1, 2) & BATNo & Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 7, 4) & l("", 174))
            
            Print #1, Trim(Header)
            Print #2, Trim(Header)
            'Debug.Print Header
        
        
        Do Until CLMBordx.EOF
        
  '      For ii = 1 To rstCount!totalrecord
            'If MBM!RELCODE = "P" Then
            '    TotalPrincipal = TotalPrincipal + 1
            '    txtTotalPrincipals = TotalPrincipal
            'Else
            '    TotalSupplementary = TotalSupplementary + 1
            '    txtTotalSupplementary1 = TotalSupplementary
            'End If
            
            ClaimNo = CLMBordx!CasNumber
            VersionNo = CLMBordx!claimversion
            BordxNo = CLMBordx!BordxRefNo
            BordxDate = CLMBordx!BordxDate

            DataString = l(CLMBordx!CLMBordxWSStatus, 1)
            DataString = DataString & l(BordxNo, 15)
            DataString = DataString & l(Format(BordxDate, "ddmmyyyy"), 8)
            DataString = DataString & l(CLMBordx!PAYCode, 2)
            DataString = DataString & l(CLMBordx!CasNumber, 15)
            DataString = DataString & l(CLMBordx!CasNumber & "-" & CLMBordx!claimversion, 15)
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
            
            If Len(CLMBordx!MBMAgentCode) Then
                DataString = DataString & l(CLMBordx!MBMAgentCode, 15)
            Else
                DataString = DataString & l("", 15)
            End If
            
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
                        
            'If Len(CLMBordx!MBMMaritalStatus) Then
            '    DataString = DataString & l(CLMBordx!MBMMaritalStatus, 1)
            'Else
                DataString = DataString & l("", 1)
            'End If
            
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
            DataString = DataString & l(CLMBordx!PLNCode, 4)
            
            If Len(CLMBordx!MBMGroupCompany) Then
                DataString = DataString & l(CLMBordx!MBMGroupCompany, 50)
            Else
                DataString = DataString & l("", 50)
            End If
            
            If Len(CLMBordx!BRNCode) Then
                DataString = DataString & l(CLMBordx!BRNCode, 50)
            Else
                DataString = DataString & l("", 50)
            End If
                        
            If Len(CLMBordx!AGECode) Then
                DataString = DataString & l(CLMBordx!AGECode, 2)
            Else
                DataString = DataString & l("", 2)
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
            
            If Len(CLMBordx!Sex) Then
                DataString = DataString & l(CLMBordx!Sex, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            If Len(CLMBordx!RACCode) Then
                DataString = DataString & l(CLMBordx!RACCode, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            If Len(CLMBordx!AgeBand) Then
                DataString = DataString & l(CLMBordx!AgeBand, 2)
            Else
                DataString = DataString & l("", 2)
            End If

            DataString = DataString & l(Format(CLMBordx!DOB, "ddmmyyyy"), 8)
            
            If Len(CLMBordx!HOSName) Then
                DataString = DataString & l(CLMBordx!HOSName, 80)
            Else
                DataString = DataString & l("", 80)
            End If
            
            If CLMBordx!FinalDiag <> "" Then
                DataString = DataString & l(CLMBordx!FinalDiag, 5)
            Else
                DataString = DataString & l("", 5)
            End If
            
            If CLMBordx!FinalDiag2 <> "" Then
                DataString = DataString & l(CLMBordx!FinalDiag2, 5)
            Else
                DataString = DataString & l("", 5)
            End If
            
            If CLMBordx!FinalDiag3 <> "" Then
                DataString = DataString & l(CLMBordx!FinalDiag3, 5)
            Else
                DataString = DataString & l("", 5)
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
            
            If CLMBordx!CLMTotalActual <> "" Then
                DataString = DataString & l(Format(CLMBordx!CLMTotalActual, "#,##0.00"), 12)
            Else
                DataString = DataString & l("", 12)
            End If
            
            If CLMBordx!CLMTotalApproved <> "" Then
                DataString = DataString & l(Format(CLMBordx!CLMTotalApproved, "#,##0.00"), 12)
            Else
                DataString = DataString & l("", 12)
            End If
            
            'DataString = DataString & l(Format(CLMBordx!CASReserveAmount, "#,##0.00"), 12)
            
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
            
            If CLMBordx!AttendDoctor <> "" Then
                DataString = DataString & l(CLMBordx!AttendDoctor, 100)
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
            
            If CLMBordx!CASPAYType <> "" Then
                DataString = DataString & l(CLMBordx!CASPAYType, 1)
            Else
                DataString = DataString & l("", 1)
            End If
            
            If CLMBordx!CASUDReasons1 <> "" Then
            'DataString = DataString & l(CLMBordx!CASUDCode1, 5)
                DataString = DataString & l(CLMBordx!CASUDReasons1, 100)
            Else
                DataString = DataString & l("", 100)
            End If
            'If CLMBordx!CASORemarks <> "" Then
            '    DataString = DataString & l(CLMBordx!CASORemarks, 3000)
            'Else
            '    DataString = DataString & l("", 3000)
            'End If
            
            DataString = DataString & l("", 984) & Chr(13)
            
            '--- Start (Do not delete this)--'
                Call UpdateWSAscii(ClaimNo, VersionNo)
            '-- end --'
            
            CLMBordx.MoveNext
            
            'Debug.Print DataString
            Print #1, DataString
            Print #2, DataString
            
   '     Next ii
     
        Loop
            
            Trailer = "T" & l(CLMBordx.recordCount, 10) & l("", 164)
            'Trailer = "T" & l(rstCount!totalrecord, 10) & l("", 164)
                'Debug.Print Trailer
                Print #1, Trailer
                Print #2, Trailer
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

Private Sub UpdateWSAscii(ClmNo As String, VerNo As String)
Dim WSAscii As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim strsqlWSAscii As String
    
    'strsql = "select CLMBordxWSASCII from Claim where CASNumber = '" & ClmNo & "' AND ClaimVersion = '" & VerNo & "'"
    'WSAscii.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    strAppendCase = "2"
    strAppend = " AND CASNumber = '" & ClmNo & "' AND ClaimVersion = '" & VerNo & "'"
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "Case_ClaimAscii"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    WSAscii.CursorLocation = adUseClient
    WSAscii.CursorType = adOpenForwardOnly
    WSAscii.CacheSize = 100
    Set WSAscii = cmd.Execute
    
    If Not WSAscii.EOF Then
        'With WSAscii
            '!CLMBordxWSASCII = "Y"
        'End With
        'WSAscii.Update
        strsqlWSAscii = "EXEC Case_CLMAscUpdate '" & ClmNo & "', Y"
        medixmasterconnWS.Execute strsqlWSAscii
    End If
    
    WSAscii.Close
    Set WSAscii = Nothing
End Sub

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub Form_Load()
        
    txtPath = SSRVPath & "ClaimASCI"
    txtLocalFileLocation = "C:\"
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
        
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
    CboExportStatus.AddItem "N - New"
    CboExportStatus.AddItem "R - Reprint"
    CboExportStatus.Text = "N - New"
    
    cboBordxWSStatus.AddItem "0 - New & Amendment"
    cboBordxWSStatus.AddItem "1 - Deleted"
    cboBordxWSStatus.Text = "0 - New & Amendment"
End Sub
