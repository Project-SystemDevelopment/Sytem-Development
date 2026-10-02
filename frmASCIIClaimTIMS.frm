VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmASCIIClaimTIMS 
   Caption         =   "General Claims ASCII TIMS"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   10050
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   9975
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "General Claims ASCII TIMS"
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
         TabIndex        =   19
         Top             =   60
         Width           =   6615
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9735
      Begin VB.CommandButton Command1 
         Caption         =   "Upload"
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
         Left            =   7680
         TabIndex        =   33
         Top             =   480
         Width           =   1440
      End
      Begin VB.TextBox txtImportFilename 
         Enabled         =   0   'False
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
         Left            =   1920
         TabIndex        =   31
         Top             =   480
         Width           =   4905
      End
      Begin VB.CommandButton cmdFileLocation 
         Caption         =   " >"
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
         Left            =   6960
         TabIndex        =   30
         Top             =   480
         Width           =   320
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
         ItemData        =   "frmASCIIClaimTIMS.frx":0000
         Left            =   4860
         List            =   "frmASCIIClaimTIMS.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3600
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
         Top             =   3840
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.ComboBox cboClient 
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
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   3660
         Width           =   3855
      End
      Begin VB.TextBox txtBordxNoFr 
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox txtBordxNoTo 
         Height          =   285
         Left            =   3960
         TabIndex        =   27
         Top             =   3960
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtReceiptDateTo 
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   4200
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox txtReceiptDateFr 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   4200
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5040
         Top             =   1800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label6 
         Caption         =   "File Name"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   1695
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
         Left            =   3600
         TabIndex        =   13
         Top             =   4245
         Width           =   195
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   4020
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Client"
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
         Left            =   480
         TabIndex        =   25
         Top             =   3675
         Width           =   915
      End
      Begin VB.Label txtFilNameWS 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   3240
         Width           =   1905
      End
      Begin VB.Label txtFilName 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   3240
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
         Left            =   4080
         TabIndex        =   12
         Top             =   3600
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
         Left            =   1800
         TabIndex        =   11
         Top             =   2940
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
         Left            =   480
         TabIndex        =   10
         Top             =   4215
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
         Left            =   480
         TabIndex        =   9
         Top             =   2715
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
         Left            =   1800
         TabIndex        =   8
         Top             =   2640
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
         Left            =   480
         TabIndex        =   7
         Top             =   3015
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
         TabIndex        =   6
         Top             =   3825
         Visible         =   0   'False
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
         Left            =   480
         TabIndex        =   5
         Top             =   3300
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   4320
      Width           =   9735
      Begin VB.CommandButton Command3 
         Caption         =   "&Print Result"
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
         Left            =   4920
         TabIndex        =   36
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Check"
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
         Left            =   7080
         TabIndex        =   35
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&OK"
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
         Left            =   6240
         TabIndex        =   23
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
         Left            =   8760
         TabIndex        =   22
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
         Left            =   7920
         TabIndex        =   21
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Rejected Cases"
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Accepted Cases"
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   840
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ListView lstCheck 
      Height          =   2415
      Left            =   120
      TabIndex        =   34
      Top             =   1560
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Case Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Version"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   8819
      EndProperty
   End
End
Attribute VB_Name = "frmASCIIClaimTIMS"
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
Dim BATNo As String
Dim SqlQry As String
Dim SaveClaim As New ADODB.Recordset
Dim FileName As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFileLocation_Click()
 CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  CommonDialog1.InitDir = "\\svr-doc\Bor_In"
  CommonDialog1.Flags = cdlOFNHideReadOnly
  CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
  CommonDialog1.ShowOpen
  txtImportFilename = CommonDialog1.FileName
  Exit Sub
ErrHandler:
  Exit Sub
End Sub

Private Sub CmdGenerate_Click()
On Error Resume Next
Dim TotalPrincipal As Integer
Dim TotalSupplementary As Integer
Dim RecordSaved
    TotalRecord = "0"
    Screen.MousePointer = vbHourglass
            Call PrincipalASCII
            If Len(FileName) Then
                Call STMBEmailNotification
            End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub STMBEmailNotification()
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim objMessage As Object
On Error GoTo debugs
Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "In-Patient Claims File to Upload"
objMessage.From = "notifications@medix.com.my"
objMessage.To = "norbaizura@takaful-ikhlas.com.my;norshah@takaful-ikhlas.com.my;"
objMessage.AddAttachment FileName
objMessage.CC = "trinath@medix.com.my; bryan@medix.com.my;thanuja@medix.com.my;hanis@medix.com.my;mahadhir@takaful-ikhlas.com.my;yngpyng@takaful-ikhlas.com.my;mjeffri@medix.com.my;"
objMessage.TextBody = "Dear Norbaizura" & vbCrLf & vbCrLf & _
                            "Please Upload Attached Claims In-Patinet File. " & vbCrLf & vbCrLf & _
                            " " & vbCrLf & vbCrLf & _
                            "Thank you." & vbCrLf & vbCrLf & _
                            "Regards, " & vbCrLf & vbCrLf & _
                            "MediExpress (Malaysia) Sdn Bhd" & vbCrLf & vbCrLf
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "notifications@medix.com.my"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "medix123"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update
objMessage.Send
debugs:
If Err.Description <> "" Then
    MsgBox Err.Description
Else
    MsgBox ("Claims File Submitted to Client")
End If
End Sub

Private Sub PrincipalASCII()
On Error Resume Next
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
Dim xlApp As excel.Application
Dim xlBook As excel.Workbook
Dim xlsheet As excel.WORKSHEET
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
Dim tmpNotCoPayLmt As String
Dim tmpNotCoPayAct As String
Dim tmpNotCoPayApp As String
Dim tmpICUTaxAct As String
Dim tmpICUTaxApp As String
Dim tmpAdmissionReason As String
Dim tmpSurgicalProc1 As String
Dim tmpSurgicalProc2 As String
Dim tmpSurgicalProc3 As String
Dim tmpSurgicalProc4 As String
Dim tmpSurgicalProc5 As String
Dim tmpCLMExcess As Double
Dim tmpMgmtNote As String
Dim tmpCLMOutpatient As Double
Dim CLMInv As New ADODB.Recordset
Dim strsqlInv As String
Dim tmpPostDate As String
Dim tmpOthersLmt As String

        BordxDateFr = txtReceiptDateFr
        BordxDateTo = txtReceiptDateTo
        PAYCode = Mid(cboPAYCode, 1, 2)
        strAppend = ""
        
        If Len(Trim(cboClient)) Then
            strAppend = strAppend + " AND MBMGroupCompany = '" & Trim(cboClient) & "'"
        End If
        
        If Len(Trim(TxtBordxNoFr)) Then
            strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordxNoFr)) & "'"
        End If
        
        If Len(Trim(TxtBordxNoTo)) Then
            strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordxNoTo)) & "'"
        End If
        
        If IsDate(txtReceiptDateFr) = True Then
            strAppend = strAppend + " And CLMBordxDate >= '" & Format(txtReceiptDateFr, "mm/dd/yyyy") & "'"
        End If
          
        If IsDate(txtReceiptDateTo) = True Then
            strAppend = strAppend + " And CLMBordxDate <= '" & Format(txtReceiptDateTo, "mm/dd/yyyy") & "'"
        End If
        '----- End -----'
        If Option1.Value = True Then
            strAppendCase = 3
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
            txtFilNameWS = Trim("W" & "TK" & Format(Date, "YYMMDD"))
            FileName = "\\svr-doc\TK Claims\" & txtFilNameWS & "(" & BATNo & ")"
            Open FileName For Output As #1

        Do Until CLMBordx.EOF
        
            strsqlbnf = "Select *  from BNF where INSCode = '" & Trim(CLMBordx!INSCode) & "' AND PLNCode = '" & Trim(CLMBordx!PLNCode) & "'"
            BNF.Open strsqlbnf, medixmasterconn, adOpenForwardOnly, adLockOptimistic
                      
            Do While Not BNF.EOF
                If BNF!BNFCode = "101" Then
                    tmpRBBnf = BNF!BNFClaimAbleAmount
                End If
                If BNF!BNFCode = "102" Then
                    If IsNumeric(BNF!BNFClaimAbleAmount) Then
                        tmpICUBnf = BNF!BNFClaimAbleAmount
                    Else
                        tmpICUBnf = IIf(BNF!BNFClaimAbleAmount = "FULL", "As Charge", 0)
                    End If
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
                    tmpGovHosIncBnf = IIf(Len(BNF!BNFClaimAbleAmount), BNF!BNFClaimAbleAmount, 0)
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
            ClaimNo = CLMBordx!CAsNumber
            VersionNo = IIf(Len(CLMBordx!claimversion), CLMBordx!claimversion, "")
            BordxNo = IIf(Len(CLMBordx!BordxRefNo), CLMBordx!BordxRefNo, "")
            BordxDate = IIf(Len(CLMBordx!CLMBordxDate), CLMBordx!CLMBordxDate, "")
            Rcounter = Rcounter + 1
            row = row + 1
            'xlApp.DisplayAlerts = False
            DataString = CLMBordx!CAsNumber & "|" '0
            DataString = DataString & CLMBordx!MBMPolicyNo & "|" '1
            DataString = DataString & CLMBordx!MBMGroupCompany & "|" '2
            DataString = DataString & CLMBordx!MBMIcBcPp & "|" '3
            DataString = DataString & CLMBordx!MBMName & "|" '4
            DataString = DataString & CLMBordx!ICNo & "|" '5
            DataString = DataString & CLMBordx!PatientName & "|" '6
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "dd/mm/yyyy") & "|" '7
            DataString = DataString & CLMBordx!CASHOSCode & "|" '8
            If CLMBordx!CASDischargeDate <> "" Then '9
                DataString = DataString & Format(CLMBordx!CASDischargeDate, "dd/mm/yyyy") & "|"
            Else
                DataString = DataString & Format(CLMBordx!CASDischargeDate, "dd/mm/yyyy") & "|"
            End If
            If CLMBordx!CASAdmissionReason <> "" Then '10
                tmpAdmissionReason = Replace(CLMBordx!CASAdmissionReason, ",", "")
                DataString = DataString & tmpAdmissionReason & "|"
            Else
                DataString = DataString & " " & "|"
            End If
            
            If CLMBordx!ICDCode1 <> "" Then '11
                DataString = DataString & CLMBordx!ICDCode1 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
            If CLMBordx!ICDCode2 <> "" Then '12
                DataString = DataString & CLMBordx!ICDCode2 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
            If CLMBordx!ICDCode3 <> "" Then '13
                DataString = DataString & CLMBordx!ICDCode3 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
                
            If CLMBordx!CASCptCode1 <> "" And Len(CLMBordx!CASCptCode1) = "5" Then '14
                tmpSurgicalProc1 = Replace(CLMBordx!CASCptCode1, ",", "")
                DataString = DataString & tmpSurgicalProc1 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
            
            If CLMBordx!CASCptCode2 <> "" And Len(CLMBordx!CASCptCode2) = "5" Then '15
                tmpSurgicalProc2 = Replace(CLMBordx!CASCptCode2, ",", "")
                DataString = DataString & tmpSurgicalProc2 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
                        
                        
            If CLMBordx!CASCptCode3 <> "" And Len(CLMBordx!CASCptCode3) = "5" Then '16
                tmpSurgicalProc3 = Replace(CLMBordx!CASCptCode3, ",", "")
                DataString = DataString & tmpSurgicalProc3 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
                
            If CLMBordx!CASCptCode4 <> "" And Len(CLMBordx!CASCptCode4) = "5" Then '17
                tmpSurgicalProc4 = Replace(CLMBordx!CASCptCode4, ",", "")
                DataString = DataString & tmpSurgicalProc4 & "|"
            Else
               DataString = DataString & " " & "|"
            End If
            
            If CLMBordx!CASCptCode5 <> "" And Len(CLMBordx!CASCptCode5) = "5" Then '18
                tmpSurgicalProc5 = Replace(CLMBordx!CASCptCode5, ",", "")
                DataString = DataString & tmpSurgicalProc5 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
                
            If CLMBordx!AttendDoctor1 <> "" Then '19
                DataString = DataString & CLMBordx!AttendDoctor1 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
            
            If CLMBordx!CASAttendDoctor2 <> "" Then '20
                DataString = DataString & CLMBordx!CASAttendDoctor2 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
            
            If CLMBordx!CASAttendDoctor3 <> "" Then '21
                DataString = DataString & CLMBordx!CASAttendDoctor3 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
                
            If CLMBordx!CASAttendDoctor4 <> "" Then '22
                DataString = DataString & CLMBordx!CASAttendDoctor4 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
            If CLMBordx!CASAttendDoctor5 <> "" Then '23
                DataString = DataString & CLMBordx!CASAttendDoctor5 & "|"
            Else
                DataString = DataString & " " & "|"
            End If
            
            DataString = DataString & VersionNo & "|"  '24
            DataString = DataString & BordxNo & "|" '25
            DataString = DataString & Format(BordxDate, "dd/mm/yyyy") & "|" '26
            
            If Len(CLMBordx!CLMBoardAmtLmt) Then
                DataString = DataString & Format(CLMBordx!CLMBoardAmtLmt, "###0.00") & "|"  '27
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMBoardAmtAct) Then '28
                DataString = DataString & Format(CLMBordx!CLMBoardAmtAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMBoardAmt) Then '29
                DataString = DataString & Format(CLMBordx!CLMBoardAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMBoardAmt < CLMBordx!CLMBoardAmtAct Then '30
                DataString = DataString & Format(CLMBordx!CLMBoardAmtAct, "###0.00") - IIf(IsNumeric(CLMBordx!CLMBoardAmt), Format(CLMBordx!CLMBoardAmt, "###0.00"), 0) & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMICUAmtLmt <> "" Then '31
                'DataString = DataString & tmpICUBnf & " PER DAY" & "|"
                DataString = DataString & CLMBordx!CLMICUAmtLmt & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMICUAmtAct) Then '32
                DataString = DataString & Format(CLMBordx!CLMICUAmtAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMICUAmt) Then '33
                DataString = DataString & Format(CLMBordx!CLMICUAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMICUAmt) = False And IsNumeric(CLMBordx!CLMICUAmtAct) = True Then '34
                DataString = DataString & Format(CLMBordx!CLMICUAmtAct, "###0.00") & "|"
            ElseIf IsNumeric(CLMBordx!CLMICUAmt) = True And IsNumeric(CLMBordx!CLMICUAmtAct) = True Then
                If CLMBordx!CLMICUAmt < CLMBordx!CLMICUAmtAct Then
                    If IsNumeric(CLMBordx!CLMICUAmt) And IsNumeric(CLMBordx!CLMICUAmtAct) Then
                        DataString = DataString & Format(CLMBordx!CLMICUAmtAct, "###0.00") - Format(CLMBordx!CLMICUAmt, "###0.00") & "|"
                    ElseIf IsNumeric(CLMBordx!CLMICUAmt) = False And IsNumeric(CLMBordx!CLMICUAmtAct) = True Then
                        DataString = DataString & Format(CLMBordx!CLMICUAmtAct, "###0.00") & "|"
                    End If
                Else
                    DataString = DataString & "0" & "|"
                End If
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMServicesLmt <> "" Then
                DataString = DataString & Format(CLMBordx!CLMServicesLmt, "###0.00") & "|" '35
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMServicesAct) Then '36
                DataString = DataString & Format(CLMBordx!CLMServicesAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMServices) Then '37
                DataString = DataString & Format(CLMBordx!CLMServices, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMServices) = False And IsNumeric(CLMBordx!CLMServicesAct) = True Then '38
                DataString = DataString & Format(CLMBordx!CLMServicesAct, "###0.00") & "|"
            ElseIf IsNumeric(CLMBordx!CLMServices) = True And IsNumeric(CLMBordx!CLMServicesAct) = True Then
                If CDbl(CLMBordx!CLMServices) < CDbl(CLMBordx!CLMServicesAct) Then
                    DataString = DataString & Format(CLMBordx!CLMServicesAct, "###0.00") - Format(CLMBordx!CLMServices, "###0.00") & "|"
                Else
                    DataString = DataString & "0" & "|"
                End If
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMPreHospLmt <> "" Then '39
                DataString = DataString & Format(CLMBordx!CLMPreHospLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMPreHospAct) Then '40
                DataString = DataString & Format(CLMBordx!CLMPreHospAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMPreHosp) Then '41
                DataString = DataString & Format(CLMBordx!CLMPreHosp, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMPreHosp < CLMBordx!CLMPreHospAct Then '42
                DataString = DataString & Format(CLMBordx!CLMPreHospAct, "###0.00") - Format(CLMBordx!CLMPreHosp, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMSurgicalAmtLmt <> "" Then '43
                DataString = DataString & Format(CLMBordx!CLMSurgicalAmtLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMSurgicalAmtAct) Then '44
                DataString = DataString & Format(CLMBordx!CLMSurgicalAmtAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMSurgicalAmt) Then '45
                DataString = DataString & Format(CLMBordx!CLMSurgicalAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMSurgicalAmt < CLMBordx!CLMSurgicalAmtAct Then '46
                DataString = DataString & Format(CLMBordx!CLMSurgicalAmtAct, "###0.00") - Format(CLMBordx!CLMSurgicalAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If

            
            If CLMBordx!CLMAnaLmt <> "" Then '47
                DataString = DataString & Format(CLMBordx!CLMAnaLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMAnaAct) Then '48
                DataString = DataString & Format(CLMBordx!CLMAnaAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMAna) Then '49
                DataString = DataString & Format(CLMBordx!CLMAna, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMAna) = False And IsNumeric(CLMBordx!CLMAnaAct) = True Then '50
                DataString = DataString & Format(CLMBordx!CLMAnaAct, "###0.00") & "|"
            ElseIf IsNumeric(CLMBordx!CLMAna) = True And IsNumeric(CLMBordx!CLMAnaAct) = True Then
                If CDbl(CLMBordx!CLMAna) < CDbl(CLMBordx!CLMAnaAct) Then
                    DataString = DataString & Format(CDbl(CLMBordx!CLMAnaAct), "###0.00") - Format(CDbl(CLMBordx!CLMAna), "###0.00") & "|"
                Else
                    DataString = DataString & "0" & "|"
                End If
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMTheatreLmt <> "" Then '51
                DataString = DataString & Format(CLMBordx!CLMTheatreLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMTheatreAct) Then '52
                DataString = DataString & Format(CLMBordx!CLMTheatreAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMTheatre) Then '53
                DataString = DataString & Format(CLMBordx!CLMTheatre, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMTheatre < CLMBordx!CLMTheatreAct Then '54
                DataString = DataString & Format(CLMBordx!CLMTheatreAct, "###0.00") - Format(CLMBordx!CLMTheatre, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            
            If CLMBordx!CLMNSPreHosLmt <> "" Then '55
                DataString = DataString & Format(CLMBordx!CLMNSPreHosLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMNSPreHosAct) Then '56
                DataString = DataString & Format(CLMBordx!CLMNSPreHosAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMNSPreHos) Then '57
                DataString = DataString & Format(CLMBordx!CLMNSPreHos, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMNSPreHos < CLMBordx!CLMNSPreHosAct Then '58
                DataString = DataString & Format(CLMBordx!CLMNSPreHosAct, "###0.00") - Format(CLMBordx!CLMNSPreHos, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMPhysicianLmt <> "" Then '59
                DataString = DataString & Format(CLMBordx!CLMPhysicianLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMPhysicianAct) Then '60
                DataString = DataString & Format(CLMBordx!CLMPhysicianAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMPhysician) Then '61
                DataString = DataString & Format(CLMBordx!CLMPhysician, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMPhysician < CLMBordx!CLMPhysicianAct Then '62
                DataString = DataString & Format(CLMBordx!CLMPhysicianAct, "###0.00") - IIf(IsNumeric(CLMBordx!CLMPhysician), Format(CLMBordx!CLMPhysician, "###0.00"), 0) & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMPostLmt <> "" Then '63
                DataString = DataString & Format(CLMBordx!CLMPostLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMPostAct) Then '64
                DataString = DataString & Format(CLMBordx!CLMPostAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMPost) Then '65
                DataString = DataString & Format(CLMBordx!CLMPost, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMPost < CLMBordx!CLMPostAct Then '66
                DataString = DataString & Format(CLMBordx!CLMPostAct, "###0.00") - Format(CLMBordx!CLMPost, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMEmergencyLmt <> "" Then '67
                DataString = DataString & Format(CLMBordx!CLMEmergencyLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMEmergencyAct) Then '68
                DataString = DataString & Format(CLMBordx!CLMEmergencyAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMEmergency) Then '69
                DataString = DataString & Format(CLMBordx!CLMEmergency, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMEmergency < CLMBordx!CLMEmergencyAct Then '70
                DataString = DataString & Format(CLMBordx!CLMEmergencyAct, "###0.00") - Format(CLMBordx!CLMEmergency, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            
            If CLMBordx!CLMAccPostAmtLmt <> "" Then '72
                DataString = DataString & Format(CLMBordx!CLMAccPostAmtLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMAccPostAmtAct) Then '72
                DataString = DataString & Format(CLMBordx!CLMAccPostAmtAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMAccPostAmt) Then '72
                DataString = DataString & Format(CLMBordx!CLMAccPostAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMAccPostAmt < CLMBordx!CLMAccPostAmtAct Then '72
                DataString = DataString & Format(CLMBordx!CLMAccPostAmtAct, "###0.00") - Format(CLMBordx!CLMAccPostAmt, "###0.00")
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMMonthlyLmt <> "" Then '75
                DataString = DataString & Format(CLMBordx!CLMMonthlyLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMMonthlyAct) Then '76
                DataString = DataString & Format(CLMBordx!CLMMonthlyAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMMonthly) Then '77
                DataString = DataString & Format(CLMBordx!CLMMonthly, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMMonthly < CLMBordx!CLMMonthlyAct Then '78
                DataString = DataString & Format(CLMBordx!CLMMonthlyAct, "###0.00") - Format(CLMBordx!CLMMonthly, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMAmbulanceLmt <> "" Then '79
                DataString = DataString & Format(CLMBordx!CLMAmbulanceLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMAmbulanceAct) Then '80
                DataString = DataString & Format(CLMBordx!CLMAmbulanceAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMAmbulance) Then '81
                DataString = DataString & Format(CLMBordx!CLMAmbulance, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMAmbulance < CLMBordx!CLMAmbulanceAct Then '82
                DataString = DataString & Format(CLMBordx!CLMAmbulanceAct, "###0.00") - Format(CLMBordx!CLMAmbulance, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            
            If CLMBordx!CLMGuardianAmtLmt <> "" Then '83
                DataString = DataString & Format(CLMBordx!CLMGuardianAmtLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMGuardianAmtAct) Then '84
                DataString = DataString & Format(CLMBordx!CLMGuardianAmtAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMGuardianAmt) Then '85
                DataString = DataString & Format(CLMBordx!CLMGuardianAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMGuardianAmt) Then
                If CLMBordx!CLMGuardianAmt < CLMBordx!CLMGuardianAmtAct Then '86
                    
                    DataString = DataString & Format(CLMBordx!CLMGuardianAmtAct, "###0.00") - Format(CLMBordx!CLMGuardianAmt, "###0.00") & "|"
                Else
                    DataString = DataString & "0" & "|"
                End If
            Else
            
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMMedicalLmt <> "" Then '87
                DataString = DataString & Format(CLMBordx!CLMMedicalLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!CLMMedicalAct) Then '88
                DataString = DataString & Format(CLMBordx!CLMMedicalAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMMedical) Then '89
                DataString = DataString & Format(CLMBordx!CLMMedical, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If CLMBordx!CLMMedical < CLMBordx!CLMMedicalAct Then '90
                If IsNumeric(CLMBordx!CLMMedical) Then
                    DataString = DataString & Format(CLMBordx!CLMMedicalAct, "###0.00") - Format(CLMBordx!CLMMedical, "###0.00") & "|"
                Else
                    DataString = DataString & Format(CLMBordx!CLMMedicalAct, "###0.00") & "|"
                End If
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CLMBordx!CLMBoardTaxAmtLmt <> "" Then '91
                DataString = DataString & Format(CLMBordx!CLMBoardTaxAmtLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
           
            tmpTaxAct = IIf(IsNumeric(CLMBordx!CLMBoardTaxAmtAct), CLMBordx!CLMBoardTaxAmtAct, 0)
            tmpTaxApp = IIf(IsNumeric(CLMBordx!CLMBoardTaxAmt), CLMBordx!CLMBoardTaxAmt, 0) '
            
            'tmpICUTaxAct = IIf(IsNumeric(CLMBordx!CLMICUTaxAmtAct), CLMBordx!CLMICUTaxAmtAct, 0)
            'tmpICUTaxApp = IIf(IsNumeric(CLMBordx!CLMICUTaxAmt), CLMBordx!CLMICUTaxAmt, 0)
            
            tmpTaxAct = CDbl(tmpTaxAct) '+ CDbl(tmpICUTaxAct)
            tmpTaxApp = CDbl(tmpTaxApp) '+ CDbl(tmpICUTaxApp)
            
            'tmpTaxAct = CDbl(tmpTaxAct) + IIf(IsNumeric(CLMBordx!CLMICUTaxAmtAct), CDbl(CLMBordx!CLMICUTaxAmtAct), 0)
            'tmpTaxApp = CDbl(tmpTaxApp) + IIf(IsNumeric(CLMBordx!CLMICUTaxAmt), CDbl(CLMBordx!CLMICUTaxAmt), 0)
            
            If IsNumeric(tmpTaxAct) Then '92
                DataString = DataString & tmpTaxAct & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            
            If IsNumeric(tmpTaxApp) Then '93
                DataString = DataString & tmpTaxApp & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If CDbl(tmpTaxApp) < CDbl(tmpTaxAct) Then '94
                DataString = DataString & Format(CDbl(tmpTaxAct), "###0.00") - Format(CDbl(tmpTaxApp), "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
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
            tmpNotCoPayLmt = 0
            tmpNotCoPayAct = 0
            tmpNotCoPayApp = 0
            
            'tmpNotCoPayLmt = IIf(Len(CLMBordx!CLMOthersNotCoPayLmt), CLMBordx!CLMOthersNotCoPayLmt, 0)
            If CLMBordx!CLMOthersNotCoPayLmt <> "" Then '91
                tmpNotCoPayLmt = CLMBordx!CLMOthersNotCoPayLmt
            Else
                tmpNotCoPayLmt = 0
            End If
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

            If IsNumeric(CLMBordx!CLMOthers1Lmt) Then
                tmpOthersLmt = CLMBordx!CLMOthers1Lmt
            Else
                tmpOthersLmt = 0
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
            'If IsNumeric(CLMBordx!CLMAdminAct) Then
            '    tmpOthersAct = CDbl(tmpOthersAct) + CDbl(CLMBordx!CLMAdminAct) + CDbl(tmpNotCoPayAct)
            'End If
            
            tmpOthersAct = CDbl(tmpOthersAct) '+ CDbl(tmpNotCoPayAct)
            
            tmpReimActOthers = CDbl(tmpReimPreH60) + CDbl(tmpReimPreGP) + CDbl(tmpReimPreMedic) + CDbl(tmpReimPreMR) + CDbl(tmpReimPreSpec) + CDbl(tmpReimPostHosp) + CDbl(tmpReimPostNotSame) + CDbl(tmpReimOthers)
           'tmpOthersAct = 0
            tmpOthersAct = CDbl(tmpOthersAct) + CDbl(tmpReimActOthers)

            
            'If IsNumeric(CLMBordx!CLMAdmin) Then
                'tmpOthersApp = CDbl(tmpOthersApp) + CDbl(CLMBordx!CLMAdmin) + CDbl(tmpNotCoPayApp)
            'End If
            
            tmpOthersApp = CDbl(tmpOthersApp) '+ CDbl(tmpNotCoPayApp)
            
            If IsNumeric(tmpOthersAct) Then '95
                DataString = DataString & Format(tmpOthersAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(tmpOthersApp) Then '96
                DataString = DataString & Format(tmpOthersApp, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If tmpOthersApp < tmpOthersAct Then '97
                If IsNumeric(tmpOthersAct) Then
                    DataString = DataString & tmpOthersAct - Format(tmpOthersApp, "###0.00") & "|"
                Else
                    DataString = DataString & Format(tmpOthersApp, "###0.00") & "|"
                End If
            Else
                DataString = DataString & "0" & "|"
            End If
            tmpCLMCoPayRB = 0
            If IsNumeric(CLMBordx!CLMCoPayRB) Then '98
                DataString = DataString & Format(CLMBordx!CLMCoPayRB, "###0.00") & "|"
                tmpCLMCoPayRB = IIf(IsNumeric(CLMBordx!CLMCoPayRB), CLMBordx!CLMCoPayRB, 0)
            Else
                DataString = DataString & "0" & "|"
            End If
            
            If IsNumeric(CLMBordx!CLMExGratiaAmt) Then '99
                DataString = DataString & Format(CLMBordx!CLMExGratiaAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            If IsNumeric(CLMBordx!LimitExc) Then '100
                DataString = DataString & Format(CLMBordx!LimitExc, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            '+ CDbl(tmpCLMCoPayRB)
            ' + CDbl(tmpReimActOthers)
            DataString = DataString & Format(CDbl(CLMBordx!CLMTotalActual), "###0.00") & "|" '101
            
            DataString = DataString & Format(CLMBordx!CLMBordxAmt, "###0.00") & "|"  '102
                        
            DataString = DataString & Format(CDbl((CLMBordx!CLMTotalActual) + CDbl(tmpReimActOthers)), "###0.00") - (Format(CLMBordx!CLMBordxAmt, "###0.00")) & "|"  '103
            
            DataString = DataString & CLMBordx!MBMTINumber & "|" '104
            
            DataString = DataString & CLMBordx!MBMTICoverID & "|" '105
            
            DataString = DataString & CLMBordx!InvPayTo & "|" '106
            
            DataString = DataString & CLMBordx!CLMExcess & "|" '107
            
            DataString = DataString & CLMBordx!INVPaidByMem & "|" '108
            
            DataString = DataString & CLMBordx!CLMPreviousBalance & "|" '109

            DataString = DataString & CLMBordx!CLMBTBGAmt & "|" '110
            
            DataString = DataString & CLMBordx!CLMBoardDay & "|" '111

            DataString = DataString & CLMBordx!CLMICUDay & "|" '112
            
            
            DataString = DataString & Format(CLMBordx!CLMTaxLmt, "###0.00") & "|" '113
            
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMTaxAct), CLMBordx!CLMTaxAct, 0) & "|" '114
            
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMTax), CLMBordx!CLMTax, 0) & "|"  '115
            
            tmpMgmtNote = ""
            
            tmpMgmtNote = IIf(Len(CLMBordx!CASORemarks1), CLMBordx!CASORemarks1, "") & IIf(Len(CLMBordx!CASORemarks2), CLMBordx!CASORemarks2, "") & IIf(Len(CLMBordx!CASORemarks3), CLMBordx!CASORemarks3, "") & IIf(Len(CLMBordx!CASORemarks4), CLMBordx!CASORemarks4, "") & IIf(Len(CLMBordx!CASORemarks5), CLMBordx!CASORemarks5, "") & IIf(Len(CLMBordx!CASORemarks6), CLMBordx!CASORemarks6, "")
            
            tmpMgmtNote = Replace(tmpMgmtNote, "|", "")
            
            If Len(tmpMgmtNote) Then '116
                DataString = DataString & tmpMgmtNote & "|"
            Else
                DataString = DataString & "" & "|"
            End If
            
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMCoPayRB), CLMBordx!CLMCoPayRB, 0) & "|"  '117
            
            'If CLMBordx!CLMTaxLmt <> "" Then '118
            '    DataString = DataString & Format(CLMBordx!CLMTaxLmt, "###0.00") & "|"
            'Else
            '    DataString = DataString & "0" & "|"
            'End If
            'If IsNumeric(CLMBordx!CLMCashAmtAct) Then '119
            '    DataString = DataString & Format(CLMBordx!CLMCashAmtAct, "###0.00") & "|"
            'Else
            '    DataString = DataString & "0" & "|"
            'End If
            'If IsNumeric(CLMBordx!CLMTax) Then '120
            '    DataString = DataString & Format(CLMBordx!CLMTax, "###0.00") & "|"
            'Else
            '    DataString = DataString & "0" & "|"
            'End If
            
            'If IsNumeric(CLMBordx!CLMTax) Then '121
            '    If CLMBordx!CLMTax < CLMBordx!CLMCashAmtAct Then '119
            '        DataString = DataString & Format(CLMBordx!CLMCashAmtAct, "###0.00") - Format(CLMBordx!CLMTax, "###0.00") & "|"
            '    Else
            '        DataString = DataString & "0" & "|"
             '   End If
            'else
            '    DataString = DataString & "0" & "|"
            'End If
            
            '118
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMBoardAmtActGST), CLMBordx!CLMBoardAmtActGST, 0) & "|"
            '119
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMICUAmtActGST), CLMBordx!CLMICUAmtActGST, 0) & "|"
            '120
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMServicesActGST), CLMBordx!CLMServicesActGST, 0) & "|"
            '121
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMTheatreActGST), CLMBordx!CLMTheatreActGST, 0) & "|"
            '122
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMPreHSAmtActGST), CLMBordx!CLMPreHSAmtActGST, 0) & "|"
            '123
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMSurgicalAmtActGST), CLMBordx!CLMSurgicalAmtActGST, 0) & "|"
            '124
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMAnaActGST), CLMBordx!CLMAnaActGST, 0) & "|"
            '125
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMNSPreHosActGST), CLMBordx!CLMNSPreHosActGST, 0) & "|"
            '126
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMPhysicianActGST), CLMBordx!CLMPhysicianActGST, 0) & "|"
            '127
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMPostActGST), CLMBordx!CLMPostActGST, 0) & "|"
            '128
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMEmergencyActGST), CLMBordx!CLMEmergencyActGST, 0) & "|"
            '129
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMAmbulanceActGST), CLMBordx!CLMAmbulanceActGST, 0) & "|"
            '130
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMOutpatientActGST), CLMBordx!CLMOutpatientActGST, 0) & "|"
            '131
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMMonthlyActGST), CLMBordx!CLMMonthlyActGST, 0) & "|"
            '132
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMOrganActGST), CLMBordx!CLMOrganActGST, 0) & "|"
            '133
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMMedicalActGST), CLMBordx!CLMMedicalActGST, 0) & "|"
            '134
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMPhoneActGST), CLMBordx!CLMPhoneActGST, 0) & "|"
            '135
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMGuardianAmtActGST), CLMBordx!CLMGuardianAmtActGST, 0) & "|"
            '136
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMAdminActGST), CLMBordx!CLMAdminActGST, 0) & "|"
            '137
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMTaxActGST), CLMBordx!CLMTaxActGST, 0) & "|"
            '138
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMOthers1ActGST), CLMBordx!CLMOthers1ActGST, 0) & "|"
            '139
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMOthers2ActGST), CLMBordx!CLMOthers2ActGST, 0) & "|"
            '140
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMCashAmtActGST), CLMBordx!CLMCashAmtActGST, 0) & "|"
            '141
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMBoardTaxAmtActGST), CLMBordx!CLMBoardTaxAmtActGST, 0) & "|"
            '142
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMICUTaxAmtActGST), CLMBordx!CLMICUTaxAmtActGST, 0) & "|"
            '143
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMLodgerTaxAmtActGST), CLMBordx!CLMLodgerTaxAmtActGST, 0) & "|"
            '144
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMMPreHSAmtActGST), CLMBordx!CLMMPreHSAmtActGST, 0) & "|"
            '145
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMCOPreHSSumActGST), CLMBordx!CLMCOPreHSSumActGST, 0) & "|"
            '146
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMCOMPreHSSumActGST), CLMBordx!CLMCOMPreHSSumActGST, 0) & "|"
            '147
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMAccPostAmtActGST), CLMBordx!CLMAccPostAmtActGST, 0) & "|"
            '148
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMPreHospActGST), CLMBordx!CLMPreHospActGST, 0) & "|"
            '149
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMOthersNotCoPayActGST), CLMBordx!CLMOthersNotCoPayActGST, 0) & "|"
            '150
            DataString = DataString & IIf(IsNumeric(CLMBordx!CLMHomeNCAmtActGST), CLMBordx!CLMHomeNCAmtActGST, 0) & "|"
            
            '151
            If CLMBordx!CLMOutpatientLmt <> "" Then
                If (IsNumeric(CLMBordx!CLMOutpatientLmt)) Then
                    DataString = DataString & Format(CLMBordx!CLMOutpatientLmt, "###0.00") & "|"
                Else
                    DataString = DataString & "0" & "|"
                End If
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '152
            If IsNumeric(CLMBordx!CLMOutpatientAct) Then
                DataString = DataString & Format(CLMBordx!CLMOutpatientAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '153
            If IsNumeric(CLMBordx!CLMOutpatientAct) Then
                DataString = DataString & Format(CLMBordx!CLMOutpatient, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '154
            If CLMBordx!CLMOutpatient < CLMBordx!CLMOutpatientAct Then
                tmpCLMOutpatient = IIf(IsNumeric(CLMBordx!CLMOutpatient), CLMBordx!CLMOutpatient, 0)
                DataString = DataString & Format(CLMBordx!CLMOutpatientAct, "###0.00") - Format(tmpCLMOutpatient, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
           

            '155
            If CLMBordx!CLMPreHSAmtLmt <> "" Then
                DataString = DataString & CLMBordx!CLMPreHSAmtLmt & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '156
            If IsNumeric(CLMBordx!CLMPreHSAmtAct) Then
                DataString = DataString & Format(CLMBordx!CLMPreHSAmtAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '157
            If IsNumeric(CLMBordx!CLMPreHSAmt) Then
                DataString = DataString & Format(CLMBordx!CLMPreHSAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '158
            'If CLMBordx!CLMPreHSAmt < CLMBordx!CLMPreHSAmtAct Then
                'tmpCLMOutpatient = IIf(IsNumeric(CLMBordx!CLMPreHSAmt), CLMBordx!CLMOutpatient, 0)
            '    DataString = DataString & Format(CLMBordx!CLMOutpatientAct, "###0.00") - Format(tmpCLMOutpatient, "###0.00") & "|"
            'Else
                DataString = DataString & "0" & "|"
            'End If
            

            '159
            If CLMBordx!CLMMPreHSAmtLmt <> "" Then
                DataString = DataString & CLMBordx!CLMMPreHSAmtLmt & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '160
            If IsNumeric(CLMBordx!CLMMPreHSAmtAct) Then
                DataString = DataString & Format(CLMBordx!CLMMPreHSAmtAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '161
            If IsNumeric(CLMBordx!CLMMPreHSAmt) Then
                DataString = DataString & Format(CLMBordx!CLMMPreHSAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            
            '162
            'If CLMBordx!CLMPreHSAmt < CLMBordx!CLMPreHSAmtAct Then
                'tmpCLMOutpatient = IIf(IsNumeric(CLMBordx!CLMPreHSAmt), CLMBordx!CLMOutpatient, 0)
            '    DataString = DataString & Format(CLMBordx!CLMOutpatientAct, "###0.00") - Format(tmpCLMOutpatient, "###0.00") & "|"
            'Else
                DataString = DataString & "0" & "|"
            'End If
            
            '163
            If CLMBordx!CLMICUTaxAmtLmt <> "" Then
                DataString = DataString & Format(CLMBordx!CLMICUTaxAmtLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '164
            If IsNumeric(CLMBordx!CLMICUTaxAmtAct) Then
                DataString = DataString & Format(CLMBordx!CLMICUTaxAmtAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '165
            If IsNumeric(CLMBordx!CLMICUTaxAmt) Then
                DataString = DataString & Format(CLMBordx!CLMICUTaxAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '166
            If CLMBordx!CLMICUTaxAmt < CLMBordx!CLMICUTaxAmtAct Then
                tmpICUTaxApp = IIf(IsNumeric(CLMBordx!CLMICUTaxAmt), CLMBordx!CLMICUTaxAmt, 0)
                DataString = DataString & Format(CLMBordx!CLMICUTaxAmtAct, "###0.00") - Format(CLMBordx!CLMICUTaxAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            
            '167
            If CLMBordx!CLMAdminLmt <> "" Then
                DataString = DataString & Format(CLMBordx!CLMAdminLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '168
            If IsNumeric(CLMBordx!CLMAdminAct) Then
                DataString = DataString & Format(CLMBordx!CLMAdminAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '169
            If IsNumeric(CLMBordx!CLMAdmin) Then
                DataString = DataString & Format(CLMBordx!CLMAdmin, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '170
            'If CLMBordx!CLMICUTaxAmt < CLMBordx!CLMICUTaxAmtAct Then
            '    tmpICUTaxApp = IIf(IsNumeric(CLMBordx!CLMICUTaxAmt), CLMBordx!CLMICUTaxAmt, 0)
            '    DataString = DataString & Format(CLMBordx!CLMICUTaxAmtAct, "###0.00") - Format(CLMICUTaxAmt, "###0.00") & "|"
            'Else
                DataString = DataString & "0" & "|"
            'End If
            
            
            '167
            If tmpNotCoPayLmt <> "" Then
                DataString = DataString & Format(tmpNotCoPayLmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '168
            If IsNumeric(tmpNotCoPayAct) Then
                DataString = DataString & Format(tmpNotCoPayAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '169
            If IsNumeric(tmpNotCoPayApp) Then
                DataString = DataString & Format(tmpNotCoPayApp, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '170
            'If CLMBordx!CLMICUTaxAmt < CLMBordx!CLMICUTaxAmtAct Then
            '    tmpICUTaxApp = IIf(IsNumeric(CLMBordx!CLMICUTaxAmt), CLMBordx!CLMICUTaxAmt, 0)
            '    DataString = DataString & Format(CLMBordx!CLMICUTaxAmtAct, "###0.00") - Format(CLMICUTaxAmt, "###0.00") & "|"
            'Else
                DataString = DataString & "0" & "|"
            'End If
            
            
            '171
            If CLMBordx!CLMLodgerTaxAmtLmt <> "" Then
                DataString = DataString & Format(IIf(IsNumeric(CLMBordx!CLMLodgerTaxAmtLmt), CLMBordx!CLMLodgerTaxAmtLmt, "###0.00"), "FULL") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '172
            If IsNumeric(CLMBordx!CLMLodgerTaxAmtAct) Then
                DataString = DataString & Format(CLMBordx!CLMLodgerTaxAmtAct, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '173
            If IsNumeric(CLMBordx!CLMLodgerTaxAmt) Then
                DataString = DataString & Format(CLMBordx!CLMLodgerTaxAmt, "###0.00") & "|"
            Else
                DataString = DataString & "0" & "|"
            End If
            
            '174
            'If CLMBordx!CLMICUTaxAmt < CLMBordx!CLMICUTaxAmtAct Then
            '    tmpICUTaxApp = IIf(IsNumeric(CLMBordx!CLMICUTaxAmt), CLMBordx!CLMICUTaxAmt, 0)
            '    DataString = DataString & Format(CLMBordx!CLMICUTaxAmtAct, "###0.00") - Format(CLMICUTaxAmt, "###0.00") & "|"
            'Else
                DataString = DataString & "0" & "|"
            'End If
            
            strsqlInv = ""
            tmpPostDate = ""
            strsqlInv = "SELECT INVFolUpDate From ClaimInvoice where CASNumber = '" & Trim(ClaimNo) & "' and INVWSVersion = '" & Trim(VersionNo) & "' "
            CLMInv.Open strsqlInv, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            
            Do While Not CLMInv.EOF
                If tmpPostDate = "" Then
                    tmpPostDate = IIf(Len(CLMInv!INVFolUpDate), CLMInv!INVFolUpDate, "")
                Else
                    If Len(CLMInv!INVFolUpDate) Then
                        tmpPostDate = tmpPostDate & "," & CLMInv!INVFolUpDate
                    End If
                End If
            CLMInv.MoveNext
            Loop
            CLMInv.Close
            Set CLMInv = Nothing
            
            '178
            DataString = DataString & tmpPostDate & "|"
            
            
           '   '179 MRI Limit
            'If CLMBordx!CLMMRILmt <> "" Then
            '    DataString = DataString & Format(IIf(IsNumeric(CLMBordx!CLMMRILmt), CLMBordx!CLMMRILmt, "###0.00"), "FULL") & "|"
            'Else
            '    DataString = DataString & "0" & "|"
            'End If
           '
           ' '180
           ' If IsNumeric(CLMBordx!CLMMRIAct1) Then
           '     DataString = DataString & Format(CLMBordx!CLMMRIAct1, "###0.00") & "|"
           ' Else
           '     DataString = DataString & "0" & "|"
           ' End If
           '
           '
           ' 'Approve  CLMMRI
           ' '181
           ' If IsNumeric(CLMBordx!CLMMRIActGST) Then
           '     DataString = DataString & Format(CLMBordx!CLMMRIActGST, "###0.00") & "|"
           ' Else
           '     DataString = DataString & "0" & "|"
           ' End If
           '
           ' '182 CLMMRIAppGST
           ' If IsNumeric(CLMBordx!CLMMRIAppGST) Then
           '     DataString = DataString & Format(CLMBordx!CLMMRIAppGST, "###0.00") & "|"
           ' Else
           '     DataString = DataString & "0" & "|"
           ' End If
            
            DataString = DataString & tmpOthersLmt
            
            tmpNotCoPayLmt = 0
            tmpNotCoPayAct = 0
            tmpNotCoPayApp = 0
            
            Print #1, DataString
            
            CLMBordx.MoveNext
        
        Loop
            
        'xlBook.SaveAs (Trim(txtLocalFileLocation) & Trim(txtFilName & ".csv"))
        
        'Call cmdExpSearch_Click
                                            
        
        'Set xlBook = Nothing
        'xlApp.Quit
        'Set xlApp = Nothing
        cmdGenerate.Enabled = True
        'MsgBox "Export successful", vbOKOnly + vbInformation, DialogBoxTitle
        
    End If
    CLMBordx.Close
    Set CLMBordx = Nothing
    Close #1
End Sub

Private Sub CaseASCII()
On Error Resume Next
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
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String
Dim CLMCount As String
Dim BatchDate As String
Dim xlApp As excel.Application
Dim xlBook As excel.Workbook
Dim xlsheet As excel.WORKSHEET
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
        DataString = ""
        strAppendCase = "6"
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
        End If
        
                    
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
            BATNo = GetCLMAsciiMedBatch(Mid("TK", 1, 2), Date)
            textField = Trim(Date)
            txtFilName = Trim("C" & "TK" & Format(Date, "YYMMDD"))
            Open txtLocalFileLocation & "\" & txtFilName & "(" & BATNo & ")" For Output As #1
        Rcounter = 0
        row = 0
        tmpString1 = "MEDIX"
        tmpString2 = "N"
        Do Until CLMBordx.EOF
        
            Rcounter = Rcounter + 1
            row = row + 1
            DataString = CLMBordx!CAsNumber & "|" '0
            DataString = DataString & CLMBordx!MBMPolicyNo & "," '1
            DataString = DataString & CLMBordx!MBMGroupCompany & "," '2
            DataString = DataString & CLMBordx!MBMIcBcPp & "," '3
            DataString = DataString & CLMBordx!MBMName & "," '4
            DataString = DataString & CLMBordx!ICNo & "," '5
            DataString = DataString & CLMBordx!PatientName & "," '6
            DataString = DataString & Format(CLMBordx!CASDateAdmitted, "dd/mm/yyyy") & "," '7
            DataString = DataString & CLMBordx!CASHOSCode & "," '8
            If CLMBordx!CASDischargeDate <> "" Then '9
                DataString = DataString & Format(CLMBordx!CASDischargeDate, "dd/mm/yyyy") & ","
            Else
                DataString = DataString & Format(CLMBordx!CASDischargeDate, "dd/mm/yyyy") & ","
            End If
            If CLMBordx!CASAdmissionReason <> "" Then '10
                tmpAdmissionReason = Replace(CLMBordx!CASAdmissionReason, ",", "")
                DataString = DataString & tmpAdmissionReason & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!FinalDiag <> "" Then '11
                DataString = DataString & CLMBordx!FinalDiag & ","
            Else
                DataString = DataString & " " & ","
            End If
            If CLMBordx!FinalDiag2 <> "" Then '12
                DataString = DataString & CLMBordx!FinalDiag2 & ","
            Else
                DataString = DataString & " " & ","
            End If
            If CLMBordx!FinalDiag3 <> "" Then '13
                DataString = DataString & CLMBordx!FinalDiag3 & ","
            Else
                DataString = DataString & " " & ","
            End If
                
            If CLMBordx!CASCptCode1 <> "" And Len(CLMBordx!CASCptCode1) = "5" Then '14
                tmpSurgicalProc1 = Replace(CLMBordx!CASCptCode1, ",", "")
                DataString = DataString & tmpSurgicalProc1 & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASCptCode2 <> "" And Len(CLMBordx!CASCptCode2) = "5" Then '15
                tmpSurgicalProc2 = Replace(CLMBordx!CASCptCode2, ",", "")
                DataString = DataString & tmpSurgicalProc2 & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASCptCode3 <> "" And Len(CLMBordx!CASCptCode3) = "5" Then '16
                tmpSurgicalProc3 = Replace(CLMBordx!CASCptCode3, ",", "")
                DataString = DataString & tmpSurgicalProc3 & ","
            Else
                DataString = DataString & " " & ","
            End If
                
            If CLMBordx!CASCptCode4 <> "" And Len(CLMBordx!CASCptCode4) = "5" Then '17
                tmpSurgicalProc4 = Replace(CLMBordx!CASCptCode4, ",", "")
                DataString = DataString & tmpSurgicalProc4 & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASCptCode5 <> "" And Len(CLMBordx!CASCptCode5) = "5" Then '18
                tmpSurgicalProc5 = Replace(CLMBordx!CASCptCode5, ",", "")
                DataString = DataString & tmpSurgicalProc5 & ","
            Else
                DataString = DataString & " " & ","
            End If
                
            If CLMBordx!AttendDoctor <> "" Then '19
                DataString = DataString & CLMBordx!AttendDoctor & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASAttendDoctor2 <> "" Then '20
                DataString = DataString & CLMBordx!CASAttendDoctor2 & ","
            Else
                DataString = DataString & " " & ","
            End If
            
            If CLMBordx!CASAttendDoctor3 <> "" Then '21
                DataString = DataString & CLMBordx!CASAttendDoctor3 & ","
            Else
                DataString = DataString & " " & ","
            End If
                
            If CLMBordx!CASAttendDoctor4 <> "" Then '22
                DataString = DataString & CLMBordx!CASAttendDoctor4 & ","
            Else
                DataString = DataString & " " & ","
            End If
            If CLMBordx!CASAttendDoctor5 <> "" Then '23
                DataString = DataString & CLMBordx!CASAttendDoctor5
            Else
                DataString = DataString & " "
            End If
            DataString = DataString & Chr(13)
            Print #1, DataString
            CLMBordx.MoveNext
        Loop
        MsgBox "Export successful", vbOKOnly + vbInformation, DialogBoxTitle
        cmdGenerate.Enabled = True
    End If
    Close #1
    CLMBordx.Close
    Set CLMBordx = Nothing
    
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim CAsNumber As String
Dim Version As String
Dim xlApp As New excel.Application
Dim xlWrkBook As New excel.Workbook
Dim xlWrkSheet As New excel.WORKSHEET
Dim Val1 As Integer
Dim IsFrst As Boolean
IsFrst = True
Val1 = 2

If Len(txtImportFilename) Then
 Set xlWrkBook = xlApp.Workbooks.Open(txtImportFilename)
 Set xlWrkSheet = xlWrkBook.Sheets.Item(1)
  Do Until Val1 = 60000
    If xlWrkSheet.Cells(Val1, 1) <> "" Then
        If IsFrst Then
            Call DeleteClaims
            IsFrst = False
        End If
        CAsNumber = CStr(xlWrkSheet.Cells(Val1, 1))
        Version = CStr(xlWrkSheet.Cells(Val1, 2))
        Call InsertClaims(CAsNumber, Version)
    Else
        xlApp.DisplayAlerts = True
        xlApp.Interactive = True
        Set xlWrkSheet = Nothing
        Set xlWrkBook = Nothing
        xlApp.Quit
        Set xlApp = Nothing
        MsgBox "Records have been uploaded!", vbOKOnly + vbCritical, DialogBoxTitle
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Val1 = Val1 + 1
  Loop
End If
End Sub

Public Sub DeleteClaims()
On Error Resume Next
    SqlQry = ""
    SqlQry = " Delete Tbl_TKClaims"
    SaveClaim.Open SqlQry, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    Set SaveClaim = Nothing
End Sub

Public Sub InsertClaims(CaseNo As String, Ver As String)
On Error Resume Next
    SqlQry = ""
    SqlQry = " Insert into Tbl_TKClaims (CLMPayorAppCode,CASNumber ) values ('" + Trim(CaseNo) + "','" + Trim(CStr(Ver)) + "')"
    SaveClaim.Open SqlQry, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    Set SaveClaim = Nothing
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim rCnt As Long
Dim Listcheck As ListItem
    rCnt = 0
    lstCheck.ListItems.Clear
    SqlQry = ""
    SqlQry = "select Tbl_TKClaims.CLMPayorAppCode as CASNumber , Tbl_TKClaims.CASNumber as Ver, (case when  View_ClaimASCII_TI.CASNumber is null then  'Not Uploaded'  else 'Uploaded' end) as Stats  from View_ClaimASCII_TI " & _
             "right outer join Tbl_TKClaims on Tbl_TKClaims.CLMPayorAppCode = View_ClaimASCII_TI.CASNumber and Tbl_TKClaims.CASNumber = View_ClaimASCII_TI.ClaimVersion"
    SaveClaim.Open SqlQry, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    Do Until SaveClaim.EOF
        With SaveClaim
             rCnt = rCnt + 1
             Set Listcheck = lstCheck.ListItems.Add(, , rCnt)
             Listcheck.SubItems(1) = SaveClaim!CAsNumber
             Listcheck.SubItems(2) = SaveClaim!Ver
             Listcheck.SubItems(3) = SaveClaim!Stats
        End With
        SaveClaim.MoveNext
    Loop
    SaveClaim.Close
    Set SaveClaim = Nothing
End Sub

Private Sub Command3_Click()
  Screen.MousePointer = vbHourglass
    If lstCheck.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstCheck)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    txtPath = SSRVPath & "ClaimASCI"
    txtLocalFileLocation = "C:\MedicaGenClaimASCII\"
    Call ComboListing
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String

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
    CboExportStatus.AddItem "N - New"
    CboExportStatus.AddItem "R - Reprint"
    CboExportStatus.Text = "N - New"
    cboClient.Clear
    strsql = "select CoGRPName from CompanyGrp where CoGRPTISM = 'Y'"
    PAYGRP.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not PAYGRP.EOF
       cboClient.AddItem Trim(PAYGRP!CoGRPName)
       PAYGRP.MoveNext
    Loop
    PAYGRP.Close
    Set PAYGRP = Nothing
    
End Sub













