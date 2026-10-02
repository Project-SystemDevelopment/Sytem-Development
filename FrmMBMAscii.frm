VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmMBMAscii 
   Caption         =   "Generate ASCII"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   7365
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   2880
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
         Height          =   435
         Left            =   5040
         TabIndex        =   4
         Top             =   260
         Width           =   960
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
         Height          =   435
         Left            =   6000
         TabIndex        =   5
         Top             =   260
         Width           =   960
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
         Height          =   435
         Left            =   4080
         TabIndex        =   3
         Top             =   260
         Width           =   960
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
         Left            =   240
         TabIndex        =   20
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label Label16 
         Caption         =   "Total Supplementary"
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
         Left            =   240
         TabIndex        =   19
         Top             =   570
         Width           =   1560
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
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   960
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
         Left            =   1800
         TabIndex        =   17
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   1440
         Width           =   5295
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
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2160
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   5280
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   1800
         Width           =   1305
         _ExtentX        =   2302
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
         Left            =   1680
         TabIndex        =   15
         Top             =   840
         Width           =   1905
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
         Left            =   360
         TabIndex        =   14
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
         Left            =   360
         TabIndex        =   13
         Top             =   1545
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
         Left            =   360
         TabIndex        =   12
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
         Left            =   1680
         TabIndex        =   11
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
         Left            =   360
         TabIndex        =   10
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
         Left            =   360
         TabIndex        =   9
         Top             =   1905
         Width           =   915
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
         Left            =   1680
         TabIndex        =   8
         Top             =   540
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Export Status"
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
         Left            =   360
         TabIndex        =   7
         Top             =   2265
         Width           =   1035
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Generate ASCII"
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
      TabIndex        =   21
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "FrmMBMAscii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    
    cmdGenerate.Enabled = False
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If cboPAYCode = "" Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPAYCode.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Bordx Date ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    Else
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation, DialogBoxTitle)
        If RecordSaved = vbYes Then
    
            Screen.MousePointer = vbHourglass
            
            If Mid(CboExportStatus, 1, 1) = "N" Then
                Call PrincipalASCII
            Else
                Call RePrintASCII
            End If
            
            Screen.MousePointer = vbDefault
            
        End If
        
        txtTotalPrincipals = l(txtTotalPrincipals, 10)
        txtTotalSupplementary1 = l(txtTotalSupplementary1, 10)
    End If
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    cmdGenerate.Enabled = True
End Sub

Private Sub PrincipalASCII()
Dim strsql As String
Dim MBM As New ADODB.Recordset
Dim BordxDate As String
Dim PAYCode As String


        BordxDate = txtReceiptDate
        PAYCode = Mid(cboPAYCode, 1, 2)
        
        strsql = " Select * from MBMExcellUpload where MBMBordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'" & _
                 " AND PAYCode = '" & PAYCode & "' AND ExportStatus = '0' ORDER BY No "
        strCount = " Select count(*) as totalRecord From MBMExcellUpload where MBMBordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'" & _
                   " AND PAYCode = '" & PAYCode & "' AND ExportStatus = '0'"
        strCount = strCount
                    
        Set rstCount = medixmasterconn.Execute(strCount)
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    If rstCount!TotalRecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
            textField = Trim(txtReceiptDate)
                  
            txtFilName = Trim(Mid(cboPAYCode, 1, 2) & Format(textField, "YYMMDD"))
                
            Open txtPath & "\" & txtFilName For Output As #1
            Open txtLocalFileLocation & txtFilName For Output As #2
            
            Header = Trim("F" & Mid(cboPAYCode, 1, 2) & Mid(txtReceiptDate, 1, 2) & Mid(txtReceiptDate, 4, 2) & Mid(txtReceiptDate, 7, 4) & l("", 174))
            
            Print #1, Trim(Header)
            Print #2, Trim(Header)
            'Debug.Print Header
        
        
        Do Until MBM.EOF
        
        For ii = 1 To rstCount!TotalRecord
            If MBM!RELCode = "P" Then
                TotalPrincipal = TotalPrincipal + 1
                txtTotalPrincipals = TotalPrincipal
            Else
                TotalSupplementary = TotalSupplementary + 1
                txtTotalSupplementary1 = TotalSupplementary
            End If
            
            
            DataString = l(MBM!RELCode, 1)
            DataString = DataString & l(MBM!MBMPolicyNo, 25)
            DataString = DataString & l(MBM!MBMAgentCode, 15)
            DataString = DataString & l(MBM!MBMSalutation, 5)
            DataString = DataString & l(MBM!MBMName, 60)
            DataString = DataString & l(MBM!MBMAddress1, 30)
            DataString = DataString & l(MBM!MBMAddress2, 30)
            DataString = DataString & l(MBM!MBMAddress3, 30)
            DataString = DataString & l(MBM!CTYCode, 20)
            DataString = DataString & l(MBM!MBMPostCode, 5)
            DataString = DataString & l(MBM!STACode, 15)
            DataString = DataString & l(MBM!MBMTelnoH, 12)
            DataString = DataString & l(MBM!MBMTelNoO, 12)
            DataString = DataString & l(MBM!MBMTelNoM, 12)
            DataString = DataString & l(MBM!MBMIcBcPp, 20)
            DataString = DataString & l(Format(MBM!MBMDOB, "ddmmyyyy"), 8)
            DataString = DataString & l(MBM!MBMSex, 1)
            DataString = DataString & l(MBM!NATCode, 10)
            DataString = DataString & l(MBM!MBMWeight, 3)
            DataString = DataString & l(MBM!MBMHeight, 3)
            DataString = DataString & l(MBM!LGNCode, 1)
            DataString = DataString & l(MBM!RACCode, 1)
            DataString = DataString & l(MBM!MBMOccupation, 30)
            DataString = DataString & l(MBM!MBMWorkNature, 1)
            DataString = DataString & l(MBM!MBMBlood, 3)
            DataString = DataString & l(MBM!MBMAllergic, 20)
            DataString = DataString & l(Format(MBM!MBMPayorEffDate, "ddmmyyyy"), 8)
            DataString = DataString & l(Format(MBM!MBMPayorEXPDate, "ddmmyyyy"), 8)
            DataString = DataString & l(0 & MBM!PLNCode, 4)
            DataString = DataString & l(MBM!INSCode, 2)
            DataString = DataString & l(MBM!MBMGroupCompany, 40)
            DataString = DataString & l(MBM!MBMRenewFlag, 1)
            DataString = DataString & l(MBM!MBMExclusion, 2500)
            DataString = DataString & l(MBM!MBMSmoking, 1)
            DataString = DataString & l(MBM!MBMAlcohol, 1)
            DataString = DataString & l(MBM!MBMMaritalStatus, 1)
            DataString = DataString & l(MBM!BRNCode, 15)
            DataString = DataString & l(MBM!MBMPrevPolNo, 25)
            DataString = DataString & l(MBM!MBMPrevMemNo, 18)
            DataString = DataString & l(MBM!MBMEmployeeNo, 15)
            DataString = DataString & l(Format(MBM!MBMJoinedDate, "ddmmyyyy"), 8)
            DataString = DataString & l(MBM!Clinical, 1)
            DataString = DataString & l(MBM!ClinicName1, 50)
            DataString = DataString & l(MBM!ClinicName2, 50)
            DataString = DataString & l(MBM!ClinicName3, 50)
            DataString = DataString & l(MBM!ClinicName4, 50)
            DataString = DataString & l(MBM!MBMDepartment, 30)
            DataString = DataString & l(MBM!MBMGRPCompanyAdd1, 30)
            DataString = DataString & l(MBM!MBMGRPCompanyAdd2, 30)
            DataString = DataString & l(MBM!MBMGRPCompanyAdd3, 30)
            DataString = DataString & l(MBM!MBMGRPCompanyAdd4, 30)
            DataString = DataString & l(MBM!MBMPrevTakeOverPlnCode, 4)
            DataString = DataString & l(MBM!MBMInstallment, 1)
            DataString = DataString & l(MBM!MBMPolSubNo1, 10)
            DataString = DataString & l(MBM!MBMPOLSubNo2, 10)
            DataString = DataString & l(Format(MBM!MBMPOLIssuedDate, "ddmmyyyy"), 8)
            DataString = DataString & l(MBM!MBMRiskNo, 4)
            DataString = DataString & l(MBM!MBMTransNo, 5)
            DataString = DataString & l(MBM!MBMPayNo, 8)
            DataString = DataString & l(MBM!MBMPAYSeqNo, 2)
            DataString = DataString & l(MBM!MBMClientNo, 8)
            DataString = DataString & l(MBM!MBMClientID, 2)
            DataString = DataString & l(MBM!MBMPrevPLNCode, 15)
            DataString = DataString & l(MBM!MBMRBAmt, 8)
            DataString = DataString & l(MBM!MBMPrevInsCom, 60)
            DataString = DataString & l(MBM!MBMPAYRemarks, 60)
            'DataString = DataString & l("", 2500)
            DataString = DataString & l("", 6) & Chr(13)
            
            MBM.MoveNext
            
            'Debug.Print DataString
            Print #1, DataString
            Print #2, DataString
                 
            
        Next ii
     
        Loop
            
            Call UpdateAsciiHIS
                        
            Trailer = "T" & l(txtTotalPrincipals, 10) & l(txtTotalSupplementary1, 10) & l("", 164)
                'Debug.Print Trailer
                Print #1, Trailer
                Print #2, Trailer
                cmdGenerate.Enabled = False
                Close #1
                Close #2
                
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
    End If
    MBM.Close
    Set MBM = Nothing
    Close #1
    Close #2
End Sub

Private Sub RePrintASCII()
Dim strsql As String
Dim MBM As New ADODB.Recordset
Dim BordxDate As String
Dim PAYCode As String

        BordxDate = txtReceiptDate
        PAYCode = Mid(cboPAYCode, 1, 2)
                
        strsql = " Select * from MBMExcellUpload where MBMBordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'" & _
                 " AND PAYCode = '" & PAYCode & "' AND ExportStatus = '1' ORDER BY No "
        strCount = " Select count(*) as totalRecord From MBMExcellUpload where MBMBordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'" & _
                   " AND PAYCode = '" & PAYCode & "' AND ExportStatus = '1'"
        strCount = strCount
                    
        Set rstCount = medixmasterconn.Execute(strCount)
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    If rstCount!TotalRecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
            textField = Trim(txtReceiptDate)
                  
            txtFilName = Trim(Mid(cboPAYCode, 1, 2) & Format(textField, "YYMMDD"))
                
            Open txtPath & "\" & txtFilName For Output As #1
            Open txtLocalFileLocation & txtFilName For Output As #2
            
            Header = Trim("F" & Mid(cboPAYCode, 1, 2) & Mid(txtReceiptDate, 1, 2) & Mid(txtReceiptDate, 4, 2) & Mid(txtReceiptDate, 7, 4) & l("", 174))
            
            Print #1, Trim(Header)
            Print #2, Trim(Header)
            'Debug.Print Header
        
        
        Do Until MBM.EOF
        
        For ii = 1 To rstCount!TotalRecord
            If MBM!RELCode = "P" Then
                TotalPrincipal = TotalPrincipal + 1
                txtTotalPrincipals = TotalPrincipal
            Else
                TotalSupplementary = TotalSupplementary + 1
                txtTotalSupplementary1 = TotalSupplementary
            End If
            
            
            DataString = l(MBM!RELCode, 1)
            DataString = DataString & l(MBM!MBMPolicyNo, 25)
            DataString = DataString & l(MBM!MBMAgentCode, 15)
            DataString = DataString & l(MBM!MBMSalutation, 5)
            DataString = DataString & l(MBM!MBMName, 60)
            DataString = DataString & l(MBM!MBMAddress1, 30)
            DataString = DataString & l(MBM!MBMAddress2, 30)
            DataString = DataString & l(MBM!MBMAddress3, 30)
            DataString = DataString & l(MBM!CTYCode, 20)
            DataString = DataString & l(MBM!MBMPostCode, 5)
            DataString = DataString & l(MBM!STACode, 15)
            DataString = DataString & l(MBM!MBMTelnoH, 12)
            DataString = DataString & l(MBM!MBMTelNoO, 12)
            DataString = DataString & l(MBM!MBMTelNoM, 12)
            DataString = DataString & l(MBM!MBMIcBcPp, 20)
            DataString = DataString & l(Format(MBM!MBMDOB, "ddmmyyyy"), 8)
            DataString = DataString & l(MBM!MBMSex, 1)
            DataString = DataString & l(MBM!NATCode, 10)
            DataString = DataString & l(MBM!MBMWeight, 3)
            DataString = DataString & l(MBM!MBMHeight, 3)
            DataString = DataString & l(MBM!LGNCode, 1)
            DataString = DataString & l(MBM!RACCode, 1)
            DataString = DataString & l(MBM!MBMOccupation, 30)
            DataString = DataString & l(MBM!MBMWorkNature, 1)
            DataString = DataString & l(MBM!MBMBlood, 3)
            DataString = DataString & l(MBM!MBMAllergic, 20)
            DataString = DataString & l(Format(MBM!MBMPayorEffDate, "ddmmyyyy"), 8)
            DataString = DataString & l(Format(MBM!MBMPayorEXPDate, "ddmmyyyy"), 8)
            DataString = DataString & l(0 & MBM!PLNCode, 4)
            DataString = DataString & l(MBM!INSCode, 2)
            DataString = DataString & l(MBM!MBMGroupCompany, 40)
            DataString = DataString & l(MBM!MBMRenewFlag, 1)
            DataString = DataString & l(MBM!MBMExclusion, 2500)
            DataString = DataString & l(MBM!MBMSmoking, 1)
            DataString = DataString & l(MBM!MBMAlcohol, 1)
            DataString = DataString & l(MBM!MBMMaritalStatus, 1)
            DataString = DataString & l(MBM!BRNCode, 15)
            DataString = DataString & l(MBM!MBMPrevPolNo, 25)
            DataString = DataString & l(MBM!MBMPrevMemNo, 18)
            DataString = DataString & l(MBM!MBMEmployeeNo, 15)
            DataString = DataString & l(Format(MBM!MBMJoinedDate, "ddmmyyyy"), 8)
            DataString = DataString & l(MBM!Clinical, 1)
            DataString = DataString & l(MBM!ClinicName1, 50)
            DataString = DataString & l(MBM!ClinicName2, 50)
            DataString = DataString & l(MBM!ClinicName3, 50)
            DataString = DataString & l(MBM!ClinicName4, 50)
            DataString = DataString & l(MBM!MBMDepartment, 30)
            DataString = DataString & l(MBM!MBMGRPCompanyAdd1, 30)
            DataString = DataString & l(MBM!MBMGRPCompanyAdd2, 30)
            DataString = DataString & l(MBM!MBMGRPCompanyAdd3, 30)
            DataString = DataString & l(MBM!MBMGRPCompanyAdd4, 30)
            DataString = DataString & l(MBM!MBMPrevTakeOverPlnCode, 4)
            DataString = DataString & l(MBM!MBMInstallment, 1)
            DataString = DataString & l(MBM!MBMPolSubNo1, 10)
            DataString = DataString & l(MBM!MBMPOLSubNo2, 10)
            DataString = DataString & l(Format(MBM!MBMPOLIssuedDate, "ddmmyyyy"), 8)
            DataString = DataString & l(MBM!MBMRiskNo, 4)
            DataString = DataString & l(MBM!MBMTransNo, 5)
            DataString = DataString & l(MBM!MBMPayNo, 8)
            DataString = DataString & l(MBM!MBMPAYSeqNo, 2)
            DataString = DataString & l(MBM!MBMClientNo, 8)
            DataString = DataString & l(MBM!MBMClientID, 2)
            DataString = DataString & l(MBM!MBMPrevPLNCode, 15)
            DataString = DataString & l(MBM!MBMRBAmt, 8)
            DataString = DataString & l(MBM!MBMPrevInsCom, 60)
            DataString = DataString & l(MBM!MBMPAYRemarks, 60)
            'DataString = DataString & l("", 2500)
            DataString = DataString & l("", 6) & Chr(13)
            
            MBM.MoveNext
            
            'Debug.Print DataString
            Print #1, DataString
            Print #2, DataString
                 
            
        Next ii
     
        Loop
            
            'Call UpdateAsciiHIS
                        
            Trailer = "T" & l(txtTotalPrincipals, 10) & l(txtTotalSupplementary1, 10) & l("", 164)
                'Debug.Print Trailer
                Print #1, Trailer
                Print #2, Trailer
                cmdGenerate.Enabled = False
                Close #1
                Close #2
                
                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
                cmdGenerate.Enabled = True
    End If
    MBM.Close
    Set MBM = Nothing
    Close #1
    Close #2
End Sub

Private Sub UpdateAsciiHIS()
Dim AsciiHIS As New ADODB.Recordset
Dim strsql As String
Dim frm As Form
    
    strsql = "select * from MBMExcellUpload where MBMBordxDate = '" & Format(txtReceiptDate, "MM/DD/YYYY") & "' AND PAYCode = '" & Mid(cboPAYCode, 1, 2) & "'"
    
    AsciiHIS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    Do Until AsciiHIS.EOF
        
        With AsciiHIS
            
            !LastUpdateUser = Logon
            !LastUpdateDate = Date
            !Filename = txtFilName
            !ExportStatus = "1"
            
            AsciiHIS.Update
        End With
        AsciiHIS.MoveNext
    Loop
    
    AsciiHIS.Close
    Set AsciiHIS = Nothing
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
        
    txtPath = DocPath & "ASCI"
    txtLocalFileLocation = "C:\"
    Call ComboListing
    
End Sub

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
    Else
        If txtReceiptDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDate.SetFocus
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
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    CboExportStatus.AddItem "N - New"
    CboExportStatus.AddItem "R - Reprint"
    CboExportStatus.Text = "N - New"
    
End Sub


