VERSION 5.00
Begin VB.Form frmDataTransferMBM 
   Caption         =   "Data Transfer - MBM"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   6075
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   240
      Width           =   6015
      Begin VB.Label Label2 
         Caption         =   "File on PC"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   15
         Top             =   220
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Export File Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblPathName 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   195
         Width           =   2175
      End
      Begin VB.Label lblFileName 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2000
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   6015
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   200
         Width           =   3735
      End
      Begin VB.Label lblSuppProc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   25
         Top             =   915
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "Total No. of Record(s) Processed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lbltotPrinc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   0
         X2              =   6000
         Y1              =   1250
         Y2              =   1250
      End
      Begin VB.Label lblSuppSel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "Supplementary"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   21
         Top             =   915
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "Supplementary"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "Principal"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   19
         Top             =   920
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "Principal"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   18
         Top             =   620
         Width           =   855
      End
      Begin VB.Label lblSuccess 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   1640
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "No. of Record(s) Selected "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   620
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "No. of Record(s) Processed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   920
         Width           =   1935
      End
      Begin VB.Label lbltotSel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblPrincProc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   915
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   3020
      Width           =   6015
      Begin VB.CommandButton CmdGenerate 
         Caption         =   "&Generate"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   1
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   2
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         Top             =   200
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "MBM Data Transfer - Generate ASCII"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmDataTransferMBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataString As String
Dim tmpFileName As String
Dim tmpPathName As String
Dim tmpPay As String
Dim MBMNumber As String
Dim MBMPolicyNo As String
Dim MBMIcBcPp As String
Dim MBMName As String
Dim PAYCode As String
Dim MBMDOB As String
Dim RACCode As String
Dim MBMSex As String
Dim MBMPayorEffDate As String
Dim MBMPayorEXPDate As String
Dim AGECode As String
Dim MBMAdultChild As String
Dim INSCode As String
Dim PLNCode As String
Dim STACode As String
Dim MBMRenewFlag As String
Dim MBMAvailableLimit As String
Dim MBMValueDate As String
Dim MBMBordxDate As String
Dim MBMMemberType As String
Dim HLTCode As String
Dim MBMBatchNo As String
Dim MBMTakeover As String
Dim MBMDataStatus As String
Dim MBMCOVID As String
Dim MBMStatus As String
Dim MBMExpiredStatus As String
Dim MBMGroupCompany As String
Dim MBMDepartment As String
Dim MBMEmployeeNo As String
Dim MBMAgentCode As String
Dim MBMPolSubNo1 As String
Dim MBMPOLSubNo2 As String
Dim MBMPremium As String
Dim MBMBranch As String
Dim RecProcS As String
    
Private Sub GenerateAscii()
Dim FileName, Header As String
    lblSuccess.Visible = False
    If Trim(tmpFileName) <> "" Then
        Header = "F" & "MBM" & tmpPay & Format(Date, "ddmmyyyy")
        Open tmpPathName & tmpFileName For Output As #1
        
        Print #1, Trim(Header)
        
        Call AssignRecord
        
        Close #1
        lblSuccess.Visible = True
        lblSuccess = "File Generated  Successfully"
        MsgBox "File Generated  Successfully...", vbOKOnly + vbInformation, DialogBoxTitle
    End If
End Sub


Private Sub WriteRecord()
    MBMDOB = Format(MBMDOB, "ddmmyyyy")
    MBMValueDate = Format(MBMValueDate, "ddmmyyyy")
    MBMBordxDate = Format(MBMBordxDate, "ddmmyyyy")
    MBMPayorEffDate = Format(MBMPayorEffDate, "ddmmyyyy")
    MBMPayorEXPDate = Format(MBMPayorEXPDate, "ddmmyyyy")

    DataString = "P"
    DataString = DataString & l(MBMNumber, 15) & l(MBMCOVID, 2)
    DataString = DataString & l(MBMPolicyNo, 25) & l(MBMIcBcPp, 30)
    DataString = DataString & l(MBMMemberType, 1) & l(MBMDOB, 8) & l(MBMName, 60)
    DataString = DataString & l(AGECode, 2) & l(MBMAdultChild, 1)
    DataString = DataString & l(MBMPayorEffDate, 8)
    DataString = DataString & l(MBMPayorEXPDate, 8) & l(PLNCode, 4) & l(INSCode, 2)
    DataString = DataString & l(MBMAvailableLimit, 8) & l(MBMSex, 1) & l(RACCode, 2)
    DataString = DataString & l(STACode, 50) & l(MBMRenewFlag, 1)
    DataString = DataString & l(MBMValueDate, 8) & l(MBMBordxDate, 8) & l(MBMBranch, 60)
    DataString = DataString & l(HLTCode, 2) & l(MBMBatchNo, 4)
    DataString = DataString & l(MBMTakeover, 1) & l(MBMDataStatus, 1)
    DataString = DataString & l(MBMStatus, 1) & l(MBMExpiredStatus, 1) & l(MBMGroupCompany, 150)
    DataString = DataString & l(MBMDepartment, 50) & l(MBMEmployeeNo, 15) & l(MBMAgentCode, 15)
    DataString = DataString & l(MBMPolSubNo1, 10) & l(MBMPOLSubNo2, 10) & l(MBMPremium, 8)

End Sub

Private Function l(stringvalue As String, stringlenght As Integer)
Dim stringformat
Dim LengLoop As Integer

    For LengLoop = 1 To stringlenght
        stringformat = stringformat + "1"
    Next LengLoop
    LSet stringformat = stringvalue
    l = stringformat
End Function


Private Sub CboPayor_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CalNoofMBM
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboPayor.Text = ValidValue
            CboPayor.SelStart = 0
            CboPayor.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CboPayor_KeyUp(KeyCode As Integer, Shift As Integer)
    Call CalNoofMBM
End Sub

Private Sub CboPayor_LostFocus()
    Call CalNoofMBM
End Sub

Private Sub cmdClear_Click()
    CboPayor.Text = ""
    lblFileName = ""
    lbltotSel = ""
    lblSuppSel = ""
    lblPrincProc = ""
    lblSuppProc = ""
    lbltotPrinc = ""
    CmdGenerate.Enabled = False
    tmpFileName = ""
    lblSuccess.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGenerate_Click()
Dim X As Boolean
    
    X = CreatePath(tmpPathName)

    If X = False Then
        MsgBox "Error Creating Folder " & tmpPathName & "  !" & vbNewLine & _
                " Please create that folder manually to store the output of leaflet in sofe copy!"
    Else
        If lbltotSel = 0 Then
            MsgBox "No Record Selected !! ", vbOKOnly + vbCritical
        Else
            Screen.MousePointer = vbHourglass
            Call GenerateAscii
            'Call Form_KeyDown(117, 0)
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Function CreatePath(NewPath) As Boolean
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        CreatePath = True
    End If
End Function

Private Sub Form_Load()
    tmpPathName = "C:\DataTransfer\"
    lblPathName = tmpPathName
    Call ComboListing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub ComboListing()
Dim SelFields As String
Dim SelCriteria As String
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
    
    SelFields = ""
    SelCriteria = ""

    SelFields = "SELECT PAYCode as tmpCode, PAYCompanyName as tmpName "
    
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHISTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, SelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, "PAY")
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, SelCriteria)
    cmd.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    CboPayor.Clear
    Do Until TmpRec.EOF
        With TmpRec
            CboPayor.AddItem !tmpCode & " - " & !tmpName
            TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing

End Sub

Private Sub AssignRecord()
Dim SelCriteria As String
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim RecProcP As String
Dim Trailer As String
    RecProcP = 0
    RecProcS = 0
    SelCriteria = tmpPay
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "DataTransferSelectMBM"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("tmpPayor", adVarChar, adParamInput, 5, SelCriteria)
    cmd.Parameters.Append Prm1

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    Do While Not TmpRec.EOF
        With TmpRec
            RecProcP = RecProcP + 1
            lblPrincProc = RecProcP
            MBMNumber = IIf(Len(TmpRec!MBMNumber), TmpRec!MBMNumber, "")
            MBMCOVID = IIf(Len(TmpRec!MBMCOVID), TmpRec!MBMCOVID, "00")
            MBMPolicyNo = IIf(Len(TmpRec!MBMPolicyNo), TmpRec!MBMPolicyNo, "")
            MBMIcBcPp = IIf(Len(TmpRec!MBMIcBcPp), TmpRec!MBMIcBcPp, "")
            MBMName = IIf(Len(TmpRec!MBMName), TmpRec!MBMName, "")
            PAYCode = IIf(Len(TmpRec!PAYCode), TmpRec!PAYCode, "")
            MBMDOB = IIf(Len(TmpRec!MBMDOB), TmpRec!MBMDOB, "")
            MBMRenewFlag = IIf(Len(TmpRec!MBMRenewFlag), TmpRec!MBMRenewFlag, "")
            STACode = IIf(Len(TmpRec!STACode), TmpRec!STACode, "")
            RACCode = IIf(Len(TmpRec!RACCode), TmpRec!RACCode, "")
            MBMSex = IIf(Len(TmpRec!MBMSex), TmpRec!MBMSex, "")
            MBMPayorEffDate = IIf(Len(TmpRec!MBMPayorEffDate), TmpRec!MBMPayorEffDate, "")
            MBMPayorEXPDate = IIf(Len(TmpRec!MBMPayorEXPDate), TmpRec!MBMPayorEXPDate, "")
            AGECode = IIf(Len(TmpRec!AGECode), TmpRec!AGECode, "")
            MBMAdultChild = IIf(Len(TmpRec!MBMAdultChild), TmpRec!MBMAdultChild, "")
            INSCode = IIf(Len(TmpRec!INSCode), TmpRec!INSCode, "")
            PLNCode = IIf(Len(TmpRec!PLNCode), TmpRec!PLNCode, "")
            MBMAvailableLimit = IIf(IsNumeric(TmpRec!MBMAvailableLimit), TmpRec!MBMAvailableLimit, 0)
            MBMValueDate = IIf(Len(TmpRec!MBMValueDate), TmpRec!MBMValueDate, "")
            MBMBordxDate = IIf(Len(TmpRec!MBMBordxDate), TmpRec!MBMBordxDate, "")
            MBMMemberType = IIf(Len(TmpRec!MBMMemberType), TmpRec!MBMMemberType, "")
            HLTCode = IIf(Len(TmpRec!HLTCode), TmpRec!HLTCode, "")
            MBMBatchNo = IIf(Len(TmpRec!MBMBatchNo), TmpRec!MBMBatchNo, "")
            MBMTakeover = IIf(Len(TmpRec!MBMTakeover), TmpRec!MBMTakeover, "")
            MBMDataStatus = IIf(Len(TmpRec!MBMDataStatus), TmpRec!MBMDataStatus, "")
            MBMStatus = IIf(Len(TmpRec!MBMStatus), TmpRec!MBMStatus, "")
            MBMExpiredStatus = IIf(Len(TmpRec!MBMExpiredStatus), TmpRec!MBMExpiredStatus, "")
            MBMGroupCompany = IIf(Len(TmpRec!MBMGroupCompany), TmpRec!MBMGroupCompany, "")
            MBMDepartment = IIf(Len(TmpRec!MBMDepartment), TmpRec!MBMDepartment, "")
            MBMEmployeeNo = IIf(Len(TmpRec!MBMEmployeeNo), TmpRec!MBMEmployeeNo, "")
            MBMAgentCode = IIf(Len(TmpRec!MBMAgentCode), TmpRec!MBMAgentCode, "")
            MBMPolSubNo1 = IIf(Len(TmpRec!MBMPolSubNo1), TmpRec!MBMPolSubNo1, "")
            MBMPOLSubNo2 = IIf(Len(TmpRec!MBMPOLSubNo2), TmpRec!MBMPOLSubNo2, "")
            MBMPremium = IIf(IsNumeric(TmpRec!MBMPremium), TmpRec!MBMPremium, 0)
            MBMBranch = IIf(Len(TmpRec!MBMBranch), TmpRec!MBMBranch, "")
            Call WriteRecord
            Print #1, DataString
            
            If Trim(INSCode) = "H" Or Trim(INSCode) = "F" Then
                Call AssignSupp(MBMNumber)
            End If
            lbltotPrinc = RecProcP + RecProcS
            TmpRec.MoveNext
        End With
    Loop
    Trailer = "T" & l(CVar(RecProcP), 10) & l(CVar(RecProcS), 10)
    Print #1, Trailer
    TmpRec.Close
    Set TmpRec = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing

End Sub

Private Sub AssignSupp(tmpMBMNo As String)
Dim SuppSql As String
Dim SuppRec As New ADODB.Recordset
Dim MBMCNumber As String
Dim MBMCCoverID As String
Dim MBMCPLNCode As String
Dim MBMCMemberType As String
Dim MBMCName As String
Dim MBMCICBCPPNo As String
Dim MBMCDOB As String
Dim MBMCAdultChild As String
Dim MBMCAgeband As String
Dim MBMCRELCode As String
Dim MBMCAvailableLimit As String
Dim MBMCEffDate As String
Dim MBMCExpDate As String
Dim MBMCSex As String
    SuppSql = "SELECT MBMCNumber, MBMCCoverID, MBMCPLNCode ,MBMCMemberType, MBMCName, " & _
              "MBMCICBCPPNo, MBMCDOB,MBMCAdultChild,MBMCAgeband,MBMCRELCode,MBMCSex, " & _
              "MBMCAvailableLimit, MBMCEffDate,MBMCExpDate From MBMCoveredPersons " & _
              "Where MBMCNumber ='" & Trim(tmpMBMNo) & "'"
    SuppRec.Open SuppSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    Do While Not SuppRec.EOF
        RecProcS = RecProcS + 1
        lblSuppProc = RecProcS
        MBMCNumber = IIf(Len(SuppRec!MBMCNumber), SuppRec!MBMCNumber, "")
        MBMCCoverID = IIf(Len(SuppRec!MBMCCoverID), SuppRec!MBMCCoverID, "")
        MBMCPLNCode = IIf(Len(SuppRec!MBMCPLNCode), SuppRec!MBMCPLNCode, "")
        MBMCMemberType = IIf(Len(SuppRec!MBMCMemberType), SuppRec!MBMCMemberType, "")
        MBMCName = IIf(Len(SuppRec!MBMCName), SuppRec!MBMCName, "")
        MBMCICBCPPNo = IIf(Len(SuppRec!MBMCICBCPPNo), SuppRec!MBMCICBCPPNo, "")
        MBMCDOB = IIf(Len(SuppRec!MBMCDOB), Format(SuppRec!MBMCDOB, "ddmmyyyy"), "")
        MBMCAdultChild = IIf(Len(SuppRec!MBMCAdultChild), SuppRec!MBMCAdultChild, "")
        MBMCAgeband = IIf(Len(SuppRec!MBMCAgeband), SuppRec!MBMCAgeband, "")
        MBMCRELCode = IIf(Len(SuppRec!MBMCRELCode), SuppRec!MBMCRELCode, "")
        MBMCAvailableLimit = IIf(Len(SuppRec!MBMCAvailableLimit), SuppRec!MBMCAvailableLimit, "")
        MBMCEffDate = IIf(Len(SuppRec!MBMCEffDate), SuppRec!MBMCEffDate, "")
        MBMCExpDate = IIf(Len(SuppRec!MBMCExpDate), SuppRec!MBMCExpDate, "")
        MBMCSex = IIf(Len(SuppRec!MBMCSex), SuppRec!MBMCSex, "")
        If Trim(MBMCExpDate) = "" Then
            MBMCExpDate = MBMPayorEXPDate
        Else
            MBMCExpDate = Format(MBMCExpDate, "ddmmyyyy")
        End If
        If Trim(MBMCEffDate) = "" Then
            MBMCEffDate = MBMPayorEffDate
        Else
            MBMCEffDate = Format(MBMCEffDate, "ddmmyyyy")
        End If
        
        DataString = MBMCRELCode
        DataString = DataString & l(MBMCNumber, 15) & l(MBMCCoverID, 2)
        DataString = DataString & Space(25) & l(MBMCICBCPPNo, 30)
        DataString = DataString & l(MBMCMemberType, 1) & l(MBMCDOB, 8) & l(MBMCName, 60)
        DataString = DataString & l(MBMCAgeband, 2) & l(MBMCAdultChild, 1)
        DataString = DataString & l(MBMCEffDate, 8)
        DataString = DataString & l(MBMCExpDate, 8) & l(MBMCPLNCode, 4) & l(INSCode, 2)
        DataString = DataString & l(MBMCAvailableLimit, 8) & l(MBMCSex, 1)
        Print #1, DataString
        SuppRec.MoveNext
    Loop
End Sub

Private Sub CalNoofMBM()
Dim MBMRec As New ADODB.Recordset
Dim SuppRec As New ADODB.Recordset
Dim tmpTotMBM, tmpTotSupp As String
Dim MBMSql As String
Dim SuppSql As String

    tmpFileName = ""
    tmpPay = Mid(CboPayor.Text, 1, 2)
    tmpTotMBM = 0
    tmpTotSupp = 0
    CmdGenerate.Enabled = False
    If Trim(tmpPay) <> "" Then
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        MBMSql = "Select Count(*) as TotMBM From MBM WHERE PAYCode='" & tmpPay & "'"
        MBMRec.Open MBMSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not MBMRec.EOF
            tmpTotMBM = MBMRec!TotMBM
            MBMRec.MoveNext
        Loop
        MBMRec.Close
        Set MBMRec = Nothing
        
        SuppSql = "Select Count(*) as TotSupp From MBMCoveredPersons WHERE substring(MBMCNumber,1,2) ='" & tmpPay & "'"
        SuppRec.Open SuppSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not SuppRec.EOF
            tmpTotSupp = SuppRec!TotSupp
            SuppRec.MoveNext
        Loop
        SuppRec.Close
        Set SuppRec = Nothing
        
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        lbltotSel = tmpTotMBM
        lblSuppSel = tmpTotSupp
        
        If tmpTotMBM > 0 Then
            CmdGenerate.Enabled = True
            tmpFileName = tmpPay & Format(Date, "YYMMDD") & "MBM.txt"
            lblFileName = tmpFileName
        Else
            'MsgBox "No Record Selected !! ", vbOKOnly + vbCritical
        End If
    End If
End Sub
