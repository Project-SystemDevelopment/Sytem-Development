VERSION 5.00
Begin VB.Form frmDataTransferCAS 
   Caption         =   "Data Transfer - CASE"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   2640
      Width           =   6015
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
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   6015
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   200
         Width           =   3735
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
         Top             =   1245
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label lblNoOfRecProcess 
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
         TabIndex        =   15
         Top             =   920
         Width           =   2175
      End
      Begin VB.Label lblNoOfRec 
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
         TabIndex        =   14
         Top             =   600
         Width           =   2175
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
         TabIndex        =   13
         Top             =   920
         Width           =   1935
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
         TabIndex        =   12
         Top             =   620
         Width           =   1815
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
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   6015
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
         TabIndex        =   9
         Top             =   480
         Width           =   2175
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
         TabIndex        =   8
         Top             =   195
         Width           =   2175
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
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
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
         TabIndex        =   6
         Top             =   220
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "CASE Data Transfer - Generate ASCII"
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
      TabIndex        =   4
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmDataTransferCAS"
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
Dim CASNumber As String
Dim MBMPatientCovID As String
Dim CASStatus As String
Dim CASDischargeDate As String
Dim CaseRegDate As String
Dim CASDateAdmitted As String
Dim CASHOSCode As String
Dim ModalityCode As String
Dim PayTypeCode As String
Dim ClaimNatureCode As String
Dim CASClaimability As String
Dim CASStayType As String
Dim CASPAYType As String
Dim CASPending As String
Dim CASORemarks As String
Dim CASAttendDoctor1 As String
Dim CASAttendDoctor2 As String
Dim CASAttendDoctor3 As String
Dim CASAdmissionReason As String
Dim ICDCode1 As String
Dim ICDCode2 As String
Dim ICDCode3 As String
Dim DIAFinalDiag1 As String
Dim DIAFinalDiag2 As String
Dim DIAFinalDiag3 As String
Dim FirstICDCode1 As String
Dim FirstICDCode2 As String
Dim FirstICDCode3 As String
Dim DIAPrelimDiag1 As String
Dim DIAPrelimDiag2 As String
Dim DIAPrelimDiag3 As String
Dim CASREPCode1 As String
Dim CASREPCode2 As String
Dim CASREPCode3 As String
Dim CASOPROCRequired1 As String
Dim CASOPROCRequired2 As String
Dim CASOPROCRequired3 As String
Dim CASOPROCRequired4 As String
Dim CASOPROCRequired5 As String
Dim CASOProcInvestigationDone1 As String
Dim CASOProcInvestigationDone2 As String
Dim CASOProcInvestigationDone3 As String
Dim CASOMedSurMgmt1 As String
Dim CASOMedSurMgmt2 As String
Dim CASOMedSurMgmt3 As String


Private Sub GenerateAscii()
Dim FileName, Header As String
    lblSuccess.Visible = False
    If Trim(tmpFileName) <> "" Then
        Header = "F" & "CAS" & tmpPay & Format(Date, "ddmmyyyy")
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
    CASDischargeDate = Format(CASDischargeDate, "ddmmyyyy")
    CaseRegDate = Format(CaseRegDate, "ddmmyyyy")
    CASDateAdmitted = Format(CASDateAdmitted, "ddmmyyyy")

    DataString = ""
    DataString = DataString & l(CASNumber, 15) & l(MBMNumber, 15) & l(MBMPatientCovID, 2)
    DataString = DataString & l(CASStatus, 1) & l(CASDischargeDate, 8) & l(CaseRegDate, 8)
    DataString = DataString & l(CASDateAdmitted, 8) & l(CASHOSCode, 10) & l(ModalityCode, 2)
    DataString = DataString & l(PayTypeCode, 2) & l(ClaimNatureCode, 5) & l(CASClaimability, 1)
    DataString = DataString & l(CASStayType, 10) & l(CASPAYType, 10) & l(CASPending, 1)
    'DataString = DataString & l(CASORemarks, 3900)
    DataString = DataString & Space(3900)
    DataString = DataString & l(CASAttendDoctor1, 100) & l(CASAttendDoctor2, 100) & l(CASAttendDoctor3, 100)
    DataString = DataString & l(CASAdmissionReason, 500)
    DataString = DataString & l(ICDCode1, 10) & l(ICDCode2, 10) & l(ICDCode3, 10)
    DataString = DataString & l(DIAFinalDiag1, 150) & l(DIAFinalDiag2, 150) & l(DIAFinalDiag3, 150)
    DataString = DataString & l(FirstICDCode1, 5) & l(FirstICDCode2, 5) & l(FirstICDCode3, 5)
    DataString = DataString & l(DIAPrelimDiag1, 150) & l(DIAPrelimDiag2, 150) & l(DIAPrelimDiag3, 150)
    DataString = DataString & l(CASREPCode1, 10) & l(CASREPCode2, 10) & l(CASREPCode3, 10)
    DataString = DataString & l(CASOPROCRequired1, 500) & l(CASOPROCRequired2, 500) & l(CASOPROCRequired3, 500)
    DataString = DataString & l(CASOPROCRequired4, 500) & l(CASOPROCRequired5, 500)
    DataString = DataString & l(CASOProcInvestigationDone1, 500) & l(CASOProcInvestigationDone2, 500) & l(CASOProcInvestigationDone3, 500)
    DataString = DataString & l(CASOMedSurMgmt1, 500) & l(CASOMedSurMgmt2, 500) & l(CASOMedSurMgmt3, 500)

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
    Call CalNoofCase
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
    Call CalNoofCase
End Sub

Private Sub CmdClear_Click()
    CboPayor.Text = ""
    lblFileName = ""
    lblNoOfRec = ""
    lblNoOfRecProcess = ""
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
        If lblNoOfRec = 0 Then
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
Dim RecProc As Integer
Dim Trailer As String
    RecProc = 0
    SelCriteria = tmpPay
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "DataTransferSelectCASE"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("tmpPayor", adVarChar, adParamInput, 2000, SelCriteria)
    cmd.Parameters.Append Prm1

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    Do Until TmpRec.EOF
        With TmpRec
            RecProc = RecProc + 1
            MBMNumber = IIf(Len(TmpRec!MBMNumber), TmpRec!MBMNumber, "")
            CASNumber = IIf(Len(TmpRec!CASNumber), TmpRec!CASNumber, "")
            MBMPatientCovID = IIf(Len(TmpRec!MBMPatientCovID), TmpRec!MBMPatientCovID, "")
            CASStatus = IIf(Len(TmpRec!CASStatus), TmpRec!CASStatus, "")
            CASDischargeDate = IIf(Len(TmpRec!CASDischargeDate), TmpRec!CASDischargeDate, "")
            CaseRegDate = IIf(Len(TmpRec!CaseRegDate), TmpRec!CaseRegDate, "")
            CASDateAdmitted = IIf(Len(TmpRec!CASDateAdmitted), TmpRec!CASDateAdmitted, "")
            CASHOSCode = IIf(Len(TmpRec!CASHOSCode), TmpRec!CASHOSCode, "")
            ModalityCode = IIf(Len(TmpRec!ModalityCode), TmpRec!ModalityCode, "")
            PayTypeCode = IIf(Len(TmpRec!PayTypeCode), TmpRec!PayTypeCode, "")
            ClaimNatureCode = IIf(Len(TmpRec!ClaimNatureCode), TmpRec!ClaimNatureCode, "")
            CASClaimability = IIf(Len(TmpRec!CASClaimability), TmpRec!CASClaimability, "")
            CASStayType = IIf(Len(TmpRec!CASStayType), TmpRec!CASStayType, "")
            CASPAYType = IIf(Len(TmpRec!CASPAYType), TmpRec!CASPAYType, "")
            CASPending = IIf(Len(TmpRec!CASPending), TmpRec!CASPending, "")
            CASORemarks = IIf(Len(TmpRec!CASORemarks), TmpRec!CASORemarks, "")
            CASORemarks = Replace(CASORemarks, Chr(13), " ")
            CASAttendDoctor1 = IIf(Len(TmpRec!CASAttendDoctor1), TmpRec!CASAttendDoctor1, "")
            CASAttendDoctor2 = IIf(Len(TmpRec!CASAttendDoctor2), TmpRec!CASAttendDoctor2, "")
            CASAttendDoctor3 = IIf(Len(TmpRec!CASAttendDoctor3), TmpRec!CASAttendDoctor3, "")
            CASAdmissionReason = IIf(Len(TmpRec!CASAdmissionReason), TmpRec!CASAdmissionReason, "")
            CASAdmissionReason = Replace(CASAdmissionReason, Chr(13), " ")
            ICDCode1 = IIf(Len(TmpRec!ICDCode1), TmpRec!ICDCode1, "")
            ICDCode2 = IIf(Len(TmpRec!ICDCode2), TmpRec!ICDCode2, "")
            ICDCode3 = IIf(Len(TmpRec!ICDCode3), TmpRec!ICDCode3, "")
            DIAFinalDiag1 = IIf(Len(TmpRec!DIAFinalDiag1), TmpRec!DIAFinalDiag1, "")
            DIAFinalDiag2 = IIf(Len(TmpRec!DIAFinalDiag2), TmpRec!DIAFinalDiag2, "")
            DIAFinalDiag3 = IIf(Len(TmpRec!DIAFinalDiag3), TmpRec!DIAFinalDiag3, "")
            FirstICDCode1 = IIf(Len(TmpRec!FirstICDCode1), TmpRec!FirstICDCode1, "")
            FirstICDCode2 = IIf(Len(TmpRec!FirstICDCode2), TmpRec!FirstICDCode2, "")
            FirstICDCode3 = IIf(Len(TmpRec!FirstICDCode3), TmpRec!FirstICDCode3, "")
            DIAPrelimDiag1 = IIf(Len(TmpRec!DIAPrelimDiag1), TmpRec!DIAPrelimDiag1, "")
            DIAPrelimDiag2 = IIf(Len(TmpRec!DIAPrelimDiag2), TmpRec!DIAPrelimDiag2, "")
            DIAPrelimDiag3 = IIf(Len(TmpRec!DIAPrelimDiag3), TmpRec!DIAPrelimDiag3, "")
            CASREPCode1 = IIf(Len(TmpRec!CASREPCode1), TmpRec!CASREPCode1, "")
            CASREPCode2 = IIf(Len(TmpRec!CASREPCode2), TmpRec!CASREPCode2, "")
            CASREPCode3 = IIf(Len(TmpRec!CASREPCode3), TmpRec!CASREPCode3, "")
            CASOPROCRequired1 = IIf(Len(TmpRec!CASOPROCRequired1), TmpRec!CASOPROCRequired1, "")
            CASOPROCRequired2 = IIf(Len(TmpRec!CASOPROCRequired2), TmpRec!CASOPROCRequired2, "")
            CASOPROCRequired3 = IIf(Len(TmpRec!CASOPROCRequired3), TmpRec!CASOPROCRequired3, "")
            CASOPROCRequired4 = IIf(Len(TmpRec!CASOPROCRequired4), TmpRec!CASOPROCRequired4, "")
            CASOPROCRequired5 = IIf(Len(TmpRec!CASOPROCRequired5), TmpRec!CASOPROCRequired5, "")
            CASOProcInvestigationDone1 = IIf(Len(TmpRec!CASOProcInvestigationDone1), TmpRec!CASOProcInvestigationDone1, "")
            CASOProcInvestigationDone2 = IIf(Len(TmpRec!CASOProcInvestigationDone2), TmpRec!CASOProcInvestigationDone2, "")
            CASOProcInvestigationDone3 = IIf(Len(TmpRec!CASOProcInvestigationDone3), TmpRec!CASOProcInvestigationDone3, "")
            CASOMedSurMgmt1 = IIf(Len(TmpRec!CASOMedSurMgmt1), TmpRec!CASOMedSurMgmt1, "")
            CASOMedSurMgmt2 = IIf(Len(TmpRec!CASOMedSurMgmt2), TmpRec!CASOMedSurMgmt2, "")
            CASOMedSurMgmt3 = IIf(Len(TmpRec!CASOMedSurMgmt3), TmpRec!CASOMedSurMgmt3, "")
            Call WriteRecord
            Print #1, DataString
            lblNoOfRecProcess = RecProc
            TmpRec.MoveNext
        End With
    Loop
    Trailer = "T" & l(CVar(RecProc), 10)
    Print #1, Trailer
    TmpRec.Close
    Set TmpRec = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing

End Sub

Private Sub CalNoofCase()
Dim CAsREc As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim tmpTotCas As Integer
    tmpFileName = ""
    tmpPay = Mid(CboPayor.Text, 1, 2)
    tmpTotCas = 0
    CmdGenerate.Enabled = False
    If Trim(tmpPay) <> "" Then
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Set cmd1.ActiveConnection = medixmasterconnMaint
        cmd1.CommandText = "CntHisCasebyPayor"
        cmd1.CommandType = adCmdStoredProc
        cmd1.CommandTimeout = 300
        Set Prm1 = cmd1.CreateParameter("tmpPayor", adVarChar, adParamInput, 20, Trim(tmpPay))
        cmd1.Parameters.Append Prm1
        CAsREc.CursorLocation = adUseClient
        CAsREc.CursorType = adOpenForwardOnly
        CAsREc.CacheSize = 100
        Set CAsREc = cmd1.Execute
        Do While Not CAsREc.EOF
            tmpTotCas = CAsREc!HISCAseCnt
            CAsREc.MoveNext
        Loop
        CAsREc.Close
        Set CAsREc = Nothing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
        lblNoOfRec = tmpTotCas
        If lblNoOfRec > 0 Then
            CmdGenerate.Enabled = True
            tmpFileName = tmpPay & Format(Date, "YYMMDD") & "CAS.txt"
            lblFileName = tmpFileName
        Else
            'MsgBox "No Record Selected !! ", vbOKOnly + vbCritical
        End If
    End If
End Sub
