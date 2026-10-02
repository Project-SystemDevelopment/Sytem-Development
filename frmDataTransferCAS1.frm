VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataTransferCAS1 
   Caption         =   "Data Transfer - CASE"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstFileName 
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3836
      View            =   3
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "No. of Rec"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "FileName1"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "FileName2"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   6015
      Begin VB.ComboBox CboPayorGrp 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   4695
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
         Top             =   525
         Width           =   735
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
         TabIndex        =   9
         Top             =   220
         Width           =   855
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
         Left            =   1200
         TabIndex        =   0
         Top             =   195
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   6015
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
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
         Left            =   3600
         TabIndex        =   4
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
         Left            =   2760
         TabIndex        =   3
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
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   200
         Width           =   735
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
      TabIndex        =   10
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmDataTransferCAS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataString1 As String
Dim tmpPathName As String
Dim MBMNumber As String
Dim CasNumber As String
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
Dim ClaimNo As String
Dim ClaimVersion As String
Dim ActualAmt As String
Dim ApprovedAmt As String
Dim BordxAmt As String
Dim ExcessAmt As String
Dim BTBGAmt As String
Dim PayArr() As Variant

Private Sub GenerateAscii()
Dim Filename, Header As String
Dim CntPay As Integer
Dim tmpPay As String
Dim SelCriteria As String
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim RecProc As String
Dim Trailer As String
Dim tmpFileName1 As String
Dim tmpFileName2 As String
Dim DataString2 As String
Dim FileNameList As ListItem
Dim CntNoOFfile As Integer
Dim FirstRec As Boolean

    CntNoOFfile = 0
    'medixneptuneconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"
   ' medixneptuneconnWS.Open "File Name=" & BOSPath & "Conn\medineptuneconnectWS.UDL"
    For CntPay = 0 To UBound(PayArr)
        RecProc = 0
        tmpPay = PayArr(CntPay)
        FirstRec = False
        SelCriteria = tmpPay
        tmpFileName1 = tmpPay & Format(Date, "YYMMDD") & "CAS1.txt"
        tmpFileName2 = tmpPay & Format(Date, "YYMMDD") & "CAS2.txt"
        
        Set cmd = Nothing
        Set Prm1 = Nothing
        Set cmd.ActiveConnection = medixneptuneconn
        cmd.CommandText = "DataTransferSelectCASENew"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 1000
        Set Prm1 = cmd.CreateParameter("tmpPayor", adVarChar, adParamInput, 2000, SelCriteria)
        cmd.Parameters.Append Prm1
    
        TmpRec.CursorLocation = adUseClient
        TmpRec.CursorType = adOpenForwardOnly
        TmpRec.CacheSize = 1000
        Set TmpRec = cmd.Execute
        
        Do While Not TmpRec.EOF
            
            RecProc = RecProc + 1
            If FirstRec = False Then
                CntNoOFfile = CntNoOFfile + 1
                Header = "F" & "CAS" & tmpPay & Format(Date, "ddmmyyyy")
                Open tmpPathName & tmpFileName1 For Output As #1
                Open tmpPathName & tmpFileName2 For Output As #2
                Print #1, Trim(Header)
                Print #2, Trim(Header)
                FirstRec = True
            End If
            With TmpRec
                MBMNumber = IIf(Len(TmpRec!MBMNumber), TmpRec!MBMNumber, "")
                CasNumber = IIf(Len(TmpRec!CasNumber), TmpRec!CasNumber, "")
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
                DataString2 = ""
                Call GetRemarks(CasNumber, DataString2)
                Print #1, DataString1
            End With
            TmpRec.MoveNext
        Loop
        If RecProc > 0 Then
            Trailer = "T" & l(CVar(RecProc), 10)
            Print #1, Trailer
            Print #2, Trailer
            Close #1
            Close #2
        End If
        TmpRec.Close
        Set TmpRec = Nothing
        If FirstRec = True Then
            Set FileNameList = lstFileName.ListItems.Add(, , CntNoOFfile)
            FileNameList.SubItems(1) = tmpPay
            FileNameList.SubItems(2) = RecProc
            FileNameList.SubItems(3) = tmpFileName1
            FileNameList.SubItems(4) = tmpFileName2
        End If
    Next CntPay
    Call GenerateWS
    'medixneptuneconn.Close
   ' Set medixneptuneconn = Nothing
   ' medixneptuneconnWS.Close
   ' Set medixneptuneconnWS = Nothing
    If CntNoOFfile = 0 Then
        MsgBox "No Record Selected ! ", vbOKOnly + vbInformation, DialogBoxTitle
    Else
        MsgBox "File Generated  Successfully...", vbOKOnly + vbInformation, DialogBoxTitle
    End If
End Sub


Private Sub WriteRecord()
    CASDischargeDate = Format(CASDischargeDate, "ddmmyyyy")
    CaseRegDate = Format(CaseRegDate, "ddmmyyyy")
    CASDateAdmitted = Format(CASDateAdmitted, "ddmmyyyy")

    DataString1 = ""
    DataString1 = DataString1 & l(CasNumber, 15) & l(MBMNumber, 15) & l(MBMPatientCovID, 2)
    DataString1 = DataString1 & l(CASStatus, 1) & l(CASDischargeDate, 8) & l(CaseRegDate, 8)
    DataString1 = DataString1 & l(CASDateAdmitted, 8) & l(CASHOSCode, 10) & l(ModalityCode, 2)
    DataString1 = DataString1 & l(PayTypeCode, 2) & l(ClaimNatureCode, 5) & l(CASClaimability, 1)
    DataString1 = DataString1 & l(CASStayType, 10) & l(CASPAYType, 10) & l(CASPending, 1)
    'DataString1 = DataString1 & l(CASORemarks, 3900)
    DataString1 = DataString1 & l(CASAttendDoctor1, 100) & l(CASAttendDoctor2, 100) & l(CASAttendDoctor3, 100)
    DataString1 = DataString1 & l(CASAdmissionReason, 500)
    DataString1 = DataString1 & l(ICDCode1, 10) & l(ICDCode2, 10) & l(ICDCode3, 10)
    DataString1 = DataString1 & l(DIAFinalDiag1, 150) & l(DIAFinalDiag2, 150) & l(DIAFinalDiag3, 150)
    DataString1 = DataString1 & l(FirstICDCode1, 5) & l(FirstICDCode2, 5) & l(FirstICDCode3, 5)
    DataString1 = DataString1 & l(DIAPrelimDiag1, 150) & l(DIAPrelimDiag2, 150) & l(DIAPrelimDiag3, 150)
    DataString1 = DataString1 & l(CASREPCode1, 10) & l(CASREPCode2, 10) & l(CASREPCode3, 10)
    DataString1 = DataString1 & l(CASOPROCRequired1, 500) & l(CASOPROCRequired2, 500) & l(CASOPROCRequired3, 500)
    DataString1 = DataString1 & l(CASOPROCRequired4, 500) & l(CASOPROCRequired5, 500)
    DataString1 = DataString1 & l(CASOProcInvestigationDone1, 500) & l(CASOProcInvestigationDone2, 500) & l(CASOProcInvestigationDone3, 500)
    DataString1 = DataString1 & l(CASOMedSurMgmt1, 500) & l(CASOMedSurMgmt2, 500) & l(CASOMedSurMgmt3, 500)

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


Private Sub CboPayorGrp_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayorGrp.Text, CboPayorGrp.SelStart) + Chr(KeyAscii) + Mid(CboPayorGrp.Text, CboPayorGrp.SelStart + CboPayorGrp.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayorGrp.ListCount - 1
 
        If NewText = LCase(Left(CboPayorGrp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayorGrp.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboPayorGrp.Text = ValidValue
            CboPayorGrp.SelStart = 0
            CboPayorGrp.SelLength = Len(ValidValue)
        End If
    End If
End Sub


Private Sub CboPayorGrp_LostFocus()
Dim SelFields As String
Dim SelCriteria As String
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim PayCnt As Integer
Dim tmpPayGrp As String
    
    PayCnt = 0
    SelFields = ""
    tmpPayGrp = Mid(CboPayorGrp.Text, 1, 3)

    SelCriteria = "  PAYGRPCode='" & tmpPayGrp & "'"

    SelFields = "SELECT PAYCode as tmpCode, PAYCompanyName as tmpName "
    
'    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"

    Set cmd.ActiveConnection = medixneptuneconn
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
    Do While Not TmpRec.EOF
        With TmpRec
            ReDim Preserve PayArr(PayCnt)
            PayArr(PayCnt) = Trim(!tmpCode)
            TmpRec.MoveNext
            PayCnt = PayCnt + 1
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
  '  medixneptuneconn.Close
  '  Set medixneptuneconn = Nothing
    CmdGenerate.Enabled = True
End Sub

Private Sub CmdClear_Click()
    CboPayorGrp.Text = ""
    lstFileName.ListItems.Clear
    CmdGenerate.Enabled = False
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGenerate_Click()
Dim X As Boolean
    
    X = CreatePath(tmpPathName)

    If X = False Then
        MsgBox "Error Creating Folder " & tmpPathName & "  !" & vbNewLine & _
                " Please create that folder manually to store the output of leaflet in sofe copy!"
    Else
        Screen.MousePointer = vbHourglass
        Call GenerateAscii
        'Call Form_KeyDown(117, 0)
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Function CreatePath(NewPath) As Boolean
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        CreatePath = True
    End If
End Function

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    
    If lstFileName.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstFileName)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = Default

End Sub

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
    SelCriteria = "0=0"

    SelFields = "SELECT PAYGRPCode as tmpCode, PAYGRPName as tmpName "
    
'    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"

    Set cmd.ActiveConnection = medixneptuneconn
    cmd.CommandText = "SearchHISTableRec"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, SelFields)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("TableName", adVarChar, adParamInput, 255, "PAYGRP")
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriteria", adVarChar, adParamInput, 2000, SelCriteria)
    cmd.Parameters.Append Prm3

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd.Execute
    CboPayorGrp.Clear
    Do Until TmpRec.EOF
        With TmpRec
            CboPayorGrp.AddItem !tmpCode & " - " & !tmpName
            TmpRec.MoveNext
        End With
    Loop
    
    TmpRec.Close
    Set TmpRec = Nothing
  '  medixneptuneconn.Close
  '  Set medixneptuneconn = Nothing

End Sub

Private Sub GetRemarks(tmpCasNumber As String, DataString2 As String)
Dim CASORemarks1 As String
Dim CASORemarks2 As String
Dim CASORemarks3 As String
Dim CASORemarks4 As String
Dim CASRemarks3 As String
Dim CASRemarks4 As String
Dim CASRemarks5 As String
Dim TmpRec As New ADODB.Recordset
Dim cmd2 As New ADODB.Command
Dim Prm2 As New ADODB.Parameter
    
    Set cmd2 = Nothing
    Set Prm2 = Nothing
    Set cmd2.ActiveConnection = medixneptuneconn
    cmd2.CommandText = "SP_CASRemarks"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 300
    Set Prm2 = cmd2.CreateParameter("CasNumber", adVarChar, adParamInput, 15, tmpCasNumber)
    cmd2.Parameters.Append Prm2

    TmpRec.CursorLocation = adUseClient
    TmpRec.CursorType = adOpenForwardOnly
    TmpRec.CacheSize = 100
    Set TmpRec = cmd2.Execute
                
    Do While Not TmpRec.EOF
        CASORemarks1 = IIf(Len(TmpRec!CASORemarks1), TmpRec!CASORemarks1, "")
        CASORemarks2 = IIf(Len(TmpRec!CASORemarks2), TmpRec!CASORemarks2, "")
        CASORemarks3 = IIf(Len(TmpRec!CASORemarks3), TmpRec!CASORemarks3, "")
        CASORemarks4 = IIf(Len(TmpRec!CASORemarks4), TmpRec!CASORemarks4, "")
        
        CASRemarks3 = IIf(Len(TmpRec!CASRemarks3), TmpRec!CASRemarks3, "")
        CASRemarks4 = IIf(Len(TmpRec!CASRemarks4), TmpRec!CASRemarks4, "")
        CASRemarks5 = IIf(Len(TmpRec!CASRemarks5), TmpRec!CASRemarks5, "")
        
        DataString2 = DataString2 & l(TmpRec!CasNumber, 15)
        DataString2 = DataString2 & l(CASORemarks1, 2000)
        DataString2 = DataString2 & l(CASORemarks2, 2000)
        DataString2 = DataString2 & l(CASORemarks3, 2000)
        DataString2 = DataString2 & l(CASORemarks4, 2000)
        DataString2 = DataString2 & l(CASRemarks3, 3900)
        DataString2 = DataString2 & l(CASRemarks4, 3900)
        DataString2 = DataString2 & l(CASRemarks5, 3900)
        Print #2, DataString2
        TmpRec.MoveNext
    Loop
    TmpRec.Close
    Set TmpRec = Nothing

End Sub


Private Sub GenerateWS()
Dim Filename, Header As String
Dim CntPay As Integer
Dim tmpPay As String
Dim SelCriteria As String
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim RecProc As String
Dim Trailer As String
Dim tmpFileName1 As String
Dim tmpFileName2 As String
Dim DataString2 As String
Dim FileNameList As ListItem
Dim CntNoOFfile As Integer
Dim FirstRec As Boolean

    CntNoOFfile = 0
    For CntPay = 0 To UBound(PayArr)
        RecProc = 0
        tmpPay = PayArr(CntPay)
        FirstRec = False
        SelCriteria = tmpPay
        tmpFileName1 = tmpPay & Format(Date, "YYMMDD") & "WS.txt"
        
        Set cmd = Nothing
        Set Prm1 = Nothing
        Set cmd.ActiveConnection = medixneptuneconnWS
        cmd.CommandText = "DataTransferSelectWS"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 1000
        Set Prm1 = cmd.CreateParameter("tmpPayor", adVarChar, adParamInput, 2000, SelCriteria)
        cmd.Parameters.Append Prm1
    
        TmpRec.CursorLocation = adUseClient
        TmpRec.CursorType = adOpenForwardOnly
        TmpRec.CacheSize = 1000
        Set TmpRec = cmd.Execute
        
        Do While Not TmpRec.EOF
            
            RecProc = RecProc + 1
            If FirstRec = False Then
                CntNoOFfile = CntNoOFfile + 1
                Header = "F" & "WS" & tmpPay & Format(Date, "ddmmyyyy")
                Open tmpPathName & tmpFileName1 For Output As #1
                Print #1, Trim(Header)
                FirstRec = True
            End If
            With TmpRec
                ClaimNo = IIf(Len(!CasNumber), !CasNumber, "")
                ClaimVersion = IIf(Len(!ClaimVersion), !ClaimVersion, "")
                ActualAmt = IIf(IsNumeric(!CLMTotalActual), !CLMTotalActual, "0")
                ApprovedAmt = IIf(IsNumeric(!CLMTotalApproved), !CLMTotalApproved, "0")
                ExcessAmt = IIf(IsNumeric(!CLMExcess), !CLMExcess, "0")
                BTBGAmt = IIf(IsNumeric(!CLMBTBGAmt), !CLMBTBGAmt, "0")
                BordxAmt = IIf(IsNumeric(!CLMBordxAmt), !CLMBordxAmt, "0")
                
                Call WriteWSRecord
                Print #1, DataString1
            End With
            TmpRec.MoveNext
        Loop
        If RecProc > 0 Then
            Trailer = "T" & l(CVar(RecProc), 10)
            Print #1, Trailer
            Close #1
        End If
        TmpRec.Close
        Set TmpRec = Nothing
        If FirstRec = True Then
            Set FileNameList = lstFileName.ListItems.Add(, , CntNoOFfile)
            FileNameList.SubItems(1) = tmpPay
            FileNameList.SubItems(2) = RecProc
            FileNameList.SubItems(3) = tmpFileName1
        End If
    Next CntPay
End Sub

Private Sub WriteWSRecord()

    DataString1 = ""
    DataString1 = DataString1 & l(ClaimNo, 15)
    DataString1 = DataString1 & l(ClaimVersion, 2)
    DataString1 = DataString1 & l(ActualAmt, 12)
    DataString1 = DataString1 & l(ApprovedAmt, 12)
    DataString1 = DataString1 & l(ExcessAmt, 12)
    DataString1 = DataString1 & l(BTBGAmt, 12)
    DataString1 = DataString1 & l(BordxAmt, 12)
End Sub

