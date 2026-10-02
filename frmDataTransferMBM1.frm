VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataTransferMBM1 
   Caption         =   "Data Transfer - MBM"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   0
      TabIndex        =   7
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
         Width           =   975
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
         TabIndex        =   8
         Top             =   220
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   9
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
   End
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
         Text            =   "NoOfPrinc"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NoOfSupp"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "FileName2"
         Object.Width           =   3175
      EndProperty
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
      TabIndex        =   10
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmDataTransferMBM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataString As String
Dim tmpPathName As String
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
Dim PayArr() As Variant
    
Private Sub GenerateAscii()
Dim Header As String
Dim tmpPay As String
Dim CntPay As Integer
Dim SelCriteria As String
Dim TmpRec As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim RecProc As String
Dim Trailer As String
Dim FileNameList As ListItem
Dim CntNoOFfile As Integer
Dim FirstRec As Boolean
Dim RecProcP As String
Dim tmpFileName As String
    CntNoOFfile = 0
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"

    For CntPay = 0 To UBound(PayArr)
        RecProc = 0
        tmpPay = PayArr(CntPay)
        FirstRec = False
        SelCriteria = tmpPay
        tmpFileName = tmpPay & Format(Date, "YYMMDD") & "MBM.txt"
        RecProcP = 0
        RecProcS = 0
        SelCriteria = tmpPay
        
        Set cmd = Nothing
        Set Prm1 = Nothing
        Set cmd.ActiveConnection = medixneptuneconn
        cmd.CommandText = "DataTransferSelectMBM"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 1000
        Set Prm1 = cmd.CreateParameter("tmpPayor", adVarChar, adParamInput, 5, SelCriteria)
        cmd.Parameters.Append Prm1
    
        TmpRec.CursorLocation = adUseClient
        TmpRec.CursorType = adOpenForwardOnly
        TmpRec.CacheSize = 1000
        Set TmpRec = cmd.Execute
        Do While Not TmpRec.EOF
            With TmpRec
                RecProcP = RecProcP + 1
                If FirstRec = False Then
                    CntNoOFfile = CntNoOFfile + 1
                    Header = "F" & "MBM" & tmpPay & Format(Date, "ddmmyyyy")
                    Open tmpPathName & tmpFileName For Output As #1
                    Print #1, Trim(Header)
                    FirstRec = True
                End If

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
                TmpRec.MoveNext
            End With
        Loop
        If RecProcP > 0 Then
            Trailer = "T" & l(CVar(RecProcP), 10) & l(CVar(RecProcS), 10)
            Print #1, Trailer
            Close #1
        End If
        TmpRec.Close
        Set TmpRec = Nothing
        If FirstRec = True Then
            Set FileNameList = lstFileName.ListItems.Add(, , CntNoOFfile)
            FileNameList.SubItems(1) = tmpPay
            FileNameList.SubItems(2) = RecProcP
            FileNameList.SubItems(3) = RecProcS
            FileNameList.SubItems(4) = tmpFileName
        End If
    Next CntPay
  '  medixneptuneconn.Close
  '  Set medixneptuneconn = Nothing
    If CntNoOFfile = 0 Then
        MsgBox "No Record Selected ! ", vbOKOnly + vbInformation, DialogBoxTitle
    Else
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

    SelCriteria = " PAYGRPCode='" & tmpPayGrp & "'"

    SelFields = "SELECT PAYCode as tmpCode, PAYCompanyName as tmpName "
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixneptuneconn.Open "File Name=" & BOSPath & "Conn\medineptuneconnect.UDL"
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
 '   Set medixneptuneconn = Nothing
    CmdGenerate.Enabled = True
End Sub


Private Sub CmdClear_Click()
    CboPayorGrp.Text = ""
    CmdGenerate.Enabled = False
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
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
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

Private Sub AssignRecord()

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
    SuppRec.Open SuppSql, medixneptuneconn, adOpenForwardOnly, adLockReadOnly
    Do While Not SuppRec.EOF
        RecProcS = RecProcS + 1
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

