VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintMBMBordxBatch 
   Caption         =   "FrmLeaflet-Batch"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   12150
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   11295
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   300
         Left            =   7800
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   8640
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9480
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10320
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3600
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "Total Supp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "Total Principal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Select Excel"
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   11295
      Begin VB.ComboBox cboPrintMode 
         Height          =   315
         ItemData        =   "frmPrintMBMBordxBatch.frx":0000
         Left            =   8520
         List            =   "frmPrintMBMBordxBatch.frx":0002
         TabIndex        =   16
         Top             =   360
         Width           =   1815
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
         Left            =   1515
         TabIndex        =   14
         Top             =   360
         Width           =   3825
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
         Left            =   5640
         TabIndex        =   13
         Top             =   360
         Width           =   320
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl14 
         Caption         =   "Printing Mode"
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Filename"
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
         Left            =   240
         TabIndex        =   15
         Top             =   435
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstMemberListing 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IC/No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Plan No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Effective Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Expiry Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Group Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Department"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Employee No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Branch"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Subsidary"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Membership Reporting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Member Listing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   1200
   End
End
Attribute VB_Name = "frmPrintMBMBordxBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As New excel.Application
Dim xlsheet As excel.WORKSHEET
Dim xlwbook As excel.Workbook
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim tmpMemberNo As String
Dim Val1 As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFileLocation_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
  CommonDialog1.InitDir = "\\svr-doc\" & "Bor_In"
  CommonDialog1.Flags = cdlOFNHideReadOnly
  CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Dat Files (*.Dat)|*.Dat"
  CommonDialog1.ShowOpen
  txtImportFilename = CommonDialog1.FileTitle
  Exit Sub
ErrHandler:
  Exit Sub
End Sub

Private Sub cmdPrint_Click()
Dim FileExists As Integer
Dim lSize As Long
Dim VPolicyNo As String
Dim VMBMNumber As String

      If txtImportFilename <> "" Then
        FileExists = 0
        lSize = -1
        lSize = FileLen(CommonDialog1.FileName)
        If lSize <= 0 Then
            FileExists = 0
        ElseIf lSize > 0 Then
            FileExists = 1
        End If
        
            If FileExists = 1 Then
                Set xlwbook = xl.Workbooks.Open(CommonDialog1.FileName)
                Set xlsheet = xlwbook.Sheets.Item(1)
                Val1 = 2
                Do Until Val1 = 60000
                    If xlsheet.Cells(Val1, 1) <> "" Then
                         VPolicyNo = xlsheet.Cells(Val1, 3)
                         VMBMNumber = xlsheet.Cells(Val1, 4)
                           Call CallLEafLet(VMBMNumber, VPolicyNo)
                    Else
                     Exit Do
                    End If
                    Val1 = CDbl(Val1) + 1
                Loop
            End If
            
      End If
End Sub

Public Sub CallLEafLet(MBMNumber As String, PolcyNo As String)
Dim tempPath As String
        tempPath = ""
        tempPath = "\\PLUTO\HISSaturn\Conn\mediconnectPluto.UDL"
        medixmasterconnPluto.Open "File Name=" & tempPath
        Call LeafLetPrint(MBMNumber, PolcyNo)
        medixmasterconnPluto.Close
        Set medixmasterconnPluto = Nothing
End Sub
Private Sub LeafLetPrint(MBMNumber As String, MBMPolNumber As String)
'On Error Resume Next
Dim strsql As String
Dim MBM As New ADODB.Recordset
Dim ii As Integer
Dim MemberNo As String
Dim tmpPrintedBDate, tmpPrintedBatchNo, tmpPrintedPaycode, tempPrintedHLTcode, tmpPrintedMemberNo As String
Dim cnt As Integer
Dim PrintM As String

 bPrintPass = False
 CancelPrinting = False
 Screen.MousePointer = vbHourglass
 'Frame1.Enabled = False
    cnt = 0 ' Sihat Malaysia
    'If PrintMode = "ALL" Then
    
    Dim MBMNum() As String
    MBMNum = Split(MBMNumber, "-")
        If Len(MBMNum(0)) = "8" Then
            strsql = "Select top 1 MBMNumber from MBMCrossReference where substring(MBMNumber,1,8) = '" & Trim(MBMNum(0)) & "' And MBMPolicyNo='" & Trim(MBMPolNumber) & "' order by MBMBordxDate  desc"
            MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        ElseIf Len(MBMNum(0)) = "8" Then
            strsql = "Select top 1 MBMNumber from MBMCrossReference where substring(MBMNumber,1,9) = '" & Trim(MBMNum(0)) & "' And MBMPolicyNo='" & Trim(MBMPolNumber) & "' order by MBMBordxDate  desc"
            MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        Else
            strsql = "Select top 1  MBMNumber from MBMCrossReference where MBMNumber like  '%" & Trim(MBMNum(0)) & "%'And MBMPolicyNo='" & Trim(MBMPolNumber) & "' order by MBMBordxDate  desc"
            MBM.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        End If
        
    If MBM.recordCount = 0 Then
        MsgBox "Record not found !", vbInformation + vbOKOnly
        Screen.MousePointer = Default
        Exit Sub
    End If
    
        Do Until MBM.EOF
            tmpMemberNo = MBM!MBMNumber
            Call InsertLealet(tmpMemberNo)
            Call MakeWordDoc(tmpMemberNo)
        MBM.MoveNext
        Loop
    MBM.Close
    Set MBM = Nothing
    
    Screen.MousePointer = Default
End Sub
Private Sub InsertLealet(MBMNumber As String)
Dim strsqlPlutoBor As String
Dim PlutoBor As New ADODB.Recordset
Dim sqlbor As String
Dim BOR As New ADODB.Recordset
Dim sqlinsert As String
Dim PLUTOStatus As Boolean
Dim tmpMBMName As String
Dim tmpMBMAddress1 As String
Dim tmpMBMAddress2 As String
Dim tmpMBMAddress3 As String
Dim tmpMBMCName As String
Dim tmpMBMGroupCompany As String
Dim tmpPrevComp As String
Dim tmpPayorRemarks As String
Dim tmpAllergy As String
Dim tmpPLNRBType As String
Dim tmpPLNDescription As String
Dim tmpSuppLimitStatus As String
Dim tmpMBMCAvailableLimit As Double
Dim tmpMBMAvailableLimit As Double
Dim tmpDepartment As String
Dim tmpBRHCode As String
Dim tmpOccupation As String
Dim tmpMBMCoverID As String

tmpMemberNo = MBMNumber

    PLUTOStatus = False
    strsqlPlutoBor = ""
    strsqlPlutoBor = "Select MBMNumber from Tmp_Leaflet where MBMNumber = '" & Trim(tmpMemberNo) & "'"
    PlutoBor.Open strsqlPlutoBor, medixmasterconnPluto, adOpenForwardOnly, adLockReadOnly
    Do While Not PlutoBor.EOF
        PLUTOStatus = True
        PlutoBor.MoveNext
    Loop
    PlutoBor.Close
    Set PlutoBor = Nothing
    If PLUTOStatus = False Then
            sqlbor = "Select * from VIEW_Leaflet where MBMNumber = '" & Trim(tmpMemberNo) & "'"
            BOR.Open sqlbor, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not BOR.EOF
            TempMBMName = Replace(Trim(BOR!MBMName), "'", "''")
            tmpMBMAddress1 = Replace(Trim(BOR!MBMAddress1), "'", "''")
            tmpMBMAddress2 = Replace(Trim(BOR!MBMAddress2), "'", "''")
            If Len(BOR!MBMAddress3) Then
                tmpMBMAddress3 = Replace(Trim(BOR!MBMAddress3), "'", "''")
            End If
            If Len(BOR!MBMCName) Then
                tmpMBMCName = Replace(Trim(BOR!MBMCName), "'", "''")
            End If
            If Len(BOR!MBMGroupCompany) Then
                tmpMBMGroupCompany = Replace(Trim(BOR!MBMGroupCompany), "'", "''")
            End If
            If Len(BOR!MBMPrevInsCom) Then
                tmpPrevComp = Replace(Trim(BOR!MBMPrevInsCom), "'", "''")
            Else
                tmpPrevComp = ""
            End If
            If Len(BOR!MBMPAYRemarks) Then
                tmpPayorRemarks = Replace(Trim(BOR!MBMPAYRemarks), "'", "''")
            Else
                tmpPayorRemarks = ""
            End If
            If Len(BOR!MBMAllergic) Then
                tmpAllergy = Replace(Trim(BOR!MBMAllergic), "'", "''")
            End If
            If Len(BOR!MBMDepartment) Then
                tmpDepartment = Replace(Trim(BOR!MBMDepartment), "'", "''")
            Else
                tmpDepartment = ""
            End If
            If Len(BOR!BRNCode) Then
                tmpBRHCode = Replace(Trim(BOR!BRNCode), "'", "''")
            Else
                tmpBRHCode = ""
            End If
            If Len(BOR!MBMOccupation) Then
                tmpOccupation = Replace(Trim(BOR!MBMOccupation), "'", "''")
            Else
                tmpOccupation = ""
            End If
            tmpMBMAvailableLimit = IIf(IsNumeric(BOR!MBMAvailableLimit), BOR!MBMAvailableLimit, "0")
            tmpMBMCAvailableLimit = IIf(IsNumeric(BOR!MBMCAvailableLimit), BOR!MBMCAvailableLimit, "0")
            If BOR!MBMCRELCode <> "P" Then
                tmpMBMCoverID = BOR!MBMCCoverID
            Else
                tmpMBMCoverID = "00"
            End If
            sqlinsert = ""
            sqlinsert = "Insert into Tmp_Leaflet(MBMNumber , MBMPolicyNo ,MBMName ,MBMAddress1  ,MBMAddress2 , " & _
                        "MBMAddress3 ,MBMCName ,PAYCompanyName,MBMAgentCode,HLTCode, " & _
                        "MBMEmployeeNo,MBMDepartment,MBMPostCode,STACode,CTYCode,MBMDataStatus, " & _
                        "MBMStatus,BRNCode ,INSCode,MBMOccupation,MBMPrevInsCom,MBMPayRemarks,MBMAllergic, " & _
                        "PLNRBType, PLNDescription, SuppLimitStatus,MBMAvailableLimit,MBMCAvailableLimit,CoverID,MBMIcBcPp  ) " & _
                        "Values('" & BOR!MBMNumber & "','" & BOR!MBMPolicyNo & "', '" & TempMBMName & "',  '" & tmpMBMAddress1 & "',  '" & tmpMBMAddress2 & "', " & _
                        "'" & tmpMBMAddress3 & "', '" & tmpMBMCName & "', '" & tmpMBMGroupCompany & "', '" & BOR!MBMAgentCode & "', '" & BOR!HLTCode & "', " & _
                        "'" & BOR!MBMEmployeeNo & "', '" & tmpDepartment & "', '" & BOR!MBMPostCode & "', '" & BOR!STACode & "', '" & BOR!CTYCode & "',  '" & BOR!MBMDataStatus & "', " & _
                        "'" & BOR!MBMStatus & "', '" & tmpBRHCode & "', '" & BOR!INSCode & "', '" & tmpOccupation & "','" & tmpPrevComp & "', '" & tmpPayorRemarks & "', '" & tmpAllergy & "', " & _
                        "'" & BOR!PLNRBType & "', '" & BOR!PLNDescription & "', '" & BOR!SuppLimitStatus & "'," & tmpMBMAvailableLimit & ", " & tmpMBMCAvailableLimit & ", '" & tmpMBMCoverID & "', '" & BOR!MBMIcBcPp & "') "
            medixmasterconnPluto.Execute sqlinsert
        BOR.MoveNext
        Loop
    BOR.Close
    Set BOR = Nothing
    End If
End Sub
'New sub for Print
Private Sub MakeWordDoc(MBMNumber As String)
'Get member details
        Dim PRSet As New ADODB.Recordset
        'If OPTNorLeaflet.Value = True Then
            sqlbor = "Select * from VIEW_Leaflet where MBMNumber = '" & Trim(MBMNumber) & "'"
            PRSet.Open sqlbor, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        'ElseIf OPTSpeLeaflet.Value = True Then
        '    sqlbor = "Select * from View_Leaflet_Clinical where MBMNumber = '" & Trim(MBMNumber) & "'"
        '    PRSet.Open sqlbor, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        'End If
        'Template Path
        Dim fname As String
         If Mid(cboPrintMode, 1, 2) = "RI" Then
            fname = crdPath & "Leflet\RITemplate.dot"
         Else
            fname = crdPath & "Leflet\PrintTemplate.dot"
         End If
        
        Dim word_app As Word.Application
        Dim word_doc As Word.Document
        Dim objSelection
        Dim SuppCount As Integer
        Dim StrArray(20) As String
        SuppCount = 0
        Dim isFirst As Boolean
        isFirst = True
        Dim MBMNo As String
        Do While Not PRSet.EOF
         MBMNo = PRSet!MBMNumber
        If InStr(1, MBMNo, "*") Then
            Dim sArr() As String
            sArr = Split(MBMNo, "*")
            MBMNo = sArr(0)
        End If
            If isFirst Then
            StrArray(6) = PRSet!MBMPolicyNo & " " & IIf(Len(PRSet!MBMAgentCode), (", " & PRSet!MBMAgentCode), "")
            StrArray(7) = PRSet!MBMName
            StrArray(8) = PRSet!MBMAddress1
            StrArray(9) = PRSet!MBMAddress2
            StrArray(10) = IIf(Len(PRSet!MBMAddress3), (PRSet!MBMAddress3 & ", "), "") & PRSet!MBMPostCode & ", " & PRSet!CTYCode
            StrArray(11) = PRSet!STACode
            StrArray(12) = "MembershipNo: " & MBMNo
            StrArray(13) = PRSet!BRNCode
            isFirst = False
            End If
            
            If SuppCount < 6 Then
                StrArray(SuppCount) = StrArray(SuppCount) + "                                                                                             " & PRSet!MBMCName
            ElseIf SuppCount = 6 And SuppCount < 14 Then
                StrArray(SuppCount) = StrArray(SuppCount) + "                                                                        " & PRSet!MBMCName
            Else
                StrArray(SuppCount) = StrArray(SuppCount) + "                                                                                            " & PRSet!MBMCName
            End If
            
            SuppCount = SuppCount + 1
            PRSet.MoveNext
        Loop
        
        Dim arrcount As Integer
        arrcount = 0
        Dim strText As String
        For arrcount = 0 To 20
                strText = strText & StrArray(arrcount) & Chr(11)
        Next
            Set word_app = New Word.Application
            Set word_doc = word_app.Documents.Open(fname)
            word_app.Selection.TypeParagraph
            word_app.Selection.TypeParagraph
            word_app.Selection.TypeText strText
            Set objSelection = word_app.Selection
            objSelection.WholeStory
            objSelection.Font.Name = "Times New Roman"
            objSelection.Font.Size = 10
            objSelection.Font.Bold = False
            word_doc.PrintOut Range:=wdPrintCurrentPage
            Sleep 1000
            word_doc.Close False
            word_app.Quit False
End Sub

Private Sub Form_Load()
    cboPrintMode.AddItem "RI - Ricoh"
    cboPrintMode.AddItem "HP - Hewlett Packard"
    cboPrintMode.Text = "RI - Ricoh"
End Sub
