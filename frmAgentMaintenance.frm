VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAgentMaintenance 
   Caption         =   "Agent"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   10605
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   7200
      Width           =   11775
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3840
         TabIndex        =   18
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   15
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   7800
         TabIndex        =   3
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   8640
         TabIndex        =   14
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   6960
         TabIndex        =   13
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   4680
         TabIndex        =   12
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2055
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   11775
      Begin VB.ComboBox CboPAYCode 
         Height          =   315
         Left            =   6120
         TabIndex        =   2
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txtAGTDesc 
         Height          =   960
         Left            =   1920
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   600
         Width           =   9000
      End
      Begin VB.TextBox txtAGTCode 
         Height          =   285
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label41 
         Caption         =   "Payor             *"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Agent Description "
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label29 
         Caption         =   "Agent Code*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   5
         Top             =   240
         Width           =   1800
      End
   End
   Begin MSComctlLib.ListView lstAgentList 
      Height          =   4320
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7620
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Agent Code"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Agent Name"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Payor"
         Object.Width           =   14993
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "AGENT LIST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Agent Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11775
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      Caption         =   "AGENT LIST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   -120
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "frmAgentMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim EditFlag As Boolean
Public AGTCriteria As String
Private m_blnDirection(1 To 3) As Boolean


Private Sub AgentComboListing()

    Dim PAY As New ADODB.Recordset
    Dim HLT As New ADODB.Recordset
        
    PAY.Open "PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    'HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    CboPAYCode.Clear
    'cboPAYHLTCode.Clear
    
    Do Until PAY.EOF
        CboPAYCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        PAY.MoveNext
    Loop
        
    
    'Do Until HLT.EOF
    '    cboPAYHLTCode.AddItem HLT!HLTCode & " - " & HLT!HLTDescription
    '    HLT.MoveNext
    'Loop
            
    PAY.Close
    'HLT.Close
    
'    CboConversion.AddItem "N" & " - " & "No"
'    CboConversion.AddItem "Y" & " - " & "Yes"
End Sub
Private Sub cboPAYCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPAYCode.Text, CboPAYCode.SelStart) + Chr(KeyAscii) + Mid(CboPAYCode.Text, CboPAYCode.SelStart + CboPAYCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPAYCode.ListCount - 1
 
        If NewText = LCase(Left(CboPAYCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPAYCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPAYCode.Text = ValidValue
                CboPAYCode.SelStart = 0
                CboPAYCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cmdAdd_Click()
EditFlag = False
Call ClearForm(Me)
Call DisableCmdButton(Me)
Call EntriesEnable(True)
CmdSave.Enabled = True
CmdAdd.Enabled = True
txtAGTCode.SetFocus
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    txtAGTDesc.SetFocus
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
    
    If txtAGTCode <> "" Then
       txtAGTCode.Enabled = False
       txtAGTDesc.SetFocus
    End If
    
End Sub

Private Sub CmdPrint_Click()
    RptAgentList.show
End Sub

Private Sub cmdSave_Click()
'On Error Resume Next
Dim RecordSaved
Dim AGT As New ADODB.Recordset
Dim SaveAGT As New ADODB.Recordset
Dim strsql As String

    Screen.MousePointer = vbHourglass
    
    If txtAGTCode = "" Then
        MsgBox "Please Enter Agent Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtAGTCode.SetFocus
    ElseIf txtAGTDesc = "" Then
        MsgBox "Please Enter Agent Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtAGTDesc.SetFocus
    ElseIf CboPAYCode = "" Then
        MsgBox "Please Enter Payor", vbOKOnly + vbCritical, DialogBoxTitle
        CboPAYCode.SetFocus
    Else
          strsql = "Select * From Agent Where AgentCode = '" & Trim(txtAGTCode) & "'"
          
          AGT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
          
            If AGT.recordCount > 0 And EditFlag = False Then
                MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
                AGT.Close
                Set AGT = Nothing
                txtAGTCode.SetFocus
                txtAGTCode.SelStart = 0
                txtAGTCode.SelLength = Len(txtAGTCode)
                Exit Sub
            Else
            AGT.Close
            Set AGT = Nothing
      RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
         If RecordSaved = vbYes Then
            If EditFlag = False Then
                AGT.Open "Agent", medixmasterconn, adOpenKeyset, adLockOptimistic
                AGT.AddNew
            Else
                strsql = "Select * From Agent Where AgentCode = '" & Trim(txtAGTCode) & "'"
                AGT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With AGT
                    !AgentCode = ChkStr(txtAGTCode)
                    !Remarks = ChkStr(txtAGTDesc)
                    !PayorCode = Mid(CboPAYCode, 1, 2)
                    !LastUpdateUser = Logon
                    !LastUpdateDate = Date
                End With
                AGT.Update
                Call ClearForm(Me)
                Call AGTListing(AGTCriteria)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
        AGT.Close
        Set SaveAGT = Nothing
        End If
        EditFlag = False
    End If
End If

    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = Default
    

End Sub

Private Sub cmdDelete_Click()
Dim DeleteAGT As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

Screen.MousePointer = vbHourglass
CmdAdd.Enabled = True
Screen.MousePointer = Default

EditFlag = False
If txtAGTCode = "" Then
   MsgBox "Please Enter Agent Code", vbOKOnly + vbCritical, DialogBoxTitle
   'txtAGTCode.SetFocus
ElseIf lstAgentList.listitems.count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
'Else
'   sqlMBM = "Select MBMNumber from MBMOthers where MBMAgentCode = '" & RemStr(txtAGTCode) & "'"
      
'   rsCheckMBM.MaxRecords = 1
'   rsCheckMBM.Open sqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
   
      
 '  If rsCheckMBM.recordCount Then
 '       MsgBox "This record cannot be deleted because it is used in Member Table!", vbOKOnly + vbCritical, "Error Message"
        'Call ClearForm(Me)
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes Then
           strsql = "Delete From Agent Where AgentCode = '" & Trim(txtAGTCode) & "'" 'ChkStr(lstRaceList.SelectedItem.Text) & "'"
           DeleteAGT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           'DeleteAGT.Close
           'Set DeleteAGT = Nothing
           CmdSave.Enabled = True
           Call ClearForm(Me)
           Call AGTListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           txtAGTCode.SetFocus
        End If
     'End If
End If
End Sub

Private Sub cmdSearch_Click()
Dim strAppend As String

Screen.MousePointer = vbHourglass
strAppend = ""

If Len(Trim(txtAGTCode)) Then
    strAppend = strAppend + " And AgentCode = '" & Trim(ChkStr(txtAGTCode)) & "'"
End If

If Len(Trim(CboPAYCode)) Then
    strAppend = strAppend + " AND PAYorCode like '" & RemStr(Mid(CboPAYCode, 1, IIf(InStr(1, CboPAYCode, "-") = 0, 4, InStr(1, CboPAYCode, "-") - 1))) & "%'"
    'strAppend = strAppend + "PAYCode = '" & Trim(Mid(CboPAYCode, 1, 2)) & "'"
End If
    
    'PAYCriteria = strAppend
    Call AGTListing(strAppend)
    Screen.MousePointer = Default


'If Len(Trim(txtAGTDesc)) Then
'strAppend = strAppend + " And Remarks like '%" & ChkStr(txtAGTDesc) & "%'"
'End If

End Sub


Private Sub AGTListing(Optional strAppend As String)
Dim AGT3 As New ADODB.Recordset
Dim AGTMaster As ListItem
Dim sql As String
Dim TotalAGTCount As Integer
Dim PAYRec As New ADODB.Recordset
Dim PAYSql As String
TotalAGTCount = 0
lstAgentList.listitems.Clear


    sql = "SELECT AgentCode, " & _
    " Remarks," & _
    " PayorCode" & _
    " from Agent where 0=0 "
    
    If Len(Trim(strAppend)) Then
    sql = sql + strAppend
    End If
    
    AGT3.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If AGT3.recordCount = 0 Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until AGT3.EOF
        Set AGTMaster = lstAgentList.listitems.Add(, , AGT3!AgentCode)
            AGTMaster.SubItems(1) = Trim(AGT3!Remarks)
            
            PAYSql = "SELECT PayCompanyName FROM PAY WHERE PAYCode = '" & AGT3!PayorCode & "'"
            PAYRec.Open PAYSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            AGTMaster.SubItems(2) = Trim(AGT3!PayorCode) & " - " & PAYRec!PAYCompanyName
            PAYRec.Close
            Set PAYRec = Nothing
            AGT3.MoveNext
        TotalAGTCount = TotalAGTCount + 1
        txtTotalRecord = TotalAGTCount
        Loop
    AGT3.Close
    Set AGT3 = Nothing
End If
End Sub


Private Sub CmdClear_Click()
Call ClearForm(Me)
    lstAgentList.listitems.Clear
     CmdEdit.Enabled = False
End Sub

Private Sub CmdExit_Click()
Unload Me
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

Private Sub Form_Load()
Call AgentComboListing
End Sub
Private Sub lstAgentList_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String

If lstAgentList.listitems.count Then
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    strsql = "SELECT AgentCode, Remarks, PayorCode " & _
    "from Agent Where AgentCode = '" & lstAgentList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    txtAGTCode = rst3!AgentCode
    txtAGTDesc = Trim(rst3!Remarks)
    CboPAYCode = rst3!PayorCode
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    rst3.Close
    Set rst3 = Nothing
End If
End Sub
Private Sub lstAgentList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstAgentList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstAgentList, ColumnHeader.Index, ldtString, m_blnDirection(2)
   Case 3
        SortListView lstAgentList, ColumnHeader.Index, ldtString, m_blnDirection(3)
   End Select
End Sub

'************************Start EntriesEnable*********************************

Private Sub EntriesEnable(status As Boolean)
     
    txtAGTCode.Enabled = status
    txtAGTDesc.Enabled = status
    CboPAYCode.Enabled = status
   
End Sub
'************************End EntriesEnable*********************************


