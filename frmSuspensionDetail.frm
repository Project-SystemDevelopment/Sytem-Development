VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSuspensionDetail 
   Caption         =   "Suspension Detail"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   9870
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Suspension Category"
      TabPicture(0)   =   "frmSuspensionDetail.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label25"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstSuspensionList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Suspension Details"
      TabPicture(1)   =   "frmSuspensionDetail.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstAgentList"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   10860
         Begin VB.TextBox txtCATcode 
            Height          =   285
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   1
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtCATdescription 
            Height          =   1095
            Left            =   2040
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   720
            Width           =   8535
         End
         Begin VB.Label Label12 
            Caption         =   "Suspension Code    *"
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
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label13 
            Caption         =   "Suspension Description *"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   1770
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   11415
         Begin VB.ComboBox cboSelection 
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtAGTCode 
            Height          =   285
            Left            =   4800
            MaxLength       =   15
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtAGTDesc 
            Height          =   960
            Left            =   1200
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   600
            Width           =   10080
         End
         Begin VB.ComboBox CboPAYCode 
            Height          =   315
            Left            =   6960
            TabIndex        =   13
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label Label2 
            Caption         =   "Selection"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label29 
            Caption         =   "Member Number*"
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
            Left            =   3240
            TabIndex        =   20
            Top             =   240
            Width           =   1680
         End
         Begin VB.Label Label14 
            Caption         =   "Description "
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label41 
            Caption         =   "Payor *"
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
            Left            =   6240
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView lstAgentList 
         Height          =   4320
         Left            =   -74880
         TabIndex        =   21
         Top             =   2400
         Width           =   11415
         _ExtentX        =   20135
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Suspension Code"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Member Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Payor"
            Object.Width           =   14993
         EndProperty
      End
      Begin MSComctlLib.ListView lstSuspensionList 
         Height          =   3525
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   6218
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Suspension Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "SUSPENSION  LIST"
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
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "SUSPENSION LISTING"
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
         Left            =   -74880
         TabIndex        =   22
         Top             =   2160
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   11655
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   6960
         TabIndex        =   9
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   8640
         TabIndex        =   10
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
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Suspension Detail"
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
      Height          =   220
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmSuspensionDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim EditFlag As Boolean
Public AGTCriteria As String
Public CATCriteria As String
Private m_blnDirection(1 To 4) As Boolean


Private Sub AgentComboListing()

    Dim PAY As New ADODB.Recordset
   ' Dim HLT As New ADODB.Recordset
        
    PAY.Open "PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    'HLT.Open "HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    CboPAYCode.Clear
    'cboPAYHLTCode.Clear
    
    Do Until PAY.EOF
        CboPAYCode.AddItem PAY!PAYCode & " - " & PAY!PAYCompanyName
        PAY.MoveNext
    Loop
        
    cboSelection.AddItem "01" & " - " & "Member"
    cboSelection.AddItem "02" & " - " & "Agent"
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

Private Sub cboSelection_Change()
    If Mid(cboSelection, 1, 2) = "01" Then
        Label29.Caption = "Member Number*"
        'txtAGTCode.Visible = False
        'txtMBMNo.Visible = True
        lstAgentList.ColumnHeaders(2).Text = " Member No"
    ElseIf Mid(cboSelection, 1, 2) = "02" Then
        Label29.Caption = "Agent*"
        lstAgentList.ColumnHeaders(2).Text = "Agent Code"
        'txtMBMNo.Visible = False
        'txtAGTCode.Visible = True
    End If
End Sub

Private Sub cboSelection_Click()
    If Mid(cboSelection, 1, 2) = "01" Then
        Label29.Caption = "Member Number*"
        'txtAGTCode.Visible = False
        'txtMBMNo.Visible = True
        lstAgentList.ColumnHeaders(2).Text = "Member No"

    ElseIf Mid(cboSelection, 1, 2) = "02" Then
        Label29.Caption = "Agent*"
        'txtMBMNo.Visible = False
        'txtAGTCode.Visible = True
        lstAgentList.ColumnHeaders(2).Text = "Agent Code"
    End If

End Sub

Private Sub cboSelection_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboSelection.Text, cboSelection.SelStart) + Chr(KeyAscii) + Mid(cboSelection.Text, cboSelection.SelStart + cboSelection.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To cboSelection.ListCount - 1
        
        If NewText = LCase(Left(cboSelection.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboSelection.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
            If ValidCount = 1 Then
                cboSelection.Text = ValidValue
                cboSelection.SelStart = 0
                cboSelection.SelLength = Len(ValidValue)
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
   
    Select Case SSTab1.Tab
        Case 0
            txtCATcode.SetFocus
        Case 1
            cboSelection.SetFocus
    End Select
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    'txtAGTDesc.SetFocus
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
     
    If txtCATcode <> "" Then
       txtCATcode.Enabled = False
       txtCATdescription.Enabled = True
       txtCATdescription.SetFocus
    End If
    
    If txtAGTCode <> "" Then
       txtAGTCode.Enabled = False
       cboSelection.Enabled = False
       CboPAYCode.Enabled = False
       txtAGTDesc.SetFocus
    End If
    
   

End Sub


Private Sub cmdprint_Click()
    Screen.MousePointer = vbHourglass
    'Select Case SSTab1.Tab
        'Case 0
            'Call PrintCAT
        'Case 1
            'Call PrintDetail
     'End Select
    Select Case SSTab1.Tab
        Case 0
            If lstSuspensionList.listitems.count <> 0 Then
                Call ExportListViewtoExcel(lstSuspensionList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
        Case 1
            If lstAgentList.listitems.count <> 0 Then
                Call ExportListViewtoExcel(lstAgentList)
            Else
                MsgBox "Report cannot be printed without any record.", vbInformation
            End If
    End Select
     
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
    Screen.MousePointer = Default
End Sub

'Private Sub PrintDetail()
    'RptAgentList.Show
'End Sub

Private Sub cmdSave_Click()
  Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SaveCAT
        Case 1
            Call SaveDetail
        End Select
    CmdAdd.Enabled = True
    CmdEdit.Enabled = False
        

    Screen.MousePointer = Default
End Sub

Private Sub SaveDetail()
'On Error Resume Next
Dim RecordSaved
Dim AGT As New ADODB.Recordset
Dim SaveAGT As New ADODB.Recordset
Dim strsql As String

    Screen.MousePointer = vbHourglass
    
    If cboSelection = "" Then
        MsgBox "Please Enter Selection", vbOKOnly + vbCritical, DialogBoxTitle
        cboSelection.SetFocus
    ElseIf txtAGTCode = "" Then
        MsgBox "Please Enter Member Number/Agent Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtAGTCode.SetFocus
    ElseIf txtAGTDesc = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtAGTDesc.SetFocus
    ElseIf CboPAYCode = "" Then
        MsgBox "Please Enter Payor", vbOKOnly + vbCritical, DialogBoxTitle
        CboPAYCode.SetFocus
    ElseIf Mid(cboSelection, 1, 2) = "01" And Mid(txtAGTCode, 1, 2) <> Mid(ChkStr(CboPAYCode), 1, 2) Then
        MsgBox "Payor Code Not Match", vbOKOnly + vbCritical, DialogBoxTitle
      
    Else
        If Mid(cboSelection, 1, 2) = "01" Then
            strsql = "Select * From SuspensionDetail Where MembershipNo = '" & Trim(txtAGTCode) & "'"
        Else
          strsql = "Select * From SuspensionDetail Where AgentCode = '" & Trim(txtAGTCode) & "'"
        End If
    
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
                AGT.Open "SuspensionDetail", medixmasterconn, adOpenKeyset, adLockOptimistic
                AGT.AddNew
            Else
                strsql = "SELECT * from SuspensionDetail where SuspensionCode = '" & Mid(cboSelection, 1, 2) & "'"
    
                If Mid(cboSelection, 1, 2) = "02" Then
                    strsql = strsql & " and AgentCode = '" & Trim(txtAGTCode) & "'"
                Else
                    strsql = strsql & " and MembershipNo = '" & Trim(txtAGTCode) & "'"
                End If

                'strsql = "Select * From SuspensionDetail Where MembershipNo = '" & Trim(txtAGTCode) & "'"
                AGT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With AGT
                    
                If Mid(cboSelection, 1, 2) = "01" Then
                    !MembershipNo = ChkStr(txtAGTCode)
                Else
                    !AgentCode = ChkStr(txtAGTCode)
                End If
                    
                    !Remarks = ChkStr(txtAGTDesc)
                    !PayorCode = Mid(CboPAYCode, 1, 2)
                    !SuspensionCode = Mid(cboSelection, 1, 2)
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
Screen.MousePointer = vbHourglass
    If SSTab1.Tab = 0 Then
       Call DeleteCat
       CmdAdd.Enabled = True
    ElseIf SSTab1.Tab = 1 Then
      Call DeleteDETAIL
      CmdAdd.Enabled = True
    End If
Screen.MousePointer = Default
End Sub

Private Sub DeleteDETAIL()
Dim DeleteAGT As New ADODB.Recordset
Dim rsCheckMBM As New ADODB.Recordset
Dim strsql, sqlMBM As String
Dim DeleteRecord

Screen.MousePointer = vbHourglass
CmdAdd.Enabled = True
Screen.MousePointer = Default

EditFlag = False
If cboSelection = "" Then
   MsgBox "Please Enter Selection", vbOKOnly + vbCritical, DialogBoxTitle
End If

If Mid(cboSelection, 1, 2) = "" Then
       MsgBox "Please Enter Member Number or Agent Code", vbOKOnly + vbCritical, DialogBoxTitle
'ElseIf Mid(cboSelection, 1, 2) = "" Then
'   MsgBox "Please Enter Agent Code", vbOKOnly + vbCritical, DialogBoxTitle
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
        If Mid(cboSelection, 1, 2) = "01" Then
            strsql = "Delete From SuspensionDetail Where MembershipNo = '" & Trim(txtAGTCode) & "'"
        Else
            strsql = "Delete From SuspensionDetail Where AgentCode = '" & Trim(txtAGTCode) & "'"
        End If

'           strsql = "Delete From SuspensionDetail Where AgentCode = '" & Trim(txtAGTCode) & ChkStr(lstAgentList.SelectedItem.Text) & "'"
           
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
Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SearchCAT
            Call ClearForm(Me)
            'txtHLTcode.Text = ""
            'txtHLTdescription.Text = ""
            txtCATcode.Enabled = True
            txtCATdescription.Enabled = True
            txtCATcode.SetFocus
        Case 1
            Call SearchDetail
            'txtINScode.Text = ""
            'txtINSdescription.Text = ""
            txtAGTCode.Enabled = True
            txtAGTDesc.Enabled = True
            txtAGTCode.SetFocus
        End Select
End Sub
Private Sub SearchDetail()
Dim StrAppend As String

Screen.MousePointer = vbHourglass
StrAppend = ""

If Len(Trim(cboSelection)) Then
    StrAppend = StrAppend & " AND SuspensionCode = '" & Mid(cboSelection, 1, 2) & "'"
End If
    
If Len(Trim(txtAGTCode)) And Mid(cboSelection, 1, 2) = "02" Then
    StrAppend = StrAppend & " And AgentCode = '" & Trim(ChkStr(txtAGTCode)) & "'"
ElseIf Len(Trim(txtAGTCode)) And Mid(cboSelection, 1, 2) = "01" Then
    StrAppend = StrAppend & " And MembershipNo = '" & Trim(ChkStr(txtAGTCode)) & "'"
End If

If Len(Trim(CboPAYCode)) Then
    'strAppend = strAppend + " AND PayorCode = '" & RemStr(Mid(CboPAYCode, 1, IIf(InStr(1, CboPAYCode, "-") = 0, 4, InStr(1, CboPAYCode, "-") - 1))) & "%'"
    StrAppend = StrAppend & " AND PAYORCode = '" & Trim(Mid(CboPAYCode, 1, 2)) & "'"
End If
    
    
    AGTCriteria = StrAppend
    Call AGTListing(StrAppend)
    Screen.MousePointer = Default

'If Len(Trim(txtAGTDesc)) Then
'strAppend = strAppend + " And Remarks like '%" & ChkStr(txtAGTDesc) & "%'"
'End If

End Sub


Private Sub AGTListing(Optional StrAppend As String)
Dim AGT3 As New ADODB.Recordset
Dim AGTMaster As ListItem
Dim Sql As String
Dim TotalAGTCount As Integer
Dim PAYREC As New ADODB.Recordset
Dim paysql As String
TotalAGTCount = 0
lstAgentList.listitems.Clear


    Sql = "SELECT AgentCode, " & _
    " Remarks," & _
    " PayorCode," & _
    " SuspensionCode," & _
    " MembershipNo" & _
    " from SuspensionDetail where 0=0 "
    
    If Len(Trim(StrAppend)) Then
    Sql = Sql + StrAppend
    End If
    
    AGT3.Open Sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If AGT3.recordCount = 0 Then
       MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    Do Until AGT3.EOF
        Set AGTMaster = lstAgentList.listitems.Add(, , AGT3!SuspensionCode)
            If Len(Trim(AGT3!MembershipNo)) Then
                AGTMaster.SubItems(1) = Trim(AGT3!MembershipNo)
            Else
                AGTMaster.SubItems(1) = Trim(AGT3!AgentCode)
            End If
            AGTMaster.SubItems(2) = Trim(AGT3!Remarks)
            paysql = "SELECT PayCompanyName FROM PAY WHERE PAYCode = '" & AGT3!PayorCode & "'"
            PAYREC.Open paysql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            AGTMaster.SubItems(3) = Trim(AGT3!PayorCode) & " - " & PAYREC!PAYCompanyName
            PAYREC.Close
            Set PAYREC = Nothing
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
    CmdAdd.Enabled = True
    lstAgentList.listitems.Clear
    lstSuspensionList.listitems.Clear
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    Call EntriesEnable(True)
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

    If lstAgentList.listitems.count <> 0 Then
   
    
    strsql = "SELECT * from SuspensionDetail where SuspensionCode = '" & lstAgentList.SelectedItem.Text & "'"
    
    If lstAgentList.SelectedItem.Text = "02" Then
        strsql = strsql & " and AgentCode = '" & lstAgentList.SelectedItem.SubItems(1) & "'"
    Else
        strsql = strsql & " and MembershipNo = '" & lstAgentList.SelectedItem.SubItems(1) & "'"
    End If
    
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    
    If Len(Trim(rst3!MembershipNo)) Then
        txtAGTCode = rst3!MembershipNo
    Else
        txtAGTCode = rst3!AgentCode
    End If
    
    'If Mid(cboSelection, 1, 2) = "02" Then
    'Else
    '    txtAGTCode = ""
    'End If
    
    txtAGTDesc = Trim(rst3!Remarks)
    CboPAYCode = rst3!PayorCode
    cboSelection = rst3!SuspensionCode
    rst3.Close
    Set rst3 = Nothing
    
    Call EntriesEnable(False)
    CmdSave.Enabled = False
    
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
   Case 4
        SortListView lstAgentList, ColumnHeader.Index, ldtString, m_blnDirection(4)
   
   End Select
End Sub

'************************Start EntriesEnable*********************************

Private Sub EntriesEnable(status As Boolean)
     
    txtCATcode.Enabled = status
    cboSelection.Enabled = status
    txtAGTCode.Enabled = status
    txtAGTDesc.Enabled = status
    CboPAYCode.Enabled = status
    CmdSearch.Enabled = status
    
End Sub
'************************End EntriesEnable*********************************
    
    
    ' '''''''''''''''''' Save, edit , Add, delete CAT..... ''''''''''''''
Private Sub DeleteCat()
Dim rsDeleteCAT As New ADODB.Recordset
Dim rsCheckCAT As New ADODB.Recordset
Dim strsql, sqlCAT As String
Dim DeleteRecord
    
EditFlag = False
       
If txtCATcode = "" Then
   MsgBox "Please Enter Category Code", vbOKOnly + vbCritical, DialogBoxTitle
   'TxtCATCode.SetFocus
ElseIf lstSuspensionList.listitems.count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCAT = "Select SuspensionCode From SuspensionDetail where SuspensionCode = '" & Trim(lstSuspensionList.SelectedItem.Text) & "'"
    
   rsCheckCAT.MaxRecords = 1
   rsCheckCAT.Open sqlCAT, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rsCheckCAT.recordCount Then
        MsgBox "This record cannot be deleted because it is used in Suspension Detail table!", vbCritical, "Error in deleting record..."
        Call ClearForm(Me)
        CmdDelete.Enabled = False
        CmdEdit.Enabled = False
    
    Else
        DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes And rsCheckCAT.recordCount = 0 Then
             strsql = "Delete From SuspensionCat Where SuspensionCode = '" & ChkStr(lstSuspensionList.SelectedItem.Text) & "'"
             rsDeleteCAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
             CmdSave.Enabled = True
             Call ClearForm(Me)
             Call CATListing
             Call DisableCmdButton(Me)
             Call EntriesEnable(True)
             txtCATcode.SetFocus
        Else
            CmdDelete.Enabled = False
            CmdEdit.Enabled = False
        End If
    End If
  End If
  Exit Sub
End Sub

    
Private Sub SearchCAT()
Dim StrAppend As String
    
  Screen.MousePointer = vbHourglass
  StrAppend = ""
  
  If Len(Trim(txtCATcode)) Then
  StrAppend = StrAppend + " And SuspensionCode = '" & Trim(Mid(txtCATcode, 1, 2)) & "'"
  'HLTCode like '%" & Mid(ChkStr(txtHLTcode), 1, 2) & "%'"
  End If
    
  CATCriteria = StrAppend
  Call CATListing(StrAppend)
  Screen.MousePointer = Default
  
End Sub

'Private Sub PrintCAT()
    'RptCATList.Show
'End Sub

Private Sub SaveCAT()
   On Error Resume Next
   Dim RecordSaved
   Dim CAT As New ADODB.Recordset
   Dim SaveCAT As New ADODB.Recordset
   Dim strsql As String
   
    If txtCATcode = "" Then
        MsgBox "Please Enter Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtCATcode.SetFocus
    ElseIf txtCATdescription = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtCATdescription.SetFocus
    Else
        strsql = "Select * From SuspensionCat where SuspensionCode =  '" & Trim(txtCATcode) & "'"
        CAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If CAT.recordCount > 0 And EditFlag = False Then
           MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
           CAT.Close
           txtCATcode.SetFocus
           txtCATcode.SelStart = 0
           txtCATcode.SelLength = Len(txtCATcode)
           Set CAT = Nothing
           Exit Sub
        Else
           RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
           If RecordSaved = vbYes Then
              If EditFlag = False Then
                 SaveCAT.Open "SuspensionCat", medixmasterconn, adOpenKeyset, adLockOptimistic
                 SaveCAT.AddNew
              Else
                 strsql = "Select * From SuspensionCat Where SuspensionCode = '" & ChkStr(txtCATcode) & "'"
                 SaveCAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
              End If
              
              With SaveCAT
                !SuspensionCode = ChkStr(txtCATcode)
                !SuspensionDesc = ChkStr(txtCATdescription)
                !CATLastUpdateUser = Logon
                !CATLastUpdateDate = Date
                
              End With
              SaveCAT.Update
            
             Call ClearForm(Me)
             Call CATListing(CATCriteria)
             
             If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
             Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
             End If
             CAT.Close
             Set SaveCAT = Nothing
        End If
        EditFlag = False
     End If
   End If
End Sub


Private Sub CATListing(Optional StrAppend As String)
    Dim rslistCAT As New ADODB.Recordset
    Dim CATMaster As ListItem
    Dim sqlCAT As String
    Dim TotalCATCount As Integer
         
    
    TotalCATCount = 0
    lstSuspensionList.listitems.Clear
    sqlCAT = "select SuspensionCode, SuspensionDesc From SuspensionCat where 0=0 "
    
    If Len(Trim(StrAppend)) Then
        sqlCAT = sqlCAT + StrAppend
    End If
    
    rslistCAT.Open sqlCAT, medixmasterconn, adOpenKeyset, adLockOptimistic
    If rslistCAT.recordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
    
    Do Until rslistCAT.EOF
        With rslistCAT
            Set CATMaster = lstSuspensionList.listitems.Add(, , !SuspensionCode)
                CATMaster.SubItems(1) = !SuspensionDesc
            .MoveNext
        End With
        TotalCATCount = TotalCATCount + 1
        txtTotalRecord = TotalCATCount
    Loop
    End If
    rslistCAT.Close
    Set rslistCAT = Nothing
        
End Sub



Private Sub lstSuspensionList_Click()
Dim strsql As String
    
    If lstSuspensionList.listitems.count <> 0 Then
        Call EnableCmdButton(Me)
        Call ClearForm(Me)
              
        txtCATcode = lstSuspensionList.SelectedItem.Text
        txtCATcode.Enabled = False
        txtCATdescription = lstSuspensionList.SelectedItem.SubItems(1)
        txtCATdescription.Enabled = False
            
        CmdDelete.Enabled = True
        'CmdSave.Enabled = False
    End If
    
End Sub

Private Sub lstSuspensionList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstSuspensionList.SortKey = ColumnHeader.Index - 1
    lstSuspensionList.Sorted = True
    If lstSuspensionList.SortOrder = lvwAscending Then
       lstSuspensionList.SortOrder = lvwDescending
    Else
       lstSuspensionList.SortOrder = lvwAscending
    End If
    
    txtCATcode.Enabled = False
End Sub



Private Sub txtAGTCode_LostFocus()
Dim paysql As String
Dim PAYREC As New ADODB.Recordset

    txtAGTCode = ChkStr(txtAGTCode)
    If Mid(cboSelection, 1, 2) = "01" And Len(Trim(txtAGTCode)) Then
        paysql = "SELECT PAYCODE, PayCompanyName FROM PAY WHERE PAYCode = '" & Mid(txtAGTCode, 1, 2) & "'"
            PAYREC.Open paysql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If PAYREC.recordCount = 0 Then
            MsgBox "Invalid MembershipNumber", vbOKOnly + vbInformation, DialogBoxTitle
            txtAGTCode.SetFocus
        Else
            CboPAYCode = PAYREC!PAYCode & " - " & PAYREC!PAYCompanyName
        End If
        PAYREC.Close
        Set PAYREC = Nothing
    End If
End Sub
