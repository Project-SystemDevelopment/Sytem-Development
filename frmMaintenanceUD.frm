VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenanceUD 
   ClientHeight    =   7695
   ClientLeft      =   2160
   ClientTop       =   1695
   ClientWidth     =   10140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   10140
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   735
      Left            =   960
      TabIndex        =   17
      Top             =   6840
      Width           =   6555
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   3600
         TabIndex        =   24
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   5520
         TabIndex        =   23
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   4680
         TabIndex        =   22
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10125
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Underlying Information"
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
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3285
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6045
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   10663
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   741
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Underlying Info"
      TabPicture(0)   =   "frmMaintenanceUD.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label43"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstUDList"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Case Exclusion"
      TabPicture(1)   =   "frmMaintenanceUD.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame12"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstCaseExList"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label40"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame13 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   6525
         Begin VB.TextBox TxtUDCode 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   12
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox TxtUDDescription 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   1680
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   720
            Width           =   4545
         End
         Begin VB.Label Label42 
            Caption         =   "UD Description *"
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
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label41 
            Caption         =   "UD Code             *"
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
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame12 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   -74640
         TabIndex        =   3
         Top             =   720
         Width           =   6525
         Begin VB.TextBox TxtCasExDescription 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1050
            Left            =   2640
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   720
            Width           =   3585
         End
         Begin VB.TextBox TxtCasExCode 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2640
            MaxLength       =   3
            TabIndex        =   4
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label39 
            Caption         =   "Case Exclusion Code             *"
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
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label38 
            Caption         =   "Case Exclusion Description *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Width           =   2475
         End
      End
      Begin MSComctlLib.ListView lstCaseExList 
         Height          =   2715
         Left            =   -74640
         TabIndex        =   8
         Top             =   3120
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   4789
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case Exclusion Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Case Exclusion Description"
            Object.Width           =   8996
         EndProperty
      End
      Begin MSComctlLib.ListView lstUDList 
         Height          =   2835
         Left            =   360
         TabIndex        =   15
         Top             =   3120
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   5001
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "UD Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "UD Description"
            Object.Width           =   8996
         EndProperty
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "UD CODE LIST"
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
         Left            =   360
         TabIndex        =   16
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Caption         =   "CASE EXCLUSION LIST"
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
         Left            =   -74640
         TabIndex        =   9
         Top             =   2880
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmMaintenanceUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Public CASECriteria As String
Public UDCriteria As String

Private Sub Form_Activate()
    SSTab1.Tab = 0
End Sub

Private Sub Form_Load()
    Frame8.Width = Screen.Width
    cmdDelete.Enabled = True
    Call CenterForm(Me)
    CASECriteria = ""
    UDCriteria = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Call EnterToTab(KeyAscii)
End Sub

'*****************************Start Tab Clicking**********************************
Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Select Case SSTab1.Tab
    Case 0 'UD
         lstUDList.ListItems.Clear
         Call ClearForm(Me)
         
    Case 1 'CASE
         lstCaseExList.ListItems.Clear
         Call ClearForm(Me)
    End Select
End Sub
'*****************************End Tab Clicking**********************************

'********************Start Add, Delete, Edit, Save, Search, Cancel buttons******************
Private Sub cmdAdd_Click()
EditFlag = False
Call ClearForm(Me)
Call DisableCmdButton(Me)
Call EntriesEnable(True)
cmdSave.Enabled = True

    Select Case SSTab1.Tab
        Case 0
            TxtUDCode.SetFocus
        Case 1
            TxtCasExCode.SetFocus
    End Select

End Sub

Private Sub cmdDelete_Click()
Screen.MousePointer = vbHourglass
    If SSTab1.Tab = 0 Then
        Call DeleteUD
    ElseIf SSTab1.Tab = 1 Then
        Call DeleteCASE
    End If
cmdAdd.Enabled = True
Screen.MousePointer = Default
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    
    If TxtCasExCode <> "" Then
       TxtCasExCode.Enabled = False
       TxtCasExDescription.SetFocus
    End If
        
    If TxtUDCode <> "" Then
       TxtUDCode.Enabled = False
       TxtUDDescription.SetFocus
    End If
    
End Sub

Private Sub cmdSave_Click()
   Screen.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0
            Call SaveUD
        Case 1
            Call SaveCASE
    End Select
    cmdAdd.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub cmdSearch_Click()

   Screen.MousePointer = vbHourglass
   Select Case SSTab1.Tab
        Case 0
            Call SearchCASE
            TxtCasExCode.Enabled = True
            TxtCasExDescription.Enabled = True
            TxtCasExCode.SetFocus
        
        Case 1
            Call SearchUD
            TxtUDCode.Enabled = True
            TxtUDDescription.Enabled = True
            TxtUDCode.SetFocus
    End Select
Screen.MousePointer = Default
cmdAdd.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'********************End Add, Delete, Edit, Save, Search, Cancel buttons******************

Private Sub EntriesEnable(status As Boolean)
        
    TxtCasExCode.Enabled = status
    TxtCasExDescription.Enabled = status
    TxtUDCode.Enabled = status
    TxtUDDescription.Enabled = status

End Sub

'***************************Start Case Exclusion**********************************

Private Sub SaveCASE()
On Error Resume Next
Dim RecordSaved
Dim CasEx As New ADODB.Recordset
Dim SaveCasEx As New ADODB.Recordset
Dim strsql As String
    
    
    If TxtCasExCode = "" Then
        MsgBox "Please Enter Case Exclusion Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtCasExCode.SetFocus
    ElseIf TxtCasExDescription = "" Then
        MsgBox "Please Enter Case Exclusion Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtCasExDescription.SetFocus
    Else
        strsql = "Select * From CASExclusion Where CASExclusionCode = '" & Trim(TxtCasExCode) & "'"
        CasEx.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If CasEx.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            CasEx.Close
            TxtCasExCode.SetFocus
            TxtCasExCode.SelStart = 0
            TxtCasExCode.SelLength = Len(TxtCasExCode)
            Set CasEx = Nothing
            Exit Sub
            Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
            If EditFlag = False Then
      
             SaveCasEx.Open "CASExclusion", medixmasterconn, adOpenKeyset, adLockOptimistic
             SaveCasEx.AddNew
             Else
                strsql = "Select * From CASExclusion Where CASExclusionCode = '" & Trim(TxtCasExCode) & "'"
                SaveCasEx.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With SaveCasEx
                    !CASExclusionCode = TxtCasExCode
                    !CASExclustionDescription = TxtCasExDescription
                End With
                SaveCasEx.Update
                Call ClearForm(Me)
                Call CASEListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
      
            SaveCasEx.Close
            Set SaveCasEx = Nothing
        End If
        EditFlag = False
    End If
End If
End Sub

Private Sub DeleteCASE()
Dim DeleteCASE As New ADODB.Recordset
Dim strsql As String
Dim DeleteRecord

EditFlag = False

If TxtCasExCode = "" Then
   MsgBox "Please Enter Case Exclusion Code", vbOKOnly + vbCritical, DialogBoxTitle
   TxtCasExCode.SetFocus
ElseIf lstCaseExList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   DeleteRecord = MsgBox("Do you want to delete this record?", vbYesNo + vbExclamation)
        If DeleteRecord = vbYes Then
           'strsql = " Delete Form CASExclusion Where CASExclusionCode = '" & ChkStr(lstCaseExList.SelectedItem.Text) & "'"
           strsql = "Delete From CASExclusion Where CASExclusionCode = '" & TxtCasExCode & "'"
           DeleteCASE.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           Call ClearForm(Me)
           Call CASEListing
           Call setCmdButton(Me, False)
           Call EntriesEnable(True)
           TxtCasExCode.SetFocus
           cmdAdd.Enabled = True
        End If
End If
Exit Sub
End Sub

Private Sub SearchCASE()
Dim strAppend As String
   
    strAppend = ""
    
    If Len(Trim(TxtCasExCode)) Then
        strAppend = strAppend + " And CASExclusionCode like '%" & Mid(ChkStr(TxtCasExCode), 1, 2) & "%'"
    End If
    If Len(Trim(TxtCasExDescription)) Then
        strAppend = strAppend + " And CASExclustionDescription like '%" & Mid(ChkStr(TxtCasExDescription), 1, 2) & "%'"
    End If
    
    CASECriteria = strAppend
    Call CASEListing(strAppend)
End Sub

Private Sub lstCASEExList_Click()
    On Error Resume Next
    Dim rst3 As New ADODB.Recordset
    Dim strsql As String
    
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    strsql = "Select CASExclusionCode, CASExclustionDescription " & _
    "from CASExclusion Where CASExclusionCode = '" & lstCaseExList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TxtCasExCode = Trim(rst3!CASExclusionCode)
    'TxtCasExDescription = Trim(rst3!CASExclustionDescription)
    TxtCasExDescription = rst3!CASExclustionDescription
    
    Call EntriesEnable(False)
    cmdSave.Enabled = False
    rst3.Close
End Sub

Private Sub CASEListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim CASEMaster As ListItem
    Dim sql As String
     
    lstCaseExList.ListItems.Clear
        sql = " SELECT CASExclusionCode, CASExclustionDescription " & _
        " from CASExclusion where 0=0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If rst2.recordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set CASEMaster = lstCaseExList.ListItems.Add(, , rst2!CASExclusionCode)
                CASEMaster.SubItems(1) = rst2!CASExclustionDescription
                'CASEMaster.ListSubItems(1) = rst2!CASExclustionDescription
                rst2.MoveNext
        Loop
    rst2.Close
    End If
End Sub
'***************************End Case Exclusion**********************************

'***************************Start UD Exclusion**********************************
Private Sub SaveUD()
On Error Resume Next
Dim RecordSaved
Dim UD As New ADODB.Recordset
Dim SaveUD As New ADODB.Recordset
Dim strsql As String
    
    
    If TxtUDCode = "" Then
        MsgBox "Please Enter UD Code", vbOKOnly + vbCritical, DialogBoxTitle
        TxtUDCode.SetFocus
    ElseIf TxtUDDescription = "" Then
        MsgBox "Please Enter UD Description", vbOKOnly + vbCritical, DialogBoxTitle
        TxtUDDescription.SetFocus
    Else
        strsql = "Select * From UD Where UDCode = '" & Trim(TxtUDCode) & "'"
        UD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If UD.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please Change The Corresponding Field", vbCritical, "Error"
            UD.Close
            TxtUDCode.SetFocus
            TxtUDCode.SelStart = 0
            TxtUDCode.SelLength = Len(TxtUDCode)
            
            Set UD = Nothing
            Exit Sub
            Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
            If EditFlag = False Then
      
             SaveUD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
             'UD.Open "UD", medixmasterconn, adOpenKeyset, adLockOptimistic
             SaveUD.AddNew
             Else
                strsql = "Select * From UD Where UDCode = '" & Trim(TxtUDCode) & "'"
                SaveUD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
                With SaveUD
                    !UDCode = TxtUDCode
                    !UDDescription = TxtUDDescription
                End With
                SaveUD.Update
                Call ClearForm(Me)
                Call UDListing
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
      
            SaveUD.Close
            Set SaveUD = Nothing
        End If
        EditFlag = False
    End If
End If
End Sub


Private Sub DeleteUD()
Dim DeleteUD As New ADODB.Recordset
Dim rsCheckCAS As New ADODB.Recordset
Dim strsql, sqlCAS As String
Dim DeleteRecord

EditFlag = False
If TxtUDCode = "" Then
   MsgBox "Please Enter UD Code", vbOKOnly + vbCritical, DialogBoxTitle
   TxtUDCode.SetFocus
ElseIf lstUDList.ListItems.Count = 0 Then
   MsgBox "Please search the record", vbOKOnly + vbCritical, DialogBoxTitle
Else
   sqlCAS = "Select CASNumber from CAS where UDCode = '" & ChkStr(lstUDList.SelectedItem.Text) & "'"
   
   
   rsCheckCAS.MaxRecords = 1
   rsCheckCAS.Open sqlCAS, medixmasterconn, adOpenKeyset, adLockOptimistic
   
     
   If rsCheckCAS.recordCount Then
        MsgBox "This record cannot be deleted because it is used in CAS Table!", vbOKOnly + vbCritical, "Error Message"
   Else
   DeleteRecord = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
     If DeleteRecord = vbYes And rsCheckCAS.recordCount = 0 Then
           strsql = "Delete From UD Where UDCode = '" & ChkStr(lstUDList.SelectedItem.Text) & "'"
           DeleteUD.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
           MsgBox "Record Deleted...", vbOKOnly + vbInformation, DialogBoxTitle
           cmdSave.Enabled = True
           Call ClearForm(Me)
           Call UDListing
           Call DisableCmdButton(Me)
           Call EntriesEnable(True)
           TxtUDCode.SetFocus
         End If
     End If
End If
Exit Sub

End Sub

Private Sub SearchUD()
    Dim strAppend As String
   
    strAppend = ""
    
    If Len(Trim(TxtUDCode)) Then
        strAppend = strAppend + " And UDCode LIKE '%" & TxtUDCode & "%'"
    End If
    If Len(Trim(TxtUDDescription)) Then
        strAppend = strAppend + " And UDDescription LIKE '%" & TxtUDDescription & "%'"
    End If
    
    UDCriteria = strAppend
    Call UDListing(strAppend)
    
End Sub

Private Sub lstUDList_Click()
    On Error Resume Next
    Dim rst3 As New ADODB.Recordset
    Dim strsql As String
    
    Call EnableCmdButton(Me)
    Call ClearForm(Me)
    strsql = "Select UDCode, UDDescription " & _
    "from UD Where UDCode = '" & lstUDList.SelectedItem.Text & "'"
    rst3.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    TxtUDCode = Trim(rst3!UDCode)
    TxtUDDescription = Trim(rst3!UDDescription)
    
    
    Call EntriesEnable(False)
    cmdSave.Enabled = False
    rst3.Close
End Sub

Private Sub UDListing(Optional strAppend As String)
    Dim rst2 As New ADODB.Recordset
    Dim UDMaster As ListItem
    Dim sql As String
     
    lstUDList.ListItems.Clear
        sql = "SELECT UDCode, UDDescription From UD where 0=0"
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If rst2.recordCount = 0 Then
        MsgBox "Record not found", vbInformation + vbOKOnly, DialogBoxTitle
    Else
        Do Until rst2.EOF
            Set UDMaster = lstUDList.ListItems.Add(, , rst2!UDCode)
                UDMaster.SubItems(1) = Trim(rst2!UDDescription)
                rst2.MoveNext
        Loop
    rst2.Close
    End If
End Sub
'***************************End UD Exclusion**********************************

