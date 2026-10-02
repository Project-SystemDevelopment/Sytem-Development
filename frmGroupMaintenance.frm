VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGroupMaintenance 
   Caption         =   "Security Maintenance "
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
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
      Height          =   930
      Index           =   1
      Left            =   840
      TabIndex        =   9
      Top             =   840
      Width           =   9885
      Begin VB.TextBox txtGRPDescription 
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
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   1
         Top             =   540
         Width           =   5310
      End
      Begin VB.TextBox txtGRPCode 
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   495
         TabIndex        =   11
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Department Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   495
         TabIndex        =   10
         Top             =   225
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   1800
      Width           =   9855
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   315
         Left            =   7980
         TabIndex        =   6
         Top             =   285
         Width           =   960
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7005
         TabIndex        =   5
         Top             =   285
         Width           =   960
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   285
         Width           =   960
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   285
         Width           =   960
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   285
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView lstGRPList 
      Height          =   4545
      Left            =   840
      TabIndex        =   8
      Top             =   2640
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   8017
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
         Text            =   "Group Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   7937
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6975
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12303
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Department"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Security Maintenance - Department"
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
      TabIndex        =   13
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmGroupMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean

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

Private Sub EntriesEnable(status As Boolean)
    Dim ctl As Control

    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        End If
    Next ctl
End Sub

Private Sub GRPListing()
    'On Error Resume Next
    Dim GRPLisrst As New ADODB.Recordset
    Dim GRPMaster As ListItem
    
    lstGRPList.listitems.Clear
    GRPLisrst.Open "GRP", medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly ', adOpenKeyset, adLockOptimistic
        Do Until GRPLisrst.EOF
            Set GRPMaster = lstGRPList.listitems.Add(, , GRPLisrst!GrpCode)
                GRPMaster.SubItems(1) = Trim(GRPLisrst!GRPDescription)
                GRPLisrst.MoveNext
            Loop
    GRPLisrst.Close
    
End Sub

Private Sub cmdAdd_Click()
     EditFlag = False
    Call ClearForm(Me)
    Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    txtGRPCode.SetFocus
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
End Sub

Private Sub cmdDelete_Click()
 Dim GRPrst As New ADODB.Recordset
    Dim strsql As String
    Dim RecordSaved
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    EditFlag = False
    
    If txtGRPCode = "" Then
        MsgBox "Please Enter Group Code ...", vbOKOnly + vbInformation, DialogBoxTitle
    Else
      RecordSaved = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            strsql = "Delete From GRP Where GRPCode = '" & txtGRPCode & "'"
            GRPrst.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
            cmdSave.Enabled = True
            Call ClearForm(Me)
            Call GRPListing
            Call EnableCmdButton(Me)
            Call EntriesEnable(True)
            txtGRPCode.SetFocus
        End If
    End If
    
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    cmdSave.Enabled = True
    cmdDelete.Enabled = True

End Sub

Private Sub cmdEdit_Click()
   EditFlag = True
    Call EntriesEnable(True)
    txtGRPCode.SetFocus
    cmdSave.Enabled = True
    
    If txtGRPCode <> "" Then
       txtGRPCode.Enabled = False
       txtGRPDescription.SetFocus
    End If

End Sub

Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
        'bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    Dim RecordSaved
    Dim GRP As New ADODB.Recordset
    Dim SaveGRP As New ADODB.Recordset
    Dim strsql As String
    Dim ListCheckBox As Integer
    Dim LoopCount As Integer
    Dim AccessString As String
    
    
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If txtGRPCode = "" Then
        MsgBox "Please Enter Group", vbOKOnly + vbCritical, DialogBoxTitle
        txtGRPCode.SetFocus
    ElseIf txtGRPDescription = "" Then
        MsgBox "Please Enter Description", vbOKOnly + vbCritical, DialogBoxTitle
        txtGRPDescription.SetFocus
    Else
        strsql = "Select * From GRP Where GRPCode = '" & Trim(txtGRPCode) & "'"
        GRP.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        If GRP.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please use edit button to commit the changes", vbOKOnly + vbCritical, DialogBoxTitle
            GRP.Close
            cmdEdit.Enabled = True
            Call ClearForm(Me)
            Set GRP = Nothing
            Set SaveGRP = Nothing
            Exit Sub
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    GRP.Open "GRP", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                    GRP.AddNew
                Else
                    strsql = "Select * From GRP Where GRPCode = '" & Trim(txtGRPCode) & "'"
                    GRP.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                End If
            
                With GRP
                    !GrpCode = ChkStr(txtGRPCode)
                    !GRPDescription = ChkStr(txtGRPDescription)
                End With
                GRP.Update
                Call ClearForm(Me)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                cmdAdd.Enabled = True
                txtGRPCode.SetFocus
                GRP.Close
                Call GRPListing
            End If
            EditFlag = False
        End If
    End If
    
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub Form_Load()
    
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Call GRPListing
    
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub lstGRPList_Click()
    Dim strsql As String
    Dim GRPs As New ADODB.Recordset
    Dim listloop As Integer
    
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    strsql = "SELECT * from GRP "
    strsql = strsql & "Where GRPCode ='" & Trim(lstGRPList.SelectedItem.Text) & "'"
    GRPs.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
        With GRPs
            txtGRPCode = !GrpCode
            txtGRPDescription = !GRPDescription
        End With
   
    Call EnableCmdButton(Me)
    cmdDelete.Enabled = True
    Call EntriesEnable(False)
    
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
    
End Sub

Private Sub lstGRPList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strsql As String
    Dim GRPs As New ADODB.Recordset
    Dim listloop As Integer
    
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    strsql = "SELECT * from GRP "
    strsql = strsql & "Where GRPCode ='" & Trim(lstGRPList.SelectedItem.Text) & "'"
    GRPs.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
        With GRPs
            txtGRPCode = !GrpCode
            txtGRPDescription = !GRPDescription
        End With
   
    Call EnableCmdButton(Me)
    cmdDelete.Enabled = True
    Call EntriesEnable(False)
    
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub lstGRPList_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim strsql As String
    Dim GRPs As New ADODB.Recordset
    Dim listloop As Integer
    
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    strsql = "SELECT * from GRP "
    strsql = strsql & "Where GRPCode ='" & Trim(lstGRPList.SelectedItem.Text) & "'"
    GRPs.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
        With GRPs
            txtGRPCode = !GrpCode
            txtGRPDescription = !GRPDescription
        End With
   
    Call EnableCmdButton(Me)
    cmdDelete.Enabled = True
    Call EntriesEnable(False)
    
    
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
End Sub
