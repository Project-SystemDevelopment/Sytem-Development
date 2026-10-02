VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserMaintenance 
   Caption         =   "Security Maintenance "
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   ScaleHeight     =   7740
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstUsersList 
      Height          =   3720
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   6562
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
         Text            =   "User ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Department"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   8565
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   225
         Width           =   960
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1050
         TabIndex        =   7
         Top             =   225
         Width           =   960
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2025
         TabIndex        =   8
         Top             =   225
         Width           =   960
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6075
         TabIndex        =   9
         Top             =   225
         Width           =   960
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7065
         TabIndex        =   10
         Top             =   225
         Width           =   960
      End
   End
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
      Height          =   1935
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   8565
      Begin VB.ComboBox txtGRPCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   3
         Top             =   1440
         Width           =   2625
      End
      Begin VB.TextBox txtUSRName 
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
         Left            =   1485
         MaxLength       =   200
         TabIndex        =   1
         Top             =   720
         Width           =   5835
      End
      Begin VB.TextBox txtUSRPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1485
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2640
      End
      Begin VB.ComboBox txtUSRStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Text            =   "A"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtUSRCode 
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
         IMEMode         =   3  'DISABLE
         Left            =   1485
         MaxLength       =   10
         TabIndex        =   0
         Top             =   315
         Width           =   1005
      End
      Begin VB.TextBox txtUSRDateCreated 
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
         IMEMode         =   3  'DISABLE
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2040
      End
      Begin VB.Label Label6 
         Caption         =   "Password"
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
         Left            =   225
         TabIndex        =   19
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "Group"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "User ID"
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
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "User Name"
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
         Left            =   225
         TabIndex        =   16
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Date Created"
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
         Left            =   4200
         TabIndex        =   15
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Status"
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
         Left            =   4200
         TabIndex        =   14
         Top             =   1440
         Width           =   1005
      End
   End
   Begin MSComctlLib.TabStrip TabUser 
      Height          =   7455
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   13150
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "USER"
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
      Caption         =   "Security Maintenance - User"
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
      Width           =   9525
   End
End
Attribute VB_Name = "frmUserMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean

Private Sub EntriesEnable(status As Boolean)
    Dim listloop As Integer
        txtUSRCode.Enabled = status
        txtUSRName.Enabled = status
        txtUSRPassword.Enabled = status
        txtGRPCode.Enabled = status
        txtUSRDateCreated.Enabled = status
        txtUSRStatus.Enabled = status
    
   ' For listloop = 1 To ACS.Count
   '     ACS(listloop).Enabled = status
   ' Next listloop
    
End Sub

Private Sub USRListing()
On Error Resume Next
    Dim GRPrst As New ADODB.Recordset
    Dim UserMaster As ListItem
    
    lstUsersList.ListItems.Clear
    GRPrst.Open "USR", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        Do Until GRPrst.EOF
            Set UserMaster = lstUsersList.ListItems.Add(, , GRPrst!USRCode)
                UserMaster.SubItems(1) = Trim(GRPrst!USRName)
                UserMaster.SubItems(2) = GRPrst!GrpCode
                UserMaster.SubItems(3) = GRPrst!USRStatus
                GRPrst.MoveNext
        Loop
    GRPrst.Close
End Sub

Private Sub ComboListing()
    Dim GRP As New ADODB.Recordset
  
    GRP.Open "GRP", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
    Do Until GRP.EOF
        txtGRPCode.AddItem GRP!GrpCode & " - " & Trim(GRP!GRPDescription)
        GRP.MoveNext
    Loop
    
    txtUSRStatus.AddItem "A"
    txtUSRStatus.AddItem "S"
    GRP.Close
    
End Sub

Private Sub cmdAdd_Click()
    EditFlag = False
    Call ClearForm(Me)
    Call DisableCmdButton(Me)
    Call EntriesEnable(True)
    txtUSRDateCreated = Date
    txtUSRStatus = "A"
    txtUSRCode.SetFocus
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
End Sub

Private Sub cmdDelete_Click()
    Dim USRrst As New ADODB.Recordset
    Dim strsql As String
    Dim RecordSaved
    
    EditFlag = False
    cmdDelete.Enabled = False
    cmdSave.Enabled = False

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

      RecordSaved = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            strsql = "Delete From USR Where USRCode = '" & txtUSRCode & "'"
            USRrst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
            cmdSave.Enabled = True
            Call ClearForm(Me)
            Call USRListing
            Call EnableCmdButton(Me)
           ' Call EntriesEnable(True)
            'txtUSRCode.SetFocus
        End If
    'End If
    
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    cmdDelete.Enabled = True
    cmdSave.Enabled = True

End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Call EntriesEnable(True)
    txtUSRCode.SetFocus
    cmdSave.Enabled = True
    
    If txtUSRCode <> "" Then
       txtUSRCode.Enabled = False
       txtUSRName.SetFocus
    End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
   On Error Resume Next
    Dim RecordSaved
    Dim USR As New ADODB.Recordset
    Dim SaveUSR As New ADODB.Recordset
    Dim strsql As String
    Dim ListCheckBox As Integer
    Dim LoopCount As Integer
    Dim AccessString As String
    
    cmdDelete.Enabled = False
    cmdSave.Enabled = False
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If txtUSRCode = "" Then
        MsgBox "Please Enter User Code", vbOKOnly + vbCritical, DialogBoxTitle
        txtUSRCode.SetFocus
    ElseIf txtUSRName = "" Then
        MsgBox "Please Enter User Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtUSRName.SetFocus
    ElseIf txtUSRPassword = "" Then
        MsgBox "Please Enter Password", vbOKOnly + vbCritical, DialogBoxTitle
        txtUSRPassword.SetFocus
    'ElseIf txtDPTCode = "" Then
    '    MsgBox "Please Enter Department", vbOKOnly + vbCritical, DialogBoxTitle
    '    txtDPTCode.SetFocus
    ElseIf txtGRPCode = "" Then
        MsgBox "Please Enter Group", vbOKOnly + vbCritical, DialogBoxTitle
        txtGRPCode.SetFocus
    ElseIf txtUSRStatus = "" Then
        MsgBox "Please Enter Status", vbOKOnly + vbCritical, DialogBoxTitle
        'txtUSRStatus.SetFocus
    Else
        strsql = "Select * From USR Where USRCode = '" & Trim(txtUSRCode) & "'"
        USR.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If USR.recordCount > 0 And EditFlag = False Then
            MsgBox "Record Duplicated! Please use edit button to commit the changes", vbOKOnly + vbCritical, DialogBoxTitle
            USR.Close
            cmdEdit.Enabled = True
            Call ClearForm(Me)
            Set USR = Nothing
            Set SaveUSR = Nothing
            Exit Sub
        Else
            RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If EditFlag = False Then
                    USR.Open "USR", medixmasterconn, adOpenKeyset, adLockOptimistic
                    USR.AddNew
                Else
                    strsql = "Select * From USR Where USRCode = '" & Trim(txtUSRCode) & "'"
                    USR.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                End If
                
               ' ListCheckBox = ACS.Count
               '
               ' For LoopCount = 1 To ListCheckBox
               '     AccessString = AccessString + CStr(ACS(LoopCount))
               ' Next LoopCount

                With USR
                    !USRCode = txtUSRCode
                    !USRName = txtUSRName
                    !USRPassword = txtUSRPassword
                    !USRDateCreated = txtUSRDateCreated
                    !USRStatus = txtUSRStatus
                    '!DPTCode = Trim(Mid(txtDPTCode, 1, 3))
                    !GrpCode = Trim(Mid(txtGRPCode, 1, 3))
                    '!USRAccess = AccessString
                End With
                USR.Update
                Call ClearForm(Me)
                If EditFlag = False Then
                    MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
                Else
                    MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                txtUSRCode.SetFocus
                USR.Close
                Call USRListing
            End If
            EditFlag = False
        End If
    End If
    
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    cmdDelete.Enabled = True
    cmdSave.Enabled = True
    
End Sub

Private Sub Form_Load()
    'Frame3.Width = Screen.Width
    txtUSRDateCreated = Date
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Call ComboListing
    Call USRListing
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub lstUsersList_Click()
On Error Resume Next
    Dim strsql As String
    Dim USR As New ADODB.Recordset
    Dim listloop As Integer
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    strsql = "SELECT USR.*, GRP.GRPDescription "
    strsql = strsql & "FROM USR INNER JOIN GRP ON USR.GRPCode = GRP.GRPCode "
    'strsql = strsql & "GRP ON USR.GRPCode = GRP.GRPCode "
    strsql = strsql & "Where USRCode ='" & Trim(lstUsersList.SelectedItem.Text) & "'"
    USR.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        With USR
            txtUSRCode = !USRCode
            txtUSRName = !USRName
            txtUSRPassword = !USRPassword
            txtUSRDateCreated = !USRDateCreated
            txtUSRStatus = !USRStatus
     '       txtDPTCode = !DPTCode & " - " & Trim(!DPTDescription)
            txtGRPCode = !GrpCode & " - " & Trim(!GRPDescription)
        End With
        
    'For listloop = 1 To Len(USR!USRAccess)
    '    If Mid(USR!USRAccess, listloop, 1) = 1 Then
    '        ACS(listloop).Value = 1
    '    Else
    '        ACS(listloop).Value = 0
    '    End If
    'Next listloop
    
    Call EnableCmdButton(Me)
    cmdDelete.Enabled = True
    Call EntriesEnable(False)
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub



Private Sub txtUSRStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtUSRStatus.Text, txtUSRStatus.SelStart) + Chr(KeyAscii) + Mid(txtUSRStatus.Text, txtUSRStatus.SelStart + txtUSRStatus.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To txtUSRStatus.ListCount - 1
        
        If NewText = LCase(Left(txtUSRStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = txtUSRStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
            If ValidCount = 1 Then
                txtUSRStatus.Text = ValidValue
                txtUSRStatus.SelStart = 0
                txtUSRStatus.SelLength = Len(ValidValue)
            End If
        
        End If

End Sub

Private Sub txtGRPCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtGRPCode.Text, txtGRPCode.SelStart) + Chr(KeyAscii) + Mid(txtGRPCode.Text, txtGRPCode.SelStart + txtGRPCode.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To txtGRPCode.ListCount - 1
        
        If NewText = LCase(Left(txtGRPCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = txtGRPCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
            If ValidCount = 1 Then
                txtGRPCode.Text = ValidValue
                txtGRPCode.SelStart = 0
                txtGRPCode.SelLength = Len(ValidValue)
            End If
        
        End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
    
End Sub

