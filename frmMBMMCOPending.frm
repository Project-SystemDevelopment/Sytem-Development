VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMMCOPending 
   Caption         =   "Membership MCO Pending"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   7170
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstMCOPending 
      Height          =   3255
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MBMNumber"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "New Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   6975
      Begin VB.ComboBox cboSearchStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5280
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtMBMNumberTO 
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtMBMNumberFR 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtBordxDateFR 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTO 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         Caption         =   "Status"
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
         Left            =   4680
         TabIndex        =   19
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "To"
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
         Left            =   2880
         TabIndex        =   16
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "To"
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
         Left            =   2880
         TabIndex        =   15
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label26 
         Caption         =   "Bordx Date"
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
         TabIndex        =   13
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "MBM Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   6975
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
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
         Left            =   5400
         TabIndex        =   8
         Top             =   210
         Width           =   720
      End
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "P - Pending"
         Top             =   180
         Width           =   1695
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
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
         Left            =   4680
         TabIndex        =   7
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
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
         Left            =   3960
         TabIndex        =   6
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
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
         Left            =   6120
         TabIndex        =   9
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label4 
         Caption         =   "Status"
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
         TabIndex        =   18
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Membership MCO Pending"
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
      TabIndex        =   14
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmMBMMCOPending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim searchCriteria As String

Sub ComboListing()
    cboStatus.AddItem "P" & " - " & "Pending"
    cboStatus.AddItem "R" & " - " & "Realease"
    
    cboSearchStatus.AddItem "B" & " - " & "Billed"
    cboSearchStatus.AddItem "P" & " - " & "Pending"
    cboSearchStatus.AddItem "R" & " - " & "Realease"
End Sub

Private Sub cboSearchStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboSearchStatus.Text, cboSearchStatus.SelStart) + Chr(KeyAscii) + Mid(cboSearchStatus.Text, cboSearchStatus.SelStart + cboSearchStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboSearchStatus.ListCount - 1
 
        If NewText = LCase(Left(cboSearchStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboSearchStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboSearchStatus.Text = ValidValue
                cboSearchStatus.SelStart = 0
                cboSearchStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboStatus.Text, cboStatus.SelStart) + Chr(KeyAscii) + Mid(cboStatus.Text, cboStatus.SelStart + cboStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboStatus.ListCount - 1
 
        If NewText = LCase(Left(cboStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboStatus.Text = ValidValue
                cboStatus.SelStart = 0
                cboStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
    
    If lstMCOPending.listitems.Count = 0 Then
        MsgBox "There is no record on listing!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Call ExportListViewtoExcel(lstMCOPending)
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
    cmdSearch.Enabled = False
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Call MBMInquiryList
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
    cmdSearch.Enabled = True
End Sub


Private Sub MBMInquiryList(Optional str As String)
Dim sqlstr, strCriteria As String

Screen.MousePointer = vbHourglass
CmdExit.Enabled = False
strCriteria = ""
lstMCOPending.listitems.Clear

    If Len(Trim(txtMBMNumberFR)) Then
        strCriteria = strCriteria & " And MBMNumber >= '" & Trim(txtMBMNumberFR) & "' "
    End If
    
    If Len(Trim(txtMBMNumberTO)) Then
         strCriteria = strCriteria & " AND MBMNumber <= '" & Trim(txtMBMNumberTO) & "' "
    End If
    
    If IsDate(txtBordxDateFR) = True Then
      strCriteria = strCriteria + " And MBMBordxDate >= '" & Format(txtBordxDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If IsDate(txtBordxDateTO) = True Then
      strCriteria = strCriteria + " And MBMBordxDate <= '" & Format(txtBordxDateTO, "mm/dd/yyyy") & "'"
    End If
    
    If Len(cboSearchStatus) Then
        strCriteria = strCriteria + " And Status = '" & Mid(cboSearchStatus, 1, 1) & "'"
    End If
    searchCriteria = strCriteria
    Call Listing(strCriteria)
    
CmdExit.Enabled = True
Screen.MousePointer = vbDefault
End Sub
Private Sub Listing(StrAppend As String)
Dim StrSql As String
Dim rst As New ADODB.Recordset
Dim ListMaster As ListItem

    
    StrSql = "Select MBMNumber,MBMBordxDate,MBMNewBordxDate,Status,PendingBordxDate from VIEW_MCOPending where 0 = 0"


    If Len(Trim(StrAppend)) Then
        StrSql = StrSql + StrAppend
    End If

    
    rst.Open StrSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic

    
    If rst.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Do Until rst.EOF
            With rst
                Set ListMaster = lstMCOPending.listitems.Add(, , "")
                ListMaster.SubItems(1) = !MBMNumber
                If Len(!PendingBordxDate) Then
                    ListMaster.SubItems(2) = !PendingBordxDate
                Else
                    ListMaster.SubItems(2) = !MBMBordxDate
                End If
                If Len(!MBMNewBordxDate) Then
                    ListMaster.SubItems(3) = !MBMNewBordxDate
                End If
                If Len(!status) Then
                    ListMaster.SubItems(4) = !status
                End If
            End With
        rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing

End Sub

Private Sub cmdUpdate_Click()
Dim MBMNo As String
Dim BordxDate As String
Dim NewBordxDate As String
Dim status As String
Dim i As Integer
Dim i2 As Integer
Dim ItemChecked As Boolean
Dim Update
    
    
    'TxtWksNo1.SetFocus
    
    ItemChecked = False
     
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

        If cboStatus = "__/__/____" Then
            MsgBox "Please select in status", vbCritical
            cboStatus.SetFocus
        Else
            Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            
            If Update = vbYes Then
                For i = 1 To lstMCOPending.listitems.Count
                    If lstMCOPending.listitems(i).Checked = True Then
                        ItemChecked = True
                    Exit For
                    End If
                Next i
                
            If ItemChecked = True Then
                
                For i = 1 To lstMCOPending.listitems.Count
                   If lstMCOPending.listitems(i).Checked = True Then
                      
                      MBMNo = Trim(lstMCOPending.listitems(i).SubItems(1))
                      BordxDate = Trim(lstMCOPending.listitems(i).SubItems(2))
                      NewBordxDate = Trim(lstMCOPending.listitems(i).SubItems(3))
                      status = Mid(cboStatus, 1, 1)
                      Call UpdateMCOPending(MBMNo, BordxDate, status)
                      
                   End If
                Next i
                
                MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                
                lstMCOPending.listitems.Clear
            
                Call Listing(searchCriteria)
            End If
        End If
    End If
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub UpdateMCOPending(MBMNo As String, BordxDate As String, status As String)
Dim strsqlMBM As String
Dim MBM As New ADODB.Recordset

    strsqlMBM = " Select MBMNumber,MBMBordxDate,status,LastUpdateUser,LastUpdateDate from MCOPending " & _
                " Where MBMNumber = '" & Trim(MBMNo) & "'"
    
    MBM.Open strsqlMBM, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    If MBM.recordCount = 0 Then
        MBM.AddNew
        MBM!MBMBordxDate = BordxDate
        MBM!MBMNumber = MBMNo
        MBM!MBMBordxDate = BordxDate
        MBM!status = status
        MBM!LastUpdateUser = Logon
        MBM!LastUpdateDate = Date
        MBM.Update
    Else
        If status = "R" Then
            MBM.Delete
        End If
    End If
    MBM.Close
    Set MBM = Nothing
End Sub

Private Sub Form_Load()
    Call ComboListing
End Sub

Private Sub lstMCOPending_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        If cboStatus = "" Then
            MsgBox "Please select status", vbCritical
            Item.Checked = False
        ElseIf Trim(Item.SubItems(4)) = "B" Then
            MsgBox "This Bordx has been billed", vbCritical
            Item.Checked = False
        Else
            Item.Checked = True
        End If
    
    End If

End Sub
