VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmITDevelopment 
   Caption         =   "IT Development Activities"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   9510
   Begin VB.Frame FrameInfo 
      Height          =   3255
      Left            =   80
      TabIndex        =   13
      Top             =   240
      Width           =   9375
      Begin VB.ComboBox cboJobitems 
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
         ItemData        =   "frmITDevelopment.frx":0000
         Left            =   1200
         List            =   "frmITDevelopment.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   8055
      End
      Begin VB.ComboBox cboStatus 
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
         Left            =   1200
         TabIndex        =   1
         Top             =   520
         Width           =   3495
      End
      Begin MSMask.MaskEdBox txtStartdate 
         Height          =   300
         Left            =   1200
         TabIndex        =   3
         Top             =   810
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEnddate 
         Height          =   300
         Left            =   3600
         TabIndex        =   4
         Top             =   810
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboUser 
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
         Left            =   5880
         TabIndex        =   2
         Top             =   530
         Width           =   3375
      End
      Begin MSMask.MaskEdBox txtDateLine 
         Height          =   300
         Left            =   5880
         TabIndex        =   5
         Top             =   810
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtAdjFrmNo 
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
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   6
         Top             =   810
         Width           =   1455
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   1200
         MaxLength       =   3800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1080
         Width           =   8055
      End
      Begin VB.Label Label1 
         Caption         =   "Target date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4800
         TabIndex        =   26
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "IT Personnel"
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
         Index           =   1
         Left            =   4800
         TabIndex        =   23
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Completed Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Start Date"
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
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Date of Admission"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
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
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Form No"
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
         Left            =   7080
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Job Category"
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
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lsItemtLisitng 
      Height          =   2745
      Left            =   75
      TabIndex        =   22
      Top             =   3555
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4842
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Job Item"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Start Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "End Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date Line"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Adj Form No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "IT Personnel"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   80
      TabIndex        =   18
      Top             =   6240
      Width           =   9375
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
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
         Left            =   7800
         TabIndex        =   24
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
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
         Left            =   7080
         TabIndex        =   10
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   5640
         TabIndex        =   9
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
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
         Left            =   4920
         TabIndex        =   8
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
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
         Left            =   8520
         TabIndex        =   11
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Left            =   6360
         TabIndex        =   19
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmpUpdate 
         Caption         =   "Update"
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
         Left            =   5640
         TabIndex        =   25
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "IT Development Activities"
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
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10485
   End
End
Attribute VB_Name = "frmITDevelopment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpItemNo As String

Private Sub combolist()
Dim rstItem As New ADODB.Recordset
Dim USR As New ADODB.Recordset

    Screen.MousePointer = vbHourglass
    cboJobitems.Clear
    cboStatus.Clear
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    rstItem.Open "select ITCatDesc from ITCategory", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    

    Do Until rstItem.EOF
         cboJobitems.AddItem rstItem!ITCatDesc
         rstItem.MoveNext
    Loop
    
    rstItem.Close
    Set rstItem = Nothing
    
    
    USR.Open "select USRName from USR where GRPCode = 'ITD' And USRStatus = 'A'", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    Do Until USR.EOF
        cboUser.AddItem USR!USRName
        USR.MoveNext
    Loop
    USR.Close
    Set USR = Nothing
    
    cboStatus.AddItem "No Started"
    cboStatus.AddItem "In Progress"
    cboStatus.AddItem "Completed"
    

    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    Screen.MousePointer = Default
End Sub

Private Sub cboJobitems_KeyPress(KeyAscii As Integer)
   Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboJobitems.Text, cboJobitems.SelStart) & _
                Chr(KeyAscii) + Mid(cboJobitems.Text, cboJobitems.SelStart + cboJobitems.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboJobitems.ListCount - 1
            
            If NewText = ChkStr(Left(cboJobitems.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboJobitems.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboJobitems.Text = ValidValue
              cboJobitems.SelStart = 0
              cboJobitems.SelLength = Len(ValidValue)
            End If
            

End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
   Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboStatus.Text, cboStatus.SelStart) & _
                Chr(KeyAscii) + Mid(cboStatus.Text, cboStatus.SelStart + cboStatus.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboStatus.ListCount - 1
            
            If NewText = ChkStr(Left(cboStatus.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboStatus.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboStatus.Text = ValidValue
              cboStatus.SelStart = 0
              cboStatus.SelLength = Len(ValidValue)
            End If

End Sub

Private Sub cboUser_KeyPress(KeyAscii As Integer)
   Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii > 32 And KeyAscii < 127 Then
        NewText = ChkStr(Left(cboUser.Text, cboUser.SelStart) & _
                Chr(KeyAscii) + Mid(cboUser.Text, cboUser.SelStart + cboUser.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboUser.ListCount - 1
            
            If NewText = ChkStr(Left(cboUser.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboUser.List(i)
            End If
        Next
            If ValidCount <= 1 Then KeyAscii = 0
            End If
            If ValidCount = 1 Then
              cboUser.Text = ValidValue
              cboUser.SelStart = 0
              cboUser.SelLength = Len(ValidValue)
            End If

End Sub

Private Sub CmdAdd_Click()
    Call ClearItems
    cmdEdit.Visible = False
    cmpUpdate.Visible = True
    cmpUpdate.Caption = "Save"
End Sub

Private Sub cmdClear_Click()
    Call ClearItems
    cmdEdit.Visible = True
    cmpUpdate.Visible = False
End Sub

Private Sub cmdDelete_Click()
Dim ASKMsg
    
    If tmpItemNo = "" Then
        MsgBox "Please Select Item to be deleted", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        ASKMsg = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
        If ASKMsg = vbYes Then
            Call DeleteItem
            Call SearchRecord
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    cmdEdit.Visible = False
    cmpUpdate.Visible = True
    cmpUpdate.Caption = "Update"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Call SearchRecord
End Sub

Private Sub cmpUpdate_Click()
    If cboJobitems = "" Then
        MsgBox "Please Enter Job Category", vbOKOnly + vbCritical, DialogBoxTitle
        cboJobitems.SetFocus
    ElseIf cboStatus = "" Then
        MsgBox "Please Enter Status", vbOKOnly + vbCritical, DialogBoxTitle
        cboStatus.SetFocus
    ElseIf txtStartdate = "__/__/____" Then
        MsgBox "Please Enter Start Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtStartdate.SetFocus
    ElseIf Trim(cboStatus) = "Completed" And txtEnddate = "__/__/____" Then
        MsgBox "Please Enter End Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtEnddate.SetFocus
    ElseIf txtDateLine = "__/__/____" Then
        MsgBox "Please Enter Target Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtStartdate.SetFocus
    ElseIf cboUser = "" Then
        MsgBox "Please Enter Person Incharge", vbOKOnly + vbCritical, DialogBoxTitle
        cboUser.SetFocus
    Else
        Call SaveData
        cmdEdit.Visible = True
        cmpUpdate.Visible = False
    End If
    
End Sub

Private Sub Form_Load()
    Call combolist
End Sub

Private Sub ClearItems()
    cboJobitems = ""
    cboStatus = ""
    txtStartdate.mask = ""
    txtStartdate.Text = ""
    txtStartdate.mask = "##/##/####"
    txtEnddate.mask = ""
    txtEnddate.Text = ""
    txtEnddate.mask = "##/##/####"
    txtDateLine.mask = ""
    txtDateLine.Text = ""
    txtDateLine.mask = "##/##/####"
    txtAdjFrmNo = ""
    cboUser = ""
    txtRemark = ""
End Sub
Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
    Frame2.Enabled = False
    strAppend = ""
    
    If Len(Trim(cboJobitems)) Then
        strAppend = strAppend + " AND ITCategory = '" & Trim(cboJobitems) & "'"
    End If

    If Len(Trim(cboStatus)) Then
        strAppend = strAppend + " AND ITStatus = '" & Trim(cboStatus) & "'"
    End If
  
    If Len(Trim(txtStartdate)) > 0 And IsDate(txtStartdate) = True Then
        strAppend = strAppend + " And ITStartDate >= '" & Format(txtStartdate, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtEnddate)) > 0 And IsDate(txtEnddate) = True Then
        strAppend = strAppend + " And ITEndDate <= '" & Format(txtEnddate, "mm/dd/yyyy") & "'"
    End If
       
    
    If Len(Trim(txtAdjFrmNo)) Then
        strAppend = strAppend + " AND ITAdjFormNo = '" & Trim(txtAdjFrmNo) & "'"
    End If
    
    If Len(Trim(cboUser)) Then
        strAppend = strAppend + " AND ITPersonnel = '" & Trim(cboUser) & "'"
    End If
      
    Call ItemListing(strAppend)
    Frame2.Enabled = True
End Function

Private Sub ItemListing(Optional strAppend As String)
Dim strsql As String
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    lsItemtLisitng.ListItems.Clear
        strsql = " SELECT * From ITDevelopment where 0 = 0"
        strsql = strsql + strAppend
        rst2.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic

        Do Until rst2.EOF
            With rst2
                Set Record = lsItemtLisitng.ListItems.Add(, , !ItemNo)
                
                If Len(!ITCategory) Then
                    Record.SubItems(1) = Trim(!ITCategory)
                End If
                If Len(!ITStatus) Then
                    Record.SubItems(2) = !ITStatus
                End If
                If Len(!ITStartDate) Then
                    Record.SubItems(3) = !ITStartDate
                End If
                
                If Len(!ITEndDate) Then
                    Record.SubItems(4) = !ITEndDate
                End If
                If Len(!ITdateLine) Then
                    Record.SubItems(5) = !ITdateLine
                End If
                
                If Len(!ITAdjFormNo) Then
                    Record.SubItems(6) = !ITAdjFormNo
                End If
                If Len(!ITPersonnel) Then
                    Record.SubItems(7) = !ITPersonnel
                End If
                
                If Len(!ITRemarks) Then
                    Record.SubItems(8) = !ITRemarks
                End If
                
            End With
            rst2.MoveNext
        Loop
    
        rst2.Close
        Set rst2 = Nothing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
    
    If lsItemtLisitng.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub lsItemtLisitng_Click()
    Call ListItems
End Sub

Private Sub SaveData()
Dim strsql As String
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    strsql = ""
    If cmpUpdate.Caption = "Update" Then
        If txtEnddate <> "__/__/____" Then
            strsql = "Update ITDevelopment set ITCategory = '" & Trim(cboJobitems) & "', ITStatus = '" & Trim(cboStatus) & _
                        "', ITStartDate = '" & Format(txtStartdate, "yyyy/mm/dd") & "', ITEndDate ='" & Format(txtEnddate, "yyyy/mm/dd") & "'" & _
                        ", ITdateLine = '" & Format(txtDateLine, "yyyy/mm/dd") & "' , ITAdjFormNo = '" & txtAdjFrmNo & "', ITPersonnel ='" & cboUser & "' , ITRemarks ='" & txtRemark & "'" & _
                        " Where ItemNo = '" & Trim(tmpItemNo) & "'"
            medixmasterconnMaint.Execute strsql
        Else
            strsql = "Update ITDevelopment set ITCategory = '" & Trim(cboJobitems) & "', ITStatus = '" & Trim(cboStatus) & _
                        "', ITStartDate = '" & Format(txtStartdate, "yyyy/mm/dd") & "' " & _
                        ", ITdateLine = '" & Format(txtDateLine, "yyyy/mm/dd") & "', ITAdjFormNo = '" & txtAdjFrmNo & "', ITPersonnel ='" & cboUser & "' , ITRemarks ='" & txtRemark & "'" & _
                        " Where ItemNo = '" & Trim(tmpItemNo) & "'"
            medixmasterconnMaint.Execute strsql
        End If
    Else
        If txtEnddate <> "__/__/____" Then
            strsql = "Insert Into ITDevelopment ( ITCategory, ITStatus,ITStartDate, ITEndDate, " & _
                        " ITdateLine, ITAdjFormNo ,ITPersonnel , ITRemarks ) " & _
                        " Values ('" & Trim(cboJobitems) & "', '" & Trim(cboStatus) & _
                        "', '" & Format(txtStartdate, "yyyy/mm/dd") & "', '" & Format(txtEnddate, "yyyy/mm/dd") & "'" & _
                        ",'" & Format(txtDateLine, "yyyy/mm/dd") & "', '" & txtAdjFrmNo & "', '" & cboUser & "' , '" & txtRemark & "')"
            medixmasterconnMaint.Execute strsql
        Else
            strsql = "Insert Into ITDevelopment ( ITCategory, ITStatus,ITStartDate," & _
                        " ITdateLine, ITAdjFormNo ,ITPersonnel , ITRemarks ) " & _
                        " Values ('" & Trim(cboJobitems) & "', '" & Trim(cboStatus) & _
                        "', '" & Format(txtStartdate, "yyyy/mm/dd") & "'" & _
                        ",'" & Format(txtDateLine, "YYYY/MM/DD") & "', '" & txtAdjFrmNo & "', '" & cboUser & "' , '" & txtRemark & "')"
            medixmasterconnMaint.Execute strsql
        End If
    End If
    If cmpUpdate.Caption = "Update" Then
        MsgBox "Record Updated Successfully!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        MsgBox "Record Saved Successfully!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    Call SearchRecord
End Sub

Private Sub ListItems()
    Call ClearItems
    cmdEdit.Visible = True
    cmpUpdate.Visible = False
    If lsItemtLisitng.ListItems.Count > 0 Then
        tmpItemNo = lsItemtLisitng.SelectedItem.Text
        cboJobitems = lsItemtLisitng.SelectedItem.SubItems(1)
        cboStatus = lsItemtLisitng.SelectedItem.SubItems(2)
        txtStartdate = lsItemtLisitng.SelectedItem.SubItems(3)
        If lsItemtLisitng.SelectedItem.SubItems(4) <> "" Then
            txtEnddate = lsItemtLisitng.SelectedItem.SubItems(4)
        End If
        If lsItemtLisitng.SelectedItem.SubItems(5) <> "" Then
            txtDateLine = lsItemtLisitng.SelectedItem.SubItems(5)
        End If
        
        txtAdjFrmNo = lsItemtLisitng.SelectedItem.SubItems(6)
        cboUser = lsItemtLisitng.SelectedItem.SubItems(7)
        txtRemark = lsItemtLisitng.SelectedItem.SubItems(8)
    End If
End Sub

Private Sub lsItemtLisitng_KeyDown(KeyCode As Integer, Shift As Integer)
    Call ListItems
End Sub

Private Sub lsItemtLisitng_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ListItems
End Sub

Private Sub DeleteItem()
Dim strsql As String
Dim rst As New ADODB.Recordset
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

        strsql = "Delete from ITDevelopment Where ItemNo = '" & tmpItemNo & "'"
        rst.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    MsgBox "Record Deleted Successfully!", vbOKOnly + vbCritical, DialogBoxTitle
End Sub

Private Sub txtDateLine_LostFocus()
    If IsDate(txtDateLine) Then
    Else
        If txtDateLine <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateLine.SetFocus
        End If
    End If

End Sub

Private Sub txtEnddate_LostFocus()
    If IsDate(txtEnddate) Then
    Else
        If txtEnddate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtEnddate.SetFocus
        End If
    End If
End Sub

Private Sub txtStartdate_LostFocus()
    If IsDate(txtStartdate) Then
    Else
        If txtStartdate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtStartdate.SetFocus
        End If
    End If
End Sub
