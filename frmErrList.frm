VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmErrList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Listing Information"
   ClientHeight    =   6450
   ClientLeft      =   945
   ClientTop       =   1440
   ClientWidth     =   10065
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   10065
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   90
      TabIndex        =   6
      Top             =   7320
      Width           =   11730
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
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
         Left            =   10230
         TabIndex        =   8
         Top             =   220
         Width           =   960
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
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
         Left            =   9240
         TabIndex        =   7
         Top             =   220
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
      Height          =   615
      Left            =   90
      TabIndex        =   2
      Top             =   450
      Width           =   11775
      Begin VB.ComboBox txtSearch 
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
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   2505
      End
      Begin VB.Label Label6 
         Caption         =   "Search For Filename"
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
         Left            =   375
         TabIndex        =   4
         Top             =   270
         Width           =   1530
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11865
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Error Listing Information"
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
         Left            =   225
         TabIndex        =   1
         Top             =   45
         Width           =   6615
      End
   End
   Begin MSComctlLib.ListView lstERRList 
      Height          =   6135
      Left            =   90
      TabIndex        =   5
      Top             =   1125
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   10821
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Record Line"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Policy No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Error Message"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Filename"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmErrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DisplayList()
    On Error Resume Next
    Dim ERR As New ADODB.Recordset
    Dim sqlstr As String
    Dim ErrList As ListItem
    
    If txtSearch.Text = "" Then
       MsgBox "Please Enter File Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtSearch.SetFocus
    Else
       sqlstr = "Select * From ERR Where ERRFilename LIKE '%" & txtSearch.Text & "%' Order By ErrIndex"
       ERR.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
       lstERRList.listitems.Clear

    If ERR.recordCount = 0 Then
       MsgBox "No Record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Do Until ERR.EOF
            Set ErrList = lstERRList.listitems.Add(, , ERR!ERRRecordLine)
                ErrList.SubItems(1) = ERR!ERRPolicyNo
                ErrList.SubItems(2) = ERR!ERRName
                ErrList.SubItems(3) = ERR!ERRMessage
                ErrList.SubItems(4) = ERR!ErrFileName
                ERR.MoveNext
        Loop
        End If
        ERR.Close
        txtSearch.SetFocus
End If
End Sub

Private Sub cmdExit_Click()
'    Call EnableListBar(True)
'    Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
        'bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdSearch_Click()
Call DisplayList
End Sub

Private Sub Form_Load()
    'Frame1.Width = Screen.Width - 1800
    'Frame3.Width = Screen.Width
    'lstERRList.Width = Screen.Width - 2000
    'lstERRList.Height = Screen.Height - 4400
    Call ComboListing
    'Call ErrorListing
    'Call CenterForm(Me)
End Sub

Private Sub ErrorListing()
On Error Resume Next
    Dim rst2 As New ADODB.Recordset
    Dim ErrList As ListItem
    
    lstERRList.listitems.Clear
    rst2.Open "ERR", medixmasterconn, adOpenKeyset, adLockOptimistic
        Do Until rst2.EOF
            Set ErrList = lstERRList.listitems.Add(, , rst2!ERRRecordLine)
                ErrList.SubItems(1) = rst2!ERRPolicyNo
                ErrList.SubItems(2) = rst2!ERRName
                ErrList.SubItems(3) = rst2!ERRMessage
                ErrList.SubItems(4) = rst2!ErrFileName
                rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
Private Sub lstERRList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   lstERRList.SortKey = ColumnHeader.Index - 1
   lstERRList.Sorted = True

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(txtSearch.Text, txtSearch.SelStart) + Chr(KeyAscii) + Mid(txtSearch.Text, txtSearch.SelStart + txtSearch.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To txtSearch.ListCount - 1
        
        If NewText = LCase(Left(txtSearch.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = txtSearch.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
        
            If ValidCount = 1 Then
                txtSearch.Text = ValidValue
                txtSearch.SelStart = 0
                txtSearch.SelLength = Len(ValidValue)
            End If
        
        End If

If KeyAscii = 13 Then
    Call cmdSearch_Click
End If

End Sub

Private Sub ComboListing()
    Dim ERR As New ADODB.Recordset
    Dim strsql As String
    
    strsql = "Select DISTINCT ERRFileName From ERR"
    
    ERR.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until ERR.EOF
        txtSearch.AddItem ERR!ErrFileName
        ERR.MoveNext
    Loop
    
    ERR.Close
    Set ERR = Nothing
    
End Sub


