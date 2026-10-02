VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGSTItems 
   Caption         =   "frmGSTItems"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   7335
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   4650
         TabIndex        =   10
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   5500
         TabIndex        =   9
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   150
         Width           =   855
      End
      Begin VB.Label lblWSStatus 
         Caption         =   "Label1"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   80
      TabIndex        =   1
      Top             =   240
      Width           =   7335
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   6000
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboGSTItem 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label65 
         Caption         =   "Amount"
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
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label65 
         Caption         =   "GST Items"
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
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstListing 
      Height          =   2415
      Left            =   80
      TabIndex        =   6
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "GST Item"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "INVIndex"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "WSDone"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "GST"
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
      TabIndex        =   0
      Top             =   0
      Width           =   7605
   End
End
Attribute VB_Name = "frmGSTItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboGSTItem_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboGSTItem.Text, cboGSTItem.SelStart) + Chr(KeyAscii) + Mid(cboGSTItem.Text, cboGSTItem.SelStart + cboGSTItem.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboGSTItem.ListCount - 1
 
        If NewText = LCase(Left(cboGSTItem.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboGSTItem.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
                cboGSTItem.Text = ValidValue
                cboGSTItem.SelStart = 0
                cboGSTItem.SelLength = Len(ValidValue)
             End If
        End If

End Sub

Private Sub cmdAdd_Click()
    Dim INVList As ListItem
    
    Set INVList = lstListing.ListItems.Add(, , cboGSTItem)
    
    With INVList
        .SubItems(1) = txtAmount
    End With
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdRemove_Click()
On Error GoTo errorRemove
Dim ForLoop As Integer
Dim ListCount As Integer
Dim strsql, strActual, strNot As String
Dim RecordSaved
Dim RemoveInvoice  As New ADODB.Recordset
Dim RemoveActualInvoice  As New ADODB.Recordset
Dim RemoveNotCover  As New ADODB.Recordset
    
   ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    'If lstListing.SelectedItem.SubItems(2) = False Then
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
          If RecordSaved = vbYes Then
              ListCount = lstListing.ListItems.Count
              For ForLoop = 1 To ListCount
                  If lstListing.ListItems(ForLoop).Selected = True Then
                      lstListing.ListItems.Remove (ForLoop)
                      Exit For
                  End If
              Next ForLoop
              
            'strsql = "delete  from ClaimInvoice where INVIndex =" & intIndex
            'medixmasterconnWS.beginTrans
            'beginTrans = True
            'RemoveInvoice.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            'strActual = "delete from ClaimInvoiceActual where INVIndex =" & intIndex
            'RemoveActualInvoice.Open strActual, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            'strNot = "delete from ClaimInvoiceNotCover where INVIndex =" & intIndex
            'RemoveNotCover.Open strNot, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            
           ' medixmasterconnWS.commitTrans
          End If
          'checking = False
          'Call clearInvoice
          'checking = True
   ' Else
    '    MsgBox "Selected Invoice cannot be deleted!" & vbNewLine & "It has been used in Worksheet!", vbCritical
        
    'End If
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    Exit Sub
errorRemove:
    'If beginTrans = True Then
    '    medixmasterconnWS.RollbackTrans
    'End If
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
Dim rst3 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strsql As String
        
    strsql = "SELECT * from GSTItems "
    rst3.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    Do Until rst3.EOF
        With rst3
            cboGSTItem.AddItem !GSTName
            rst3.MoveNext
        End With
    Loop
    rst3.Close
    Set rst3 = Nothing

End Sub



Public Function KeyNumber(KeyAscii As Integer) As Integer
    If KeyAscii > 26 Then
       If InStr("-.0123456789", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    KeyNumber = KeyAscii
End Function


Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
