VERSION 5.00
Begin VB.Form frmRepudiationSelection 
   Caption         =   "frmRepudiationSelection"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   6975
      Begin VB.TextBox txtunderlying 
         Height          =   315
         Left            =   1200
         TabIndex        =   20
         Top             =   1680
         Width           =   5400
      End
      Begin VB.TextBox txtRemarks 
         Height          =   795
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   5400
      End
      Begin VB.TextBox txtAgent 
         Height          =   315
         Left            =   1200
         TabIndex        =   18
         Top             =   240
         Width           =   5400
      End
      Begin VB.ComboBox cboLetterType 
         Height          =   315
         Left            =   5520
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.OptionButton optUnderlying2 
         Caption         =   "No"
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   2880
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.OptionButton optUnderlying1 
         Caption         =   "Yes"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   2880
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.ComboBox cboAttention 
         Height          =   315
         ItemData        =   "frmRepudiationSelection.frx":0000
         Left            =   5760
         List            =   "frmRepudiationSelection.frx":0002
         TabIndex        =   2
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboPersonnel 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   2160
         Visible         =   0   'False
         Width           =   5400
      End
      Begin VB.ComboBox cboDocument 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   2520
         Visible         =   0   'False
         Width           =   5400
      End
      Begin VB.Label Label6 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Agent Name"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Letter Format"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Underlying"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Attention"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Personnel"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Documents"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   6855
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   275
         Left            =   5160
         TabIndex        =   6
         Top             =   150
         Width           =   780
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   275
         Left            =   5925
         TabIndex        =   7
         Top             =   150
         Width           =   800
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Repudiation Categories"
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
      TabIndex        =   8
      Top             =   0
      Width           =   6850
   End
End
Attribute VB_Name = "frmRepudiationSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Source As Integer
Public Document As String
Public Personnel As String
Public Attention As String
Public CASUnderlying As Boolean
Public LetterStatus As String


Private Sub ComboListing()
Dim cmd As New ADODB.Command
Dim REP As New ADODB.Recordset
Dim strsql As String

    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    strsql = "SELECT REPCode, REPDescription FROM REPSelection"
    REP.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockOptimistic
    
    Do Until REP.EOF
    With REP
        If Mid(Trim(!REPCode), 1, 1) = "D" Then
            cboDocument.AddItem !REPCode + " - " + !RepDescription
        ElseIf Mid(Trim(!REPCode), 1, 1) = "P" Then
            cboPersonnel.AddItem !REPCode + " - " + !RepDescription
        ElseIf Mid(Trim(!REPCode), 1, 1) = "A" Then
            cboAttention.AddItem !REPCode + " - " + !RepDescription
        End If
        .MoveNext
    End With
    Loop
    
    'cboLetterType.AddItem "A - Admission"
    'cboLetterType.AddItem "D - Discharge"
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
End Sub

Private Sub cboAttention_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboAttention, cboAttention.SelStart) + Chr(KeyAscii) + Mid(cboAttention, cboAttention.SelStart + cboAttention.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
           
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboAttention = ValidValue
            cboAttention.SelStart = 0
            cboAttention.SelLength = Len(ValidValue)
        End If
    End If

End Sub

Private Sub cboDocument_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDocument, cboDocument.SelStart) + Chr(KeyAscii) + Mid(cboDocument, cboDocument.SelStart + cboDocument.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
        
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboDocument = ValidValue
            cboDocument.SelStart = 0
            cboDocument.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub cboPersonnel_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Clipboard.Clear

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPersonnel, cboPersonnel.SelStart) + Chr(KeyAscii) + Mid(cboPersonnel, cboPersonnel.SelStart + cboPersonnel.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboPersonnel = ValidValue
            cboPersonnel.SelStart = 0
            cboPersonnel.SelLength = Len(ValidValue)
        End If
    End If

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    'If cboDocument = "" Then
     '   MsgBox "Please select Document type", vbOKOnly + vbCritical
      '  cboDocument.SetFocus
    'ElseIf cboPersonnel = "" Then
     '   MsgBox "Please select Personnel type", vbOKOnly + vbCritical
      '  cboPersonnel.SetFocus
    'ElseIf cboAttention = "" Then
    '    MsgBox "Please select Attention type", vbOKOnly + vbCritical
    '    cboAttention.SetFocus
    'Else
        If Source = 1 Then
            'frmCaseAdjustment2016.Document = cboDocument
            'frmCaseAdjustment2016.Personnel = cboPersonnel
            'frmCaseAdjustment2016.Attention = cboAttention
            frmCaseAdjustment2016.tmpAgent = txtAgent
            frmCaseAdjustment2016.repremarks = txtRemarks
            frmCaseAdjustment2016.tmpunderlying = txtunderlying
            
            'If optUnderlying1.Value = True Then
             '   frmCaseAdjustment2016.Underlying = True
              '  frmCaseAdjustment2016.CASUnderlying = "Y"
            'Else
             '   frmCaseAdjustment2016.Underlying = False
              '  frmCaseAdjustment2016.CASUnderlying = "N"
            'End If
            
        ElseIf Source = 2 Then
            frmCaseAdmission.Document = cboDocument
            frmCaseAdmission.Personnel = cboPersonnel
            frmCaseAdmission.Attention = cboAttention
            'If optUnderlying1.Value = True Then
             '   frmCaseAdmission.Underlying = True
              '  frmCaseAdmission.CASUnderlying = "Y"
            'Else
             '   frmCaseAdmission.Underlying = False
              '  frmCaseAdmission.CASUnderlying = "N"
            'End If
        End If
        Unload Me
    'End If
End Sub

Private Sub Form_Load()
    'Call ComboListing
    'Call ChkListing
    'optUnderlying2.value = True
End Sub

Private Sub ChkListing()
Dim i As Integer
    If Document <> "" And Personnel <> "" Then
        For i = 0 To cboDocument.ListCount - 1
            If StrComp(Mid(cboDocument.List(i), 1, InStr(1, cboDocument.List(i), "-") - 2), Document) = 0 Then
                cboDocument.ListIndex = i
            End If
        Next
        For i = 0 To cboPersonnel.ListCount - 1
            If StrComp(Mid(cboPersonnel.List(i), 1, InStr(1, cboPersonnel.List(i), "-") - 2), Personnel) = 0 Then
                cboPersonnel.ListIndex = i
            End If
        Next
        For i = 0 To cboAttention.ListCount - 1
            If StrComp(Mid(cboAttention.List(i), 1, InStr(1, cboAttention.List(i), "-") - 2), Attention) = 0 Then
                cboAttention.ListIndex = i
            End If
        Next
    Else
        cboDocument = ""
        cboPersonnel = ""
        cboAttention = ""
    End If
    
    If CASUnderlying = True Then
        optUnderlying1.Value = True
        optUnderlying2.Value = False
    Else
        optUnderlying1.Value = False
        optUnderlying2.Value = True
    End If
        
End Sub
