VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMemberEntry 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Registration"
   ClientHeight    =   2130
   ClientLeft      =   2940
   ClientTop       =   3690
   ClientWidth     =   6780
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   6780
   Begin VB.Frame Frame2 
      Height          =   1875
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   6810
      Begin VB.ComboBox cboHLTCode 
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
         ItemData        =   "frmMemberEntry.frx":0000
         Left            =   2205
         List            =   "frmMemberEntry.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   130
         Width           =   4305
      End
      Begin VB.ComboBox cboPAYCode 
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
         Left            =   2205
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   400
         Width           =   4305
      End
      Begin MSMask.MaskEdBox txtMBMBordxDate 
         Height          =   285
         Left            =   2205
         TabIndex        =   2
         Top             =   680
         Width           =   1560
         _ExtentX        =   2752
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
      Begin VB.TextBox txtLastBatch 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         TabIndex        =   11
         Text            =   "0"
         Top             =   920
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
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
         Left            =   4680
         TabIndex        =   4
         Top             =   1320
         Width           =   960
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
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
         Left            =   5640
         TabIndex        =   5
         Top             =   1320
         Width           =   960
      End
      Begin VB.TextBox txtMBMBatchNo 
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
         Left            =   2205
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "1"
         Top             =   920
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Last Batch No"
         Height          =   255
         Left            =   4680
         TabIndex        =   12
         Top             =   935
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
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
         Left            =   180
         TabIndex        =   10
         Top             =   430
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Batch No"
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
         Left            =   180
         TabIndex        =   9
         Top             =   935
         Width           =   1140
      End
      Begin VB.Label Label17 
         Caption         =   "Bordx. Date (dd/mm/yyyy)"
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
         Left            =   180
         TabIndex        =   8
         Top             =   690
         Width           =   2265
      End
      Begin VB.Label Label4 
         Caption         =   "Health Type"
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
         Left            =   180
         TabIndex        =   7
         Top             =   160
         Width           =   1680
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Membership Header"
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
      Width           =   6855
   End
End
Attribute VB_Name = "frmMemberEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub Form_Load()
    Me.Width = 7000
    Call ComboListing
End Sub

Private Sub ComboListing()
    Dim HLT As New ADODB.Recordset
    Dim cmd As New ADODB.Command

    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHLTType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HLT.CursorLocation = adUseClient
    HLT.CursorType = adOpenForwardOnly
    HLT.CacheSize = 100
    Set HLT = cmd.Execute
    
    Do Until HLT.EOF
        With HLT
            cboHLTCode.AddItem ChkStr(!HLTCode) & " - " & ChkStr(!HLTDescription)
            HLT.MoveNext
        End With
    Loop
    HLT.Close
    Set HLT = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing

    cboHLTCode.ListIndex = 1
    
    
End Sub

Private Sub CmdNext_Click()
On Error GoTo ErrorDone
    If Len(Trim(cboHLTCode)) = 0 Then
        MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboHLTCode.SetFocus
    ElseIf Len(Trim(cboPAYCode)) = 0 Then
        MsgBox "Please Enter Payor Code", vbOKOnly + vbCritical, DialogBoxTitle
        cboPAYCode.SetFocus
    ElseIf Len(Trim(txtMBMBordxDate)) = 0 Then
        MsgBox "Please Enter Bordx Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMBordxDate.SetFocus
    ElseIf Len(Trim(txtMBMBatchNo)) = 0 Then
        MsgBox "Please Enter Batch No", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMBatchNo.SetFocus
    ElseIf IsDate(Trim(txtMBMBordxDate)) = False Then
        MsgBox "Please Enter a Valid Bordx Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMBordxDate.SetFocus
    Else
        
'        If FormLoadedByName("frmNewMemberRegistration") = True Then
        If FormLoadedByName("frmNewMBMReg05") = True Then
'            frmNewMemberRegistration.SetFocus
            frmNewMBMReg05.SetFocus
        Else
'            frmNewMemberRegistration.show
            frmNewMBMReg05.Show
        End If
        
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        frmNewMBMReg05.txtHLTCode = ChkStr(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1))
        frmNewMBMReg05.txtPAYCode = ChkStr(Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1))
        frmNewMBMReg05.txtMBMBordxDate = txtMBMBordxDate
        frmNewMBMReg05.txtMBMBatchNo = txtMBMBatchNo
        frmNewMBMReg05.txtTempBatch = txtMBMBatchNo
        frmNewMBMReg05.txtTempBordx = txtMBMBordxDate
        
        'medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        Unload frmMemberEntry
        Set frmMemberEntry = Nothing
    End If
    Exit Sub
ErrorDone:
    MsgBox Err.Description
End Sub



Private Sub cboHLTCode_Click()
    Dim PAY As New ADODB.Recordset
    Dim strsql As String
    cboPAYCode.Clear
    
    Dim strAppend As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    
    strAppend = strAppend & "And HLTCOde = '" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayorHLTType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 255, strAppend)
    cmd.Parameters.Append Prm1
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    
    cboPAYCode.Clear
    Do Until PAY.EOF
       cboPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub cboHLTCode_Change()
    Dim PAY As New ADODB.Recordset
    Dim strAppend As String
    Dim cmd As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    
    strAppend = strAppend & "where HLTCOde = '" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 255, strAppend)
    cmd.Parameters.Append Prm1
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    
    Do Until PAY.EOF
       cboPAYCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
       PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
End Sub



Private Sub txtMBMBordxDate_LostFocus()
    Dim BAT As New ADODB.Recordset
    Dim strsql As String
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If IsDate(txtMBMBordxDate) Then
        strsql = "Select  BATSequence  , bordxType From BAT " & _
                " where BATBordxDate = '" & Format(txtMBMBordxDate, "mm/dd/yyyy") & "'"
        
        If Len(Trim(cboPAYCode)) Then
            strsql = strsql + " AND PAYCode ='" & Trim(Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1)) & "'"
        End If
        
        BAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If BAT.recordCount Then
            Do While Not BAT.EOF
                If BAT!BordxType = "N" Then
                    txtLastBatch = BAT!BATSequence
                End If
                BAT.MoveNext
            Loop
        Else
            txtLastBatch = 0
        End If
        BAT.Close
        Set BAT = Nothing
    Else
        If txtMBMBordxDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! Please key in 'dd/mm/yyyy' format!", vbCritical
            txtMBMBordxDate.SetFocus
            txtMBMBordxDate.SelStart = 0
        End If
    End If
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub


Private Sub cmdExit_Click()
'    Call EnableListBar(True)
    Unload Me
    Set frmMemberEntry = Nothing
End Sub

Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
'    If KeyAscii >= 32 And KeyAscii <> 127 Then
'    NewText = LCase(Left(cboHLTCode.text, cboHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode.text, cboHLTCode.SelStart + cboHLTCode.SelLength + 1))
    
'    ValidCount = 0
'    ValidValue = ""
    
'    For i = 0 To cboHLTCode.ListCount - 1
        
'        If NewText = LCase(Left(cboHLTCode.list(i), Len(NewText))) Then
'                ValidCount = ValidCount + 1
'                ValidValue = cboHLTCode.list(i)
'        End If
'        Next
'        If ValidCount <= 1 Then KeyAscii = 0
'            If ValidCount = 1 Then
 '              cboHLTCode.text = ValidValue
 '              cboHLTCode.SelStart = 0
 '              cboHLTCode.SelLength = Len(ValidValue)
 '           End If
 '       End If

' ENable the CAP button
'GetKeyboardState kbArray
'kbArray.kbByte(VK_CAPITAL) = 1
'SetKeyboardState kbArray
End Sub

Private Sub cboPAYCode_KeyPress(KeyAscii As Integer)
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer

'    If KeyAscii >= 32 And KeyAscii <> 127 Then
'    NewText = LCase(Left(cboPAYCode, cboPAYCode.SelStart) + Chr(KeyAscii) + Mid(cboPAYCode, cboPAYCode.SelStart + cboPAYCode.SelLength + 1))
    
'    ValidCount = 0
'    ValidValue = ""
    
'    For i = 0 To cboPAYCode.ListCount - 1
        
'        If NewText = LCase(Left(cboPAYCode.list(i), Len(NewText))) Then
'                ValidCount = ValidCount + 1
'                ValidValue = cboPAYCode.list(i)
'        End If
'        Next
'        If ValidCount <= 1 Then KeyAscii = 0
'            If ValidCount = 1 Then
'               cboPAYCode = ValidValue
'               cboPAYCode.SelStart = 0
'               cboPAYCode.SelLength = Len(ValidValue)
'            End If
'        End If
End Sub

