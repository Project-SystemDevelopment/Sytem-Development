VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmReturnPayorCase 
   Caption         =   "Bordx Delivery"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Claim Detail"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   11655
      Begin VB.TextBox TxtClaimNo 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtBordxStatus 
         Height          =   285
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TxtBordxAmt 
         Height          =   285
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TxtBordxNo 
         Height          =   285
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtAppStatus 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TxtAppAmt 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TxtPayor 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtBordxDate 
         Height          =   285
         Left            =   4200
         TabIndex        =   7
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Bordx Status"
         Height          =   255
         Left            =   3120
         TabIndex        =   29
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Bordx Amount"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   3120
         TabIndex        =   27
         Top             =   765
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   3120
         TabIndex        =   26
         Top             =   405
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "App. Status"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1485
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "App. Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   405
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   795
      Left            =   120
      TabIndex        =   20
      Top             =   3030
      Width           =   11655
      Begin VB.ComboBox CboStatus 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   280
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Approved Status"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox cboVersion 
         Height          =   315
         Left            =   2910
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxtCaseNo 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "--"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   405
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   11655
      Begin VB.CommandButton CmdToPayor 
         Caption         =   "&Update"
         Height          =   375
         Left            =   8400
         TabIndex        =   12
         Top             =   220
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7560
         TabIndex        =   11
         Top             =   220
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10080
         TabIndex        =   14
         Top             =   220
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9240
         TabIndex        =   13
         Top             =   220
         Width           =   855
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Returned From Payor (Not App)"
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
      TabIndex        =   18
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmReturnPayorCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim searchCriteria As String


Private Sub Form_KeyPress(KeyAscii As Integer)
    'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
End Sub

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
    Call ComboListing
    cboVersion.Text = "1"
End Sub

'**********************Start Command Button *******************************
Private Sub CmdClear_Click()
    TxtCaseNo.Text = ""
    TxtCaseNo.SetFocus
    cboVersion.Text = "1"
    CboStatus.Clear
    TxtPayor.Text = ""
    TxtClaimNo.Text = ""
    TxtAppAmt.Text = ""
    TxtAppStatus.Text = ""
    TxtBordxNo.Text = ""
    txtBordxDate.Text = "__/__/____"
    TxtBordxAmt.Text = ""
    TxtBordxStatus.Text = ""
    Call ComboListing
End Sub

Private Sub CmdSearch_Click()
On Error GoTo errRun

    TxtPayor.Text = ""
    TxtClaimNo.Text = ""
    TxtAppAmt.Text = ""
    TxtAppStatus.Text = ""
    TxtBordxNo.Text = ""
    txtBordxDate.Text = "__/__/____"
    TxtBordxAmt.Text = ""
    TxtBordxStatus.Text = ""
    CmdSearch.Enabled = False
    If TxtCaseNo = "" Then
        MsgBox "Please Enter Claim No", vbOKOnly + vbCritical, DialogBoxTitle
        TxtCaseNo.SetFocus
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        CmdExit.Enabled = True
    Else
        Screen.MousePointer = vbHourglass
            medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            Call SearchRecord
            medixmasterconnWS.Close
            Set medixmasterconnWS = Nothing
        Screen.MousePointer = Default
        
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        CmdExit.Enabled = True
    End If
    CmdSearch.Enabled = True
    Exit Sub
    
errRun:
    MsgBox ERR.Description
    Screen.MousePointer = vbDefault
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
End Sub

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim Sql As String
Dim CAS As New ADODB.Recordset

    strAppend = ""
        
    If Len(Trim(TxtCaseNo)) Then
        strAppend = strAppend + " AND CASNumber = '" & Trim(TxtCaseNo) & "'"
    End If
    
    If Len(Trim(cboVersion)) Then
        strAppend = strAppend + " AND ClaimVersion = '" & Trim(cboVersion) & "'"
    End If
    
    searchCriteria = strAppend
    
    Call BordxDeliveryListing(strAppend)
       
End Function

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdToPayor_Click()
Dim ClmNo As String
Dim ClmVer As String
Dim Update

    TxtCaseNo.SetFocus
    CmdToPayor.Enabled = False
    If CboStatus = "" Then
        MsgBox "Please Enter Approved Status", vbCritical
        CboStatus.SetFocus
    Else
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            
        If Update = vbYes Then
            medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            ClmNo = Trim(Mid(TxtClaimNo.Text, 1, IIf(InStr(1, TxtClaimNo.Text, "-") = 0, 10, InStr(1, TxtClaimNo.Text, "-") - 1)))
            ClmVer = Trim(Mid(TxtClaimNo.Text, InStr(1, TxtClaimNo.Text, "-") + 1, Len(TxtClaimNo.Text) - InStr(1, TxtClaimNo.Text, "-")))
                                                            
            Call UpdateCLM(ClmNo, CInt(ClmVer))
                
            MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
            CboStatus.Text = "N - No"
            Call BordxDeliveryListing(searchCriteria)
            medixmasterconnWS.Close
            Set medixmasterconnWS = Nothing
        End If
    End If
    CmdToPayor.Enabled = True
End Sub
'**********************End Command Button *******************************

'**********************Start BordxDelivery Listing  *******************************
Private Sub BordxDeliveryListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
    
    'Sql = " Select BordxRefNo, BordxPAYCode, BordxDate, BordxClaimAmt, BordxPAYToStatus, " & _
          " CASNumber, ClaimVersion, CLMTotalApproved, BordxStatus, CLMAppStatus " & _
          " From VIEW_BORDX_ClaimApp where BordxPAYToStatus = '1' "
    
    'If Len(Trim(StrAppend)) Then
    '    Sql = Sql + StrAppend
    'End If
    
    'rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "Bordx_ReturnPayorCase"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    cmd.Parameters.Append prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    If rst2.EOF = True Then
        MsgBox "Bordx not send to Payor!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        CmdExit.Enabled = True
    Else
            
        With rst2
                            
            If Len(rst2!BordxPAYCode) Then
                TxtPayor = rst2!BordxPAYCode
            End If
            If Len(rst2!CasNumber & "-" & rst2!claimversion) Then
                TxtClaimNo = (rst2!CasNumber & "-" & rst2!claimversion)
            End If
            If Len(rst2!CLMTotalApproved) Then
                TxtAppAmt = Format(rst2!CLMTotalApproved, "#,##0.00")
            End If
            If Len(rst2!CLMAppStatus) Then
                TxtAppStatus = rst2!CLMAppStatus
            End If
            If Len(rst2!BordxRefNo) Then
                TxtBordxNo = rst2!BordxRefNo
            End If
            If Len(rst2!BordxDate) Then
                txtBordxDate = Format(rst2!BordxDate, "dd/mm/yyyy")
            End If
            If Len(rst2!BordxClaimAmt) Then
                TxtBordxAmt = Format(rst2!BordxClaimAmt, "#,##0.00")
            End If
            If Len(rst2!Bordxstatus) Then
                TxtBordxStatus = rst2!Bordxstatus
            End If
            
        End With
            
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    CmdExit.Enabled = True
End Sub
'**********************End BordxDelivery Listing *******************************

'**********************Start Update all Record functions*******************************

Private Function UpdateCLM(ClmNo As String, ClmVer As String) As Boolean
Dim strsql As String

    'update claim table
    strsql = " update Claim set " & _
             " CLMAppStatus = '" & Mid(CboStatus, 1, 1) & "'" & _
             " Where CASNumber = '" & Trim(ClmNo) & "' AND ClaimVersion = '" & Trim(ClmVer) & "'"
    
    medixmasterconnWS.Execute strsql
  
End Function
'**********************End Update all Record functions*******************************


'**********************Start Combo Listing*******************************
Private Sub ComboListing()

   With CboStatus
       .AddItem "N - No"
       .AddItem "Y - Yes"
        .Text = "N - No"
   End With
   
End Sub
'**********************End Combo Listing*******************************

Private Sub TxtCaseNo_LostFocus()
Dim Version As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim strAppend As String

    medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    If Trim(TxtCaseNo.Text) <> "" Then
        cboVersion.Clear
        TxtCaseNo = ChkStr(TxtCaseNo)
    
        'Version.Open "Select CASNumber, ClaimVersion from Claim Where CASNumber = '" & Trim(TxtCaseNo) & "'", medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
        strAppendCase = "2"
        strAppend = " AND CASNumber = '" & Trim(TxtCaseNo) & "'"
        Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "Bordx_ReturnPayorCase"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        Version.CursorLocation = adUseClient
        Version.CursorType = adOpenForwardOnly
        Version.CacheSize = 100
        Set Version = cmd.Execute
        
        Do Until Version.EOF
            cboVersion.AddItem Version!claimversion
            Version.MoveNext
        Loop
        Version.Close
        Set Version = Nothing
    
    End If
    
    If cboVersion = "" Then
        cboVersion.Text = "1"
    End If
    
    Call SearchRecord
    
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
    
End Sub

Private Sub cboVersion_Change()
    If InStr(cboVersion.Text, "null") Then
       cboVersion.Enabled = False
    'Else
        'Call SearchRecord
    End If
End Sub

Private Sub cboVersion_Click()
    If InStr(cboVersion.Text, "null") Then
       cboVersion.Enabled = False
    'Else
        'Call SearchRecord
    End If
End Sub

Private Sub cboVersion_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
Dim strsql As String


  If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboVersion.Text, cboVersion.SelStart) + Chr(KeyAscii) + Mid(cboVersion.Text, cboVersion.SelStart + cboVersion.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboVersion.ListCount - 1
 
        If NewText = LCase(Left(cboVersion.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboVersion.List(i)
        End If
    Next i
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboVersion.Text = ValidValue
            cboVersion.SelStart = 0
            cboVersion.SelLength = Len(ValidValue)
        End If
  End If
End Sub
