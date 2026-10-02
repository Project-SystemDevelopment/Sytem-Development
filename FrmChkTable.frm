VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmChkTable 
   Caption         =   "Check Claim Table"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox CboPayor 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton CmdChk 
      Caption         =   "COMPARE TABLE"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
   Begin MSComctlLib.ListView lstClmList 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2990
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clm No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CasNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Clm Version"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Payor"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   225
      Width           =   615
   End
End
Attribute VB_Name = "FrmChkTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Private CASNo As String
Private clmversion As String
Private BordxNo As String


Private Function CheckRecord()
Dim CLMINV As New ADODB.Recordset
Dim strsql As String
    
    strsql = " Select CASNumber, INVWSVersion, " & _
             " INVBordxrefNo, INVBordxed " & _
             " From ClaimInvoice Where CASNumber = '" & Trim(CASNo) & "' And INVWSVersion = '" & Trim(clmversion) & "'"
    CLMINV.Open strsql, medixmasterconnWS, adOpenStatic, adLockOptimistic

    If CLMINV.recordCount = 0 Then
        With CLMINV
            .AddNew
            !CASNumber = CASNo
            !INVWSVersion = clmversion
            !INVBordxRefNo = BordxNo
            If Trim(Mid(Bordx, 1, 1)) = "N" Then
                !INVBordxed = "1"
            Else
                !INVBordxed = "0"
            End If
            .Update
       End With
    End If

    CLMINV.Close
    Set CLMINV = Nothing
End Function

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
    
    strAppend = ""

    If Len(Trim(CboPayor)) Then
        strAppend = " And Substring(Casnumber,1,2) = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
    End If
    
    Call ClmListing(strAppend)
    
End Function

Private Sub ClmListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strCount

    lstClmList.listitems.Clear
    CASNo = ""
    clmversion = ""

    sql = " Select CASNumber, ClaimVersion, BordxRefNo From Claim where 0 = 0 "
    
    strCount = "select count(*) as totalRecord From Claim where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
        strCount = strCount + strAppend
    End If

    Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rstCount!totalrecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
         
    Do
        Do Until rst2.EOF
                
            With rst2
                
                Set Record = lstClmList.listitems.Add(, , rst2!CASNumber & "-" & rst2!claimversion)
                
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(1) = Trim(rst2!BordxRefNo)
                    BordxNo = Trim(rst2!BordxRefNo)
                Else
                    BordxNo = ""
                End If
                
                If Len(rst2!CASNumber) Then
                    Record.SubItems(2) = Trim(rst2!CASNumber)
                    CASNo = Trim(rst2!CASNumber)
                End If
                
                If Len(rst2!claimversion) Then
                    Record.SubItems(3) = Trim(rst2!claimversion)
                    clmversion = Trim(rst2!claimversion)
                End If
                
                Call CheckRecord '(CASNo, clmversion, BordxNo)
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    
End Sub

Private Sub CmdChk_Click()
Call SearchRecord
End Sub

Private Sub Form_Load()
    intExit = 0
    Call ComboListing
End Sub


Private Sub ComboListing()
Dim PAY As New ADODB.Recordset

PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic

    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

End Sub

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub
