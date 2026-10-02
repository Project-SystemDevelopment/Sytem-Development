VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepudiation 
   Caption         =   "Repudiation Letter"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtICcode 
      Height          =   375
      Left            =   1800
      TabIndex        =   27
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtReasonCode 
      Height          =   285
      Left            =   2520
      TabIndex        =   26
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtLettercode 
      Height          =   285
      Left            =   2520
      TabIndex        =   24
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtRepudiationCode 
      Height          =   285
      Left            =   2520
      TabIndex        =   21
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdGenerateCode 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8805
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Repudiation Letter"
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
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.TextBox txtPatientCode 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtDateCode 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtPolicyCode 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtMBMcode 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtINScode 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtCAScode 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "MA010000011"
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdSearchcode 
      Caption         =   "&Search"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtHospitalCode 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Case Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Insured Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "MBM No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Policy No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Admission Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Hospital"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Repudiation Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Admission Reason"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label11 
      Caption         =   "Admission Reason"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Letter to generate"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Policy No"
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Repudiation Code"
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Case Number"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Insured Name"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Membership No"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Date Of Admission"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Hospital"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Patient Name"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmRepudiation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objword As word.Application
Dim objdoc As word.Document
Dim bookmarktype As String
Dim RCcode As String 'to save repcode such as RC2.01, RC4.01 etc


Private Sub Clear()
    'txtCAScode = ""
    txtPatientCode = ""
    txtINScode = ""
    ttpatientcode = ""
    txtMBMcode = ""
    txtPolicyCode = ""
    txtDateCode = ""
    txtHospitalCode = ""
    txtRepudiationCode = ""
    txtLettercode = ""
    ListView1.listitems.Clear
    txtReasonCode = ""
    
End Sub

Private Sub cmdClear_Click()
Call Clear

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdGenerateCode_Click()
Dim objword As word.Application
Dim objdoc As word.Document

If ListView1.listitems.Count = 0 Or txtPatientCode = "" Or txtMBMcode = "" _
Or txtPolicyCode = "" Then
    MsgBox "Some fields are empty or imcomplete data. Please select a record from the list or search again", vbExclamation + vbOKOnly, DialogBoxTitle
ElseIf txtRepudiationCode = "" Then
    MsgBox "Repudiation code is empty. System unable to generate appropriate document", vbExclamation + vbOKOnly, DialogBoxTitle
ElseIf Trim(UCase(txtRepudiationCode)) = "RC6" Then
    cmdGenerateCode.Enabled = False
    frmRepudiation.Hide
    frmRepudiationOthers.show
Else
    cmdGenerateCode.Enabled = False
    Call funcword
    
End If

End Sub

Private Sub cmdSearchcode_Click()
If Trim(txtCAScode) = "" Then
    MsgBox "Please enter search code", vbExclamation + vbOKOnly, DialogBoxTitle
Else
    Call Clear
    Call search
    cmdGenerateCode.Enabled = True
End If

End Sub


Private Sub Form_Load()
    Call Clear
    bookmarktype = ""
    RCcode = ""
    txtPatientCode.Locked = True
    txtINScode.Locked = True
    txtPatientCode.Locked = True
    txtMBMcode.Locked = True
    txtPolicyCode.Locked = True
    txtDateCode.Locked = True
    txtHospitalCode.Locked = True
    txtRepudiationCode.Locked = True
    txtLettercode.Locked = True
    txtReasonCode.Locked = True
    cmdGenerateCode.Enabled = False
    
End Sub
Private Sub Form_activate()
txtCAScode.SetFocus
End Sub



Private Sub ListView1_Click()
On Error Resume Next
    txtCAScode = ListView1.SelectedItem.Text
    txtINScode = ListView1.SelectedItem.SubItems(1)
    txtPatientCode = ListView1.SelectedItem.SubItems(2)
    txtMBMcode = ListView1.SelectedItem.SubItems(3)
    txtPolicyCode = ListView1.SelectedItem.SubItems(4)
    txtDateCode = ListView1.SelectedItem.SubItems(5)
    txtHospitalCode = ListView1.SelectedItem.SubItems(6)
    txtRepudiationCode = ListView1.SelectedItem.SubItems(7)
    txtReasonCode = ListView1.SelectedItem.SubItems(8)

Select Case UCase(Left(txtRepudiationCode, 3))
    Case "RC1"
        txtLettercode = "Pre-Existing illness"
        bookmarktype = "RC1"
    Case "RC2"
        txtLettercode = "Specific Illness"
        bookmarktype = "RC2"
        RCcode = txtRepudiationCode 'we get RC2.01 and so on to get description from reccode table
    Case "RC3"
        txtLettercode = "Waiting Period Specification Not Met"
        bookmarktype = "RC3"
    Case "RC4"
        txtLettercode = "Exclusion"
        bookmarktype = "RC4"
        RCcode = txtRepudiationCode
        
    Case "RC5"
        txtLettercode = "Policy Issues"
        bookmarktype = "RC5"
        RCcode = txtRepudiationCode
    Case "RC6"
        txtLettercode = "Others (Specify)"
        bookmarktype = "RC6"
        
End Select


End Sub



Private Sub search()

On Error Resume Next
Dim list As New ADODB.Recordset
Dim sql As String
Dim View As ListItem
ListView1.listitems.Clear

sql = "select * from VIEW_CAS_REPUDIATION where CASnumber ='" & txtCAScode & "'"

list.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic

    
    If list.recordCount = 0 Then
        MsgBox "Record not found", vbExclamation + vbOKOnly, DialogBoxTitle
    Else
    
        Do Until list.EOF()
                Set View = ListView1.listitems.Add(, , list!CASNumber)
                View.SubItems(1) = list!MBMName
            If Not list!mbmpatientcovid = 0 Then
               View.SubItems(2) = list!mbmcname
            Else
               View.SubItems(2) = list!MBMName
            End If
                 View.SubItems(3) = list!MBMNumber
                 View.SubItems(4) = list!MBMPolicyNo
                 View.SubItems(5) = list!casdateadmitted
                 View.SubItems(6) = list!HOSName
                
            If Not IsNull(list!REPCode) = True Then
               View.SubItems(7) = list!REPCode
            End If
            If Not IsNull(list!CasAdmissionReason) = True Then
               View.SubItems(8) = list!CasAdmissionReason
            End If
         
                list.MoveNext
        
        Loop
    
    End If
    list.Close
End Sub



Private Sub funcword()
On Error GoTo openerror 'if you generate a doc then close it then generate again error occurs
                        ' so this trap will rectify this problem
Dim path As String


If objword Is Nothing Then
    Set objword = New word.Application
    objword.Visible = True
    path = openword(bookmarktype)
    Set objdoc = objword.Documents.Add(App.path & "\template\" & path & ".dot")


ElseIf Not objdoc Is Nothing Then
        If MsgBox("Previous document " & objdoc & " is not closed. Press OK to save and close it now or cancel to abort", _
            vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
            objdoc.Save 'save file to the location where it was opened
            objdoc.Close
        ElseIf vbCancel Then
          Exit Sub
        End If
       Set objdoc = Nothing
       path = openword(bookmarktype)
       Set objdoc = objword.Documents.Add(App.path & "\template\" & path & ".dot")
End If
    Call bookmark
    
    
form_load_exit:
Exit Sub

    
openerror:
Me.show
If ERR.Number = "5356" Then
MsgBox ERR.Description & ". Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle

ElseIf ERR.Number = "462" Then
 Set objword = New word.Application
    objword.Visible = True
      path = openword(bookmarktype)
      Set objdoc = objword.Documents.Add(App.path & "\template\" & path & ".dot")
     Call bookmark
       
ElseIf ERR.Number = "-2147417848" Then 'this is automation error..when you close the doc(not word) then generate again
        path = openword(bookmarktype)
        Set objdoc = objword.Documents.Add(App.path & "\template\" & path & ".dot")
        Call bookmark
        'path = openword()
        'Set objdoc = objword.Documents.Add(App.path & "\template\" & path & ".dot")
        'Call bookmark
End If



       
End Sub
Private Sub bookmark()
Select Case bookmarktype
    Case "RC1"
        If bookmarkRC1 = True Then
        Call printDoc
        End If
    Case "RC2"
        If bookmarkRC2 = True Then
        Call printDoc
        End If
    Case "RC3"
        If bookmarkRC3 = True Then
        Call printDoc
        End If
    Case "RC4"
        If bookmarkRC4 = True Then
        Call printDoc
        End If
    Case "RC5"
        If bookmarkRC5 = True Then
        Call printDoc
        End If
  
    End Select
End Sub


Private Function openword(Optional bookmarktype As String) As String
'this is to determine which document to open based on repcode
Select Case bookmarktype
    Case "RC1"
        openword = "Pre existing illness"
    Case "RC2"
        openword = "Specific Illness"
    Case "RC3"
        openword = "Waiting period"
    Case "RC4"
        openword = "Exclusion"
    Case "RC5"
        openword = "Policy Cancel"
End Select

End Function

Private Function bookmarkRC1() 'pre-exitsting illness
    objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMcode
    objdoc.Bookmarks("policynumber").Range.Text = txtPolicyCode
    objdoc.Bookmarks("insuredname").Range.Text = txtINScode
    objdoc.Bookmarks("patientname").Range.Text = txtPatientCode
    objdoc.Bookmarks("hospital").Range.Text = txtHospitalCode
    objdoc.Bookmarks("reason").Range.Text = txtReasonCode

bookmarkRC1 = True
End Function
Private Function bookmarkRC2() 'specific illness
  Dim list As New ADODB.Recordset
  Dim sql As String
  
  sql = "select * from repcode where repcode='" & RCcode & "'"
  list.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
  

    objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMcode
    objdoc.Bookmarks("policynumber").Range.Text = txtPolicyCode
    objdoc.Bookmarks("insuredname").Range.Text = txtINScode
    objdoc.Bookmarks("patientname").Range.Text = txtPatientCode
    objdoc.Bookmarks("hospital").Range.Text = txtHospitalCode
    objdoc.Bookmarks("reason").Range.Text = txtReasonCode
    objdoc.Bookmarks("specific").Range.Text = list!repdescription
bookmarkRC2 = True
    list.Close
    End Function
Private Function bookmarkRC3() 'waiting specification not met
    
    objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMcode
    objdoc.Bookmarks("policynumber").Range.Text = txtPolicyCode
    objdoc.Bookmarks("insuredname").Range.Text = txtINScode
    objdoc.Bookmarks("patientname").Range.Text = txtPatientCode
    objdoc.Bookmarks("hospital").Range.Text = txtHospitalCode
    objdoc.Bookmarks("reason").Range.Text = txtReasonCode
    objdoc.Bookmarks("date").Range.Text = txtDateCode

bookmarkRC3 = True
End Function
Private Function bookmarkRC4() 'other exclusions
Dim list As New ADODB.Recordset
Dim sql As String

sql = "select * from repcode where repcode ='" & txtRepudiationCode & "'"
list.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic

    objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMcode
    objdoc.Bookmarks("policynumber").Range.Text = txtPolicyCode
    objdoc.Bookmarks("insuredname").Range.Text = txtINScode
    objdoc.Bookmarks("patientname").Range.Text = txtPatientCode
    objdoc.Bookmarks("hospital").Range.Text = txtHospitalCode
    objdoc.Bookmarks("reason").Range.Text = txtReasonCode
    objdoc.Bookmarks("date").Range.Text = txtDateCode
    objdoc.Bookmarks("date2").Range.Text = txtDateCode
    objdoc.Bookmarks("exclusion").Range.Text = list!repdescription
    
bookmarkRC4 = True
list.Close




End Function
Private Function bookmarkRC5() 'policy issues(policy cancel)
Dim list As New ADODB.Recordset
Dim sql As String
sql = "select * from repcode where repcode ='" & txtRepudiationCode & "'"
list.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic

    objdoc.Bookmarks("mbmnumber").Range.Text = txtMBMcode
    objdoc.Bookmarks("policynumber").Range.Text = txtPolicyCode
    objdoc.Bookmarks("insuredname").Range.Text = txtINScode
    objdoc.Bookmarks("patientname").Range.Text = txtPatientCode
    objdoc.Bookmarks("hospital").Range.Text = txtHospitalCode
    objdoc.Bookmarks("reason").Range.Text = txtReasonCode
    objdoc.Bookmarks("cancel").Range.Text = list!repdescription
bookmarkRC5 = True
list.Close



End Function

Private Sub printDoc()

            objdoc.SaveAs (App.path & "\template\repudiation\" & Trim(txtCAScode) & "-" & Trim(txtRepudiationCode) & ".doc")
           Me.show
           If MsgBox("Do you want to print this document?", vbQuestion + vbYesNo, DialogBoxTitle) = vbYes Then
                'objdoc.SaveAs (App.path & "\repudiation\" & Trim(txtCAScode) & ".doc")
                'objdoc.Close
                'objdoc.PrintOut
                MsgBox "printed"
                
                Exit Sub
           Else
               Me.WindowState = 1
               Exit Sub
           End If
End Sub


